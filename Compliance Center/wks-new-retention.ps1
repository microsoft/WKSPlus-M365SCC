################ Define Variables ###################
$LogPath = "c:\temp\"
$LogCSV = "C:\temp\LabelLog.csv"
$global:nextPhase = 1
$global:recovery = $false

#label
$labelDisplayName = "WKS Highly Confidential$(get-date)"
$labelName = "WKS-Highly-Confidential$(get-date)"
$labelTooltip = "Contains Highly confidential info"
$labelComment = "Documents with this label contain sensitive data."

#label policy
$labelPolicyName = "WKS-Highly-confidential-publish$(get-date)"

################ Functions ###################
function logWrite([int]$phase, [bool]$result, [string]$logstring)
{
    if ($result)
    {
        Add-Content -Path $LogCSV -Value "$phase,$result,$(Get-Date),$logString"
        Write-Host -ForegroundColor Green "$(Get-Date) - Phase $phase : $logstring"
    } else {
        Write-Host -ForegroundColor Red "$(Get-Date) - Phase $phase : $logstring"
    }
}

function initialization
{
    $pathExists = Test-Path($LogPath)

    if (!$pathExists)
    {
        New-Item -ItemType "directory" -Path $LogPath -ErrorAction SilentlyContinue | Out-Null
    }
        Add-Content -Path $LogCSV -Value '"Phase","Result","DateTime","Status"'
        logWrite 0 $true "Initialization completed"
}

function recovery
{
    Write-host "Starting recovery..."
    $global:recovery = $true
    $savedLog = Import-Csv $LogCSV
    $lastEntry = (($savedLog.Count) - 1)
    $lastEntry2 = (($savedLog.Count) - 2)
    $lastEntryPhase = [int]$savedLog[$lastEntry].Phase
    $lastEntryResult = $savedLog[$lastEntry].Result

    if ($lastEntryResult -eq $false){
        if ($lastEntryPhase -eq $savedLog[$lastEntry2].Phase){
            WriteHost -ForegroundColor Red "The script has failed at Phase $lastEntryPhase repeatedly.  PLease check with your instructor."
            exit
        } else {
            Write-Host "There was a problem with Phase $lastEntryPhase, so trying again...."
            $global:nextPhase = $lastEntryPhase
        }
    } else {
        # set the phase
        Write-Host "Phase $lastEntryPhase was successful, so picking up where we left off...."
        $global:nextPhase = $lastEntryPhase + 1
    }
}


function createLabel
{
    <#
    TO DO:
    Need to check to see if label exists in case the failure occured after cmd was successful, such as if they close the PS window. Maybe just check if label exists, and use Set-Label if so.
    #>
    $domainName = (Get-AcceptedDomain | ?{$_.Default -eq $true}).DomainName
    $Encpermission = $domainname + ":VIEW,VIEWRIGHTSDATA,DOCEDIT,EDIT,PRINT,EXTRACT,REPLY,REPLYALL,FORWARD,OBJMODEL"
    Write-Host "Creating label with permissions: $Encpermission..."
    try {
        $labelStatus = New-Label -DisplayName $labelDisplayName -Name $labelName -ToolTip $labelTooltip -Comment $labelComment -ContentType "file","Email","Site","UnifiedGroup" -EncryptionEnabled:$true -SiteAndGroupProtectionEnabled:$true -EncryptionPromptUser:$true -EncryptionRightsDefinitions $Encpermission -SiteAndGroupProtectionPrivacy "private" -EncryptionDoNotForward:$true -SiteAndGroupProtectionAllowLimitedAccess:$true -ErrorAction Stop | Out-Null
    } catch {
        logWrite 1 $false "Error creating label"
        exit
    }
    logWrite 1 $true "Successfully created label"
    $global:nextPhase++

    #sleeping for 30 seconds
    for ($i = 1; $i -le 30; $i++ )
    {
        $p = ([Math]::Round($i/30, 2) * 100)
        Write-Progress -Activity "Allowing time for label to be created on backend..." -Status "$p% Complete:" -PercentComplete $p
        Start-Sleep -Seconds 1
    }
}

function createPolicy
{
        <#
    TO DO:
    - Need to check to see if label policy exists in case the failure occured after cmd was successful, such as if they close the PS window. Maybe just check if label exists, and use Set-Label if so.
    - Need to make sure the labele exists
    #>

    try {
        New-LabelPolicy -name $labelPolicyName -Settings @{mandatory=$false} -AdvancedSettings @{requiredowngradejustification= $true} -Labels $labelName -ErrorAction Stop | Out-Null
    } catch {
        logWrite 2 $false "Error creating label policy"
        exit
    }
    logWrite 2 $true "Successfully created label policy"
    $global:nextPhase++
}

function exitScript
{
    #Get-PSSession | Remove-PSSession
    logWrite 6 $true "Session removed successfully."
}

################ main Script start ###################

if(!(Test-Path($logCSV))){
    # if log doesn't exist then must be first time we run this, so go to initialization
    initialization
} 
else {
    # if log already exists, check if we need to recover
    recovery
    createLabel
    createPolicy
}

#use variable to control phases



if($nextPhase -eq 1){
createLabel
}

if ($nextPhase -eq 2){
createPolicy
}
