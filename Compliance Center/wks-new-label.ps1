Param (
    [switch]$debug
)

################ Define Variables ###################
$LogPath = "c:\temp\"
$LogCSV = "C:\temp\labellog.csv"
#$LogPath = "$env:UserProfile\Desktop\SCLabFiles\Scripts\"
#$LogCSV = "$env:UserProfile\Desktop\SCLabFiles\Scripts\Progress_Label_Log.csv"
$global:nextPhase = 1
$global:recovery = $false

#label
$labelDisplayName = "WKS Highly Confidential-01"
$labelName = "WKS-Highly-Confidential-01"
$labelTooltip = "Contains Highly confidential info"
$labelComment = "Documents with this label contain sensitive data."

#label policy
$labelPolicyName = "WKS-Highly-confidential-publish-01"

###DEBUG###
$oldDebugPreference = $DebugPreference
if($debug){
    write-debug "Debug Enabled"
    $DebugPreference = "Continue"
    Start-Transcript -Path "$($LogPath)sensitivitylabel-debug.txt"
}

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
        cd $LogPath
        Add-Content -Path $LogCSV -Value '"Phase","Result","DateTime","Status"'
        logWrite 0 $true "Initialization completed"
}

function recovery
{
    Write-host "Starting recovery..."
    cd $LogPath
    $global:recovery = $true
    $savedLog = Import-Csv $LogCSV
    $lastEntry = (($savedLog.Count) - 1)
    $lastEntry2 = (($savedLog.Count) - 2)
    $lastEntryPhase = [int]$savedLog[$lastEntry].Phase
    $lastEntryResult = $savedLog[$lastEntry].Result

    if ($lastEntryResult -eq $false){
        if ($lastEntryPhase -eq $savedLog[$lastEntry2].Phase){
            WriteHost -ForegroundColor Red "The script has failed at Phase $lastEntryPhase repeatedly.  PLease check with your instructor."
            exitScript
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

function goToSleep ([int]$seconds){
    for ($i = 1; $i -le $seconds; $i++ )
    {
        $p = ([Math]::Round($i/$seconds, 2) * 100)
        Write-Progress -Activity "Allowing time for label to be created on backend..." -Status "$p% Complete:" -PercentComplete $p
        Start-Sleep -Seconds 1
    }
}

function checkModule 
{
    try {
        Get-Command Connect-ExchangeOnline -ErrorAction SilentlyContinue | Out-Null
    } catch {
        logWrite 1 $false "ExchangeOnlineManagement module is not installed! Exiting."
        exitScript
    }
    logWrite 1 $True "ExchangeOnlineManagement module is installed."
    $global:nextPhase++
}

function connectExo
{
    try {
        Get-Command Set-Mailbox -ErrorAction stop | Out-Null
    } catch {
        Write-Host "Connecting to Exchange Online..."
        Connect-ExchangeOnline
        try {
            Get-Command Get-Mailbox -ErrorAction stop | Out-Null
        } catch {
            logWrite 2 $false "Couldn't connect to Exchange Online.  Exiting."
            exitScript
        }
    }
    
    if($global:recovery -eq $false){
        logWrite 2 $true "Connected to Exchange Online"
        $global:nextPhase++
    }
}

function connectSCC
{
    try {
        Get-Command Set-Label -ErrorAction:stop | Out-Null
    } catch {
        Write-Host "Connecting to Compliance Center..."
        Connect-IPPSSession
        try {
            Get-Command Set-Label -ErrorAction:stop | Out-Null
        } catch {
            logWrite 3 $false "Couldn't connect to Compliance Center.  Exiting."
            exitScript
        }
    }
    if($global:recovery -eq $false){
        logWrite 3 $true "Connected to Compliance Center"
        $global:nextPhase++
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
    try {
            if (Get-label -Identity $labelName)
        {
            write-host " The $labelName already Exists " 
      
            logWrite 4 $true "The $labelName already Exists"
            $global:nextPhase++
        }

            else{
                    $labelStatus = New-Label -DisplayName $labelDisplayName -Name $labelName -ToolTip $labelTooltip -Comment $labelComment -ContentType "file","Email","Site","UnifiedGroup" -EncryptionEnabled:$true -SiteAndGroupProtectionEnabled:$true -EncryptionPromptUser:$true -EncryptionRightsDefinitions $Encpermission -SiteAndGroupProtectionPrivacy "private" -EncryptionDoNotForward:$true -SiteAndGroupProtectionAllowLimitedAccess:$true -ErrorAction stop | Out-Null
                } 
        }
    catch {
        logWrite 4 $true "Error creating label"
        exitScript
    }
    logWrite 4 $true "Successfully created label"
    $global:nextPhase++

    #sleeping for 30 seconds
    
    goToSleep 30
}

function createPolicy
{
        <#
    TO DO:
    - Need to check to see if label policy exists in case the failure occured after cmd was successful, such as if they close the PS window. Maybe just check if label exists, and use Set-Label if so.
    - Need to make sure the labele exists
    #>

    try
    {

        if(get-labelpolicy -identity $labelpolicyname)
            {
                write-host " The $labelpolicyname already Exists " 
      
                logWrite 5 $true "The $labelpolicyname already Exists"
                $global:nextPhase++
            } 

        else
        {
        New-LabelPolicy -name $labelPolicyName -Settings @{mandatory=$false} -AdvancedSettings @{requiredowngradejustification= $true} -Labels $labelName -ErrorAction stop | Out-Null
        }
    }
    
    catch {
        logWrite 5 $false "Error creating label policy ($error)"
        exitScript
    }
    logWrite 5 $true "Successfully created label policy"
    $global:nextPhase++
}

function exitScript
{
    #remove psession if fails only
    if ($debug){
        $DebugPreference = $oldDebugPreference
        Stop-Transcript
    }
    exit
}

################ main Script start ###################
Write-Host "Starting Sensitivity Label Configuration Script...."
if(!(Test-Path($logCSV))){
    # if log doesn't exist then must be first time we run this, so go to initialization
    initialization
} else {
    # if log already exists, check if we need to recover
    recovery
    connectExo
    connectSCC
}

#use variable to control phases

if($nextPhase -eq 1){
checkModule
}

if($nextPhase -eq 2){
connectExo
}

if($nextPhase -eq 3){
connectSCC
}

if($nextPhase -eq 4){
createLabel
}

if ($nextPhase -eq 5){
createPolicy
}

if ($nextPhase -ge 6){
    $nextScript = $LogPath + ".\wks-new-retention.ps1"
    logwrite 6 $true "Launching $nextScript"
    if ($debug){
        Stop-Transcript
        .$nextScript -$debug
    } else {
        .$nextScript
    }
}