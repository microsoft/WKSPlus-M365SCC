Param (
    [switch]$debug
)

################ Define Variables ###################
$LogPath = "$env:UserProfile\Desktop\SCLabFiles\Scripts\"
$LogCSV = "$env:UserProfile\Desktop\SCLabFiles\Scripts\Progress_Label_Log.csv"
$global:nextPhase = 1
$global:recovery = $false

#label
$labelDisplayName = "WKS Highly Confidential"
$labelName = "WKS-Highly-Confidential"
$labelTooltip = "Contains Highly confidential info"
$labelComment = "Documents with this label contain sensitive data."

#label policy
$labelPolicyName = "WKS-Highly-confidential-publish"

###DEBUG###
$oldDebugPreference = $DebugPreference
if($debug){
    write-debug "Debug Enabled"
    $DebugPreference = "Continue"
    Start-Transcript -Path "$($LogPath)sensitivitylabel-debug.txt"
}

# -----------------------------------------------------------
# Write log function
# -----------------------------------------------------------
function logWrite([int]$phase, [bool]$result, [string]$logstring)
{
    if ($result)
        {
            Add-Content -Path $LogCSV -Value "$phase,$result,$(Get-Date),$logString"
            Write-Host -ForegroundColor Green "$(Get-Date) - Phase $phase : $logstring"
        } 
    else 
        {
            Write-Host -ForegroundColor Red "$(Get-Date) - Phase $phase : $logstring"
        }
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

# -----------------------------------------------------------
# Test the log path (Step 0)
# -----------------------------------------------------------
function initialization
{
    $pathExists = Test-Path($LogPath)
    if (!$pathExists)
        {
            New-Item -ItemType "directory" -Path $LogPath -ErrorAction SilentlyContinue | Out-Null
        }
        Set-Location -Path $LogPath
        Add-Content -Path $LogCSV -Value '"Phase","Result","DateTime","Status"'
        logWrite 0 $true "Initialization completed"
}

# -----------------------------------------------------------
# Connect to Exchange Online (Step 1)
# -----------------------------------------------------------
function ConnectEXO
{
    try 
    {
        Write-Debug "Get-OrganizationConfig -ErrorAction stop"
        $testConnection = Get-OrganizationConfig -ErrorAction stop | Out-Null #if true (Already Connected)
    }
    catch
        {
            try
                {
                    write-Debug $error[0].Exception
                    Write-Host "Connecting to Exchange Online..."
                    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction stop | Out-Null
                }
                catch    
                    {
                        try
                            {
                                write-Debug $error[0].Exception
                                Write-Host "Installing Exchange Online PowerShell Module..."
                                Install-Module ExchangeOnlineManagement -Force -AllowClobber -ErrorAction stop | Out-Null
                                Connect-ExchangeOnline -ShowBanner:$false -ErrorAction stop | Out-Null
                            }
                            catch
                                {
                                    write-Debug $error[0].Exception
                                    logWrite 1 $false "Couldn't connect to Exchange Online. Exiting."
                                    exitScript
                                }
                   
                    }
        }
        if($global:recovery -eq $false)
            {
                logWrite 1 $true "Successfully connected to Exchange Online"
                $global:nextPhase++
                Write-Debug "nextPhase set to $global:nextPhase"
            }
}

# -----------------------------------------------------------
# Connect to Compliance Center (Step 2)
# -----------------------------------------------------------
function ConnectSCC
{
    try 
    {
        Write-Debug "Get-Label -ErrorAction stop"
        $testConnection = Get-Label -ErrorAction stop | Out-Null #if true (Already Connected)
    }
    catch
        {
            try
                {
                    write-Debug $error[0].Exception
                    Write-Host "Connecting to Compliance Center..."
                    Connect-IPPSSession -ErrorAction stop | Out-Null
                }
                catch    
                    {
                        try
                            {
                                write-Debug $error[0].Exception
                                Write-Host "Installing Compliance Center PowerShell Module..."
                                #Install-Module ExchangeOnlineManagement -Force -AllowClobber #Not required, but it was already installed on the previous step
                                Connect-IPPSSession -ErrorAction stop | Out-Null
                            }
                            catch
                                {
                                    write-Debug $error[0].Exception
                                    logWrite 2 $false "Couldn't connect to Compliance Center. Exiting."
                                    exitScript
                                }
                   
                    }
        }
        if($global:recovery -eq $false)
            {
                logWrite 2 $true "Successfully connected to Compliance Center"
                $global:nextPhase++
                Write-Debug "nextPhase set to $global:nextPhase"
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
        $labelStatus = New-Label -DisplayName $labelDisplayName -Name $labelName -ToolTip $labelTooltip -Comment $labelComment -ContentType "file","Email","Site","UnifiedGroup" -EncryptionEnabled:$true -SiteAndGroupProtectionEnabled:$true -EncryptionPromptUser:$true -EncryptionRightsDefinitions $Encpermission -SiteAndGroupProtectionPrivacy "private" -EncryptionDoNotForward:$true -SiteAndGroupProtectionAllowLimitedAccess:$true -ErrorAction stop | Out-Null
    } catch {
        logWrite 3 $false "Error creating label"
        exitScript
    }
    logWrite 3 $true "Successfully created label"
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

    try {
        New-LabelPolicy -name $labelPolicyName -Settings @{mandatory=$false} -AdvancedSettings @{requiredowngradejustification= $true} -Labels $labelName -ErrorAction stop | Out-Null
    } catch {
        logWrite 4 $false "Error creating label policy ($error)"
        exitScript
    }
    logWrite 4 $true "Successfully created label policy"
    $global:nextPhase++
}

# -------------------------------------------------------
# Exit function
# -------------------------------------------------------
function exitScript
{
    # Get-PSSession | Remove-PSSession
    if ($debug)
        {
            $DebugPreference = $oldDebugPreference
            Stop-Transcript
        }
    exit
}

# -------------------------------------------------------
# Script start here
# -------------------------------------------------------
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
connectExo
}

if($nextPhase -eq 2){
connectSCC
}

if($nextPhase -eq 3){
createLabel
}

if ($nextPhase -eq 4){
createPolicy
}

if ($nextPhase -ge 5){

    write-debug "Phase $nextPhase"
    Set-Location -Path $LogPath
    $nextScript = "wks-new-retention.ps1"
    logwrite 5 $true "Launching $nextScript script"
    if ($debug){
        Stop-Transcript
        .$nextScript -$debug
    } else {
        .$nextScript
    }
}