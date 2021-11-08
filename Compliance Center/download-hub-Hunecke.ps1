Param (
    [switch]$debug
)

################ Define Variables ###################
#$LogPath = "c:\temp\"
#$LogCSV = "C:\temp\download1.csv"
$LogPath = "$env:UserProfile\Desktop\SCLabFiles\Scripts"
$LogCSV = "$env:UserProfile\Desktop\SCLabFiles\Scripts\download1.csv"
$global:nextPhase = 1
$global:recovery = $false

###DEBUG###
$oldDebugPreference = $DebugPreference
if($debug){
    write-debug "Debug Enabled"
    $DebugPreference = "Continue"
    Start-Transcript -Path "$($LogPath)download-debug.txt"
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
    Write-Debug "Last Entry #: $lastEntry"
    $lastEntry2 = (($savedLog.Count) - 2)
    Write-Debug "Entry Before Last: $lastEntry2"
    $lastEntryPhase = [int]$savedLog[$lastEntry].Phase
    Write-Debug "Last Phase: $lastEntryPhase"
    $lastEntryResult = $savedLog[$lastEntry].Result
    Write-Debug "Last Entry Result: $lastEntryResult"

    if ($lastEntryResult -eq $false){
        if ($lastEntryPhase -eq $savedLog[$lastEntry2].Phase){
            WriteHost -ForegroundColor Red "The script has failed at Phase $lastEntryPhase repeatedly.  PLease check with your instructor."
            exitScript
        } else {
            Write-Host "There was a problem with Phase $lastEntryPhase, so trying again...."
            $global:nextPhase = $lastEntryPhase
            Write-Debug "nextPhase set to $global:nextPhase"
        }
    } else {
        # set the phase
        Write-Host "Phase $lastEntryPhase was successful, so picking up where we left off...."
        $global:nextPhase = $lastEntryPhase + 1
        write-Debug "nextPhase set to $global:nextPhase"
    }
}

function checkModule
{
    try {
        Write-Debug "Get-Command Connect-ExchangeOnline -ErrorAction stop"
        $testModule = Get-Command Connect-ExchangeOnline -ErrorAction stop | Out-Null
    } catch {
        write-Debug $error[0].Exception
        logWrite 1 $false "ExchangeOnlineManagement module is not installed! Exiting."
        exitScript
    }
    logWrite 1 $True "ExchangeOnlineManagement module is installed."
    $global:nextPhase++
    Write-Debug "nextPhase set to $global:nextPhase"
}

function checkModuleMSOL
{
    try {
        Write-Debug "Get-Command Connect-MsolService -ErrorAction stop"
        $testModule = Get-Command Connect-MsolService -ErrorAction stop | Out-Null
    } catch {
        write-Debug $error[0].Exception
        logWrite 2 $false "MSOL module is not installed! Exiting."
        exitScript
    }
    logWrite 2 $True "MSOL module is installed."
    $global:nextPhase++
    Write-Debug "nextPhase set to $global:nextPhase"
}

function connectExo
{
    try {
        Write-Debug "Get-Command Set-Mailbox -ErrorAction stop"
        $testConnection = Get-Command Set-Mailbox -ErrorAction stop | Out-Null
    } catch {
        write-Debug $error[0].Exception
        Write-Host "Connecting to Exchange Online..."
        Connect-ExchangeOnline
        try {
            Write-Debug "Get-Command Set-Mailbox -ErrorAction stop"
            $testConnection = Get-Command Set-Mailbox -ErrorAction stop | Out-Null
            
        } catch {
            write-Debug $error[0].Exception
            logWrite 3 $false "Couldn't connect to Exchange Online.  Exiting."
            exitScript
        }
    }
    if($global:recovery -eq $false){
        logWrite 3 $true "Successfully connected to Exchange Online"
        $global:nextPhase++
        Write-Debug "nextPhase set to $global:nextPhase"
    }
}
# ----------------------------------------
# Connect to Microsoft Compliance center
# ----------------------------------------
function connectSCC
{
    try {
        Write-Debug "Get-Command Set-Label -ErrorAction:stop"
        Get-Command Set-Label -ErrorAction:Stop | Out-Null
    }
    catch {
        write-Debug $error[0].Exception
        Write-Host "Connecting to Compliance Center..."
        Connect-IPPSSession
        try {
            Write-Debug "Get-Command Set-Label -ErrorAction:Stop"
            Get-Command Set-Label -ErrorAction:Stop | Out-Null
        } catch {
            write-Debug $error[0].Exception
            logWrite 4 $false "Couldn't connect to Compliance Center.  Exiting."
            exitScript
        }
    }
    if($global:recovery -eq $false){
        logWrite 4 $true "Successfully connected to Compliance Center"
        $global:nextPhase++
        Write-Debug "nextPhase set to $global:nextPhase"
    }
}

# ------------------------------------
# Connect to Microsoft Online Service
# ------------------------------------
function ConnectMsolService
{
    try {
        write-debug "$testConnection = Get-MsolDomain -ErrorAction stop"
        $testConnection = Get-MsolDomain -ErrorAction stop | out-null
    } catch {
        write-Debug $error[0].Exception
        Write-Host "Connecting to msol Service..."
        Connect-MsolService
        try {
            write-Debug "Get-MsolContact -ErrorAction SilentlyContinue"
            $testContact = Get-MsolContact -ErrorAction stop | Out-Null
        } catch {
            write-Debug $error[0].Exception
            logWrite 5 $false "Couldn't connect to MSOL Service.  Exiting."
            exitScript
        }
    }
    if($global:recovery -eq $false){
        logWrite 5 $true "Successfully connected to MSOL Service"
        $global:nextPhase++
        Write-Debug "nextPhase set to $global:nextPhase"
    }
}


# ------------------------------------
# get accepted domains
# ------------------------------------

Function getdomain
{
    try{
        Write-Debug "$InitialDomain = Get-MsolDomain -ErrorAction stop | Where-Object {$_.IsInitial -eq $true}"
        $InitialDomain = Get-MsolDomain -ErrorAction stop | Where-Object {$_.IsInitial -eq $true}
   }catch {
        write-Debug $error[0].Exception
        logWrite 6 $false "unable to fetch all accepted Domains."
        exitScript
    }
    Write-Debug "Initial domain: $InitialDomain"
    if($global:recovery -eq $false){
        logWrite 6 $True "Able to get all accepted Domains."
        $global:nextPhase++
        Write-Debug "nextPhase set to $global:nextPhase"
    }
    return $InitialDomain.name.split(".")[0]
}



function connectspo([string]$tenantName)
{
    $AdminURL = "https://$tenantName-admin.sharepoint.com"
    Try {
        write-Debug "Get-Sposite -ErrorAction stop"
        $testConnection = Get-Sposite -ErrorAction stop | Out-Null
    } catch {
        write-Debug $error[0].Exception
        Try{
        #Connect to Office 365
            write-Debug "Connect-sposervice -Url $AdminURL"
            Connect-sposervice -Url $AdminURL
        } catch {
            write-Debug $error[0].Exception
            logWrite 7 $false "Unable to connect to Sharepoint using $adminURL."
            exitScript
        }
    }
    if($global:recovery -eq $false){
        logWrite 7 $True "Successfully connected to Sharepoint using $adminURL."
        $global:nextPhase++
        Write-Debug "nextPhase set to $global:nextPhase"
    }
}

function downloadscriptlabel
{
    Try{
        Write-Debug "Invoke-WebRequest -Uri https://raw.githubusercontent.com/microsoft/WKSPlus-M365SCC/main/Compliance%20Center/wks-new-label.ps1 -OutFile $LogPath\wks-new-label.ps1 -ErrorAction Stop"
        Invoke-WebRequest -Uri https://raw.githubusercontent.com/microsoft/WKSPlus-M365SCC/main/Compliance%20Center/wks-new-label.ps1 -OutFile $LogPath\wks-new-label.ps1 -ErrorAction Stop
    } catch {
        write-Debug $error[0].Exception
        logWrite 8 $false "Unable to download the Script! Exiting."
        exitScript
    }
    logWrite 8 $True "The Script has been downloaded ."
    $global:nextPhase++
    Write-Debug "nextPhase set to $global:nextPhase"
}

function downloadscriptDLP
{
    Try{
        Write-Debug "Invoke-WebRequest -Uri https://raw.githubusercontent.com/microsoft/WKSPlus-M365SCC/main/Compliance%20Center/wks-new-DLP.ps1 -OutFile $LogPath\wks-new-DLP.ps1 -ErrorAction Stop"
        Invoke-WebRequest -Uri https://raw.githubusercontent.com/microsoft/WKSPlus-M365SCC/main/Compliance%20Center/wks-new-DLP.ps1 -OutFile $LogPath\wks-new-DLP.ps1 -ErrorAction Stop
        } catch {
            write-Debug $error[0].Exception
            logWrite 9 $false "Unable to download the Script! Exiting."
            exitScript
        }
    logWrite 9 $True "The Script has been downloaded ."
    $global:nextPhase++
    Write-Debug "nextPhase set to $global:nextPhase"
}

function downloadscriptRetention
{
    try{
        Write-Debug "Invoke-WebRequest -Uri https://raw.githubusercontent.com/microsoft/WKSPlus-M365SCC/main/Compliance%20Center/wks-new-retention.ps1 -OutFile $LogPath\wks-new-retention.ps1 -ErrorAction Stop"
        Invoke-WebRequest -Uri https://raw.githubusercontent.com/microsoft/WKSPlus-M365SCC/main/Compliance%20Center/wks-new-retention.ps1 -OutFile $LogPath\wks-new-retention.ps1 -ErrorAction Stop
    }
    catch {
        write-Debug $error[0].Exception
        logWrite 10 $false "Unable to download the Script! Exiting."
        exitScript
    }
    logWrite 10 $True "The Script has been downloaded ."
    $global:nextPhase++
    Write-Debug "nextPhase set to $global:nextPhase"
}

function downloadscriptInsiderRisks01
{
    try{
        Write-Debug "Invoke-WebRequest -Uri https://raw.githubusercontent.com/microsoft/WKSPlus-M365SCC/main/Compliance%20Center/wks-new-HRConnector.ps1 -OutFile $LogPath\wks-new-HRConnector.ps1 -ErrorAction Stop"
        Invoke-WebRequest -Uri https://raw.githubusercontent.com/microsoft/WKSPlus-M365SCC/main/Compliance%20Center/wks-new-HRConnector.ps1 -OutFile $LogPath\wks-new-HRConnector.ps1 -ErrorAction Stop
    }
    catch {
        write-Debug $error[0].Exception
        logWrite 11 $false "Unable to download the Script! Exiting."
        exitScript
    }
    logWrite 11 $True "The Script has been downloaded ."
    $global:nextPhase++
    Write-Debug "nextPhase set to $global:nextPhase"
}

function downloadscriptInsiderRisks02
{
    try{
        Write-Debug "Invoke-WebRequest -Uri https://raw.githubusercontent.com/microsoft/m365-hrconnector-sample-scripts/master/upload_termination_records.ps1 -OutFile $LogPath\upload_termination_records.ps1-ErrorAction Stop"
        Invoke-WebRequest -Uri https://raw.githubusercontent.com/microsoft/m365-hrconnector-sample-scripts/master/upload_termination_records.ps1 -OutFile $LogPath\upload_termination_records.ps1 -ErrorAction Stop
    }
    catch {
        write-Debug $error[0].Exception
        logWrite 12 $false "Unable to download the Script! Exiting."
        exitScript
    }
    logWrite 12 $True "The Script has been downloaded ."
    $global:nextPhase++
    Write-Debug "nextPhase set to $global:nextPhase"
}


function exitScript
{
    # Get-PSSession | Remove-PSSession
    if ($debug){
        $DebugPreference = $oldDebugPreference
        Stop-Transcript
    }
    exit
}

################ main Script start ###################
if(!(Test-Path($logCSV))){
    # if log doesn't exist then must be first time we run this, so go to initialization
    Write-Debug "Entering Initialization"
    initialization
} else {
    # if log already exists, check if we need to recover
    Write-Debug "Entering Recovery"
    recovery
    connectExo
    connectSCC
    ConnectMsolService
    $tenantName = getdomain
    Write-Debug "$tenantName Returned"
    connectspo $tenantName
}

#use variable to control phases

if($nextPhase -eq 1){
    write-debug "Phase $nextPhase"
    checkModule
}

if($nextPhase -eq 2){
    write-debug "Phase $nextPhase"
    checkModulemsol
}

if($nextPhase -eq 3){
    write-debug "Phase $nextPhase"
    connectExo
}

if($nextPhase -eq 4){
    write-debug "Phase $nextPhase"
    connectSCC
}

if($nextPhase -eq 5){
    write-debug "Phase $nextPhase"
    ConnectMsolService
}

if($nextPhase -eq 6){
    write-debug "Phase $nextPhase"
    $tenantName = getdomain
    write-debug "$tenantName Returned"
}

if($nextPhase -eq 7){
    write-debug "Phase $nextPhase"
    connectspo $tenantName
}

if($nextPhase -eq 8){
    write-debug "Phase $nextPhase"
    downloadscriptlabel
}

if($nextPhase -eq 9){
    write-debug "Phase $nextPhase"
    downloadscriptDLP
}

if($nextPhase -eq 10){
    write-debug "Phase $nextPhase"
    downloadscriptRetention
}

if($nextPhase -eq 11){
    write-debug "Phase $nextPhase"
    downloadscriptInsiderRisks01
}

if($nextPhase -eq 12){
    write-debug "Phase $nextPhase"
    downloadscriptInsiderRisks02
}

if ($nextPhase -ge 11){
    write-debug "Phase $nextPhase"
    $nextScript = $LogPath + "wks-new-label.ps1"
    logWrite 11 $true "Launching $nextScript"
    if ($debug){
        Stop-Transcript
        .$nextScript -$debug
    } else {
        $nextScript
    }
}