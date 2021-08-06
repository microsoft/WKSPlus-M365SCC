################ Define Variables ###################
$LogPath = "c:\temp\"
$LogCSV = "C:\temp\download1.csv"
$global:nextPhase = 1
$global:recovery = $false

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

function checkModule
{
    try {
        Get-Command Connect-ExchangeOnline -ErrorAction Stop | Out-Null
    } catch {
        logWrite 1 $false "ExchangeOnlineManagement module is not installed! Exiting."
        exit
    }
    logWrite 1 $True "ExchangeOnlineManagement module is installed."
    $global:nextPhase++
}

function checkModuleMSOL
{
    try {
        Get-Command Connect-MsolService -ErrorAction Stop | Out-Null
        
    } catch {
        logWrite 2 $false "MSOL module is not installed! Exiting."
        exit
    }
    logWrite 2 $True "MSOL module is installed."
    $global:nextPhase++
}

function connectExo
{
    try {
        Get-Command Set-Mailbox -ErrorAction Stop | Out-Null
    }
    catch {
        Write-Host "Connecting to Exchange Online..."
        Connect-ExchangeOnline
        try {
            Get-Command Set-Mailbox -ErrorAction Stop | Out-Null
        } catch {
            logWrite 3 $false "Couldn't connect to Exchange Online.  Exiting."
            exit
        }
        if($global:recovery -eq $false){
            logWrite 3 $true "Successfully connected to Exchange Online"
            $global:nextPhase++
        }
    }
}
# ----------------------------------------
# Connect to Microsoft Compliance center
# ----------------------------------------
function connectSCC
{
    try {
        Get-Command Set-Label -ErrorAction:Stop | Out-Null
    }
    catch {
        Write-Host "Connecting to Compliance Center..."
        Connect-IPPSSession
        try {
            Get-Command Set-Label -ErrorAction:Stop | Out-Null
        } catch {
            logWrite 4 $false "Couldn't connect to Compliance Center.  Exiting."
            exit
        }
        if($global:recovery -eq $false){
            logWrite 4 $true "Successfully connected to Compliance Center"
            $global:nextPhase++
        }
    }
}

# ------------------------------------
# Connect to Microsoft Online Service
# ------------------------------------
function ConnectMsolService
{
    try {
        Get-MsolDomain -ErrorAction Stop
    }
    catch {
        Write-Host "Connecting to msol Service..."
        Connect-MsolService
        try {
        $testContact = Get-MsolContact -ErrorAction Stop | Out-Null
        } catch {
            logWrite 5 $false "Couldn't connect to MSOL Service.  Exiting."
            exit
        }
        if($global:recovery -eq $false){
            logWrite 5 $true "Successfully connected to MSOL Service"
            $global:nextPhase++
        }
    }
}


# ------------------------------------
# get accepted domains
# ------------------------------------

Function getdomain
{
    try{
        $InitialDomain = Get-MsolDomain | Where-Object {$_.IsInitial -eq $true}
   }catch {
        logWrite 6 $false "unable to fetch all accepted Domains."
        exit
    }
    logWrite 6 $True "Able to get all accepted Domains."
    $global:nextPhase++
    return $InitialDomain.name.split(".")[0]
}



function connectspo([string]$tenantName)
{
    $AdminURL = "https://$tenantName-admin.sharepoint.com"
    Try{
        #Connect to Office 365
        Connect-sposervice -Url $AdminURL
        }
        catch {
            logWrite 7 $false "Unable to create the SharePoint Website."
            exit
        }
        logWrite 7 $True "Able to create the SharePoint Website."
        $global:nextPhase++
  
}

function downloadscriptlabel
{

Try{
    Invoke-WebRequest -Uri https://raw.githubusercontent.com/microsoft/WKSPlus-M365SCC/main/Compliance%20Center/wks-new-label.ps1 -OutFile c:\temp\wks-new-label.ps1
}
catch {
    logWrite 8 $false "Unable to download the Script! Exiting."
    exit
}
logWrite 8 $True "The Script has been downloaded ."
$global:nextPhase++
}


function downloadscriptDLP
{

Try{
    Invoke-WebRequest -Uri https://raw.githubusercontent.com/microsoft/WKSPlus-M365SCC/main/Compliance%20Center/wks-new-DLP.ps1 -OutFile c:\temp\wks-new-DLP.ps1
}
catch {
    logWrite 9 $false "Unable to download the Script! Exiting."
    exit
}
logWrite 9 $True "The Script has been downloaded ."
$global:nextPhase++
}

function downloadscriptRetention
{

    try{
        Invoke-WebRequest -Uri https://raw.githubusercontent.com/microsoft/WKSPlus-M365SCC/main/Compliance%20Center/wks-new-retention.ps1 -OutFile c:\temp\wks-new-retention.ps1
    }
    catch {
        logWrite 10 $false "Unable to download the Script! Exiting."
        exit
    }
    logWrite 10 $True "The Script has been downloaded ."
    $global:nextPhase++
}

################ main Script start ###################
if(!(Test-Path($logCSV))){
    # if log doesn't exist then must be first time we run this, so go to initialization
    initialization
} else {
    # if log already exists, check if we need to recover
    recovery
    connectExo
    connectSCC
    ConnectMsolService
    $tenantName = getdomain
    connectspo $tenantName
}

#use variable to control phases

if($nextPhase -eq 1){
    checkModule
}

if($nextPhase -eq 2){
    checkModulemsol
}

if($nextPhase -eq 3){
    connectExo
}

if($nextPhase -eq 4){
    connectSCC
}

if($nextPhase -eq 5){
    ConnectMsolService
}

if($nextPhase -eq 6){
    $tenantName = getdomain
}

if($nextPhase -eq 7){
    connectspo $tenantName
}

if($nextPhase -eq 8){
    downloadscriptlabel
}

if($nextPhase -eq 9){
    downloadscriptDLP
}

if($nextPhase -eq 10){
    downloadscriptRetention
}

if ($nextPhase -ge 11){
    $nextScript = $LogPath + "wks-new-label.ps1"
    logWrite 11 $true "Launching $nextScript"
    .$nextScript
}