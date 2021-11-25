Param (
    [switch]$debug
)

################ Define Variables ###################
$LogPath = "c:\temp\"
$LogCSV = "C:\temp\download.csv"
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

function checkModule
{
    try {
        Get-Command Connect-ExchangeOnline -ErrorAction Stop | Out-Null
    

    if (Get-Command Connect-ExchangeOnline -ErrorAction Stop | Out-Null) {
     write-host "Exchange online Management module is not installed" 
    }

    else {
        install-module -name ExchangeOnlineManagement
    }
    {
        logWrite 1 $false "ExchangeOnlineManagement module is not installed! Exiting." 
      }
    }
             catch {
            logWrite 1 $true "ExchangeOnlineManagement module is now installed! Exiting." -foregroundcolor green
    }
    logWrite 1 $True "ExchangeOnlineManagement module is installed."
    $global:nextPhase++
}



function downloadscriptlabel
{

Try{
    Invoke-WebRequest -Uri https://raw.githubusercontent.com/microsoft/WKSPlus-M365SCC/main/Compliance%20Center/wks-new-label.ps1 -OutFile c:\temp\wks-new-label.ps1
}
catch {
    logWrite 1 $false "Unable to download the Script! Exiting."
    exit
}
logWrite 1 $True "The Script has been downloaded ."
$global:nextPhase++
}


function downloadscriptDLP
{

Try{
    Invoke-WebRequest -Uri https://raw.githubusercontent.com/microsoft/WKSPlus-M365SCC/main/Compliance%20Center/wks-new-DLP.ps1 -OutFile c:\temp\wks-new-DLP.ps1
}
catch {
    logWrite 2 $false "Unable to download the Script! Exiting."
    exit
}
logWrite 2 $True "The Script has been downloaded ."
$global:nextPhase++
}

function downloadscriptRetention
{

    try{
        Invoke-WebRequest -Uri https://raw.githubusercontent.com/microsoft/WKSPlus-M365SCC/main/Compliance%20Center/wks-new-retention.ps1 -OutFile c:\temp\wks-new-retention.ps1
    }
    catch {
        logWrite 3 $false "Unable to download the Script! Exiting."
        exit
    }
    logWrite 3 $True "The Script has been downloaded ."
    $global:nextPhase++
}

function downloadscriptadaptivescope
{

    try{
        Invoke-WebRequest -Uri https://raw.githubusercontent.com/microsoft/WKSPlus-M365SCC/main/Compliance%20Center/wks-new-Adaptive-scope.ps1 -OutFile c:\temp\wks-new-Adaptive-scope.ps1
    }
    catch {
        logWrite 4 $false "Unable to download the Script! Exiting."
        exit
    }
    logWrite 4 $True "The Script has been downloaded ."
    $global:nextPhase++
}

Function downloadinsiderrisks
{

    try{
        Invoke-WebRequest -Uri https://raw.githubusercontent.com/microsoft/WKSPlus-M365SCC/main/Compliance%20Center/InsiderRisks.ps1 -OutFile c:\temp\InsiderRisks.ps1
    }
    catch {
        logWrite 5 $false "Unable to download the Script! Exiting."
        exit
    }
    logWrite 5 $True "The Script has been downloaded ."
    $global:nextPhase++
}


function exitScript
{
    #remove psession if fails only
    #Get-PSSession | Remove-PSSession
    logWrite 6 $true "Session removed successfully."
}

################ main Script start ###################
cd C:\temp\

if(!(Test-Path($logCSV))){
    # if log doesn't exist then must be first time we run this, so go to initialization
    initialization
} else {
    # if log already exists, check if we need to recover
    recovery
    downloadscriptlabel
    downloadscriptDLP
    downloadscriptRetention
    downloadscriptadaptivescope
    downloadinsiderrisks

    exitScript
}

#use variable to control phases

if($nextPhase -eq 1){
    downloadscriptlabel
}

if($nextPhase -eq 2){
    downloadscriptDLP
}
if($nextPhase -eq 3){
    downloadscriptRetention
}

if ($nextPhase -ge 4){
    ./wks-new-label.ps1
}