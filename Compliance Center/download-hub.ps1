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

function downloadscriptlabel
{

Try{
    Invoke-WebRequest -Uri https://raw.githubusercontent.com/microsoft/WKSPlus-M365SCC/main/Compliance%20Center/new-label.ps1 -OutFile c:\temp\wks-new-label.ps1
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

function runLabelScript
{
    .\wks-new-label.ps1
}

function runDlpScript
{

}

function runRetentionScript
{
    ./wks-new-retention.ps1
}

function exitScript
{
    #remove psession if fails only
    #Get-PSSession | Remove-PSSession
    logWrite 4 $true "Session removed successfully."
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
    exitScript
}

#use variable to control phases

if($nextPhase -eq 1){
downloadscriptlabel
runLabelScript
}

if($nextPhase -eq 2){
downloadscriptDLP
runDlpScript
}
if($nextPhase -eq 3){
downloadscriptRetention
runRetentionScript
}

if ($nextPhase -eq 4){
exitScript
}