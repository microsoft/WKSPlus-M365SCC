################ Define Variables ###################
$LogPath = "c:\temp\"
$LogCSV = "C:\temp\retentionlog.csv"
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

function ConnectMsolService
{
    try {
        Get-Command Get-MsolDomain -ErrorAction Stop
    }
    catch {
        Write-Host "Connecting to msol Service..."
        Connect-MsolService
        try {
            Get-Command Get-MsolContact -ErrorAction Stop
        } catch {
            logWrite 5 $false "Couldn't connect to MSOL Service.  Exiting."
            exit
        }
        if($global:recovery -eq $false)
        {
            logWrite 5 $true "Successfully connected to MSOL Service"
            $global:nextPhase++
        }
    }
}

Function getdomain
{
    try{
        $InitialDomain = Get-MsolDomain -TenantId $customer.TenantId | Where-Object {$_.IsInitial -eq $true}
        $SharepointURL = "$($InitialDomain.name.split(".")[0]).sharepoint.com"
   }catch {
        logWrite 6 $false "unable to fetch all accepted Domains."
        exit
    }
    logWrite 6 $True "Able to get all accepted Domains."
    $global:nextPhase++


}


function connectSPO
{
    try {
        Get-Command Connect-SPOService -ErrorAction:Stop | Out-Null
    }
    catch {
        Write-Host "Connecting to Compliance Center..."
        Connect-SPOService "https://$SharepointURL-admin.sharepoint.com"
        try {
            get-command Connect-SPOService -ErrorAction:Stop | Out-Null
        } catch {
            logWrite 7 $false "Couldn't connect to Sharepoint Online.  Exiting."
            exit
        }
        if($global:recovery -eq $false){
            logWrite 7 $true "Successfully connected to Sharepoint Online"
            $global:nextPhase++
        }
    }
}


function exitScript
{
    Get-PSSession | Remove-PSSession
    logWrite 8 $true "Session removed successfully"
}

################ main Script start ###################

if(!(Test-Path($logCSV))){
    # if log doesn't exist then must be first time we run this, so go to initialization
    initialization
} else {
    # if log already exists, check if we need to recover#
    recovery
    checkModule
    checkModuleMSOL
    connectExo
    connectSCC
    ConnectMsolService
    getdomain
    connectSPO
    

}

#use variable to control phases

if($nextPhase -eq 0){
initialization
}

if($nextPhase -eq 1){
checkModule
}

if($nextPhase -eq2){
checkModuleMSOL    
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
getdomain
}

if($nextPhase -eq 7){
connectSPO
}

if ($nextPhase -eq 8){
exitScript
}