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
            logWrite 2 $false "Couldn't connect to Exchange Online.  Exiting."
            exit
        }
        if($global:recovery -eq $false){
            logWrite 2 $true "Successfully connected to Exchange Online"
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
            logWrite 3 $false "Couldn't connect to Compliance Center.  Exiting."
            exit
        }
        if($global:recovery -eq $false){
            logWrite 3 $true "Successfully connected to Compliance Center"
            $global:nextPhase++
        }
    }
}

###### get domain name

$domainName = (Get-AcceptedDomain | ?{$_.Default -eq $false}).DomainName
$SPO = $domainName.Split('.')[0]

function checkModule 
{
    try {
        Get-Command Connect-SPOService -ErrorAction Stop | Out-Null
    } catch {
        logWrite 1 $false "SharePointOnlineManagement module is not installed! Exiting."
        exit
    }
    logWrite 1 $True "SharePointOnlineManagement module is installed."
    $global:nextPhase++
}

function connectSPO
{
    try {
        Get-Command Set-SPOHomeSite -ErrorAction:Stop | Out-Null
    }
    catch {
        Write-Host "Connecting to Compliance Center..."
        Connect-SPOService "https://$spo-admin.sharepoint.com"
        try {
            get-command Set-SPOHomeSite -ErrorAction:Stop | Out-Null
        } catch {
            logWrite 3 $false "Couldn't connect to Compliance Center.  Exiting."
            exit
        }
        if($global:recovery -eq $false){
            logWrite 3 $true "Successfully connected to Compliance Center"
            $global:nextPhase++
        }
    }
}

function createsite
{
    $siteName = "WKS-Compliance-Center"
    $siteUrl = "https://$spo.sharepoint.com/sites/$siteName"
    $owner = "admin@$spo.onmicrosoft.com"
    $template = "STS#3"
    $Storagequota = "1024" #MB

    Write-Host "Creating Sharepoint Site"

    try {
        New-SPOSite –url $siteUrl -Owner $owner –Template $template –Storagequota $Storagequota -Title $siteName
        } 
        catch {
        logWrite 5 $false "Error creating SharePoint Site"
        exit
    }
    logWrite 5 $true "Successfully created SharePoint Site"
    $global:nextPhase++
    write-host "Sleeping for 30 seconds..."
    Start-Sleep -Seconds 30

}

function exitScript
{
    Get-PSSession | Remove-PSSession
    logWrite 6 $true "Session removed successfully"
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
connectSPO
}

if($nextPhase -eq 5){
createsite
    }

if ($nextPhase -eq 6){
exitScript
}