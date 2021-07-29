################ Define Variables ###################
$LogPath = "c:\temp\"
$LogCSV = "C:\temp\DLPLog.csv"
$global:nextPhase = 1
$global:recovery = $false
############### Functions ###################
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


function getdomain
{
    try{
        [Array]$DomainName = Get-AcceptedDomain
        $SuffixDomain = $DomainName[0]
    }
    catch {
        logWrite 4 $false "Not able to get all Accepted Domains.  Exiting."
        exit
    }
        if($global:recovery -eq $false){
        logWrite 4 $true "Successfully collected all Accepted Domains."
        $global:nextPhase++
        }
}

function createDPLPolicy
{
    
    Try{
        $params = @{
            "Name" = "WKS-Credit Card Number";
            "ExchangeLocation" ="All";
            "OneDriveLocation" = "All";
            "SharePointLocation" = "All";
            "EndpointDlpLocation" = "all";
            "Mode" = "Enable"
            }
            new-dlpcompliancepolicy @params
        }
        catch {
            logWrite 5 $false "Not able to Create DLP Policy.  Exiting."
            exit
        }
            if($global:recovery -eq $false){
            logWrite 5 $true "Successfully created the DLP Policy."
            $global:nextPhase++
            }
}

function SensitiveTypes
{
    try{
        $SensitiveTypes = @( 
            @{Name="Credit Card Number"; minCount="1"; maxcount="5"}    
            )
    
        
            $SensitiveTypesHigh = @( 
            @{Name="Credit Card Number"; minCount="6";}    
            )   
    }
    catch {
        logWrite 6 $false "Unable to set Sensitivity Types.  Exiting."
        exit
     }
        if($global:recovery -eq $false){
        logWrite 6 $true "Successfully set sensitivity Types"
        $global:nextPhase++
        }

}

function DLPCompliancerule
{
    try{
    Start-Sleep -Seconds 5
    
     New-DlpComplianceRule -Name "WKS-Credit Card Number-low" -Policy "WKS-Credit Card Number" -ContentContainsSensitiveInformation $SensitiveTypes -NotifyUser "lastmodifier"


    New-DlpComplianceRule -Name "WKS-Credit Card Number-High" -Policy "WKS-Credit Card Number" -ContentContainsSensitiveInformation $SensitiveTypesHigh -NotifyUser "LastModifier","owner" -blockaccess:$true -BlockAccessScope "All" -GenerateIncidentReport $email 
    }
        catch {
        logWrite 7 $false "Unable to create DLP Rules.  Exiting."
        exit
    }
        if($global:recovery -eq $false){
        logWrite 7 $true "Successfully created DLP Rules"
        $global:nextPhase++
        }
    
}
function exitScript
{
    Get-PSSession | Remove-PSSession
    logWrite 8 $true "Session removed successfully."
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
    getdomain
    createDPLPolicy
    SensitiveTypes
    DLPCompliancerule
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
getdomain
}

if($nextPhase -eq 5){
createDPLPolicy
}

if($nextPhase -eq 6){
SensitiveTypes
}

if($nextPhase -eq 7){
DLPCompliancerule
}

if ($nextPhase -eq 8){
exitScript
}