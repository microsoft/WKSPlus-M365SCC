Param (
    [switch]$debug
)

################ Define Variables ###################
$LogPath = "c:\temp\"
$LogCSV = "C:\temp\DLPLog.csv"
#$LogPath = "$env:UserProfile\Desktop\SCLabFiles\Scripts\"
#$LogCSV = "$env:UserProfile\Desktop\SCLabFiles\Scripts\Progress_DLP_Log.csv"
$global:nextPhase = 1
$global:recovery = $false

################ global ###################

$policyname = "WKS Compliance Policy-01"
$rulesetname = "WKS-Copmpliance-Rule-set"

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
        Add-Content -Path $LogCSV -Value 'Phase,Result,DateTime,Status'
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

#### check for existing DLP Policy. ####

function createDLPpolicy
{
    if ($SkipDLP -eq $false)
        {        
            try
                {
                    if (Get-DlpCompliancePolicy -Identity $policyname)
                        {
                            write-host " The DLP Compliance Policy already Exists "
                        }
                    else
                        {
                            $params = @{
                                "Name" = $policyname;
                                "Comment" = "Helps detect the presence of information commonly considered to be financial information in United States, including information like credit card, account information, and debit card numbers."
                                "ExchangeLocation" ="All";
                                "OneDriveLocation" = "All";
                                "SharePointLocation" = "All";
                                "Teamslocation" = "All";
                                "Mode" = "Enable"
                                }
                            }
                        New-dlpcompliancepolicy @params
                }
                catch 
                    {
                        write-Debug $Error[0].Exception
                        logWrite 1 $false "Unable to create DLP Policy."
                        exitScript
                    }
        }
}

    ###### Create DLP Compliance Rule ############
function createDLPComplianceRuleLow
{
    if ($SkipDLP -eq $false)
        {
            try
                {
                    $senstiveinfo = @(@{Name ="Credit Card Number"; minCount = "1"},@{Name ="International Banking Account Number (IBAN)"; minCount = "1"},@{Name ="U.S. Bank Account Number"; minCount = "1"})
                    $Rulevalue = @{
                    "Name" = $rulesetname;
                    "Comment" = "Helps detect the presence of information commonly considered to be subject to the GLBA act in America. like driver's license and passport number.";
                    "Policy" = "WKS Compliance DLP Policy-01";
                    "ContentContainsSensitiveInformation"=$senstiveinfo;
                    "AccessScope"= "NotInOrganization";
                    "Disabled" =$false;
                    "ReportSeverityLevel"="High";
                    "GenerateIncidentReport" = "SiteAdmin";
                    "IncidentReportContent" = "DocumentLastModifier", "Detections", "Severity", "DetectionDetails", "OriginalContent";
                    "NotifyUser" = "LastModifier","owner";
                    "BlockAccess" = $true;
                    "BlockAccessScope" = "All";
                        }
                    New-DlpComplianceRule @rulevalue 
                }
                catch 
                    {
                        write-Debug $Error[0].Exception
                        logWrite 2 $false "Unable to create DLP Rule."
                        exitScript
                    }
    
        }
   }
function exitScript
{
   logWrite 3 $true "Session removed successfully."
}
################ main Script start ###################

if(!(Test-Path($logCSV))){
    # if log doesn't exist then must be first time we run this, so go to initialization
    initialization
} else {
    # if log already exists, check if we need to recover#
    recovery
    createDLPpolicy
    createDLPComplianceRuleLow
    
}

#use variable to control phases
  
if($nextPhase -eq 1){
    createDLPpolicy
    }
    
if ($nextPhase -eq 2){
    createDLPComplianceRuleLow
    }
   
    if ($nextPhase -eq 3){
    exitScript
    }