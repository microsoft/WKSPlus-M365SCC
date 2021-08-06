################ Standard Variables ###################
$LogPath = "c:\temp\"
$LogCSV = "C:\temp\retentionlog.csv"
$global:nextPhase = 1
$global:recovery = $false
################ Site Variables ###################
$siteName = "wks-compliance-center-test-jorg-01"
$siteOwner = "" #need to set this after we connect to msonline
$siteStorageQuota = 1024
$siteResourceQuota = 1024
$siteTemplate = "STS#3"
################ Tag Variables ###################
$retentionTagName = "WKS-Compliance-Tag-test-jorg-01"
$retentionTagComment = "Keep and delete tag - 3 Days"
$retentionTagAction = "KeepAndDelete"
$retentionTagDuration = 3
$retentionTagType = "ModificationAgeInDays"
################ Policy Variables ###################
$retentionPolicyName = "WKS-Compliance-policy-test-jorg-01"



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
        Add-Content -Path $LogCSV -Value 'Phase,Result,DateTime,Status'
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

function goToSleep ([int]$seconds){
    for ($i = 1; $i -le $seconds; $i++ )
    {
        $p = ([Math]::Round($i/$seconds, 2) * 100)
        Write-Progress -Activity "Allowing time for label to be created on backend..." -Status "$p% Complete:" -PercentComplete $p
        Start-Sleep -Seconds 1
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

# -------------------
# Retrive all accepted Domains
# -------------------

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

function connectpnp([string]$tenantName)
{
    $connectionURL = "https://$tenantName.sharepoint.com/sites/$global:siteName"
    # add try/catch
    Connect-PnpOnline -Url $connectionURL -useWebLogin
}

# ------------------------------
# Create Sharepoint Online Site
# ------------------------------
function createSPOSite([string]$tenantName)
{
    param
      (
          [string]$Title  = "wks-compliance-center-test-jorg-01",
          [string]$URL = "https://$global:sharepoint.sharepoint.com/sites/wks-compliance-center-test-jorg-01",
          [string]$Owner = "adm-jorg@myjorg.be",
          [int]$StorageQuota = "1024",
          [int]$ResourceQuota = "1024",
          [string]$Template = "STS#3"
      )
   
  #Connection parameters 
  $AdminURL = "https://$global:Sharepoint-admin.sharepoint.com"
   
  Try{
      #Connect to Office 365
      Connect-sposervice -Url $AdminURL
      

             #sharepoint online create site collection powershell
          $spoSiteCreationStatus = New-spoSite -Url $URL -title $Title -Owner $Owner -StorageQuota $StorageQuota -ResourceQuota $ResourceQuota -Template $Template | Out-Null
          #write-host "Site Collection $($url) Created Successfully!" -foregroundcolor Green
      }
  catch {
          logWrite 2 $false "Unable to create the SharePoint Website."
          exit
      }
      logWrite 2 $True "Able to create the SharePoint Website."
      $global:nextPhase++
}

# -------------------
# Create Compliance Tag
# -------------------
Function CreateComplianceTag
{
    try {
        #get-compliancetag -Identity "$global:name" 
        #Write-Host "The Compliance Tag already Exists!" -ForegroundColor red
        complianceTagStatus = new-ComplianceTag -Name "$global:name" -Comment 'Keep and delete tag - 3 Days' -IsRecordLabel $false -RetentionAction "$global:retentionaction" -RetentionDuration "$global:retentionduration" -RetentionType ModificationAgeInDays | Out-Null
        
     }

     catch {
        logWrite 3 $false "unable to create Retention Tag"
        exit
    }
    logWrite 3 $True "Able to Create Retention Tag."
    $global:nextPhase++

}


# -------------------
# Create Retention Policy
# -------------------
function NewRetentionPolicy
{
    Try
    {
     
       #Create compliance retention Policy
          New-RetentionCompliancePolicy -Name "$global:Policy" -SharePointLocation "https://$global:Sharepoint.sharepoint.com/sites/WKS-compliance-center-test-jorg-01" -Enabled $true -ExchangeLocation All -ModernGroupLocation All -OneDriveLocation All
          New-RetentionComplianceRule -Policy "$global:Policy" -publishComplianceTag "$global:name"
          write-host "Retention policy and rule are Created Successfully!" -foregroundcolor Green
      
  }
  catch {
          logWrite 4 $false "Unable to create the Retention Policy and Rule."
          exit
      }
      logWrite 4 $True "The Retention policy and rule has been created."
      $global:nextPhase++
}



function setlabelsposite
{
    #sleep for 240 seconds
    
    goToSleep 240

    try{
        connect-pnponline -url "https://M365x576146.sharepoint.com/sites/wks-compliance-center-test-jorg-01" -useWebLogin
        
        Set-PnPLabel -List "Shared Documents" -Label $global:name -SyncToItems $true
    }
    catch {
        logWrite 5 $false "Unable to set the Retention label to $URL."
        exit
    }
    logWrite 5 $True "Able to set the Retention label to $URL."
    $global:nextPhase++
}


function exitScript
{
    Get-PSSession | Remove-PSSession
    logWrite 6 $true "Session removed successfully."
}
################ main Script start ###################

if(!(Test-Path($logCSV))){
    # if log doesn't exist then must be first time we run this, so go to initialization
    initialization
} else {
    # if log already exists, check if we need to recover#
    recovery
    getdomain
    createSPOSite
    CreateComplianceTag
    NewRetentionPolicy
    setlabelsposite
}

#use variable to control phases

if($nextPhase -eq 1){
getdomain
}

if($nextPhase -eq 2){
createSPOSite
}

if ($nextPhase -eq 3){
CreateComplianceTag
}

if ($nextPhase -eq 4){
NewRetentionPolicy
}

if ($nextPhase -eq 5){
setlabelsposite
}

if ($nextPhase -eq 6){
exitScript
}