################ Standard Variables ###################
$LogPath = "c:\temp\"
$LogCSV = "C:\temp\retentionlog.csv"
$global:nextPhase = 1
$global:recovery = $false
################ Site Variables ###################
$siteName = "wks-compliance-center-test-jorg-01"
$siteStorageQuota = 1024
$siteResourceQuota = 1024
$siteTemplate = "STS#3"
################ Tag Variables ###################
$retentionTagName = "WKS-Compliance-Tag-test-jorg-01"
$retentionTagComment = "Keep and delete tag - 3 Days"
$retentionTagAction = "KeepAndDelete"
$retentionTagDuration = 3
$retentionTagType = "ModificationAgeInDays"
$isRecordLabel = $false
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

function connectSCC
{
    try {
        Get-Command Set-Label -ErrorAction:SilentlyContinue | Out-Null
    } catch {
        Write-Host "Connecting to Compliance Center..."
        Connect-IPPSSession
        try {
            Get-Command Set-Label -ErrorAction:SilentlyContinue | Out-Null
        } catch {
            logWrite 1 $false "Couldn't connect to Compliance Center.  Exiting."
            exitScript
        }
    }
    if($global:recovery -eq $false){
        logWrite 1 $true "Connected to Compliance Center"
        $global:nextPhase++
    }
}

function ConnectMsolService
{
    try {
        $testConnection = Get-MsolDomain -ErrorAction SilentlyContinue
    }
    catch {
        Write-Host "Connecting to msol Service..."
        Connect-MsolService
        try {
        $testContact = Get-MsolContact -ErrorAction SilentlyContinue | Out-Null
        } catch {
            logWrite 2 $false "Couldn't connect to MSOL Service.  Exiting."
            exitScript
        }
    }
    if($global:recovery -eq $false){
        logWrite 2 $true "Successfully connected to MSOL Service"
        $global:nextPhase++
    }
}

function getSiteOwner
{
    # shoudl be connected to MSOL Service to set site owner
    $siteOwner = (Get-MsolUser -ErrorAction SilentlyContinue | ?{$_.UserPrincipalName -like "admin@*"}).UserPrincipalName
    #then verify
    if($siteOwner -eq $null){
        logWrite 3 $false "Failed to get or set siteOwner variable."
        exitScript
    } else {
        if($global:recovery -eq $false){
            logWrite 3 $true "siteOwner set as $siteOwner"
            $global:nextPhase++
        }
        return $siteOwner
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
        logWrite 4 $false "unable to fetch all accepted Domains."
        exitScript
    }
    if($global:recovery -eq $false){
        logWrite 4 $True "Able to get all accepted Domains."
        $global:nextPhase++
    }
    return $InitialDomain.name.split(".")[0]
}

function connectspo([string]$tenantName)
{
    $AdminURL = "https://$tenantName-admin.sharepoint.com"
    Try{
        $testConnection = Get-SpoSite -ErrorAction SilentlyContinue | Out-Null
    } catch {
        Try{
        #Connect to Office 365
        Connect-sposervice -Url $AdminURL
        } catch {
            logWrite 5 $false "Unable to connect to SPOusing $adminURL"
            exitScript
        }
    }
    if($global:recovery -eq $false){
        logWrite 5 $True "Connected to SPO using $adminURL successfully."
        $global:nextPhase++
    }
}

function connectpnp([string]$tenantName)
{
    $connectionURL = "https://$tenantName.sharepoint.com/sites/$global:siteName"
    Try
    {
        Connect-PnpOnline -Url $connectionURL -useWebLogin
    } catch {
        logWrite 6 $false "Failed to connect to PnP Powershell for $connectionURL"
        exitScript
    }
    if($global:recovery -eq $false){
        logWrite 6 $true "Connected successfully to PnP Powershell for $connectionURL"
        $global:nextPhase++
    }
}

# ------------------------------
# Create Sharepoint Online Site
# ------------------------------
function createSPOSite([string]$tenantName, [string]$siteName, [string]$siteOwner, [int]$siteStorageQuota, [int]$siteResourceQuota, [string]$siteTemplate)
{
    $url = "https://$tenantName.sharepoint.com/sites/$siteName"
    Try{
        $spoSiteCreationStatus = New-spoSite -Url $url -title $siteName -Owner $siteOwner -StorageQuota $siteStorageQuota -ResourceQuota $siteResourceQuota -Template $siteTemplate | Out-Null
        } catch {
            logWrite 7 $false "Unable to create the SharePoint site $siteName."
            exitScript
        }
        logWrite 7 $True "$siteName site created successfully."
        $global:nextPhase++
}

# -------------------
# Create Compliance Tag
# -------------------
Function CreateComplianceTag([string]$retentionTagName, [string]$retentionTagComment, [bool]$isRecordLabel, [string]$retentionTagAction, [int]$retentionTagDuration, [string]$retentionTagType)
{
    try {
        $complianceTagStatus = new-ComplianceTag -Name $retentionTagName -Comment $retentionTagComment -IsRecordLabel $isRecordLabel -RetentionAction $retentionTagAction -RetentionDuration $retentionTagDuration -RetentionType $retentionTagType | Out-Null
        } catch {
        logWrite 8 $false "Unable to create Retention Tag $retentionTagName"
        exitScript
    }
    logWrite 8 $True "Retention Tag $retentionTagName created successfully."
    $global:nextPhase++
}


# -------------------
# Create Retention Policy
# -------------------
function NewRetentionPolicy([string]$retentionPolicyName, [string]$tenantName, [string]$siteName, [string]$retentionTagName)
{
    $url = "https://$tenantName.sharepoint.com/sites/$siteName"

    #try to create policy first
    Try
    {
        #Create compliance retention Policy
        $policyStatus = New-RetentionCompliancePolicy -Name $retentionPolicyName -SharePointLocation $url -Enabled $true -ExchangeLocation All -ModernGroupLocation All -OneDriveLocation All | Out-Null
    } catch {
        #failed to create policy
        logWrite 9 $false "Unable to create the Retention Policy $retentionPolicyName"
        exitScript
    }
    
    #then, if successfull, create rule in policy
    try {
        $policyRuleStatus = New-RetentionComplianceRule -Policy $retentionPolicyName -publishComplianceTag $retentionTagName | Out-Null
    }
    catch {
         #failed to create policy
         logWrite 9 $false "Unable to create the Retention Policy Rule."
         exitScript
    }
    
    #if successful, move on
    logWrite 9 $True "Retention Policy $retentionPolicyName and Rule created successfully."
    $global:nextPhase++
}

function setlabelsposite([string]$tenantName, [string]$siteName, [string]$retentionTagName)
{
    $url = "https://$tenantName.sharepoint.com/sites/$siteName"
    #sleep for 240 seconds
    goToSleep 240

    try{
        Set-PnPLabel -List "Shared Documents" -Label $retentionTagName -SyncToItems $true
    } catch {
        logWrite 10 $false "Unable to set the Retention label to $URL."
        exitScript
    }
    logWrite 10 $True "Able to set the Retention label to $URL."
    $global:nextPhase++
}


function exitScript
{
    Get-PSSession | Remove-PSSession
    exit
}

################ main Script start ###################
Write-Host "Starting Retention Script...."

if(!(Test-Path($logCSV))){
    # if log doesn't exist then must be first time we run this, so go to initialization
    initialization
} else {
    # if log already exists, check if we need to recover#
    recovery
    connectSCC
    ConnectMsolService
    $siteOwner = getSiteOwner
    $tenantName = getdomain
    connectspo $tenantName
    connectpnp $tenantName
}

#use variable to control phases

if($nextPhase -eq 1){
    connectSCC
}

if($nextPhase -eq 2){
    ConnectMsolService
}

if($nextPhase -eq 3){
    $siteOwner = getSiteOwner
}

if($nextPhase -eq 4){
    $tenantName = getdomain
}

if($nextPhase -eq 5){
    connectspo $tenantName
}

if($nextPhase -eq 6){
    connectpnp $tenantName
}

if($nextPhase -eq 7){
    createSPOSite $tenantName $siteName $siteOwner $siteStorageQuota $siteResourceQuota $siteTemplate
}

if ($nextPhase -eq 8){
    CreateComplianceTag $retentionTagName $retentionTagComment $isRecordLabel $retentionTagAction $retentionTagDuration $retentionTagType
}

if ($nextPhase -eq 9){
    NewRetentionPolicy $retentionPolicyName $tenantName $siteName $retentionTagName
}

if ($nextPhase -eq 10){
    setlabelsposite $tenantName $siteName $retentionTagName
}

if ($nextPhase -eq 11){
exitScript
}