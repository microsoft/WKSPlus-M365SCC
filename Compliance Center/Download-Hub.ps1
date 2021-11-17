<#
 .Synopsis
  Pré-configuration for WorkshopPLUS: Securiy and Compliance: Compliance Center

 .Description
  Displays a visual representation of a calendar. This function supports multiple months
  and lets you highlight specific date ranges or days.

    ##################################################################################################
    # This sample script is not supported under any Microsoft standard support program or service.   #
    # This sample script is provided AS IS without warranty of any kind.                             #
    # Microsoft further disclaims all implied warranties including, without limitation, any implied  #
    # warranties of merchantability or of fitness for a particular purpose. The entire risk arising  #
    # out of the use or performance of the sample script and documentation remains with you. In no   #
    # event shall Microsoft, its authors, or anyone else involved in the creation, production, or    #
    # delivery of the scripts be liable for any damages whatsoever (including, without limitation,   #
    # damages for loss of business profits, business interruption, loss of business information,     #
    # or other pecuniary loss) arising out of the use of or inability to use the sample script or    #
    # documentation, even if Microsoft has been advised of the possibility of such damages.          #
    ##################################################################################################

 .Parameter Start
  The first month to display.

 .Parameter End
  The last month to display.

 .Parameter FirstDayOfWeek
  The day of the month on which the week begins.

 .Parameter HighlightDay
  Specific days (numbered) to highlight. Used for date ranges like (25..31).
  Date ranges are specified by the Windows PowerShell range syntax. These dates are
  enclosed in square brackets.

 .Parameter HighlightDate
  Specific days (named) to highlight. These dates are surrounded by asterisks.

 .Example
   # Show a default display of this month.
   Show-Calendar

 .Example
   # Display a date range.
   Show-Calendar -Start "March, 2010" -End "May, 2010"

 .Example
   # Highlight a range of days.
   Show-Calendar -HighlightDay (1..10 + 22) -HighlightDate "December 25, 2008"
#>

##
## New-ModuleManifest -Path .\Scripts\TestModule.psd1 -Author 'Marcelo Hunecke' -CompanyName 'Microsoft' -RootModule 'WorkshopSnC.psm1' -FunctionsToExport @('Get-RegistryKey','Set-RegistryKey') -Description 'This is a Workshop Security and Compliance module.'
##

Param (
    [CmdletBinding()]
    [switch]$debug,
    [switch]$SkipSensitivityLabels,
    [switch]$SkipRetentionPolicies,
    [switch]$SkipDLP,
    [switch]$InsiderRisksOnly
)

# -----------------------------------------------------------
# Write the log
# -----------------------------------------------------------
function logWrite([int]$phase, [bool]$result, [string]$logstring)
{
    if ($result)
        {
            Add-Content -Path $LogCSV -Value "$phase,$result,$(Get-Date),$logString"
            Write-Host -ForegroundColor Green "$(Get-Date) - Phase $phase : $logstring"
        } 
    else 
        {
            Write-Host -ForegroundColor Red "$(Get-Date) - Phase $phase : $logstring"
        }
}

# -----------------------------------------------------------
# Sleep x seconds
# -----------------------------------------------------------
function goToSleep ([int]$seconds){
    for ($i = 1; $i -le $seconds; $i++ )
    {
        $p = ([Math]::Round($i/$seconds, 2) * 100)
        Write-Progress -Activity "Allowing time for the creation on backend..." -Status "$p% Complete:" -PercentComplete $p
        Start-Sleep -Seconds 1
    }
}

# -----------------------------------------------------------
# Start the recovery steps
# -----------------------------------------------------------
function recovery
{
    Write-host "Starting recovery..."
    Set-Location -Path $LogPath
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

    if ($lastEntryResult -eq $false)
        {
            if ($lastEntryPhase -eq $savedLog[$lastEntry2].Phase)
                {
                    WriteHost -ForegroundColor Red "The script has failed at Phase $lastEntryPhase repeatedly.  PLease check with your instructor."
                    exitScript
                }
                else 
                    {
                        Write-Host "There was a problem with Phase $lastEntryPhase, so trying again...."
                        $global:nextPhase = $lastEntryPhase
                        Write-Debug "nextPhase set to $global:nextPhase"
                    }
        }
            else
                {
                    # set the phase
                    Write-Host "Phase $lastEntryPhase was successful, so picking up where we left off...."
                    $global:nextPhase = $lastEntryPhase + 1
                    write-Debug "nextPhase set to $global:nextPhase"
                }
}


# -----------------------------------------------------------
# Test the log path (Step 0)
# -----------------------------------------------------------
function initialization
{
    $pathExists = Test-Path($LogPath)
    if (!$pathExists)
        {
            New-Item -ItemType "directory" -Path $LogPath -ErrorAction SilentlyContinue | Out-Null
        }
        Set-Location -Path $LogPath
        Add-Content -Path $LogCSV -Value '"Phase","Result","DateTime","Status"'
        logWrite 0 $true "Initialization completed"
}

# -----------------------------------------------------------
# Connect to AzureAD (Step 1)
# -----------------------------------------------------------
function ConnectAzureAD
{
    try 
        {
            Write-Debug "Get-AzureADDirectoryRole -ErrorAction stop"
            $testConnection = Get-AzureADDirectoryRole -ErrorAction stop | Out-Null #if true (Already Connected)
        }
        catch
            {
                try
                    {
                        write-Debug $error[0].Exception
                        Write-Host "Connecting to Azure AD..."
                        Connect-AzureAD -ErrorAction stop | Out-Null
                    }
                    catch    
                        {
                            try
                                {
                                    write-Debug $error[0].Exception
                                    Write-Host "Installing Azure AD PowerShell Module..."
                                    Install-Module AzureAD -Force -AllowClobber
                                    Connect-AzureAD -ErrorAction stop | Out-Null
                                }
                                catch
                                    {
                                        write-Debug $error[0].Exception
                                        logWrite 1 $false "Couldn't connect to Azure AD. Exiting."
                                        exitScript
                                    }
                       
                        }
            }
    if($global:recovery -eq $false)
        {
            logWrite 1 $true "Successfully connected to Azure AD."
            if ($InsiderRisksOnly -eq $true)
            {
                $global:nextPhase = 41
            }
            else 
                {
                    $global:nextPhase++
                }
            Write-Debug "nextPhase set to $global:nextPhase"
        }
}

# -----------------------------------------------------------
# Connect to Microsoft Online (Step 2)
# -----------------------------------------------------------
function ConnectMsol
{
    try 
    {
        Write-Debug "Get-MSOLCompanyInformation -ErrorAction stop"
        $testConnection = Get-MSOLCompanyInformation -ErrorAction stop | Out-Null #if true (Already Connected)
    }
    catch
        {
            try
                {
                    write-Debug $error[0].Exception
                    Write-Host "Connecting to Microsoft Online..."
                    Connect-MSOLService -ErrorAction stop | Out-Null
                }
                catch    
                    {
                        try
                            {
                                write-Debug $error[0].Exception
                                Write-Host "Installing Microsoft Online PowerShell Module..."
                                Install-Module MSOnline -Force -AllowClobber -ErrorAction stop | Out-Null
                                Connect-MSOLService -ErrorAction stop | Out-Null
                            }
                            catch
                                {
                                    write-Debug $error[0].Exception
                                    logWrite 2 $false "Couldn't connect to Microsoft Online. Exiting."
                                    exitScript
                                }
                   
                    }
        }
        if($global:recovery -eq $false)
            {
                logWrite 2 $true "Successfully connected to Microsoft Online"
                $global:nextPhase++
                Write-Debug "nextPhase set to $global:nextPhase"
            }
}

# -----------------------------------------------------------
# Connect to Exchange Online (Step 3)
# -----------------------------------------------------------
function ConnectEXO
{
    try 
    {
        Write-Debug "Get-OrganizationConfig -ErrorAction stop"
        $testConnection = Get-OrganizationConfig -ErrorAction stop | Out-Null #if true (Already Connected)
    }
    catch
        {
            try
                {
                    write-Debug $error[0].Exception
                    Write-Host "Connecting to Exchange Online..."
                    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction stop | Out-Null
                }
                catch    
                    {
                        try
                            {
                                write-Debug $error[0].Exception
                                Write-Host "Installing Exchange Online PowerShell Module..."
                                Install-Module ExchangeOnlineManagement -Force -AllowClobber -ErrorAction stop | Out-Null
                                Connect-ExchangeOnline -ShowBanner:$false -ErrorAction stop | Out-Null
                            }
                            catch
                                {
                                    write-Debug $error[0].Exception
                                    logWrite 3 $false "Couldn't connect to Exchange Online. Exiting."
                                    exitScript
                                }
                   
                    }
        }
        if($global:recovery -eq $false)
            {
                logWrite 3 $true "Successfully connected to Exchange Online"
                $global:nextPhase++
                Write-Debug "nextPhase set to $global:nextPhase"
            }
}

# -----------------------------------------------------------
# Connect to Compliance Center (Step 4)
# -----------------------------------------------------------
function ConnectSCC
{
    try 
    {
        Write-Debug "Get-Label -ErrorAction stop"
        $testConnection = Get-Label -ErrorAction stop | Out-Null #if true (Already Connected)
    }
    catch
        {
            try
                {
                    write-Debug $error[0].Exception
                    Write-Host "Connecting to Compliance Center..."
                    Connect-IPPSSession -ErrorAction stop -WarningAction SilentlyContinue | Out-Null
                }
                catch    
                    {
                        try
                            {
                                write-Debug $error[0].Exception
                                Write-Host "Installing Compliance Center PowerShell Module..."
                                #Install-Module ExchangeOnlineManagement -Force -AllowClobber #Not required, but it was already installed on the previous step
                                Connect-IPPSSession -ErrorAction stop -WarningAction SilentlyContinue | Out-Null
                            }
                            catch
                                {
                                    write-Debug $error[0].Exception
                                    logWrite 4 $false "Couldn't connect to Compliance Center. Exiting."
                                    exitScript
                                }
                   
                    }
        }
        if($global:recovery -eq $false)
            {
                logWrite 4 $true "Successfully connected to Compliance Center"
                $global:nextPhase++
                Write-Debug "nextPhase set to $global:nextPhase"
            }
}

# -------------------------------------------------------
# Connect to Microsoft Teams (Step 5)
# -------------------------------------------------------
function ConnectTeams
{
    try 
    {
        Write-Debug "Get-CsTenant -ErrorAction stop"
        $testConnection = Get-CsTenant -ErrorAction stop | Out-Null #if true (Already Connected)
    }
    catch
        {
            try
                {
                    write-Debug $error[0].Exception
                    Write-Host "Connecting to Microsoft Teams..."
                    Connect-MicrosoftTeams -ErrorAction stop | Out-Null
                }
                catch    
                    {
                        try
                            {
                                write-Debug $error[0].Exception
                                Write-Host "Installing Microsoft Teams PowerShell Module..."
                                Install-Module MicrosoftTeams -Force -AllowClobber -ErrorAction stop | Out-Null
                                Connect-MicrosoftTeams -ErrorAction stop | Out-Null
                            }
                            catch
                                {
                                    write-Debug $error[0].Exception
                                    logWrite 5 $false "Couldn't connect to Microsoft Teams. Exiting."
                                    exitScript
                                }
                   
                    }
        }
        if($global:recovery -eq $false)
            {
                logWrite 5 $true "Successfully connected to Microsoft Teams"
                $global:nextPhase++
                Write-Debug "nextPhase set to $global:nextPhase"
            }
}

# -------------------------------------------------------
# Get Tenant Name (Step 6)
# Funcion required for SharePoint and PNP connections
# -------------------------------------------------------
Function getdomain
{
    try
        {
            Write-Debug "$InitialDomain = Get-MsolDomain -ErrorAction stop | Where-Object {$_.IsInitial -eq $true}"
            $InitialDomain = Get-MsolDomain -ErrorAction stop | Where-Object {$_.IsInitial -eq $true}
        }
        catch
            {
                write-Debug $error[0].Exception
                logWrite 6 $false "Unable to fetch Tenant name."
                exitScript
            }
    Write-Debug "Initial domain: $InitialDomain"
    if($global:recovery -eq $false)
        {
            logWrite 6 $True "Successfully got Tenant Name."
            $global:nextPhase++
            Write-Debug "nextPhase set to $global:nextPhase"
        }
    return $InitialDomain.name.split(".")[0]
}

# -------------------------------------------------------
# Connect to SharePoint Online (Step 7)
# -------------------------------------------------------
function ConnectSPO([string]$tenantName)
{
    $AdminURL = "https://$tenantName-admin.sharepoint.com"
    try 
    {
        Write-Debug "Get-SPOTenant -ErrorAction stop"
        $testConnection = Get-SPOTenant -ErrorAction stop | Out-Null #if true (Already Connected)
    }
    catch
        {
            try
                {
                    write-Debug $error[0].Exception
                    Write-Host "Connecting to SharePoint Online..."
                    Connect-SPOService -Url $AdminURL -ErrorAction stop -WarningAction SilentlyContinue | Out-Null
                }
                catch    
                    {
                        try
                            {
                                write-Debug $error[0].Exception
                                Write-Host "Installing SharePoint Online PowerShell Module..."
                                Install-Module Microsoft.Online.SharePoint.PowerShell -Force -AllowClobber -ErrorAction stop | Out-Null
                                Connect-SPOService -Url $AdminURL -ErrorAction stop -WarningAction SilentlyContinue | Out-Null
                            }
                            catch
                                {
                                    write-Debug $error[0].Exception
                                    logWrite 7 $false "Couldn't connect to SharePoint Online. Exiting."
                                    exitScript
                                }
                   
                    }
        }
        if($global:recovery -eq $false)
            {
                logWrite 7 $true "Successfully connected to SharePoint Online"
                $global:nextPhase++
                Write-Debug "nextPhase set to $global:nextPhase"
            }
}

# -------------------------------------------------------
# Connect to PNP Online (Step 8)
# -------------------------------------------------------
function ConnectPNP([string]$tenantName)
{
    $connectionURL = "https://$tenantName.sharepoint.com/sites/$global:siteName"
    try 
    {
        Write-Debug "Get-PNPChangeLog -ErrorAction stop"
        $testConnection = Get-PNPChangeLog -ErrorAction stop | Out-Null #if true (Already Connected)
    }
    catch
        {
            try
                {
                    write-Debug $error[0].Exception
                    Write-Host "Connecting to PNP Online..."
                    Connect-PnpOnline -Url $connectionURL -UseWebLogin -ErrorAction stop -WarningAction SilentlyContinue | Out-Null
                }
                catch    
                    {
                        try
                            {
                                write-Debug $error[0].Exception
                                Write-Host "Installing PNP Online PowerShell Module..."
                                Install-Module PNP.PowerShell -Force -AllowClobber -ErrorAction stop | Out-Null
                                Connect-PnpOnline -Url $connectionURL -UseWebLogin -ErrorAction stop -WarningAction SilentlyContinue | Out-Null
                            }
                            catch
                                {
                                    write-Debug $error[0].Exception
                                    logWrite 8 $false "Couldn't connect to PNP Online. Exiting."
                                    exitScript
                                }
                   
                    }
        }
        if($global:recovery -eq $false)
            {
                logWrite 8 $true "Successfully connected to PNP Online"
                $global:nextPhase++
                Write-Debug "nextPhase set to $global:nextPhase"
            }
}


# -------------------------------------------------------
# Download Workshop Script (Step 9)
# -------------------------------------------------------
function downloadscripts
{
    try
        {
            #General scripts
            Invoke-WebRequest -Uri https://raw.githubusercontent.com/microsoft/WKSPlus-M365SCC/main/Compliance%20Center/Update-Hub.ps1 -OutFile "$($LogPath)Update-Hub.ps1" -ErrorAction Stop
            #Labels scritp
            Write-Debug "Invoke-WebRequest -Uri https://raw.githubusercontent.com/microsoft/WKSPlus-M365SCC/main/Compliance%20Center/wks-new-label.ps1 -OutFile $($LogPath)wks-new-label.ps1 -ErrorAction Stop"
            Invoke-WebRequest -Uri https://raw.githubusercontent.com/microsoft/WKSPlus-M365SCC/main/Compliance%20Center/wks-new-label.ps1 -OutFile "$($LogPath)wks-new-label.ps1" -ErrorAction Stop
            #DLP Script
            Write-Debug "Invoke-WebRequest -Uri https://raw.githubusercontent.com/microsoft/WKSPlus-M365SCC/main/Compliance%20Center/wks-new-DLP.ps1 -OutFile $($LogPath)wks-new-DLP.ps1 -ErrorAction Stop"
            Invoke-WebRequest -Uri https://raw.githubusercontent.com/microsoft/WKSPlus-M365SCC/main/Compliance%20Center/wks-new-DLP.ps1 -OutFile "$($LogPath)wks-new-DLP.ps1" -ErrorAction Stop
            #Retention script
            Write-Debug "Invoke-WebRequest -Uri https://raw.githubusercontent.com/microsoft/WKSPlus-M365SCC/main/Compliance%20Center/wks-new-retention.ps1 -OutFile $($LogPath)wks-new-retention.ps1 -ErrorAction Stop"
            Invoke-WebRequest -Uri https://raw.githubusercontent.com/microsoft/WKSPlus-M365SCC/main/Compliance%20Center/wks-new-retention.ps1 -OutFile "$($LogPath)wks-new-retention.ps1" -ErrorAction Stop
            #InsiderRisk scripts
            Write-Debug "Invoke-WebRequest -Uri https://raw.githubusercontent.com/microsoft/WKSPlus-M365SCC/main/Compliance%20Center/wks-new-HRConnector.ps1 -OutFile $($LogPath)wks-new-HRConnector.ps1 -ErrorAction Stop"
            Invoke-WebRequest -Uri https://raw.githubusercontent.com/microsoft/WKSPlus-M365SCC/main/Compliance%20Center/wks-new-HRConnector.ps1 -OutFile "$($LogPath)wks-new-HRConnector.ps1" -ErrorAction Stop
            Write-Debug "Invoke-WebRequest -Uri https://raw.githubusercontent.com/microsoft/m365-hrconnector-sample-scripts/master/upload_termination_records.ps1 -OutFile $($LogPath)upload_termination_records.ps1 -ErrorAction Stop"
            Invoke-WebRequest -Uri https://raw.githubusercontent.com/microsoft/m365-hrconnector-sample-scripts/master/upload_termination_records.ps1 -OutFile "$($LogPath)upload_termination_records.ps1" -ErrorAction Stop
        } 
        catch 
            {
                write-Debug $error[0].Exception
                logWrite 9 $false "Unable to download the workshop scripts from GitHub! Exiting."
                exitScript
            }
    if($global:recovery -eq $false)
        {
            logWrite 9 $True "Successfully downloaded the workshop scripts."
            $global:nextPhase++ #10
            $global:nextPhase++ #11
            Write-Debug "nextPhase set to $global:nextPhase"
        }
}       

#######################################################################################
#########                S E N S I T I V I T Y      L A B E L                ##########
#######################################################################################

# -------------------------------------------------------
# Create Sensitivity label (Step 11)
# -------------------------------------------------------
function SensitivityLabel_Label
{
    <#
    TO DO:
    Need to check to see if label exists in case the failure occured after cmd was successful, such as if they close the PS window. Maybe just check if label exists, and use Set-Label if so.
    #>

    #label
    $labelDisplayName = "WKS Highly Confidential"
    $global:labelName = "WKS-Highly-Confidential"
    $labelTooltip = "Contains Highly confidential info"
    $labelComment = "Documents with this label contain sensitive data."

    if ($SkipSensitivityLabels -eq $false)
        {
            $domainName = (Get-AcceptedDomain | Where-Object{$_.Default -eq $true}).DomainName
            $Encpermission = $domainname + ":VIEW,VIEWRIGHTSDATA,DOCEDIT,EDIT,PRINT,EXTRACT,REPLY,REPLYALL,FORWARD,OBJMODEL"
            try 
                {
                    write-Debug "New-Label -DisplayName $labelDisplayName -Name $global:labelName -ToolTip $labelTooltip -Comment $labelComment -ContentType file,Email,Site,UnifiedGroup -EncryptionEnabled:$true -SiteAndGroupProtectionEnabled:$true -EncryptionPromptUser:$true -EncryptionRightsDefinitions $Encpermission -SiteAndGroupProtectionPrivacy private -EncryptionDoNotForward:$true -SiteAndGroupProtectionAllowLimitedAccess:$true -ErrorAction stop | Out-Null"
                    $labelStatus = New-Label -DisplayName $labelDisplayName -Name $global:labelName -ToolTip $labelTooltip -Comment $labelComment -ContentType "file","Email","Site","UnifiedGroup" -EncryptionEnabled:$true -SiteAndGroupProtectionEnabled:$true -EncryptionPromptUser:$true -EncryptionRightsDefinitions $Encpermission -SiteAndGroupProtectionPrivacy "private" -EncryptionDoNotForward:$true -SiteAndGroupProtectionAllowLimitedAccess:$true -ErrorAction stop | Out-Null
                } 
                catch 
                    {
                        write-Debug $error[0].Exception
                        logWrite 11 $false "Error creating Sensitivity label"
                        exitScript
                    }
            if($global:recovery -eq $false)
                {
                    logWrite 11 $True "Successfully created Sensitivity label."
                    $global:nextPhase++
                    Write-Debug "nextPhase set to $global:nextPhase"
                }

            goToSleep 30
        }
        else
            {
                logWrite 11 $True "Skipped Sensitivity label."
                $global:nextPhase++
                Write-Debug "nextPhase set to $global:nextPhase"
            }
}

# -------------------------------------------------------
# Create Sensitivity policy (Step 12)
# -------------------------------------------------------
function SensitivityLabel_Policy
{
    <#
    TO DO:
    - Need to check to see if label policy exists in case the failure occured after cmd was successful, such as if they close the PS window. Maybe just check if label exists, and use Set-Label if so.
    - Need to make sure the labele exists
    #>

    #label policy
    $labelPolicyName = "WKS-Highly-confidential-publish"

    if ($SkipSensitivityLabels -eq $false)
        {
            try 
                {
                   write-Debug "New-LabelPolicy -name $labelPolicyName -Settings @{mandatory=$false} -AdvancedSettings @{requiredowngradejustification= $true} -Labels $global:labelName -ErrorAction stop | Out-Null"
                   New-LabelPolicy -name $labelPolicyName -Settings @{mandatory=$false} -AdvancedSettings @{requiredowngradejustification= $true} -Labels $global:labelName -ErrorAction stop | Out-Null
                } 
                catch 
                    {
                        write-Debug $error[0].Exception
                        logWrite 12 $false "Error creating Sensitivity label policy"
                        exitScript
                    }
            
            if($global:recovery -eq $false)
                {
                    logWrite 12 $True "Successfully created Sensitivity label policy."
                    $global:nextPhase++ #13
                    $global:nextPhase++ #14
                    $global:nextPhase++ #15
                    $global:nextPhase++ #16
                    $global:nextPhase++ #17
                    $global:nextPhase++ #18
                    $global:nextPhase++ #19
                    $global:nextPhase++ #20
                    $global:nextPhase++ #21
                    Write-Debug "nextPhase set to $global:nextPhase"
                }
        }
        else
            {
                logWrite 12 $True "Skipped Sensitivity label."
                $global:nextPhase++ #13
                $global:nextPhase++ #14
                $global:nextPhase++ #15
                $global:nextPhase++ #16
                $global:nextPhase++ #17
                $global:nextPhase++ #18
                $global:nextPhase++ #19
                $global:nextPhase++ #20
                $global:nextPhase++ #21
                Write-Debug "nextPhase set to $global:nextPhase"
            }
}

#######################################################################################
#########                  R E T E N T I O N     P O L I C Y                 ##########
#######################################################################################

# -------------------------------------------------------
# Retention policy - Get the Site Owner (Step 21)
# -------------------------------------------------------
function RetentionPolicy_GetSiteOwner
{
    if ($SkipRetentionPolicies -eq $false)
        {
            try 
                {
                    # should be connected to MSOL Service to set site owner
                    write-debug "$global:siteOwner = (Get-MsolUser -ErrorAction SilentlyContinue | Where-Object{$_.UserPrincipalName -like admin@*}).UserPrincipalName"
                    $global:siteOwner = (Get-MsolUser -ErrorAction SilentlyContinue | Where-Object{$_.UserPrincipalName -like "admin@*"}).UserPrincipalName
                }
                catch 
                    {
                        write-Debug $error[0].Exception
                        logWrite 21 $false "Failed to get or set siteOwner variable."
                        exitScript
                    }
            
            if($global:recovery -eq $false)
                {
                    logWrite 21 $True "Successfully got the Site Owner."
                    $global:nextPhase++
                    Write-Debug "nextPhase set to $global:nextPhase"
                    return $global:siteOwner | Out-Null
                }
        }
        else 
            {
                logWrite 21 $True "Skipped Retention Policy."
                $global:nextPhase++
                Write-Debug "nextPhase set to $global:nextPhase" 
            }
}

# -------------------------------------------------------
# Retention policy - Create Sharepoint Online Site (Step 22)
# -------------------------------------------------------
function RetentionPolicy_CreateSPOSite([string]$tenantName, [string]$global:siteName, [string]$global:siteOwner, [int]$siteStorageQuota, [int]$siteResourceQuota, [string]$siteTemplate)
{
    #Site Variables
    $global:siteName = "wks-compliance-center6"
    $siteStorageQuota = 1024
    $siteResourceQuota = 1024
    $siteTemplate = "STS#3"
        
    if ($SkipRetentionPolicies -eq $false)
        {
            $url = "https://$tenantName.sharepoint.com/sites/$global:siteName"
            try
                {
                    write-debug "New-spoSite -Url $url -title $global:siteName -Owner $global:siteOwner -StorageQuota $siteStorageQuota -ResourceQuota $siteResourceQuota -Template $siteTemplate -ErrorAction Stop | Out-Null"
                    $spoSiteCreationStatus = New-spoSite -Url $url -title $global:siteName -Owner $global:siteOwner -StorageQuota $siteStorageQuota -ResourceQuota $siteResourceQuota -Template $siteTemplate -ErrorAction Stop | Out-Null
                } 
                catch 
                    {
                        write-Debug $error[0].Exception
                        logWrite 22 $false "Unable to create the SharePoint site $global:siteName."
                        exitScript
                    }
            if($global:recovery -eq $false)
                {
                    logWrite 22 $True "$global:siteName site created successfully."
                    $global:nextPhase++
                    Write-Debug "nextPhase set to $global:nextPhase"
                }
        }
        else 
            {
                logWrite 22 $True "Skipped Retention Policy."
                $global:nextPhase++
                Write-Debug "nextPhase set to $global:nextPhase" 
            }

}

# -------------------------------------------------------
# Retention Policy - Create Compliance Tag (Step 23)
# -------------------------------------------------------
Function RetentionPolicy_CreateComplianceTag([string]$global:retentionTagName, [string]$retentionTagComment, [bool]$global:isRecordLabel, [string]$retentionTagAction, [int]$retentionTagDuration, [string]$retentionTagType)
{
    # Tag Variables
    $global:retentionTagName = "WKS-Compliance-Tag5"
    $retentionTagComment = "Keep and delete tag - 3 Days"
    $retentionTagAction = "KeepAndDelete"
    $retentionTagDuration = 3
    $retentionTagType = "ModificationAgeInDays"
    $global:isRecordLabel = $false
    
    if ($SkipRetentionPolicies -eq $false)
        {
            try {
                    write-Debug "New-ComplianceTag -Name $global:retentionTagName -Comment $retentionTagComment -IsRecordLabel $global:isRecordLabel -RetentionAction $retentionTagAction -RetentionDuration $retentionTagDuration -RetentionType $retentionTagType -ErrorAction Stop | Out-Null"
                    $complianceTagStatus = New-ComplianceTag -Name $global:retentionTagName -Comment $retentionTagComment -IsRecordLabel $global:isRecordLabel -RetentionAction $retentionTagAction -RetentionDuration $retentionTagDuration -RetentionType $retentionTagType -ErrorAction Stop | Out-Null
                }
                catch 
                    {
                        write-Debug $Error[0].Exception
                        logWrite 23 $false "Unable to create Retention Tag $global:retentionTagName"
                        exitScript
                    }

        if($global:recovery -eq $false)
                {
                    logWrite 23 $True "Retention Tag $global:retentionTagName created successfully."
                    $global:nextPhase++
                    Write-Debug "nextPhase set to $global:nextPhase" 
                    }
        }
        else 
            {
                logWrite 23 $True "Skipped Retention Policy."
                $global:nextPhase++
                Write-Debug "nextPhase set to $global:nextPhase"     
            }
}

# -------------------------------------------------------
# Retention Policy - Create Retention Policy (Step 24)
# -------------------------------------------------------
function RetentionPolicy_NewRetentionPolicy([string]$retentionPolicyName, [string]$tenantName, [string]$global:siteName, [string]$global:retentionTagName)
{
    # Policy Variables
    $retentionPolicyName = "WKS-Compliance-policy"
    
    if ($SkipRetentionPolicies -eq $false)
        {
            $url = "https://$tenantName.sharepoint.com/sites/$global:siteName"
            #try to create policy first
            Try
                {
                    #Create compliance retention Policy
                    write-Debug "New-RetentionCompliancePolicy -Name $retentionPolicyName -SharePointLocation $url -Enabled $true -ExchangeLocation All -ModernGroupLocation All -OneDriveLocation All -ErrorAction Stop | Out-Null"
                    $policyStatus = New-RetentionCompliancePolicy -Name $retentionPolicyName -SharePointLocation $url -Enabled $true -ExchangeLocation All -ModernGroupLocation All -OneDriveLocation All -ErrorAction Stop | Out-Null
                } 
                catch 
                    {
                        #failed to create policy
                        write-Debug $Error[0].Exception
                        logWrite 24 $false "Unable to create the Retention Policy $retentionPolicyName"
                        exitScript
                    }
            
            #then, if successfull, create rule in policy
            try 
                {
                    write-Debug "New-RetentionComplianceRule -Policy $retentionPolicyName -publishComplianceTag $global:retentionTagName -ErrorAction Stop | Out-Null"
                    $policyRuleStatus = New-RetentionComplianceRule -Policy $retentionPolicyName -publishComplianceTag $global:retentionTagName -ErrorAction Stop | Out-Null
                    #sleep for 240 seconds
                    #goToSleep 240
                }
            catch 
                {
                    #failed to create policy
                    write-Debug $Error[0].Exception
                    logWrite 24 $false "Unable to create the Retention Policy Rule."
                    exitScript
                }
            
            if($global:recovery -eq $false)
                {
                    logWrite 24 $True "Retention Policy $retentionPolicyName and Rule created successfully."
                    $global:nextPhase++ #25
                    $global:nextPhase++ #26
                    $global:nextPhase++ #27
                    $global:nextPhase++ #28
                    $global:nextPhase++ #29
                    $global:nextPhase++ #30
                    $global:nextPhase++ #31
                    Write-Debug "nextPhase set to $global:nextPhase" 
                }
        }
        else 
            {
                logWrite 24 $True "Skipped Retention Policy."
                $global:nextPhase++ #25
                $global:nextPhase++ #26
                $global:nextPhase++ #27
                $global:nextPhase++ #28
                $global:nextPhase++ #29
                $global:nextPhase++ #30
                $global:nextPhase++ #31
                Write-Debug "nextPhase set to $global:nextPhase"    
            }
}

#######################################################################################
#########                         D     L     P                              ##########
#######################################################################################

# -------------------------------------------------------
# DLP - Create DLP Policy (Step 31)
# -------------------------------------------------------
function DLP_CreateDLPCompliancePolicy
{
    if ($SkipDLP -eq $false)
        {        
            try
                {
                    if (Get-DlpCompliancePolicy -Identity "WKS Compliance Policy")
                        {
                            write-host " The DLP Compliance Policy already Exists "
                        }
                    else
                        {
                            $params = @{
                                "Name" = "WKS Compliance Policy";
                                "ExchangeLocation" ="All";
                                "OneDriveLocation" = "All";
                                "SharePointLocation" = "All";
                                "EndpointDlpLocation" = "all";
                                "Teamslocation" = "All";
                                "Mode" = "Enable"
                                }
                            New-dlpcompliancepolicy @params
                        }
                }
                catch 
                    {
                        write-Debug $Error[0].Exception
                        logWrite 31 $false "Unable to create DLP Policy."
                        exitScript
                    }
            if($global:recovery -eq $false)
                {
                    logWrite 31 $True "Able to Create DLP Policy."
                    $global:nextPhase++
                    Write-Debug "nextPhase set to $global:nextPhase" 
                }
        }
        else 
            {
                logWrite 31 $True "Skipped DLP."
                $global:nextPhase++
                Write-Debug "nextPhase set to $global:nextPhase"       
            }
}

# -------------------------------------------------------
# DLP - Create DLP Rule (Step 32)
# -------------------------------------------------------
function DLP_CreateDLPComplianceRule
{
    if ($SkipDLP -eq $false)
        {
            try
                {
                    $senstiveinfo = @(@{Name =”Credit Card Number”; minCount = “1”},@{Name =”International Banking Account Number (IBAN)”; minCount = “1”},@{Name =”U.S. Bank Account Number”; minCount = “1”})
                    $Rulevalue = @{
                        "Name" = "WKS-Copmpliance-Ruleset";
                        "Comment" = "Helps detect the presence of information commonly considered to be subject to the GLBA act in America. like driver's license and passport number.";
                        "Policy" = "WKS Compliance Policy";
                        "ContentContainsSensitiveInformation"=$senstiveinfo;
                        "AccessScope"= "NotInOrganization";
                        "Disabled" =$false;
                        'ReportSeverityLevel'='High'
                        }
                    New-DlpComplianceRule @rulevalue 
                }
                catch 
                    {
                        write-Debug $Error[0].Exception
                        logWrite 32 $false "Unable to create DLP Rule."
                        exitScript
                    }
            if($global:recovery -eq $false)
                {
                    logWrite 32 $True "Able to Create DLP Rule."
                    $global:nextPhase++ #33
                    $global:nextPhase++ #34
                    $global:nextPhase++ #35
                    $global:nextPhase++ #36
                    $global:nextPhase++ #37
                    $global:nextPhase++ #38
                    $global:nextPhase++ #39
                    $global:nextPhase++ #40
                    $global:nextPhase++ #41
                    Write-Debug "nextPhase set to $global:nextPhase" 
                }
        }
        else 
            {
                logWrite 32 $True "Skipped DLP."
                $global:nextPhase++ #33
                $global:nextPhase++ #34
                $global:nextPhase++ #35
                $global:nextPhase++ #36
                $global:nextPhase++ #37
                $global:nextPhase++ #38
                $global:nextPhase++ #39
                $global:nextPhase++ #40
                $global:nextPhase++ #41
                Write-Debug "nextPhase set to $global:nextPhase"   
            }    
}

#######################################################################################
#########                    I N S I D E R     R I S K S                     ##########
#######################################################################################

# -------------------------------------------------------
# InsiderRisks - Create an Azure App (Step 41)
# -------------------------------------------------------
function InsiderRisks_CreateAzureApp
{
    try
        {
            $AzureADAppReg = New-AzureADApplication -DisplayName HRConnector -AvailableToOtherTenants $false -ErrorAction Stop
            $appname = $AzureADAppReg.DisplayName
            $global:appid = $AzureADAppReg.AppID
            $AzureTenantID = Get-AzureADTenantDetail
            $global:tenantid = $AzureTenantID.ObjectId
            $AzureSecret = New-AzureADApplicationPasswordCredential -CustomKeyIdentifier PrimarySecret -ObjectId $azureADAppReg.ObjectId -EndDate ((Get-Date).AddMonths(6)) -ErrorAction Stop
            $global:Secret = $AzureSecret.value

            write-host "##################################################################" -ForegroundColor Green
            write-host "##                                                              ##" -ForegroundColor Green
            write-host "##   Microsoft 365 Security and Compliance: Compliance Center   ##" -ForegroundColor Green
            write-host "##                                                              ##" -ForegroundColor Green
            write-host "##   App name  : $appname                                    ##" -ForegroundColor Green
            write-host "##   App ID    : $global:appid           ##" -ForegroundColor Green
            write-host "##   Tenant ID : $global:tenantid           ##" -ForegroundColor Green
            write-host "##   App Secret: $global:secret   ##" -ForegroundColor Green
            write-host "##                                                              ##" -ForegroundColor Green
            write-host "##################################################################" -ForegroundColor Green
            write-host
            Write-host "Return to the lab instructions" -ForegroundColor Yellow
            Write-host "When requested, press ENTER to continue." -ForegroundColor Yellow
            write-host
        }
        catch 
        {
            write-Debug $error[0].Exception
            logWrite 41 $false "Error creating the Azure App for HR Connector"
            exitScript
        }
    if($global:recovery -eq $false)
        {
            logWrite 41 $True "Successfully created the Azure App for HR Connector."
            $global:nextPhase++
            Write-Debug "nextPhase set to $global:nextPhase"
        }
}

# -------------------------------------------------------
# InsiderRisks - Create the CSV file (Step 42)
# -------------------------------------------------------
function InsiderRisks_CreateCSVFile
{
    $CurrentPath = Get-Location
    write-host "##################################################################" -ForegroundColor Green
    write-host "##                                                              ##" -ForegroundColor Green
    write-host "##   Microsoft 365 Security and Compliance: Compliance Center   ##" -ForegroundColor Green
    write-host "##                                                              ##" -ForegroundColor Green
    write-host "##   The CSV file was created on $CurrentPath\wks-new-HRConnector.csv" -ForegroundColor Green
    write-host "##                                                              ##" -ForegroundColor Green
    write-host "##################################################################" -ForegroundColor Green
    write-host
    Write-host "Return to the lab instructions" -ForegroundColor Yellow
    Write-host "When requested, press ENTER to continue." -ForegroundColor Yellow
    write-host

    try 
        {
            $global:HRConnectorCSVFile = "$($LogPath)HRConnector.csv"
            "HRScenarios,EmailAddress,ResignationDate,LastWorkingDate,EffectiveDate,YearsOnLevel,OldLevel,NewLevel,PerformanceRemarks,PerformanceRating,ImprovementRemarks,ImprovementRating" | out-file $HRConnectorCSVFile -Encoding utf8
            $Users = Get-AzureADuser | where-object {$null -ne $_.AssignedLicenses} | Select-Object UserPrincipalName -ErrorAction Stop

            foreach ($User in $Users)
                {
                    $EmailAddress = $User.UserPrincipalName
                    #Resignation block
                    $RandResignationDate  = Get-Random -Minimum 20 -Maximum 30
                    $ResignationDate = (Get-Date).AddDays(-$RandResignationDate).ToString("yyyy-MM-ddTH:mm:ssZ")
                    $RandLastWorkingDate = Get-Random -Minimum 10 -Maximum 20
                    $LastWorkingDate = (Get-Date).AddDays(-$RandLastWorkingDate).ToString("yyyy-MM-ddTH:mm:ssZ")
                    $RandEffectiveDate = Get-Random -Minimum 365 -Maximum 1000
                    $EffectiveDate = (Get-Date).AddDays(-$RandEffectiveDate).ToString("yyyy-MM-ddTH:mm:ssZ")
                    #Employee level block
                    $YearsOnLevel = Get-Random -Minimum 1 -Maximum 6
                    $OldLevel = Get-Random -Minimum 57 -Maximum 64
                    $NewLevel = $OldLevel--
                    #performance and performance review block
                    $RandRating = Get-Random -Minimum 1 -Maximum 4
                    Switch ($RandRating) 
                        {
                            1 
                                {
                                    $PerformanceRemarks = "Achieved all commitments and exceptional results that surpassed expectations"
                                    $PerformanceRating = "1 - Exceeded"
                                    $ImprovementRemarks = $null
                                    $ImprovementRating = $null
                                }
                            2 
                                {
                                    $PerformanceRemarks = "Achieved all commitments and expected results"
                                    $PerformanceRating = "2 - Achieved"
                                    $ImprovementRemarks = "Increase the team collaboration"
                                    $ImprovementRating = "1 - Exceeded"
                                }
                            3
                                {
                                    $PerformanceRemarks = "Failed to achieve commitments or expected results or both"
                                    $PerformanceRating = "3 - Underperformed"
                                    $ImprovementRemarks = "Increase overall performance"
                                    $ImprovementRating = "2 - Achieved"
                                }
                        }
                    "Resignation,$EmailAddress,$ResignationDate,$LastWorkingDate," | out-file $HRConnectorCSVFile -Encoding utf8 -Append -ErrorAction Stop
                    "Job level change,$EmailAddress,,,$EffectiveDate,$YearsOnLevel,Level $OldLevel,Level $NewLevel" | out-file $HRConnectorCSVFile -Encoding utf8 -Append -ErrorAction Stop
                    "Performance review,$EmailAddress,,,$EffectiveDate,,,,$PerformanceRemarks,$PerformanceRating" | out-file $HRConnectorCSVFile -Encoding utf8 -Append -ErrorAction Stop
                    "Performance improvement plan,$EmailAddress,,,$EffectiveDate,,,,,,$ImprovementRemarks,$ImprovementRating,"  | out-file $HRConnectorCSVFile -Encoding utf8 -Append -ErrorAction Stop
                }
        }
        catch 
        {
            write-Debug $error[0].Exception
            logWrite 42 $false "Error creating the HRConnector.csv file"
            exitScript
        }
    if($global:recovery -eq $false)
        {
            logWrite 42 $True "Successfully created the HRConnector.csv file."
            $global:nextPhase++
            Write-Debug "nextPhase set to $global:nextPhase"
        }
}

# -------------------------------------------------------
# InsiderRisks - Upload CSV file (Step 43)
# -------------------------------------------------------
function InsiderRisks_UploadCSV
{

    try   
        {
            $ConnectorJobID = Read-Host "Paste the Connector job ID"
            if ($null -eq $ConnectorJobID)
                {
                    $ConnectorJobID = Read-Host "Paste the Connector job ID"
                }
            Write-Host
            write-host "##################################################################" -ForegroundColor Green
            write-host "##                                                              ##" -ForegroundColor Green
            write-host "##   Microsoft 365 Security and Compliance: Compliance Center   ##" -ForegroundColor Green
            write-host "##                                                              ##" -ForegroundColor Green
            write-host "##   App ID    : $global:appid           ##" -ForegroundColor Green
            write-host "##   Tenant ID : $global:tenantid           ##" -ForegroundColor Green
            write-host "##   App Secret: $global:secret   ##" -ForegroundColor Green
            write-host "##   JobId     : $ConnectorJobID           ##" -ForegroundColor Green
            write-host "##   CSV File  : $global:HRConnectorCSVFile           " -ForegroundColor Green
            write-host "##                                                              ##" -ForegroundColor Green
            write-host "##################################################################" -ForegroundColor Green
            Write-Host

            Set-Location -Path "$env:UserProfile\Desktop\SCLabFiles\Scripts"
            .\upload_termination_records.ps1 -tenantId $tenantId -appId $appId -appSecret $Secret -jobId $ConnectorJobID -csvFilePath $HRConnectorCSVFile
        }
        catch 
        {
            write-Debug $error[0].Exception
            logWrite 43 $false "Error uploading the HRConnector.csv file"
            exitScript
        }
    if($global:recovery -eq $false)
        {
            logWrite 43 $True "Successfully creating the HRConnector.csv file."
            $global:nextPhase++
            Write-Debug "nextPhase set to $global:nextPhase"
        }
}

# -------------------------------------------------------
# Exit function
# -------------------------------------------------------
function exitScript
{
    # Get-PSSession | Remove-PSSession
    if ($debug)
        {
            $DebugPreference = $oldDebugPreference
            Stop-Transcript
        }
    exit
}

# -------------------------------------------------------
# FUNCTION - Start-SnCCompliance
# -------------------------------------------------------

# -------------------------------------------------------
# Variable definition - General
# -------------------------------------------------------
$LogPath = "$env:UserProfile\Desktop\SCLabFiles\Scripts\"
$LogCSV = "$env:UserProfile\Desktop\SCLabFiles\Scripts\Progress_Download_Log.csv"
$global:nextPhase = 1
$global:recovery = $false

# -----------------------------------------------------------
# Debug mode
# -----------------------------------------------------------
$oldDebugPreference = $DebugPreference
if($debug)
{
    write-debug "Debug Enabled"
    $DebugPreference = "Continue"
    Start-Transcript -Path "$($LogPath)download-debug.txt"
}


if(!(Test-Path($logCSV)))
    {
        # if log doesn't exist then must be first time we run this, so go to initialization
        Write-Debug "Entering Initialization"
        initialization
    } 
        else 
            {
                # if log already exists, check if we need to recover
                Write-Debug "Entering Recovery"
                recovery
                ConnectAzureAD
                ConnectMSOL
                ConnectEXO
                ConnectSCC
                ConnectTeams
                $tenantName = GetDomain
                Write-Debug "$tenantName Returned"
                ConnectSPO $tenantName
            }

# -------------------------------------------------------
# use variable to control phases
# -------------------------------------------------------
if($nextPhase -eq 1)
    {
        write-debug "Phase $nextPhase"
        ConnectAzureAD
    }

if($nextPhase -eq 2)
    {
        write-debug "Phase $nextPhase"
        ConnectMSOL
    }

if($nextPhase -eq 3)
    {
        write-debug "Phase $nextPhase"
        ConnectEXO
    }

if($nextPhase -eq 4)
    {
        write-debug "Phase $nextPhase"
        ConnectSCC
    }

if($nextPhase -eq 5)
    {
        write-debug "Phase $nextPhase"
        ConnectTeams
    }

if($nextPhase -eq 6)
    {
        write-debug "Phase $nextPhase"
        $tenantName = getdomain
        write-debug "$tenantName Returned"
    }

if($nextPhase -eq 7)
    {
        write-debug "Phase $nextPhase"
        ConnectSPO $tenantName
    }

if($nextPhase -eq 8)
    {
        write-debug "Phase $nextPhase"
        ConnectPNP $tenantName
    }

if($nextPhase -eq 9)
    {
        write-debug "Phase $nextPhase"
        downloadscripts
    }

if($nextPhase -eq 11)
    {
        write-debug "Phase $nextPhase"
        SensitivityLabel_Label
    }

if($nextPhase -eq 12)
    {
        write-debug "Phase $nextPhase"
        SensitivityLabel_Policy
    }

if($nextPhase -eq 21)
    {
        write-debug "Phase $nextPhase"
        RetentionPolicy_GetSiteOwner
    }

if($nextPhase -eq 22)
    {
        write-debug "Phase $nextPhase"
        RetentionPolicy_CreateSPOSite $tenantName $global:siteName $global:siteOwner $siteStorageQuota $siteResourceQuota $siteTemplate
    }

if($nextPhase -eq 23)
    {
        write-debug "Phase $nextPhase"
        RetentionPolicy_CreateComplianceTag $global:retentionTagName $retentionTagComment $global:isRecordLabel $retentionTagAction $retentionTagDuration $retentionTagType
    }

if($nextPhase -eq 24)
    {
        write-debug "Phase $nextPhase"
        RetentionPolicy_NewRetentionPolicy $retentionPolicyName $tenantName $global:siteName $global:retentionTagName
    }

if($nextPhase -eq 31)
    {
        write-debug "Phase $nextPhase"
        DLP_CreateDLPCompliancePolicy
    }

if($nextPhase -eq 32)
    {
        write-debug "Phase $nextPhase"
        DLP_CreateDLPComplianceRule
    }

if($nextPhase -eq 41)
    {
        write-debug "Phase $nextPhase"
        InsiderRisks_CreateAzureApp
        $answer = Read-Host "Press ENTER to continue"
    }

if($nextPhase -eq 42)
    {
        write-debug "Phase $nextPhase"
        InsiderRisks_CreateCSVFile
        $answer = Read-Host "Press ENTER to continue"
    }

if($nextPhase -eq 43)
    {
        write-debug "Phase $nextPhase"
        InsiderRisks_UploadCSV
    }


#if ($nextPhase -ge 9)
#    {
#        write-debug "Phase $nextPhase"
#        Set-Location -Path $LogPath
#        $nextScript = "wks-new-label.ps1"
#        logWrite 9 $true "Launching $nextScript script"
#        if ($debug)
#            {
#                Stop-Transcript
#                .\wks-new-label.ps1 -$debug
#            } 
#            else 
#                {
#                    .\wks-new-label.ps1
#                }
#    }
 

function New-WksComp-InsiderRisks
{
    ConnectAzureAD  #Call the function
    Write-Host
    InsiderRisks_CreateAzureAp #Call the function
    Write-Host
    $answer = Read-Host "Press ENTER to continue"
    Write-Host
    InsiderRisks_CreateCSVFile #call the function
    Write-Host
    $answer = Read-Host "Press ENTER to continue"
    Write-Host
    InsiderRisks_UploadCSV #call the function
    Write-Host
    $answer = Read-Host "Press ENTER to continue"
    Write-Host
}