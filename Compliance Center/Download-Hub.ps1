<#
 .Synopsis
  Displays a visual representation of a calendar.

 .Description
  Displays a visual representation of a calendar. This function supports multiple months
  and lets you highlight specific date ranges or days.

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

Param (
    [switch]$debug
)

# -----------------------------------------------------------
# Variable definition
# -----------------------------------------------------------
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
            $global:nextPhase++
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
                    Connect-IPPSSession -ErrorAction stop | Out-Null
                }
                catch    
                    {
                        try
                            {
                                write-Debug $error[0].Exception
                                Write-Host "Installing Compliance Center PowerShell Module..."
                                #Install-Module ExchangeOnlineManagement -Force -AllowClobber #Not required, but it was already installed on the previous step
                                Connect-IPPSSession -ErrorAction stop | Out-Null
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
# Connect to SharePoint Online (Step 6)
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
                    Connect-SPOService -Url $AdminURL -ErrorAction stop | Out-Null
                }
                catch    
                    {
                        try
                            {
                                write-Debug $error[0].Exception
                                Write-Host "Installing SharePoint Online PowerShell Module..."
                                Install-Module Microsoft.Online.SharePoint.PowerShell -Force -AllowClobber -ErrorAction stop | Out-Null
                                Connect-SPOService -Url $AdminURL -ErrorAction stop | Out-Null
                            }
                            catch
                                {
                                    write-Debug $error[0].Exception
                                    logWrite 6 $false "Couldn't connect to SharePoint Online. Exiting."
                                    exitScript
                                }
                   
                    }
        }
        if($global:recovery -eq $false)
            {
                logWrite 6 $true "Successfully connected to SharePoint Online"
                $global:nextPhase++
                Write-Debug "nextPhase set to $global:nextPhase"
            }
}

# -------------------------------------------------------
# Get EXO Accepted Domains (Step 7)
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
                logWrite 7 $false "Unable to fetch all accepted Domains."
                exitScript
            }
    Write-Debug "Initial domain: $InitialDomain"
    if($global:recovery -eq $false)
        {
            logWrite 7 $True "Successfully got the accepted domains."
            $global:nextPhase++
            Write-Debug "nextPhase set to $global:nextPhase"
        }
    return $InitialDomain.name.split(".")[0]
}

# -------------------------------------------------------
# Download Workshop Script (Step 8)
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
                logWrite 8 $false "Unable to download the workshop scripts from GitHub! Exiting."
                exitScript
            }
    if($global:recovery -eq $false)
        {
            logWrite 8 $True "Successfully downloaded the workshop scripts."
            $global:nextPhase++ #9
            $global:nextPhase++ #10
            $global:nextPhase++ #11
            Write-Debug "nextPhase set to $global:nextPhase"
        }
}       

# -------------------------------------------------------
# Create Sensitivity label (Step 11)
# -------------------------------------------------------
function SensitivityLabel
{
    <#
    TO DO:
    Need to check to see if label exists in case the failure occured after cmd was successful, such as if they close the PS window. Maybe just check if label exists, and use Set-Label if so.
    #>
    $domainName = (Get-AcceptedDomain | Where-Object{$_.Default -eq $true}).DomainName
    $Encpermission = $domainname + ":VIEW,VIEWRIGHTSDATA,DOCEDIT,EDIT,PRINT,EXTRACT,REPLY,REPLYALL,FORWARD,OBJMODEL"
    try 
        {
            $labelStatus = New-Label -DisplayName $labelDisplayName -Name $labelName -ToolTip $labelTooltip -Comment $labelComment -ContentType "file","Email","Site","UnifiedGroup" -EncryptionEnabled:$true -SiteAndGroupProtectionEnabled:$true -EncryptionPromptUser:$true -EncryptionRightsDefinitions $Encpermission -SiteAndGroupProtectionPrivacy "private" -EncryptionDoNotForward:$true -SiteAndGroupProtectionAllowLimitedAccess:$true -ErrorAction stop | Out-Null
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

# -------------------------------------------------------
# Create Sensitivity policy (Step 12)
# -------------------------------------------------------
function SensitivityPolicy
{
    <#
    TO DO:
    - Need to check to see if label policy exists in case the failure occured after cmd was successful, such as if they close the PS window. Maybe just check if label exists, and use Set-Label if so.
    - Need to make sure the labele exists
    #>

    try 
        {
            New-LabelPolicy -name $labelPolicyName -Settings @{mandatory=$false} -AdvancedSettings @{requiredowngradejustification= $true} -Labels $labelName -ErrorAction stop | Out-Null
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
# Script start here
# -------------------------------------------------------
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
        ConnectSPO $tenantName
    }

if($nextPhase -eq 7)
    {
        write-debug "Phase $nextPhase"
        $tenantName = getdomain
        write-debug "$tenantName Returned"
    }

if($nextPhase -eq 8)
    {
        write-debug "Phase $nextPhase"
        downloadscripts
    }

if($nextPhase -eq 11)
    {
        write-debug "Phase $nextPhase"
        SensitivityLabel
    }

if($nextPhase -eq 12)
    {
        write-debug "Phase $nextPhase"
        SensitivityPolicy
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
