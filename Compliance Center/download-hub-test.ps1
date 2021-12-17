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
            logWrite 1 $true "Successfully connected to Microsoft Online"
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
# Get Tenant Name (Step 5)
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
                logWrite 5 $false "Unable to fetch Tenant name."
                exitScript
            }
    Write-Debug "Initial domain: $InitialDomain"
    if($global:recovery -eq $false)
        {
            logWrite 5 $True "Successfully got Tenant Name."
            $global:nextPhase++
            Write-Debug "nextPhase set to $global:nextPhase"
        }
    return $InitialDomain.name.split(".")[0]
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
# Connect to PNP Online (Step 7)
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
                                    logWrite 7 $false "Couldn't connect to PNP Online. Exiting."
                                    exitScript
                                }
                   
                    }
        }
        if($global:recovery -eq $false)
            {
                logWrite 7 $true "Successfully connected to PNP Online"
                $global:nextPhase++
                Write-Debug "nextPhase set to $global:nextPhase"
            }
}


function downloadscripts
{
    try
        {
            #Labels scritp
            Write-Debug "Invoke-WebRequest -Uri https://raw.githubusercontent.com/microsoft/WKSPlus-M365SCC/main/Compliance%20Center/wks-new-Adaptive-scope.ps1 -OutFile $($LogPath)wks-new-Adaptive-scope.ps1 -ErrorAction Stop"
            Invoke-WebRequest -Uri https://raw.githubusercontent.com/microsoft/WKSPlus-M365SCC/main/Compliance%20Center/wks-new-Adaptive-scope.ps1 -OutFile "$($LogPath)wks-new-Adaptive-scope.ps1" -ErrorAction Stop
            #Adaptive Scopes
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
            Write-Debug "nextPhase set to $global:nextPhase"
        }
} 

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
        $tenantName = getdomain
        write-debug "$tenantName Returned"
    }

if($nextPhase -eq 6)
    {
        write-debug "Phase $nextPhase"
        ConnectSPO $tenantName
    }

if($nextPhase -eq 7)
    {
        write-debug "Phase $nextPhase"
        ConnectPNP $tenantName
    }

if($nextPhase -eq 8)
    {
        write-debug "Phase $nextPhase"
        downloadscripts
    }
    
if ($nextPhase -ge 8){
        ./wks-new-dlp.ps1
    }