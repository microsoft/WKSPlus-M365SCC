<#
.Synopsis
  Pre-configuration for the following Microsoft Premier offerings:
    1) Activate Microsoft 365 Security and Compliance: Purview Manage Insider Risks
    2) WorkshopPLUS: Microsoft 365 Security and Compliance - Microsoft Purview

.Description
  Prepare the required configuration for some Microsoft Unified support titles.

.Description
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

.Version
    4.0 (August 21st, 2023)
    Owners: 
        Marcelo Hunecke <mhunecke@microsoft.com>
        Eli Yang <Eli.Yang@microsoft.com>
        Ashley Wills <Ashley.Wills@microsoft.com>
#>

Param (
    [CmdletBinding()]
    [switch]$debug
)

#---------------------------------------------------------------------
# Write the log
#---------------------------------------------------------------------
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

#---------------------------------------------------------------------
# Start the Recovery steps
#---------------------------------------------------------------------
function Recovery
{
    Write-host "Starting Recovery..."
    Set-Location -Path $LogPath
    $global:Recovery = $true
    $savedLog = Import-Csv $LogCSV
    $lastEntry = (($savedLog.Count) - 1)
    Write-Debug "Last Entry #: $lastEntry"
    $lastEntry2 = (($savedLog.Count) - 2)
    Write-Debug "Entry Before Last: $lastEntry2"
    $lastEntryPhase = [int]$savedLog[$lastEntry].Phase
    #Always need to restart from phase 5 to get the app infos.
    $lastEntryPhase = 2
    #Always need to restart from phase 5 to get the app infos.
    Write-Debug "Last Phase: $lastEntryPhase"
    $lastEntryResult = $savedLog[$lastEntry].Result
    Write-Debug "Last Entry Result: $lastEntryResult"

    if ($lastEntryResult -eq $false)
        {
            if ($lastEntryPhase -eq $savedLog[$lastEntry2].Phase)
                {
                    WriteHost -ForegroundColor Red "The script has failed at Phase $lastEntryPhase repeatedly.  Please check with your instructor."
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

#---------------------------------------------------------------------
# Exit function
#---------------------------------------------------------------------
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

#######################################################################################
#########                   I N I T I A L I Z A T I O N                      ##########
#######################################################################################

#---------------------------------------------------------------------
# Test the log path (Step 0)
#---------------------------------------------------------------------
function Initialization
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

#---------------------------------------------------------------------
# Connect to AzureAD (Step 1)
#---------------------------------------------------------------------
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
                        Connect-AzureAD -WarningAction SilentlyContinue -ErrorAction stop | Out-Null
                    }
                    catch    
                        {
                            try
                                {
                                    write-Debug $error[0].Exception
                                    Write-Host "Installing Azure AD PowerShell Module..."
                                    Install-Module AzureAD -Force -AllowClobber
                                    Connect-AzureAD -WarningAction SilentlyContinue -ErrorAction stop | Out-Null
                                }
                                catch
                                    {
                                        write-Debug $error[0].Exception
                                        logWrite 1 $false "Couldn't connect to Azure AD. Exiting."
                                        exitScript
                                    }
                       
                        }
            }
    if($global:Recovery -eq $false)
    {
        logWrite 1 $true "Successfully connected to Azure AD."
        $global:nextPhase++
        Write-Debug "nextPhase set to $global:nextPhase"
    }
}

#---------------------------------------------------------------------
# Connect to Microsoft Online (Step 2)
#---------------------------------------------------------------------
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
    if($global:Recovery -eq $false)
        {
            logWrite 2 $true "Successfully connected to Microsoft Online."
            $global:nextPhase++
            Write-Debug "nextPhase set to $global:nextPhase"
        }
}

#######################################################################################
#########             I N S I D E R     R I S K S  - General                 ##########
#######################################################################################

#---------------------------------------------------------------------
# Insider Risks - Download scripts for Connectors (Step 3)
#---------------------------------------------------------------------
function DownloadScripts
{
    try
        {
            #Get the public script for HR Connector from GitHub
            write-Debug "Invoke-WebRequest -Uri https://raw.githubusercontent.com/microsoft/m365-compliance-connector-sample-scripts/master/sample_script.ps1 -OutFile $($LogPath)upload_termination_records.ps1 -ErrorAction Stop"
            Invoke-WebRequest -Uri https://raw.githubusercontent.com/microsoft/m365-compliance-connector-sample-scripts/master/sample_script.ps1 -OutFile "$($LogPath)upload_termination_records.ps1" -ErrorAction Stop
            #Get the public script for Physical Badging Connector from GitHub
            write-Debug "Invoke-WebRequest -Uri https://raw.githubusercontent.com/microsoft/m365-physical-badging-connector-sample-scripts/master/push_physical_badging_records.ps1 -OutFile $($LogPath)upload_badging_records.ps1 -ErrorAction Stop"
            Invoke-WebRequest -Uri https://raw.githubusercontent.com/microsoft/m365-physical-badging-connector-sample-scripts/master/push_physical_badging_records.ps1 -OutFile "$($LogPath)upload_badging_records.ps1" -ErrorAction Stop
            $global:Recovery = $false #There no Recover process from here. All the steps below (3, 4, and 5) will be executed.
        } 
            catch 
                {
                    write-Debug $error[0].Exception
                    logWrite 3 $false "Unable to download the script from GitHub! Exiting."
                    exitScript
                }
    if($global:Recovery -eq $false)
        {
            logWrite 3 $True "Successfully downloaded the script from GitHub."
            $global:nextPhase++
            Write-Debug "nextPhase set to $global:nextPhase"
        }
}       

########################################################################################
#########           I N S I D E R     R I S K S  -  HR Connector              ##########
########################################################################################
#---------------------------------------------------------------------
# InsiderRisks - Create the CSV file for HR Connector (Step 4)
#---------------------------------------------------------------------
function InsiderRisks_CreateCSVFile_HRConnector
{
    try 
        {
            $global:HRConnectorCSVFile = "$($LogPath)HRConnectorData.csv"
            "HRScenarios,EmailAddress,ResignationDate,LastWorkingDate" | out-file $HRConnectorCSVFile -Encoding utf8
            $Users = Get-AzureADuser | where-object {$null -ne $_.AssignedLicenses} | Select-Object UserPrincipalName -ErrorAction Stop
            foreach ($User in $Users)
                {
                    $EmailAddress = $User.UserPrincipalName
                    $RandResignationDate  = Get-Random -Minimum 20 -Maximum 30
                    $ResignationDate = (Get-Date).AddDays(-$RandResignationDate).ToString("yyyy-MM-ddTHH:mm:ssZ")
                    $RandLastWorkingDate = Get-Random -Minimum 10 -Maximum 20
                    $LastWorkingDate = (Get-Date).AddDays(-$RandLastWorkingDate).ToString("yyyy-MM-ddTHH:mm:ssZ")
                    "Resignation,$EmailAddress,$ResignationDate,$LastWorkingDate" | out-file $HRConnectorCSVFile -Encoding utf8 -Append -ErrorAction Stop
                }
        }
        catch 
            {
                write-Debug $error[0].Exception
                logWrite 4 $false "Error creating the HRConnectorData.csv file."
                exitScript
            }
    if($global:Recovery -eq $false)
        {
            logWrite 4 $True "Successfully created the HRConnectorData.csv file."
            $global:nextPhase++
            Write-Debug "nextPhase set to $global:nextPhase"
        }
}

#---------------------------------------------------------------------
# InsiderRisks - Create an Azure App for HR Connector (Step 5)
#---------------------------------------------------------------------
function InsiderRisks_CreateAzureApp_HRConnector
{
    try
        {
            $HRapp_appsecret = "$($LogPath)_HRapp_appsecret.txt"
            $BadgingApp_Name = "HRConnector01"
            $appExists = $null
            $appExists = Get-AzureADApplication -SearchString $BadgingApp_Name
            $AzureTenantID = Get-AzureADTenantDetail
            $global:tenantid = $AzureTenantID.ObjectId
            if ($null -eq $appExists)
                {
                    $AzureADAppReg = New-AzureADApplication -DisplayName $BadgingApp_Name -AvailableToOtherTenants $false -ErrorAction Stop
                    $appname = $AzureADAppReg.DisplayName
                    $global:appid = $AzureADAppReg.AppID
                    $AzureSecret = New-AzureADApplicationPasswordCredential -CustomKeyIdentifier PrimarySecret -ObjectId $azureADAppReg.ObjectId -EndDate ((Get-Date).AddMonths(6)) -ErrorAction Stop
                    $global:Secret = $AzureSecret.value
                    "Secret" | out-file $HRapp_appsecret -Encoding utf8 -ErrorAction Stop
                    $global:Secret | out-file $HRapp_appsecret -Encoding utf8 -Append -ErrorAction Stop
                    write-host
                    write-host "##########################################################################################" -ForegroundColor Green
                    write-host "##                                                                                      ##" -ForegroundColor Green
                    write-host "##     WorkshopPLUS: Microsoft 365 Security and Compliance - Microsoft Purview  and     ##" -ForegroundColor Green
                    write-host "##     Activate Microsoft 365 Security and Compliance: Purview Manage Insider Risks     ##" -ForegroundColor Green
                    write-host "##                                                                                      ##" -ForegroundColor Green            
                    write-host "##   App name  : $appname                                                          ##" -ForegroundColor Green
                    write-host "##   App ID    : $global:appid                                   ##" -ForegroundColor Green
                    write-host "##   Tenant ID : $global:tenantid                                   ##" -ForegroundColor Green
                    write-host "##   App Secret: $global:secret                           ##" -ForegroundColor Green
                    write-host "##                                                                                      ##" -ForegroundColor Green
                    write-host "##########################################################################################" -ForegroundColor Green
                    write-host
                    Write-host "Return to the lab instructions" -ForegroundColor Yellow
                    Write-host "When requested, press ENTER to continue." -ForegroundColor Yellow
                    write-host
                }
                else 
                    {
                        $appname = $appExists.DisplayName
                        $global:appid = $appExists.AppId
                        $SecretFileExists = Test-Path $HRapp_appsecret
                        if ($SecretFileExists)
                            {
                                $Secretfile = Import-Csv $HRapp_appsecret -Encoding utf8 -ErrorAction SilentlyContinue
                            }
                            else
                                {
                                    Remove-AzureADApplication -ObjectId $appExists.ObjectId
                                    lastEntryPhase = 2
                                    logWrite 5 $false "HR Azure App already exists, but the secret file was not found. Try again."
                                }
                        $global:Secret = $Secretfile.Secret
                        write-host
                        write-host "##########################################################################################" -ForegroundColor Green
                        write-host "##                                                                                      ##" -ForegroundColor Green
                        write-host "##     WorkshopPLUS: Microsoft 365 Security and Compliance - Microsoft Purview  and     ##" -ForegroundColor Green
                        write-host "##     Activate Microsoft 365 Security and Compliance: Purview Manage Insider Risks     ##" -ForegroundColor Green
                        write-host "##                                                                                      ##" -ForegroundColor Green            
                        write-host "##   App name  : $appname                                                          ##" -ForegroundColor Green
                        write-host "##   App ID    : $global:appid                                   ##" -ForegroundColor Green
                        write-host "##   Tenant ID : $global:tenantid                                   ##" -ForegroundColor Green
                        write-host "##   App Secret: $global:secret                           ##" -ForegroundColor Green
                        write-host "##                                                                                      ##" -ForegroundColor Green
                        write-host "##########################################################################################" -ForegroundColor Green
                        write-host
                        Write-host "Return to the lab instructions" -ForegroundColor Yellow
                        Write-host "When requested, press ENTER to continue." -ForegroundColor Yellow
                        write-host
                        logWrite 5 $True "Azure App for HR Connector already exists, so this step was skipped."
                        $global:HRapp_JustUploadCSV = $true #This variable will be used to skip the step 5 (Create app) if the app already exists.
                    }
        }
        catch 
            {
                write-Debug $error[0].Exception
                logWrite 5 $false "Error creating the Azure App for HR Connector. Try again."
                exitScript
            }
    if($global:Recovery -eq $false)
        {
            logWrite 5 $True "Successfully created the Azure App for HR Connector."
            $global:nextPhase++
            Write-Debug "nextPhase set to $global:nextPhase"
        }
}

#---------------------------------------------------------------------
# InsiderRisks - Upload CSV file for HR Connector (Step 6)
#---------------------------------------------------------------------
function InsiderRisks_UploadCSV_HRConnector
{
    try   
        {
            Write-Host
            $HRConnector_JobID = "$($LogPath)_HRConnector_jobID.txt"
            if ($global:HRapp_JustUploadCSV -eq $false)
                {
                    $ConnectorJobID = Read-Host "Paste the Connector job ID"
                    if ($null -eq $ConnectorJobID)
                        {
                            $ConnectorJobID = Read-Host "Paste the Connector job ID"
                        }
                    "JobID" | out-file $HRConnector_JobID -Encoding utf8 -ErrorAction Stop
                    $ConnectorJobID | out-file $HRConnector_JobID -Encoding utf8 -Append -ErrorAction Stop
                }
                else
                    {
                        $JobIDfile = Import-Csv $HRConnector_JobID -Encoding utf8 -ErrorAction SilentlyContinue
                        $ConnectorJobID = $JobIDfile.JobID
                    }
            Write-Host
            write-host "##########################################################################################" -ForegroundColor Green
            write-host "##                                                                                      ##" -ForegroundColor Green
            write-host "##     WorkshopPLUS: Microsoft 365 Security and Compliance - Microsoft Purview  and     ##" -ForegroundColor Green
            write-host "##     Activate Microsoft 365 Security and Compliance: Purview Manage Insider Risks     ##" -ForegroundColor Green
            write-host "##                                                                                      ##" -ForegroundColor Green            
            write-host "##   App ID    : $global:appid                                   ##" -ForegroundColor Green
            write-host "##   Tenant ID : $global:tenantid                                   ##" -ForegroundColor Green
            write-host "##   App Secret: $global:secret                           ##" -ForegroundColor Green
            write-host "##   JobId     : $ConnectorJobID                                   ##" -ForegroundColor Green
            write-host "##   CSV File  : $global:HRConnectorCSVFile          ##" -ForegroundColor Green
            write-host "##                                                                                      ##" -ForegroundColor Green
            write-host "##########################################################################################" -ForegroundColor Green
            Write-Host
            Set-Location -Path "$env:UserProfile\Desktop\SCLabFiles\Scripts"
            .\upload_termination_records.ps1 -tenantId $tenantId -appId $appId -appSecret $Secret -jobId $ConnectorJobID -FilePath $HRConnectorCSVFile
        }
        catch 
        {
            write-Debug $error[0].Exception
            logWrite 6 $false "Error uploading the HRConnectorData.csv file"
            exitScript
        }
    if($global:Recovery -eq $false)
        {
            logWrite 6 $True "Successfully uploading the HRConnectorData.csv file."
            $global:nextPhase++
            Write-Debug "nextPhase set to $global:nextPhase"
        }
}

#################################################################################################
#########         I N S I D E R     R I S K S  -  Physical Badging Connector          I##########
#################################################################################################
#---------------------------------------------------------------------
# InsiderRisks - Create the CSV file for Physical Badging Connector (Step 7)
#---------------------------------------------------------------------
function InsiderRisks_CreateCSVFile_BadgingConnector
{
    try 
        {
            $UsersProcessed = 0
            $global:BadgingConnectorCSVFile = "$($LogPath)BadgingConnectorData.csv"
            $Priority_physical_assets = "$($LogPath)Priority_physical_assets.csv"
            "[" | out-file $BadgingConnectorCSVFile -Encoding utf8
            $Users = Get-AzureADuser | where-object {$null -ne $_.AssignedLicenses} | Select-Object UserPrincipalName -ErrorAction Stop
            $UsersCount = $Users.Count
            foreach ($User in $Users)
                {
                    $UsersProcessed++
                    $EmailAddress = $User.UserPrincipalName
                    $RandOffice  = Get-Random -Minimum 0 -Maximum 9
                    Switch ($RandOffice) 
                    {
                        0 {$AssetID = "Tokyo_Main_01";$AssetName = "Tokyo Office - Main Building Door"}
                        1 {$AssetID = "Delhi_Main_01";$AssetName = "Delhi Office - Main Building Door"}
                        2 {$AssetID = "Shanghai_Main_01";$AssetName = "Shanghai Office - Main Building Door"}
                        3 {$AssetID = "SaoPaulo_Main_01";$AssetName = "Sao Paulo Office - Main Building Door"}
                        4 {$AssetID = "MexicoCity_Main_01";$AssetName = "Mexico City Office - Main Building Door"}
                        5 {$AssetID = "Dhaka_Main_01";$AssetName = "Dhaka Office - Main Building Door"}
                        6 {$AssetID = "Cairo_Main_01";$AssetName = "Cairo Office - Main Building Door"}
                        7 {$AssetID = "Beijing_Main_01";$AssetName = "Beijing Office - Main Building Door"}
                        8 {$AssetID = "London_Main_01";$AssetName = "London Office - Main Building Door"}
                        9 {$AssetID = "Seattle_Main_01";$AssetName = "Seattle Office - Main Building Door"}
                    }
                    $RandEventTime  = Get-Random -Minimum 1 -Maximum 31
                    $EventTime = (Get-Date).AddDays(-$RandEventTime).ToString("yyyy-MM-ddTHH:mm:ss")
                    $RandAccessStatus  = Get-Random -Minimum 0 -Maximum 2
                    if ($RandAccessStatus -eq 0)
                        {
                            $AccessStatus = "Sucess"
                        }
                        else
                            {
                                $AccessStatus = "Failed"
                            }
                    "   {" | out-file $BadgingConnectorCSVFile -Encoding utf8 -Append
                    "       " + [char]34 + "UserID" + [char]34 + ":" + [char]34 + $EmailAddress + [char]34 + "," | out-file $BadgingConnectorCSVFile -Encoding utf8 -Append
                    "       " + [char]34 + "AssetID" + [char]34 + ":" + [char]34 + $AssetID + [char]34 + "," | out-file $BadgingConnectorCSVFile -Encoding utf8 -Append
                    "       " + [char]34 + "AssetName" + [char]34 + ":" + [char]34 + $AssetName + [char]34 + "," | out-file $BadgingConnectorCSVFile -Encoding utf8 -Append
                    "       " + [char]34 + "EventTime" + [char]34 + ":" + [char]34 + $EventTime + [char]34 + "," | out-file $BadgingConnectorCSVFile -Encoding utf8 -Append
                    "       " + [char]34 + "AccessStatus" + [char]34 + ":" + [char]34 + $AccessStatus + [char]34 | out-file $BadgingConnectorCSVFile -Encoding utf8 -Append
                    if ($UsersProcessed -eq $UsersCount)
                        {
                            "   }" | out-file $BadgingConnectorCSVFile -Encoding utf8 -Append
                        }
                        else
                            {
                                "   }," | out-file $BadgingConnectorCSVFile -Encoding utf8 -Append
                            }
                }
                "]" | out-file $BadgingConnectorCSVFile -Encoding utf8 -Append
        }
        catch 
            {
                write-Debug $error[0].Exception
                logWrite 7 $false "Error creating the BadgingConnectorData.csv file."
                exitScript
            }
    #---------------------------------------------------------------------
    # InsiderRisks - Create the CSV file for Priority Physical Assests import
    #---------------------------------------------------------------------
    "Asset ID" | Out-File $Priority_physical_assets -Encoding utf8
    "Tokyo_Main_01" | Out-File $Priority_physical_assets -Encoding utf8 -Append
    "Delhi_Main_01" | Out-File $Priority_physical_assets -Encoding utf8 -Append
    "Shanghai_Main_01" | Out-File $Priority_physical_assets -Encoding utf8 -Append
    "SaoPaulo_Main_01" | Out-File $Priority_physical_assets -Encoding utf8 -Append
    "MexicoCity_Main_01" | Out-File $Priority_physical_assets -Encoding utf8 -Append
    "Dhaka_Main_01" | Out-File $Priority_physical_assets -Encoding utf8 -Append
    "Cairo_Main_01" | Out-File $Priority_physical_assets -Encoding utf8 -Append
    "Beijing_Main_01" | Out-File $Priority_physical_assets -Encoding utf8 -Append
    "London_Main_01" | Out-File $Priority_physical_assets -Encoding utf8 -Append
    "Seattle_Main_01" | Out-File $Priority_physical_assets -Encoding utf8 -Append
    if($global:Recovery -eq $false)
        {
            logWrite 7 $True "Successfully created the BadgingConnectordata.csv file."
            $global:nextPhase++
            Write-Debug "nextPhase set to $global:nextPhase"
        }
}

#---------------------------------------------------------------------
# InsiderRisks - Create an Azure App for Physical Badging Connector (Step 8)
#---------------------------------------------------------------------
function InsiderRisks_CreateAzureApp_BadgingConnector
{
    try
        {
            $BadgingApp_appsecret = "$($LogPath)_BadgingApp_appsecret.txt"
            $BadgingApp_Name = "BadgingConnector01"
            $appExists = $null
            $appExists = Get-AzureADApplication -SearchString $BadgingApp_Name
            $AzureTenantID = Get-AzureADTenantDetail
            $global:tenantid = $AzureTenantID.ObjectId
            if ($null -eq $appExists)
                {
                    $AzureADAppReg = New-AzureADApplication -DisplayName $BadgingApp_Name -AvailableToOtherTenants $false -ErrorAction Stop
                    $appname = $AzureADAppReg.DisplayName
                    $global:appid = $AzureADAppReg.AppID
                    #$AzureTenantID = Get-AzureADTenantDetail
                    #$global:tenantid = $AzureTenantID.ObjectId
                    $AzureSecret = New-AzureADApplicationPasswordCredential -CustomKeyIdentifier PrimarySecret -ObjectId $azureADAppReg.ObjectId -EndDate ((Get-Date).AddMonths(6)) -ErrorAction Stop
                    $global:Secret = $AzureSecret.value
                    "Secret" | out-file $BadgingApp_appsecret -Encoding utf8 -ErrorAction Stop
                    $global:Secret | out-file $BadgingApp_appsecret -Encoding utf8 -Append -ErrorAction Stop
                    write-host
                    write-host "##########################################################################################" -ForegroundColor Green
                    write-host "##                                                                                      ##" -ForegroundColor Green
                    write-host "##     WorkshopPLUS: Microsoft 365 Security and Compliance - Microsoft Purview  and     ##" -ForegroundColor Green
                    write-host "##     Activate Microsoft 365 Security and Compliance: Purview Manage Insider Risks     ##" -ForegroundColor Green
                    write-host "##                                                                                      ##" -ForegroundColor Green            
                    write-host "##   App name  : $appname                                                     ##" -ForegroundColor Green
                    write-host "##   App ID    : $global:appid                                   ##" -ForegroundColor Green
                    write-host "##   Tenant ID : $global:tenantid                                   ##" -ForegroundColor Green
                    write-host "##   App Secret: $global:secret                           ##" -ForegroundColor Green
                    write-host "##                                                                                      ##" -ForegroundColor Green
                    write-host "##########################################################################################" -ForegroundColor Green
                    write-host
                    Write-host "Return to the lab instructions" -ForegroundColor Yellow
                    Write-host "When requested, press ENTER to continue." -ForegroundColor Yellow
                    write-host
                }
                else 
                    {
                        $appname = $appExists.DisplayName
                        $global:appid = $appExists.AppId
                        $SecretFileExists = Test-Path $BadgingApp_appsecret
                        if ($SecretFileExists)
                            {
                                $Secretfile = Import-Csv $BadgingApp_appsecret -Encoding utf8 -ErrorAction SilentlyContinue
                            }
                            else
                                {
                                    Remove-AzureADApplication -ObjectId $appExists.ObjectId
                                    lastEntryPhase = 2
                                    logWrite 8 $false "Badging Azure App already exists, but the secret file was not found. Try again."
                                }
                        $global:Secret = $Secretfile.Secret
                        write-host
                        write-host "##########################################################################################" -ForegroundColor Green
                        write-host "##                                                                                      ##" -ForegroundColor Green
                        write-host "##     WorkshopPLUS: Microsoft 365 Security and Compliance - Microsoft Purview  and     ##" -ForegroundColor Green
                        write-host "##     Activate Microsoft 365 Security and Compliance: Purview Manage Insider Risks     ##" -ForegroundColor Green
                        write-host "##                                                                                      ##" -ForegroundColor Green            
                        write-host "##   App name  : $appname                                                     ##" -ForegroundColor Green
                        write-host "##   App ID    : $global:appid                                   ##" -ForegroundColor Green
                        write-host "##   Tenant ID : $global:tenantid                                   ##" -ForegroundColor Green
                        write-host "##   App Secret: $global:secret                           ##" -ForegroundColor Green
                        write-host "##                                                                                      ##" -ForegroundColor Green
                        write-host "##########################################################################################" -ForegroundColor Green
                        write-host
                        Write-host "Return to the lab instructions" -ForegroundColor Yellow
                        Write-host "When requested, press ENTER to continue." -ForegroundColor Yellow
                        write-host
                        logWrite 8 $True "Azure App for Badging Connector already exists, so this step was skipped."
                        $global:Badgingapp_JustUploadJSON = $true #This variable will be used to skip the step 8 (Create app) if the app already exists.
                    }
        }
        catch 
            {
                write-Debug $error[0].Exception
                logWrite 8 $false "Error creating the Azure App for Physical Badging Connector. Try again."
                exitScript
            }
    if($global:Recovery -eq $false)
        {
            logWrite 8 $True "Successfully created the Azure App for Bagde Connector."
            $global:nextPhase++
            Write-Debug "nextPhase set to $global:nextPhase"
        }
}

#---------------------------------------------------------------------
# InsiderRisks - Upload CSV file for Physical Badging Connector (Step 9)
#---------------------------------------------------------------------
function InsiderRisks_UploadCSV_BadgingConnector
{
    try   
        {
            Write-Host
            $BadgingConnector_JobID = "$($LogPath)_BadgingConnector_jobID.txt"
            if ($global:Badgingapp_JustUploadJSON -eq $false)
            {
                $ConnectorJobID = Read-Host "Paste the Connector job ID"
                if ($null -eq $ConnectorJobID)
                    {
                        $ConnectorJobID = Read-Host "Paste the Connector job ID"
                    }
                "JobID" | out-file $BadgingConnector_JobID -Encoding utf8 -ErrorAction Stop
                $ConnectorJobID | out-file $BadgingConnector_JobID -Encoding utf8 -Append -ErrorAction Stop
            }
            else
                {
                    $JobIDfile = Import-Csv $BadgingConnector_JobID -Encoding utf8 -ErrorAction SilentlyContinue
                    $ConnectorJobID = $JobIDfile.JobID
                }
            Write-Host
            write-host "##########################################################################################" -ForegroundColor Green
            write-host "##                                                                                      ##" -ForegroundColor Green
            write-host "##     WorkshopPLUS: Microsoft 365 Security and Compliance - Microsoft Purview  and     ##" -ForegroundColor Green
            write-host "##     Activate Microsoft 365 Security and Compliance: Purview Manage Insider Risks     ##" -ForegroundColor Green
            write-host "##                                                                                      ##" -ForegroundColor Green            
            write-host "##   App ID    : $global:appid                                   ##" -ForegroundColor Green
            write-host "##   Tenant ID : $global:tenantid                                   ##" -ForegroundColor Green
            write-host "##   App Secret: $global:secret                           ##" -ForegroundColor Green
            write-host "##   JobId     : $ConnectorJobID                                   ##" -ForegroundColor Green
            write-host "##   CSV File  : $global:BadgingConnectorCSVFile     ##" -ForegroundColor Green
            write-host "##                                                                                      ##" -ForegroundColor Green
            write-host "##########################################################################################" -ForegroundColor Green
            Write-Host
            Set-Location -Path "$env:UserProfile\Desktop\SCLabFiles\Scripts"
            .\upload_Badging_records.ps1 -tenantId $tenantId -appId $appId -appSecret $Secret -jobId $ConnectorJobID -jsonFilePath $BadgingConnectorCSVFile
        }
        catch 
            {
                write-Debug $error[0].Exception
                logWrite 9 $false "Error uploading the BadgingConnectorData.csv file"
                exitScript
            }
    if($global:Recovery -eq $false)
        {
            logWrite 9 $True "Successfully uploading the BadgingConnectorData.csv file."
            $global:nextPhase++
            Write-Debug "nextPhase set to $global:nextPhase"
        }
}


#######################################################################################
#########            S C R I P T    S T A R T S   H E R E                    ##########
#######################################################################################

#---------------------------------------------------------------------
# Variable definition - General
#---------------------------------------------------------------------
$LogPath = "$env:UserProfile\Desktop\SCLabFiles\Scripts\"
$LogCSV = "$env:UserProfile\Desktop\SCLabFiles\Scripts\InsiderRisks_Log.csv"
$global:nextPhase = 1
$global:Recovery = $false
$global:HRapp_JustUploadCSV = $false
$global:Badgingapp_JustUploadJSON = $false

#-------------------------------------------------------------------------
# Debug mode
#-------------------------------------------------------------------------
$oldDebugPreference = $DebugPreference
if($debug)
{
    Write-debug "Debug Enabled"
    $DebugPreference = "Continue"
    Start-Transcript -Path "$($LogPath)download-debug.txt"
}

if(!(Test-Path($logCSV)))
    {
        # if log doesn't exist then must be first time we run this, so go to initialization function
        Write-Debug "Entering Initialization"
        Initialization
    } 
        else 
            {
                # if log already exists, check if we need to recover
                Write-Debug "Entering Recovery"
                Recovery
            }

#---------------------------------------------------------------------
# use variable to control phases
#---------------------------------------------------------------------
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
        DownloadScripts
    }

if($nextPhase -eq 4)
    {
        write-debug "Phase $nextPhase"
        InsiderRisks_CreateCSVFile_HRConnector
    }

if($nextPhase -eq 5)
    {
        write-debug "Phase $nextPhase"
        InsiderRisks_CreateAzureApp_HRConnector
        if ($global:HRapp_JustUploadCSV -eq $true)
            {
                Write-Debug "nextPhase set to $global:nextPhase"
            }
            else
                {
                    $answer = Read-Host "Press ENTER to continue"
                }
    }

if($nextPhase -eq 6)
    {
        write-debug "Phase $nextPhase"
        InsiderRisks_UploadCSV_HRConnector
    }

if($nextPhase -eq 7)
    {
        write-debug "Phase $nextPhase"
        InsiderRisks_CreateCSVFile_BadgingConnector
    }

if($nextPhase -eq 8)
    {
        write-debug "Phase $nextPhase"
        InsiderRisks_CreateAzureApp_BadgingConnector
        if ($global:Badgingapp_JustUploadJSON -eq $true)
            {
                Write-Debug "nextPhase set to $global:nextPhase"
            }
            else
                {
                    $answer = Read-Host "Press ENTER to continue"
                }
    }

if($nextPhase -eq 9)
    {
        write-debug "Phase $nextPhase"
        InsiderRisks_UploadCSV_BadgingConnector
    }

if($nextPhase -eq 10)
    {
        write-debug "Phase $nextPhase"
        logWrite 10 $true "Configuration completed"
        exitScript
    }