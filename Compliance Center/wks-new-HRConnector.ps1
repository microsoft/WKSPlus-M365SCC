   <#
   .SYNOPSIS
   Short description
   
   .DESCRIPTION
   Long description
   
   .EXAMPLE
   An example
   
   .NOTES
   General notes
   #>


Function ConnectToAzureAD
    {
        Connect-AzureAD -InformationAction SilentlyContinue | Out-Null
    }

Function CreateAzureapp
    {
        $AzureADAppReg = New-AzureADApplication -DisplayName HRConnector -AvailableToOtherTenants $false
        $global:appname = $AzureADAppReg.DisplayName
        $global:appid = $AzureADAppReg.AppID
        $AzureTenantID = Get-AzureADTenantDetail
        $global:tenantid = $AzureTenantID.ObjectId
        $AzureSecret = New-AzureADApplicationPasswordCredential -CustomKeyIdentifier PrimarySecret -ObjectId $azureADAppReg.ObjectId -EndDate ((Get-Date).AddMonths(6))
        $global:Secret = $AzureSecret.value

        write-host "##################################################################" -ForegroundColor Green
        write-host "##                                                              ##" -ForegroundColor Green
        write-host "##   Microsoft 365 Security and Compliance: Compliance Center   ##" -ForegroundColor Green
        write-host "##                                                              ##" -ForegroundColor Green
        write-host "##   App name  : $global:appname                                    ##" -ForegroundColor Green
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
    #$appid = "75491c0d-51ce-445b-ac74-9d134f8c89cb"
    #$tenantid = "df19d9ab-81ab-4d8f-900c-47564c8706e5"
    #$Secret = "ZDl7Q~ZPJmgPi0zERZi~lSTZCVMYrFNum1PZJ"

    #$jobid = "4e6242d6-7c91-4b9d-9f59-981a4c84ea0c"

Function GenerateTheCSV
    {
        $CurrentPath = Get-Location
        write-host "##################################################################" -ForegroundColor Green
        write-host "##                                                              ##" -ForegroundColor Green
        write-host "##   Microsoft 365 Security and Compliance: Compliance Center   ##" -ForegroundColor Green
        write-host "##                                                              ##" -ForegroundColor Green
        write-host "##   The CSV file was created on $CurrentPath\HRConnector.csv" -ForegroundColor Green
        write-host "##                                                              ##" -ForegroundColor Green
        write-host "##################################################################" -ForegroundColor Green
        write-host
        Write-host "Return to the lab instructions" -ForegroundColor Yellow
        Write-host "When requested, press ENTER to continue." -ForegroundColor Yellow
        write-host

        $global:HRConnectorCSVFile = ".\HRConnector.csv"
        "HRScenarios,EmailAddress,ResignationDate,LastWorkingDate,EffectiveDate,YearsOnLevel,OldLevel,NewLevel,PerformanceRemarks,PerformanceRating,ImprovementRemarks,ImprovementRating" | out-file $HRConnectorCSVFile -Encoding utf8
        $Users = Get-AzureADuser | where-object {$null -ne $_.AssignedLicenses} | Select-Object UserPrincipalName

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
                "Resignation,$EmailAddress,$ResignationDate,$LastWorkingDate," | out-file $HRConnectorCSVFile -Encoding utf8 -Append
                "Job level change,$EmailAddress,,,$EffectiveDate,$YearsOnLevel,Level $OldLevel,Level $NewLevel" | out-file $HRConnectorCSVFile -Encoding utf8 -Append
                "Performance review,$EmailAddress,,,$EffectiveDate,,,,$PerformanceRemarks,$PerformanceRating" | out-file $HRConnectorCSVFile -Encoding utf8 -Append
                "Performance improvement plan,$EmailAddress,,,$EffectiveDate,,,,,,$ImprovementRemarks,$ImprovementRating,"  | out-file $HRConnectorCSVFile -Encoding utf8 -Append
            }
    }

Function RunTheConnector
    {
        $ConnectorJobID = Read-Host "Paste the Connector job ID"
        Write-Host
        write-host "##################################################################" -ForegroundColor Green
        write-host "##                                                              ##" -ForegroundColor Green
        write-host "##   Microsoft 365 Security and Compliance: Compliance Center   ##" -ForegroundColor Green
        write-host "##                                                              ##" -ForegroundColor Green
        write-host "##   App ID    : $global:appid           ##" -ForegroundColor Green
        write-host "##   Tenant ID : $global:tenantid           ##" -ForegroundColor Green
        write-host "##   App Secret: $global:secret   ##" -ForegroundColor Green
        write-host "##   JobId     : $ConnectorJobID           ##" -ForegroundColor Green
        write-host "##   CSV File  : $global:HRConnectorCSVFile           ##" -ForegroundColor Green
        write-host "##                                                              ##" -ForegroundColor Green
        write-host "##################################################################" -ForegroundColor Green
        Write-Host

        C:\temp\upload_termination_records.ps1 -tenantId $tenantId -appId $appId -appSecret $Secret -jobId $ConnectorJobID -csvFilePath $HRConnectorCSVFile

        Read-Host "Press ENTER to continue"
    }

#Script starts here

clear-host
write-host
Write-Host "           
WorkshopPLUS: Microsoft 365 Security and Compliance: Compliance Center

This script will help to create the HR Connector for the Lab 6.1 - Insider Risk Management.

Step 1 - Create Azure app
Step 2 - Generate the CSV file
Step 3 - Run the Connector

" -ForegroundColor Yellow

$answer = Read-Host "Are you ready to proceed (y/n)"

if ($answer -eq "y")
    {
        ConnectToAzureAD  #Call the function
        Write-Host
        CreateAzureapp #Call the function
        Write-Host
        Read-Host "Press ENTER to continue"
        Write-Host
        GenerateTheCSV #call the function
        Write-Host
        Read-Host "Press ENTER to continue"
        Write-Host
        RunTheConnector #call the function
        Write-Host
        Read-Host "Press ENTER to continue"
        Write-Host
    }

if ($answer -eq "1")
    {
        ConnectToAzureAD  #Call the function
        Write-Host
        CreateAzureapp #Call the function
        Write-Host
        exit
    }

if ($answer -eq "2")
    {
        ConnectToAzureAD  #Call the function
        Write-Host
        GenerateTheCSV #call the function
        Write-Host
        Exit
    }

if ($answer -eq "3")
    {
        ConnectToAzureAD  #Call the function
        Write-Host
        RunTheConnector #call the function
        Write-Host
        Exit
    }

Exit