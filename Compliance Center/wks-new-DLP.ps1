$VerbosePreference = "Continue"
$LogPath = 'c:\temp'
Get-ChildItem "$LogPath\*.log" | Where LastWriteTime -LT (Get-Date).AddDays(-15) | Remove-Item -Confirm:$false
$LogPathName = Join-Path -Path $LogPath -ChildPath "$($MyInvocation.MyCommand.Name)-$(Get-Date -Format 'MM-dd-yyyy').log"
Start-Transcript $LogPathName -Append

Write-Verbose "$(Get-Date)"

#### Conntection #####

Connect-IPPSSession
Connect-ExchangeOnline


Start-Sleep -Seconds 5
##### Settings #####

$domainname =  (Get-AcceptedDomain).domainname[0]


##### DLP Policy Parameters. #####

$params = @{
    ‘Name’ = ‘WKS-Credit Card Number -test’;
    ‘ExchangeLocation’ =’All’;
    ‘OneDriveLocation’ = ‘All’;
    ‘SharePointLocation’ = ‘All’;
    ‘Mode’ = ‘Enable’
    }
    new-dlpcompliancepolicy @params

    #### New DLP Rule Low and High volume. ######
    New-DlpComplianceRule -Name "WKS-Credit Card Number-low" -Policy "WKS-Credit Card Number -test" -ContentContainsSensitiveInformation @{Name="Credit Card Number"; minCount="2";maxCount="5"} -BlockAccess $True -NotifyUser:$true

    New-DlpComplianceRule -Name "WKS-Credit Card Number-High" -Policy "WKS-Credit Card Number -test" -ContentContainsSensitiveInformation @{Name="Credit Card Number"; minCount="5"} -BlockAccess $True -NotifyUser:$true


    Stop-Transcript