$VerbosePreference = "Continue"
$LogPath = "c:\temp"

##### check if log path excist if not will create it.

If ( !(Test-Path $LogPath) ) 
{New-Item -ItemType "directory" -Path $LogPath}
else {
    write-host " Folder excist"
}

#### write log file ####
Get-ChildItem "$LogPath\*.log" | Where LastWriteTime -LT (Get-Date).AddDays(-15) | Remove-Item -Confirm:$false
$LogPathName = Join-Path -Path $LogPath -ChildPath "$($MyInvocation.MyCommand.Name)-$(Get-Date -Format "MM-dd-yyyy").log"
Start-Transcript $LogPathName -Append

Write-Verbose "$(Get-Date)"

###### Connect & Login to ExchangeOnline and Compliance Center (MFA) ######
$getsessions = Get-PSSession | Select-Object -Property State, Name
$isconnected = (@($getsessions) -like "@{State=Opened; Name=ExchangeOnlineInternalSession*").Count -gt 0
If ($isconnected -ne "True") {
    Write-Host -ForegroundColor "red" "Will make a connection to Exchange online and Microsoft 365 Compliance Center"

    Start-Sleep -seconds 3

Connect-IPPSSession
Connect-ExchangeOnline
}
else {
   write-host -ForegroundColor "Green" " You already have a connection to Office365 compliance Center"
}
Start-Sleep -Seconds 5
##### Settings #####

[Array]$DomainName = Get-AcceptedDomain

$SuffixDomain = $DomainName[0].DomainName
$Email = "Admin@$SuffixDomain"

#### check for existing DLP Policy. ####

if (Get-DlpCompliancePolicy -Identity "WKS-Credit Card Number")
{
 ##### DLP Policy Parameters. #####
$params = @{
    "Name" = "WKS-Credit Card Number-test";
    "ExchangeLocation" ="All";
    "OneDriveLocation" = "All";
    "SharePointLocation" = "All";
    "EndpointDlpLocation" = "all";
    "Teamslocaltion" = "All";
    "Mode" = "Enable"
    }
    new-dlpcompliancepolicy @params

}

else 
{

    Write-Host " DLP Policy already excist"
}
<##### DLP Policy Parameters. #####
$params = @{
    "Name" = "WKS-Credit Card Number-test02";
    "ExchangeLocation" ="All";
    "OneDriveLocation" = "All";
    "SharePointLocation" = "All";
    "EndpointDlpLocation" = "all";
    "Teamslocaltion" = "All";
    "Mode" = "Enable"
    }
    new-dlpcompliancepolicy @params
#>


    ###### sensitivity Types low Volume ############
$SensitiveTypes = @( 
    @{Name="Credit Card Number"; minCount="1"; maxcount="5"}    
)

    ###### sensitivity Types High Volume ############
    $SensitiveTypesHigh = @( 
        @{Name="Credit Card Number"; minCount="6";}    
    )

    Start-Sleep -Seconds 5
    #### New DLP Rule Low and High volume. ######
     New-DlpComplianceRule -Name "WKS-Credit Card Number-low" -Policy "WKS-Credit Card Number" -ContentContainsSensitiveInformation $SensitiveTypes -NotifyUser "lastmodifier"


    New-DlpComplianceRule -Name "WKS-Credit Card Number-High" -Policy "WKS-Credit Card Number" -ContentContainsSensitiveInformation $SensitiveTypesHigh -NotifyUser "LastModifier","owner" -blockaccess:$true -BlockAccessScope "All" -GenerateIncidentReport $email 


    Stop-Transcript