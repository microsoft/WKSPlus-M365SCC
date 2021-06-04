$VerbosePreference = "Continue"
$LogPath = 'c:\temp'
Get-ChildItem "$LogPath\*.log" | Where LastWriteTime -LT (Get-Date).AddDays(-15) | Remove-Item -Confirm:$false
$LogPathName = Join-Path -Path $LogPath -ChildPath "$($MyInvocation.MyCommand.Name)-$(Get-Date -Format 'MM-dd-yyyy').log"
Start-Transcript $LogPathName -Append

Write-Verbose "$(Get-Date)"

#### Conntection #####

#Connect-IPPSSession
#Connect-ExchangeOnline


Start-Sleep -Seconds 5
##### Settings #####

[Array]$DomainName = Get-AcceptedDomain

$SuffixDomain = $DomainName[0].DomainName
$Email = "Admin@$SuffixDomain"


##### DLP Policy Parameters. #####
$params = @{
    'Name' = 'WKS-Credit Card Number-test01';
    'ExchangeLocation' ='All';
    'OneDriveLocation' = 'All';
    'SharePointLocation' = 'All';
    'EndpointDlpLocation' = 'all';
    'Mode' = 'Enable'
    }
    new-dlpcompliancepolicy @params

    ###### sensitivity Types low Volume ############
$SensitiveTypes = @( 
    @{Name="Credit Card Number"; minCount="1"; maxcount="5"}    
)

    ###### sensitivity Types low Volume ############
    $SensitiveTypesHigh = @( 
        @{Name="Credit Card Number"; minCount="6"}    
    )

    Start-Sleep -Seconds 5
    #### New DLP Rule Low and High volume. ######
     New-DlpComplianceRule -Name "WKS-Credit Card Number-low-01" -Policy "WKS-Credit Card Number-test01" -ContentContainsSensitiveInformation $SensitiveTypes -NotifyUser 'lastmodified' -blockaccess:$true

    New-DlpComplianceRule -Name "WKS-Credit Card Number-High-02" -Policy "WKS-Credit Card Number-test01" -ContentContainsSensitiveInformation $SensitiveTypesHigh -NotifyUser 'LastModifier','owner' -blockaccess:$true -BlockAccessScope 'All' -GenerateIncidentReport $email -GenerateAlert


    Stop-Transcript