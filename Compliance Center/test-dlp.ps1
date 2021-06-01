# Step 1 - Authenticate to SCC FIRST.

Connect-IPPSSession
Connect-ExchangeOnline

###### sensitivity Types
$SensitiveTypes = @( 
    @{Name="Credit Card Number"; minCount="1"}    
)

# Syntax: https://docs.microsoft.com/en-us/powershell/module/exchange/policy-and-compliance-dlp/new-dlpcompliancepolicy?view=exchange-ps
New-DlpCompliancePolicy `
    -Name "WKS-Credit Card Number -test" `
    -ExchangeLocation "All"`
    -SharePointLocation "All"`
    -OneDriveLocation "All"`
    -TeamsLocation "All"`



# Syntax: https://docs.microsoft.com/en-us/powershell/module/exchange/policy-and-compliance-dlp/new-dlpcompliancerule?view=exchange-ps
New-DlpComplianceRule `
    -Name "WKS-Credit Card Number-low" `
    -Policy "WKS-Credit Card Number -test" `
    -AccessScope NotInOrganization `
    -EncryptRMSTemplate "Encrypt" `
    -NotifyUser "LastModifier" `
    -NotifyPolicyTipCustomText "This email contains sensitive information and will be encrypted." `
    -NotifyEmailCustomText "This email contains sensitive information and will be encrypted." `
    -ContentContainsSensitiveInformation $SensitiveTypes