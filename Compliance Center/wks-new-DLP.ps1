$VerbosePreference = "Continue"
$LogPath = 'c:\temp'
Get-ChildItem "$LogPath\*.log" | Where LastWriteTime -LT (Get-Date).AddDays(-15) | Remove-Item -Confirm:$false
$LogPathName = Join-Path -Path $LogPath -ChildPath "$($MyInvocation.MyCommand.Name)-$(Get-Date -Format 'MM-dd-yyyy').log"
Start-Transcript $LogPathName -Append

Write-Verbose "$(Get-Date)"

$params = @{
    ‘Name’ = ‘WKS-Credit Card Number’;
    ‘ExchangeLocation’ =’All’;
    ‘OneDriveLocation’ = ‘All’;
    ‘SharePointLocation’ = ‘All’;
    ‘Mode’ = ‘Enable’
    }
    new-dlpcompliancepolicy @params

    New-DlpComplianceRule -

   