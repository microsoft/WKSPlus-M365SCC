$VerbosePreference = "Continue"
$LogPath = 'c:\temp'
Get-ChildItem "$LogPath\*.log" | Where LastWriteTime -LT (Get-Date).AddDays(-15) | Remove-Item -Confirm:$false
$LogPathName = Join-Path -Path $LogPath -ChildPath "$($MyInvocation.MyCommand.Name)-$(Get-Date -Format 'MM-dd-yyyy').log"
Start-Transcript $LogPathName -Append

Write-Verbose "$(Get-Date)"


#### install and load module Exchange online V2 #####
#Install-Module PowershellGet -Force
#Install-Module -Name ExchangeOnlineManagement -force


###### Connect & Login to ExchangeOnline and Compliance Center (MFA) ######
$getsessions = Get-PSSession | Select-Object -Property State, Name
$isconnected = (@($getsessions) -like '@{State=Opened; Name=ExchangeOnlineInternalSession*').Count -gt 0
If ($isconnected -ne "True") {
    Write-Host -ForegroundColor 'red' 'Will make a connection to Exchange online and Microsoft 365 Compliance Center'
    Start-Sleep -seconds 3

Connect-IPPSSession
Connect-ExchangeOnline
}
else {
   write-host -ForegroundColor 'Green' " You already have a connection to Office365 compliance Center"
}
Start-Sleep -Seconds 5

#### permissions variable #####
$domainname =  (Get-AcceptedDomain).domainname[0]
$Encpermission = $domainname + ":VIEW,VIEWRIGHTSDATA,DOCEDIT,EDIT,PRINT,EXTRACT,REPLY,REPLYALL,FORWARD,OBJMODEL"
$groups = (Get-UnifiedGroup).PrimarySmtpAddress

### Create new label for Highly confidential #######
New-Label -DisplayName ‘WKS Highly confidential-04’ -Name ‘WKS-Highly-confidential-04’ -ToolTip ‘Contains Highly confidential info’ -Comment ‘Documents with this label contain sensitive data.’ -ContentType "file","Email","Site","UnifiedGroup" -EncryptionEnabled:$true -SiteAndGroupProtectionEnabled:$true -EncryptionPromptUser:$true -EncryptionRightsDefinitions $Encpermission -SiteAndGroupProtectionPrivacy 'private' -SiteExternalSharingControlType 'ExistingExternalUserSharingOnly' -EncryptionDoNotForward:$true -SiteAndGroupProtectionAllowLimitedAccess:$true 

Start-Sleep -Seconds 30

###### publish labels

New-LabelPolicy -name 'WKS-Highly-confidential-publish-03' -Settings @{mandatory=' false'} -AdvancedSettings @{requiredowngradejustification= 'true'} -Labels 'WKS-Highly-confidential-04' 
Stop-Transcript