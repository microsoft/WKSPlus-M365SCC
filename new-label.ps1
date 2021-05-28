$VerbosePreference = "Continue"
$LogPath = 'c:\temp'
Get-ChildItem "$LogPath\*.log" | Where LastWriteTime -LT (Get-Date).AddDays(-15) | Remove-Item -Confirm:$false
$LogPathName = Join-Path -Path $LogPath -ChildPath "$($MyInvocation.MyCommand.Name)-$(Get-Date -Format 'MM-dd-yyyy').log"
Start-Transcript $LogPathName -Append

Write-Verbose "$(Get-Date)"
#### install and load module Exchange online V2 #####
#Install-Module PowershellGet -Force
#Install-Module -Name ExchangeOnlineManagement -force


### connect to the compliance center####
#Connect-IPPSSession
#Connect-ExchangeOnline

#### permissions variable #####
$domainname =  (Get-AcceptedDomain).domainname[0]
$Encpermission = $domainname + ":VIEW,VIEWRIGHTSDATA,DOCEDIT,EDIT,PRINT,EXTRACT,REPLY,REPLYALL,FORWARD,OBJMODEL"
$groups = (Get-UnifiedGroup).PrimarySmtpAddress

### Create new label for Highly confidential #######
New-Label -DisplayName ‘WKS Highly confidential-04’ -Name ‘WKS-Highly-confidential-04’ -ToolTip ‘Contains Highly confidential info’ -Comment ‘Documents with this label contain sensitive data.’ -ContentType "file","Email","Site","UnifiedGroup" -EncryptionEnabled:$true -SiteAndGroupProtectionEnabled:$true -EncryptionPromptUser:$true -EncryptionRightsDefinitions $Encpermission -SiteAndGroupProtectionPrivacy 'private' -SiteExternalSharingControlType 'ExistingExternalUserSharingOnly' -EncryptionDoNotForward:$true -SiteAndGroupProtectionAllowLimitedAccess:$true 

Start-Sleep -Seconds 30

###### publish labels

New-LabelPolicy -name 'WKS-Highly-confidential-publish-03' -Settings @{mandatory=' false'} -AdvancedSettings @{requiredowngradejustification= 'true'} -Labels 'WKS-Highly-confidential-04' -SkypeLocation 'all' -OneDriveLocation 'all' -ExchangeLocation 'all' -SharePointLocation 'all' -ModernGroupLocation $groups

Stop-Transcript