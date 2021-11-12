New-Item -Path "$env:UserProfile\Desktop\SCLabFiles\Scripts" -ItemType Directory -ErrorAction SilentlyContinue
Invoke-WebRequest -Uri https://raw.githubusercontent.com/microsoft/WKSPlus-M365SCC/main/Compliance%20Center/download-hub.ps1 -OutFile "$env:UserProfile\Desktop\SCLabFiles\Scripts\download-hub.ps1" -ErrorAction Stop
Set-Location -Path "$env:UserProfile\Desktop\SCLabFiles\Scripts"