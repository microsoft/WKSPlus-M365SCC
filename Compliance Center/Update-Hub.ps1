$LogPath = "$env:UserProfile\Desktop\SCLabFiles\Scripts\"
New-Item -Path "$($LogPath)Desktop\SCLabFiles\Scripts" -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
Invoke-WebRequest -Uri https://raw.githubusercontent.com/microsoft/WKSPlus-M365SCC/main/Compliance%20Center/Update-hub.ps1 -OutFile "$($LogPath)Desktop\SCLabFiles\Scripts\Update-hub.ps1" -ErrorAction Stop
Set-Location -Path "$($LogPath)"