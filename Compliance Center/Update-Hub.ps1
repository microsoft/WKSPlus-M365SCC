$LogPath = "$env:UserProfile\Desktop\SCLabFiles\Scripts\"
New-Item -Path "$($LogPath)Desktop\SCLabFiles\Scripts" -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
Invoke-WebRequest -Uri https://raw.githubusercontent.com/microsoft/WKSPlus-M365SCC/main/Compliance%20Center/Download-Hub.ps1 -OutFile "$($LogPath)Download-Hub.ps1" -ErrorAction Stop
Set-Location -Path "$($LogPath)"
Remove-Item -Path "$($LogPath)Download-Hub.ps1" -ErrorAction SilentlyContinue | Out-Null
Remove-Item -Path "$($LogPath)Progress_Download_Log.csv" -ErrorAction SilentlyContinue | Out-Null
