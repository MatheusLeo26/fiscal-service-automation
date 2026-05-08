$WshShell = New-Object -ComObject WScript.Shell
$DesktopPath = [System.Environment]::GetFolderPath('Desktop')
$Shortcut = $WshShell.CreateShortcut("$DesktopPath\Automação ITUGISS.lnk")
$Shortcut.TargetPath = "$env:USERPROFILE\.gemini\antigravity\scratch\itugiss_automation\START_AUTOMATION.bat"
$Shortcut.WorkingDirectory = "$env:USERPROFILE\.gemini\antigravity\scratch\itugiss_automation"
$Shortcut.IconLocation = "$env:USERPROFILE\.gemini\antigravity\scratch\itugiss_automation\itugiss_robot.ico"
$Shortcut.Description = "Lançador da Automação de NFS-e ITUGISS"
$Shortcut.Save()

Write-Output "Atalho criado com sucesso na sua Área de Trabalho! 🚀🤖"
pause
