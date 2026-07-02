$WshShell = New-Object -ComObject WScript.Shell
$DesktopPath = [System.Environment]::GetFolderPath('Desktop')
$Shortcut = $WshShell.CreateShortcut("$DesktopPath\Automação ITUGISS.lnk")
$Shortcut.TargetPath = "$PSScriptRoot\START_AUTOMATION.bat"
$Shortcut.WorkingDirectory = "$PSScriptRoot"
$Shortcut.IconLocation = "$PSScriptRoot\itugiss_robot.ico"
$Shortcut.Description = "Lançador da Automação de NFS-e ITUGISS"
$Shortcut.Save()

Write-Output "Atalho criado com sucesso na sua Área de Trabalho! 🚀🤖"
pause

