$hostsPath = "$env:windir\System32\drivers\etc\hosts"
try {
    Set-ItemProperty $hostsPath -Name IsReadOnly -Value $false -ErrorAction SilentlyContinue
    Add-Content -Path $hostsPath -Value "127.0.0.1 robo.itugiss" -ErrorAction Stop
    "Success" | Out-File -FilePath "C:\Users\Matheus\.gemini\antigravity\scratch\itugiss_automation\update_log.txt"
} catch {
    $_.Exception.Message | Out-File -FilePath "C:\Users\Matheus\.gemini\antigravity\scratch\itugiss_automation\update_log.txt"
}
