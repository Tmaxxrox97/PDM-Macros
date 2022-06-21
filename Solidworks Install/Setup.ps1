New-NetFirewallRule -DisplayName 'PDM TCP INBOUND TEST2' -Profile @('Domain', 'Private') -Direction Inbound -Action Allow -Protocol TCP -LocalPort @('1433', '1434', '3030', '25734', '25735')
Write-Host "Inbound Rule Created"
New-NetFirewallRule -DisplayName 'PDM TCP OUTBOUND TEST2' -Profile @('Domain', 'Private') -Direction Outbound -Action Allow -Protocol TCP -LocalPort @('1433', '1434', '3030', '25734', '25735')
Write-Host "Outbound Rule Created"

Write-Host "Editing Host File"
Add-Content -Path $env:windir\System32\drivers\etc\hosts -Value "`n10.10.2.4`tPDM-SERVER" -Force

Write-Host "Host File Edited"

