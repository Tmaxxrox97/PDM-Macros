@ Echo Off
Echo Setting Host File
CALL "%~dp0Setup.bat"
Echo Installing PDM
CALL "%~dp0InstallSolidworks.bat"
Echo Creating Vault
CALL "%~dp0CreateVault.bat"
Echo Installing VPN
CALL "%~dp0InstallVPN.lnk"
pause
