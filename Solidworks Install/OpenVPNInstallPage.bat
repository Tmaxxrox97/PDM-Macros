rem powershell -command "& {Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Force}"
xcopy /y "%~dp0AzureVPN.Msixbundle" "%localappdata%\Temp\PDMVPN\"
xcopy /y "%~dp0InstallAzureVPN.ps1" "%localappdata%\Temp\PDMVPN\"
powershell -command "& {Add-AppPackage -path "%localappdata%\Temp\PDMVPN\AzureVPN.Msixbundle"}"
xcopy /y "%~dp0azurevpnconfig.xml" "%userprofile%\AppData\Local\Packages\Microsoft.AzureVpn_8wekyb3d8bbwe\LocalState\"
azurevpn -i azurevpnconfig.xml
rem powershell -command "& {Set-ExecutionPolicy -ExecutionPolicy Restricted -Force}"
