powershell -command "& {Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Force}"
powershell -noprofile -command "&{ start-process powershell -ArgumentList '-noprofile -file \"%~dp0Setup.ps1\" ' -verb RunAs}"
