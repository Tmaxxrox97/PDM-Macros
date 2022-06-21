@ Echo Off

Echo Installing SOLIDWORKS Administrative Image

rem *****************************************************************************************************

rem **************** Installation Options Available for Administrative Images **************************

rem *****************************************************************************************************

rem *** /install — Installs the administrative image on client machines

rem *** /uninstall — Uninstalls the software from client machines, with two optional switches:

rem ******* /removedata — Removes SOLIDWORKS data files and folders during uninstall

rem ******* /removeregistry — Removes SOLIDWORKS registry entries during uninstall

rem *** /showui — Displays a progress window for the SOLIDWORKS Installation Manager (Hidden otherwise)

rem *** /now — Starts the install or uninstall immediately. No 5 minute warning dialog

rem *****************************************************************************************************

"\\Desktop-jpe5noa\solidworks 2021 sp4.1\StartSWInstall.exe" /install /now /showui
