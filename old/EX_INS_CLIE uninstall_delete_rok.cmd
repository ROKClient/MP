REM Parameters section *********************************
Set Origin=https://cloudapp.rok-solution.com/ClientBin/Bpm.Shell.xap

REM execution section **********************************
REM 1. Uninstall ROK


REM cd %ProgramFiles%\Microsoft Silverlight\
REM sllauncher.exe /uninstall /origin:%Origin%

REM 2. Delete remaining ROK files
REM 2.1 Delete ROK links in the Start Menu Folder

cd %appdata%\Microsoft\Windows\Start Menu\Programs
del ROK*

REM 2.2 Delete ROK links in the desktop

cd %USERPROFILE%\Desktop\
del ROK*

REM 2.3 Delete AppLocal folder ROK
rmdir /Q/S %USERPROFILE%\AppData\Local\ROK
rmdir /Q/S %USERPROFILE%\AppData\Local\ROK


rmdir /Q/S %USERPROFILE%\AppData\Roaming\ROKInstall
rmdir /Q/S %USERPROFILE%\AppData\Roaming\ROKInstall


REM 2.4 Delete AppLocal folder Silverlight
rmdir /Q/S %USERPROFILE%\AppData\Local\Microsoft\Silverlight
rmdir /Q/S %USERPROFILE%\AppData\Local\Microsoft\Silverlight



REM Delete AppLocalLow folder Silverlight
rmdir /Q/S %USERPROFILE%\AppData\LocalLow\Microsoft\Silverlight
rmdir /Q/S %USERPROFILE%\AppData\LocalLow\Microsoft\Silverlight
