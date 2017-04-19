setlocal

REM *************************************************************************
REM ROK Silent deployment V1.2
REM Environment customization begins here. Modify variables below.
REM This script aims to deploy the central ROK client application and the
REM webautomation module through a local network
REM The Webautomation files should be located in %DeployServer%\cloudDir
REM *************************************************************************

REM Set DeployServer to the distribution server******************************
set DeployServer=C:\share\

REM Set LogLocation to a central directory to collect log files.**************
Set LogLocation=C:\share\SilverlightLogs

REM Set DownloadSourceFromInternet******************************************
REM DownloadSourceFromInternet = Y  the sources are automatically downloaded from the internet
REM DownloadSourceFromInternet = N the sources are downloaded from the distribution server
set DownloadSourceFromInternet=Y

REM Set DistributionServerUpdate (distribution server mode):****************
REM DistributionServerUpdate = Y the ROK sources (application.xap) that is downloaded from the internet are automatically copied to the Distribution server
REM DistributionServerUpdate = N the ROK source that is downloaded from the internet are NOT copied to the Distribution server
REM Note that the other sources (silverlight.exe, configuration file) are NOT copied to the distribution server and must be set manually.
Set DistributionServerUpdate=N
REM parameter to desinstall ROK before installation
REM UninstallROK = Y ROK is uninstalled before installation
REM UninstallROK=N ROK is not uninstalled before installation
Set UninstallROK=Y


REM Set configuration file parameters: DatabaseName, UseProxy (true  false), UseSSL (true  false), UseAdfs (true false), IsWindowsAuthenticationEnabled (true false)
Set UseProxy=false
Set UseSSL=true
Set UseAdfs=false
Set IsWindowsAuthenticationEnabled=true
Set Databasename=MECAPLAST

REM Set InstallerName to the name of your copy of the Silverlight installer**
REM set InstallerName=Silverlight_x64.exe /q*** InstallerName is automatically determined

REM Set Origin for ROK (for ROK application update)**************************
Set Origin=https://cloudapp3.rok-solution.com/ClientBin/Bpm.Shell.xap

REM Set the Configuration file for ROK***************************************
Set ROKConfig=cloudapp3.rok-solution.com_443.xml



REM *************************************************************************
REM Deployment code begins here. Do not modify anything below this line.
REM *************************************************************************
set OSVersion=WXP
if exist %USERPROFILE%\AppData\Local\ (set OSVersion=W7)
echo %OSVersion%

Set DistributionServer=%DeployServer%
Set DeploySource=%DeployServer%
Set LogLocationSource=%LogLocation%
REM 000. Uninstall ROK before installation
if %UninstallROK%==Y (goto UninstallROKBeforeInstall) else (goto EndUninstallROKBeforeInstall)
:UninstallROKBeforeInstall
REM execution section **********************************
REM 4.1 Uninstall ROK

cd %ProgramFiles%\Microsoft Silverlight\
REM sllauncher.exe /uninstall /origin:%Origin%
REM 4.2 Delete remaining ROK files
REM 4.2.1 Delete ROK links in the Start Menu Folder
cd %appdata%\Microsoft\Windows\Start Menu\Programs
del ROK*
REM 4.2.2 Delete ROK links in the desktop
if %OSVersion%==W7 (cd %USERPROFILE%\Desktop\)
if %OSVersion%==WXP (cd %USERPROFILE%\Bureau\)
del ROK*
REM 4.2.3 Delete AppLocal folder ROK
if %OSVersion%==W7 (rmdir /Q/S %USERPROFILE%\AppData\Local\ROK)
if %OSVersion%==WXP (rmdir /Q/S "%USERPROFILE%\Local Settings\Application Data\ROK")

if %OSVersion%==W7 (rmdir /Q/S %USERPROFILE%\AppData\Local\ROK)
if %OSVersion%==WXP (rmdir /Q/S "%USERPROFILE%\Local Settings\Application Data\ROK")

REM 4.2.4 Delete AppLocal folder Silverlight
if %OSVersion%==W7 (rmdir /Q/S %USERPROFILE%\AppData\Local\Microsoft\Silverlight)
if %OSVersion%==WXP (rmdir /Q/S "%USERPROFILE%\Local Settings\Application Data\Microsoft\Silverlight")

if %OSVersion%==W7 (rmdir /Q/S %USERPROFILE%\AppData\Local\Microsoft\Silverlight)
if %OSVersion%==WXP (rmdir /Q/S "%USERPROFILE%\Local Settings\Application Data\Microsoft\Silverlight")

REM Delete AppLocalLow folder Silverlight
if %OSVersion%==W7 (rmdir /Q/S %USERPROFILE%\AppData\LocalLow\Microsoft\Silverlight)
if %OSVersion%==WXP (rmdir /Q/S "%USERPROFILE%\Local Settings\Application Data\Microsoft\Silverlight")

if %OSVersion%==W7 (rmdir /Q/S %USERPROFILE%\AppData\LocalLow\Microsoft\Silverlight)
if %OSVersion%==WXP (rmdir /Q/S "%USERPROFILE%\Local Settings\Application Data\Microsoft\Silverlight")
:EndUninstallROKBeforeInstall

REM 00. Determine if 32 bits system
reg Query "HKLM\Hardware\Description\System\CentralProcessor\0" | find /i "x86" > NUL && set OS=32BIT || set OS=64BIT

if %OS%==32BIT (goto SetInstallerName32Bit)else (goto EndSetInstallerName32Bit)
:SetInstallerName32Bit
echo This is a 32 bit operating system
set InstallerName=Silverlight.exe /q
:EndSetInstallerName32Bit

if %OS%==64BIT (goto SetInstallerName64Bit)else (goto EndSetInstallerName64Bit)
:SetInstallerName64Bit
echo This is a 64 bit operating system
set InstallerName=Silverlight_x64.exe /q
:EndSetInstallerName64Bit


REM 0. Check the Webauto conf and copy the files to the distribution server
if %DistributionServerUpdate%==WEBAUTO (goto WebAuto) else (goto EndWebAuto)
:WebAuto
mkdir %DeployServer%\WebAutoDir
if %OSVersion%==W7 (robocopy %USERPROFILE%\AppData\Local\ROK %DeployServer%\WebAutoDir /E)
if %OSVersion%==WXP(robocopy ""%USERPROFILE%\Local Settings\Application Data\ROK" %DeployServer%\WebAutoDir /E)
rmdir %DeployServer%\WebAutoDir\Configuration
goto EndInstall
:EndWebAuto

REM 1. Check if Silverlight is installed and install Silverlight if necessary
reg query HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Silverlight
set SilverlightNotDeployed=%errorlevel%
if %SilverlightNotDeployed%==1 (goto DeploySilverlight) else (goto EndDeploySilverLight)
:DeploySilverlight
if %DownloadSourceFromInternet%==Y (goto DownloadSourceFromInternetmentSL) else (goto EndDownloadSourceFromInternetmentSL)
:DownloadSourceFromInternetmentSL
if %OSVersion%==W7 (Set DownloadSourceFromInternetDir=%USERPROFILE%\AppData\Roaming\ROKInstall)
if %OSVersion%==WXP (Set DownloadSourceFromInternetDir="%USERPROFILE%\Local Settings")
mkdir %DownloadSourceFromInternetDir%
del %DownloadSourceFromInternetDir%\downloadSL.vbs
if %OS%==32BIT (goto 32bit) else (goto End32bit)
:32bit
echo strFileURL = "http://download.microsoft.com/download/0/3/E/03EB1393-4F4E-4191-8364-C641FAB20344/50901.00/Silverlight.exe">>%DownloadSourceFromInternetDir%\downloadSL.vbs
set InstallerNameShort=Silverlight.exe
:End32bit
if %OS%==64BIT (goto 64bit) else (goto End64bit)
:64bit
echo strFileURL = "http://download.microsoft.com/download/0/3/E/03EB1393-4F4E-4191-8364-C641FAB20344/50901.00/Silverlight_x64.exe">>%DownloadSourceFromInternetDir%\downloadSL.vbs
set InstallerNameShort=Silverlight_x64.exe
:End64bit
echo strHDLocation = "%DownloadSourceFromInternetDir%\%InstallerNameShort%">>%DownloadSourceFromInternetDir%\downloadSL.vbs
echo Set objXMLHTTP = CreateObject("MSXML2.ServerXMLHTTP")>>%DownloadSourceFromInternetDir%\downloadSL.vbs
echo objXMLHTTP.open "GET", strFileURL, false>>%DownloadSourceFromInternetDir%\downloadSL.vbs
echo objXMLHTTP.send()>>%DownloadSourceFromInternetDir%\downloadSL.vbs
echo If objXMLHTTP.Status = 200 Then>>%DownloadSourceFromInternetDir%\downloadSL.vbs
echo Set objADOStream = CreateObject("ADODB.Stream")>>%DownloadSourceFromInternetDir%\downloadSL.vbs
echo objADOStream.Open>>%DownloadSourceFromInternetDir%\downloadSL.vbs
echo objADOStream.Type = ^1>>%DownloadSourceFromInternetDir%\downloadSL.vbs
echo objADOStream.Write objXMLHTTP.ResponseBody>>%DownloadSourceFromInternetDir%\downloadSL.vbs
echo objADOStream.Position = ^0>>%DownloadSourceFromInternetDir%\downloadSL.vbs
echo Set objFSO = Createobject("Scripting.FileSystemObject")>>%DownloadSourceFromInternetDir%\downloadSL.vbs
echo If objFSO.Fileexists(strHDLocation) Then objFSO.DeleteFile strHDLocation>>%DownloadSourceFromInternetDir%\downloadSL.vbs
echo Set objFSO = Nothing>>%DownloadSourceFromInternetDir%\downloadSL.vbs
echo objADOStream.SaveToFile strHDLocation>>%DownloadSourceFromInternetDir%\downloadSL.vbs
echo objADOStream.Close>>%DownloadSourceFromInternetDir%\downloadSL.vbs
echo Set objADOStream = Nothing>>%DownloadSourceFromInternetDir%\downloadSL.vbs
echo End if>>%DownloadSourceFromInternetDir%\downloadSL.vbs
echo Set objXMLHTTP = Nothing>>%DownloadSourceFromInternetDir%\downloadSL.vbs
cscript.exe %DownloadSourceFromInternetDir%\downloadSL.vbs
set DeploySource=%DownloadSourceFromInternetDir%
set LogLocationSource=%DownloadSourceFromInternetDir%
:EndDownloadSourceFromInternetmentSL
start /wait %DeploySource%\%InstallerName%
echo %date% %time% Setup ended with error code %errorlevel%. >> %LogLocationSource%\%computername%.txt 
:EndDeploySilverLight

REM 2. Download ROK (if DownloadSourceFromInternet is activated)
if %DownloadSourceFromInternet%==Y (goto DownloadSourceFromInternetmentROK) else (goto EndDownloadSourceFromInternetmentROK)
:DownloadSourceFromInternetmentROK
if %OSVersion%==W7 (Set DownloadSourceFromInternetDir=%USERPROFILE%\AppData\Roaming\ROKInstall)
if %OSVersion%==WXP (Set DownloadSourceFromInternetDir="%USERPROFILE%\Local Settings\ROKInstall")

mkdir %DownloadSourceFromInternetDir%
del %DownloadSourceFromInternetDir%\downloadROK.vbs

echo strFileURL = "%Origin%">>%DownloadSourceFromInternetDir%\downloadROK.vbs
echo strHDLocation = "%DownloadSourceFromInternetDir%\application.xap">>%DownloadSourceFromInternetDir%\downloadROK.vbs
echo Set objXMLHTTP = CreateObject("MSXML2.ServerXMLHTTP")>>%DownloadSourceFromInternetDir%\downloadROK.vbs
echo objXMLHTTP.open "GET", strFileURL, false>>%DownloadSourceFromInternetDir%\downloadROK.vbs
echo objXMLHTTP.send()>>%DownloadSourceFromInternetDir%\downloadROK.vbs
echo If objXMLHTTP.Status = 200 Then>>%DownloadSourceFromInternetDir%\downloadROK.vbs
echo Set objADOStream = CreateObject("ADODB.Stream")>>%DownloadSourceFromInternetDir%\downloadROK.vbs
echo objADOStream.Open>>%DownloadSourceFromInternetDir%\downloadROK.vbs
echo objADOStream.Type = ^1>>%DownloadSourceFromInternetDir%\downloadROK.vbs
echo objADOStream.Write objXMLHTTP.ResponseBody>>%DownloadSourceFromInternetDir%\downloadROK.vbs
echo objADOStream.Position = ^0>>%DownloadSourceFromInternetDir%\downloadROK.vbs
echo Set objFSO = Createobject("Scripting.FileSystemObject")>>%DownloadSourceFromInternetDir%\downloadROK.vbs
echo If objFSO.Fileexists(strHDLocation) Then objFSO.DeleteFile strHDLocation>>%DownloadSourceFromInternetDir%\downloadROK.vbs
echo Set objFSO = Nothing>>%DownloadSourceFromInternetDir%\downloadROK.vbs
echo objADOStream.SaveToFile strHDLocation>>%DownloadSourceFromInternetDir%\downloadROK.vbs
echo objADOStream.Close>>%DownloadSourceFromInternetDir%\downloadROK.vbs
echo Set objADOStream = Nothing>>%DownloadSourceFromInternetDir%\downloadROK.vbs
echo End if>>%DownloadSourceFromInternetDir%\downloadROK.vbs
echo Set objXMLHTTP = Nothing>>%DownloadSourceFromInternetDir%\downloadROK.vbs
cscript.exe %DownloadSourceFromInternetDir%\downloadROK.vbs
set DeploySource=%DownloadSourceFromInternetDir%
set LogLocationSource=%DownloadSourceFromInternetDir%
:EndDownloadSourceFromInternetmentROK

REM 3. Generate ROK configuration file (if DownloadSourceFromInternet is activated)
if %DownloadSourceFromInternet%==Y (goto DownloadSourceFromInternetmentROKConf) else (goto EndDownloadSourceFromInternetmentROKConf)
:DownloadSourceFromInternetmentROKConf
del %DownloadSourceFromInternetDir%\%ROKConfig%
echo ^<?xml version="1.0" encoding="utf-8"?^>>>%DownloadSourceFromInternetDir%\%ROKConfig%
echo ^<Configuration xmlns:i="http://www.w3.org/2001/XMLSchema-instance" xmlns="http://www.opserv.fr/2009/11/"^>>>%DownloadSourceFromInternetDir%\%ROKConfig%
echo ^<DatabaseGroups^>>>%DownloadSourceFromInternetDir%\%ROKConfig%
echo ^<DatabaseGroup^>>>%DownloadSourceFromInternetDir%\%ROKConfig%
echo ^<Name^>%Databasename%^</Name^>>>%DownloadSourceFromInternetDir%\%ROKConfig%
echo ^<IsWindowsAuthenticationEnabled^>%IsWindowsAuthenticationEnabled%^</IsWindowsAuthenticationEnabled^>>>%DownloadSourceFromInternetDir%\%ROKConfig%
echo ^<UseProxy^>%UseProxy%^</UseProxy^>>>%DownloadSourceFromInternetDir%\%ROKConfig%
echo ^<UseSsl^>%UseSSL%^</UseSsl^>>>%DownloadSourceFromInternetDir%\%ROKConfig%
echo ^<UseAdfs^>%UseAdfs%^</UseAdfs^>>>%DownloadSourceFromInternetDir%\%ROKConfig%
echo ^</DatabaseGroup^>>>%DownloadSourceFromInternetDir%\%ROKConfig%
echo ^</DatabaseGroups^>>>%DownloadSourceFromInternetDir%\%ROKConfig%
echo ^<AutomaticClientUpdate^>Enabled^</AutomaticClientUpdate^>>>%DownloadSourceFromInternetDir%\%ROKConfig%
echo ^</Configuration^>>>%DownloadSourceFromInternetDir%\%ROKConfig%
:EndDownloadSourceFromInternetmentROKConf

REM 4. (Distribution server mode) Remove existing ROK version
if %DistributionServerUpdate%==Y (goto RemoveROK) else (goto EndRemoveROK)
:RemoveROK
REM execution section **********************************
REM 4.1 Uninstall ROK
cd %ProgramFiles%\Microsoft Silverlight\
REM sllauncher.exe /uninstall /origin:%Origin%
REM 4.2 Delete remaining ROK files
REM 4.2.1 Delete ROK links in the Start Menu Folder
cd %appdata%\Microsoft\Windows\Start Menu\Programs
del ROK*
REM 4.2.2 Delete ROK links in the desktop
if %OSVersion%==W7 (cd %USERPROFILE%\Desktop\)
if %OSVersion%==WXP (cd %USERPROFILE%\Bureau\)
del ROK*
REM 4.2.3 Delete AppLocal folder ROK
if %OSVersion%==W7 (rmdir /Q/S %USERPROFILE%\AppData\Local\ROK)
if %OSVersion%==WXP (rmdir /Q/S "%USERPROFILE%\Local Settings\Application Data\ROK")
if %OSVersion%==W7 (rmdir /Q/S %USERPROFILE%\AppData\Local\ROK)
if %OSVersion%==WXP (rmdir /Q/S "%USERPROFILE%\Local Settings\Application Data\ROK")

REM 4.2.4 Delete AppLocal folder Silverlight
if %OSVersion%==W7 (rmdir /Q/S %USERPROFILE%\AppData\Local\Microsoft\Silverlight)
if %OSVersion%==WXP (rmdir /Q/S "%USERPROFILE%\Local Settings\Application Data\Microsoft\Silverlight")

if %OSVersion%==W7 (rmdir /Q/S %USERPROFILE%\AppData\Local\Microsoft\Silverlight)
if %OSVersion%==WXP (rmdir /Q/S "%USERPROFILE%\Local Settings\Application Data\Microsoft\Silverlight")

REM Delete AppLocalLow folder Silverlight
if %OSVersion%==W7 (rmdir /Q/S %USERPROFILE%\AppData\LocalLow\Microsoft\Silverlight)
if %OSVersion%==WXP (rmdir /Q/S "%USERPROFILE%\Local Settings\Application Data\Microsoft\Silverlight")

if %OSVersion%==W7 (rmdir /Q/S %USERPROFILE%\AppData\LocalLow\Microsoft\Silverlight)
if %OSVersion%==WXP (rmdir /Q/S "%USERPROFILE%\Local Settings\Application Data\Microsoft\Silverlight")

:EndRemoveROK


REM 5. Check if ROK is installed and install ROK if necessary
REM Set ROK config Directory depending on OS version
if %OSVersion%==W7 (goto SetConfDirW7) else (goto EndSetConfDirW7)
:SetConfDirW7
set ROKConfdirectory=%UserProfile%\AppData\Local\ROK\Configuration\%ROKConfig%
goto EndSetConfDirWXP
:EndSetConfDirW7

if %OSVersion%==WXP (goto SetConfDirWXP) else (goto EndSetConfDirWXP)
:SetConfDirWXP
set ROKConfdirectory="%UserProfile%\Local Settings\Application Data\ROK\Configuration\%ROKConfig%"
:EndSetConfDirWXP

if exist %ROKConfdirectory%(goto EndDeployRokApp) else (goto DeployRokApp)

:DeployRokApp
if %OSVersion%==W7 (mkdir %USERPROFILE%\AppData\Local\ROK\Configuration\)
if %OSVersion%==WXP (mkdir "%USERPROFILE%\Local Settings\Application Data\ROK\Configuration\")

if %OSVersion%==W7 (copy %DeploySource%\%ROKConfig% %USERPROFILE%\AppData\Local\ROK\Configuration\%ROKConfig%)
if %OSVersion%==WXP (copy %DeploySource%\%ROKConfig% "%USERPROFILE%\Local Settings\Application Data\ROK\Configuration\%ROKConfig%")
REM if %OSVersion%==W7(xcopy /Y %DeploySource%\%ROKConfig% %USERPROFILE%\AppData\Local\ROK\Configuration\%ROKConfig%)
REM if %OSVersion%==WXP(xcopy /Y %DeploySource%\%ROKConfig% "%USERPROFILE%\Local Settings\Application Data\ROK\Configuration\%ROKConfig%")

REM copy Webautomation file
if %DownloadSourceFromInternet%==N (goto DownloadWebautoFiles) else (goto EndDownloadWebautoFiles)
:DownloadWebautoFiles
if %OSVersion%==W7 (robocopy %DeployServer%\WebAutoDir %USERPROFILE%\AppData\Local\ROK /E)
if %OSVersion%==WXP (robocopy %DeployServer%\WebAutoDir "%USERPROFILE%\Local Settings\Application Data\ROK" /E)
:EndDownloadWebautoFiles
cd %ProgramFiles%\Microsoft Silverlight\
if %UninstallROK% == N (goto SLlauncherWithoutOverWrite) else (goto SLlauncherWithOverWrite)
:SLlauncherWithoutOverWrite
sllauncher.exe /install:"%DeploySource%\application.xap" /origin:%Origin% /shortcut:desktop+startmenu
goto EndDeployRokApp
:SLlauncherWithOverWrite
sllauncher.exe /install:"%DeploySource%\application.xap" /origin:%Origin% /shortcut:desktop+startmenu /overwrite

:EndDeployRokApp

REM 6. (Distribution server mode) copy ROK xap to distribution server and copy configuration file to distribution server
if %DistributionServerUpdate%==Y (goto CopyROKxapSL) else (goto EndCopyROKxapSL)
:CopyROKxapSL
if %OSVersion%==W7 (cd %USERPROFILE%\AppData\Local\Microsoft\Silverlight\OutOfBrowser\*rok-solution.com)
if %OSVersion%==WXP (cd "%USERPROFILE%\Local Settings\Application data\Microsoft\Silverlight\OutOfBrowser\*rok-solution.com")
xcopy /Y *.xap %DistributionServer%\* 
copy /Y %DeploySource%\%ROKConfig% %DistributionServer%\%ROKConfig%


if %SilverlightNotDeployed%==1 (goto CopySilverlightExe) else (goto EndCopySilverlightExe)
:CopySilverlightExe
del %DownloadSourceFromInternetDir%\downloadSL.vbs
if %OS%==64BIT (goto 32bitDS) else (goto End32bitDS)
:32bitDS
echo strFileURL = "http://download.microsoft.com/download/0/3/E/03EB1393-4F4E-4191-8364-C641FAB20344/50901.00/Silverlight.exe">>%DownloadSourceFromInternetDir%\downloadSL.vbs
set InstallerName=Silverlight.exe /q
set InstallerNameShort=Silverlight.exe
:End32bitDS
if %OS%==32BIT (goto 64bitDS) else (goto End64bitDS)
:64bitDS
echo strFileURL = "http://download.microsoft.com/download/0/3/E/03EB1393-4F4E-4191-8364-C641FAB20344/50901.00/Silverlight_x64.exe">>%DownloadSourceFromInternetDir%\downloadSL.vbs
set InstallerName=Silverlight_x64.exe /q
set InstallerNameShort=Silverlight_x64.exe
:End64bitDS
echo strHDLocation = "%DownloadSourceFromInternetDir%\%InstallerNameShort%">>%DownloadSourceFromInternetDir%\downloadSL.vbs
echo Set objXMLHTTP = CreateObject("MSXML2.ServerXMLHTTP")>>%DownloadSourceFromInternetDir%\downloadSL.vbs
echo objXMLHTTP.open "GET", strFileURL, false>>%DownloadSourceFromInternetDir%\downloadSL.vbs
echo objXMLHTTP.send()>>%DownloadSourceFromInternetDir%\downloadSL.vbs
echo If objXMLHTTP.Status = 200 Then>>%DownloadSourceFromInternetDir%\downloadSL.vbs
echo Set objADOStream = CreateObject("ADODB.Stream")>>%DownloadSourceFromInternetDir%\downloadSL.vbs
echo objADOStream.Open>>%DownloadSourceFromInternetDir%\downloadSL.vbs
echo objADOStream.Type = ^1>>%DownloadSourceFromInternetDir%\downloadSL.vbs
echo objADOStream.Write objXMLHTTP.ResponseBody>>%DownloadSourceFromInternetDir%\downloadSL.vbs
echo objADOStream.Position = ^0>>%DownloadSourceFromInternetDir%\downloadSL.vbs
echo Set objFSO = Createobject("Scripting.FileSystemObject")>>%DownloadSourceFromInternetDir%\downloadSL.vbs
echo If objFSO.Fileexists(strHDLocation) Then objFSO.DeleteFile strHDLocation>>%DownloadSourceFromInternetDir%\downloadSL.vbs
echo Set objFSO = Nothing>>%DownloadSourceFromInternetDir%\downloadSL.vbs
echo objADOStream.SaveToFile strHDLocation>>%DownloadSourceFromInternetDir%\downloadSL.vbs
echo objADOStream.Close>>%DownloadSourceFromInternetDir%\downloadSL.vbs
echo Set objADOStream = Nothing>>%DownloadSourceFromInternetDir%\downloadSL.vbs
echo End if>>%DownloadSourceFromInternetDir%\downloadSL.vbs
echo Set objXMLHTTP = Nothing>>%DownloadSourceFromInternetDir%\downloadSL.vbs
cscript.exe %DownloadSourceFromInternetDir%\downloadSL.vbs


copy /Y %DownloadSourceFromInternetDir%\Silverlight*.* %DistributionServer%\Silverlight*.*
:EndCopySilverlightExe
:EndCopyROKxapSL

Endlocal
