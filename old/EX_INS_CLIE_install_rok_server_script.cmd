setlocal

REM *************************************************************************
REM Environment customization begins here. Modify variables below.
REM *************************************************************************

REM Set DeployServer to the distribution server******************************
set DeployServer=\\192.168.10.82\share\

REM Set LogLocation to a central directory to collect log files.**************
Set LogLocation=\\192.168.10.82\share\SilverlightLogs

REM Set DownloadSourceFromInternet******************************************
REM DownloadSourceFromInternet = Y  the sources are automatically downloaded from the internet
REM DownloadSourceFromInternet = N the sources are downloaded from the distribution server
set DownloadSourceFromInternet=Y

REM Set DistributionServerUpdate (distribution server mode):****************
REM DistributionServerUpdate = Y the ROK sources (application.xap) that is downloaded from the internet are automatically copied to the Distribution server
REM DistributionServerUpdate = N the ROK source that is downloaded from the internet are NOT copied to the Distribution server
REM Note that the other sources (silverlight.exe, configuration file) are NOT copied to the distribution server and must be set manually.
Set DistributionServerUpdate=Y

REM Set configuration file parameters: DatabaseName, UseProxy (true  false), UseSSL (true  false), UseAdfs (true false), IsWindowsAuthenticationEnabled (true false)
Set UseProxy=false
Set UseSSL=true
Set UseAdfs=true
Set IsWindowsAuthenticationEnabled=true
Set Databasename=MECAPLAST

REM Set InstallerName to the name of your copy of the Silverlight installer**
set InstallerName=Silverlight_x64.exe /q

REM Set Origin for ROK (for ROK application update)**************************
Set Origin=https://cloudapp3.rok-solution.com/ClientBin/Bpm.Shell.xap

REM Set the Configuration file for ROK***************************************
Set ROKConfig=cloudapp3.rok-solution.com_443.xml


REM *************************************************************************
REM Deployment code begins here. Do not modify anything below this line.
REM *************************************************************************
Set DistributionServer=%DeployServer%
Set DeploySource=%DeployServer%
Set LogLocationSource=%LogLocation%

REM 1. Check if Silverlight is installed and install Silverlight if necessary
reg query HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Silverlight
set SilverlightDeployed=%errorlevel%
if %SilverlightDeployed%==1 (goto DeploySilverlight) else (goto EndDeploySilverLight)

:DeploySilverlight
if %DownloadSourceFromInternet%==Y (goto DownloadSourceFromInternetmentSL) else (goto EndDownloadSourceFromInternetmentSL)
:DownloadSourceFromInternetmentSL
Set DownloadSourceFromInternetDir=%USERPROFILE%\AppData\Roaming\ROKInstall
mkdir %DownloadSourceFromInternetDir%
del %DownloadSourceFromInternetDir%\downloadSL.vbs

echo strFileURL = "http://download.microsoft.com/download/0/3/E/03EB1393-4F4E-4191-8364-C641FAB20344/50901.00/Silverlight_x64.exe">>%DownloadSourceFromInternetDir%\downloadSL.vbs
echo strHDLocation = "%DownloadSourceFromInternetDir%\Silverlight_x64.exe">>%DownloadSourceFromInternetDir%\downloadSL.vbs
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
Set DownloadSourceFromInternetDir=%USERPROFILE%\AppData\Roaming\ROKInstall
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
REM set DeployServer=%DownloadSourceFromInternetDir%
REM set LogLocation=%DownloadSourceFromInternetDir%
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
sllauncher.exe /uninstall /origin:%Origin%
REM 4.2 Delete remaining ROK files
REM 4.2.1 Delete ROK links in the Start Menu Folder
cd %appdata%\Microsoft\Windows\Start Menu\Programs
del ROK*
REM 4.2.2 Delete ROK links in the desktop
cd %USERPROFILE%\Desktop\
del ROK*
REM 4.2.3 Delete AppLocal folder ROK
rmdir /Q/S %USERPROFILE%\AppData\Local\ROK
rmdir /Q/S %USERPROFILE%\AppData\Local\ROK

REM 4.2.4 Delete AppLocal folder Silverlight
rmdir /Q/S %USERPROFILE%\AppData\Local\Microsoft\Silverlight
rmdir /Q/S %USERPROFILE%\AppData\Local\Microsoft\Silverlight
REM Delete AppLocalLow folder Silverlight
rmdir /Q/S %USERPROFILE%\AppData\LocalLow\Microsoft\Silverlight
rmdir /Q/S %USERPROFILE%\AppData\LocalLow\Microsoft\Silverlight
:EndRemoveROK


REM 5. Check if ROK is installed and install ROK if necessary
if exist %UserProfile%\AppData\Local\ROK\Configuration\*443 (goto EndDeployRokApp) else (goto DeployRokApp)

:DeployRokApp
mkdir %USERPROFILE%\AppData\Local\ROK\Configuration\
xcopy /Y %DeploySource%\*443.xml %USERPROFILE%\AppData\Local\ROK\Configuration\*443.xml
cd %ProgramFiles%\Microsoft Silverlight\
sllauncher.exe /install:"%DeploySource%\application.xap" /origin:%Origin% /shortcut:desktop+startmenu /overwrite
:EndDeployRokApp

REM 6. (Distribution server mode) copy ROK xap to distribution server and copy configuration file to distribution server
if %DistributionServerUpdate%==Y (goto CopyROKxap) else (goto EndCopyROKxap)
:CopyROKxap
cd %USERPROFILE%\AppData\Local\Microsoft\Silverlight\OutOfBrowser\*rok-solution.com
xcopy /Y *.xap %DistributionServer%\* 
copy /Y %DeploySource%\%ROKConfig% %DistributionServer%\%ROKConfig%
:EndCopyROKxap

if %SilverlightDeployed%==1 (goto CopySilverlightExe) else (goto EndCopySilverlightExe)
:CopySilverlightExe
copy /Y %DownloadSourceFromInternetDir%\Silverlight_x64.exe %DistributionServer%\Silverlight_x64.exe
:EndCopySilverlightExe


Endlocal
