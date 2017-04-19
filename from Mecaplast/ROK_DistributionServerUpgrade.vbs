'
' ROK Distribution Server Upgrade
'
'		. check presence of existing files
'		. download files from Internet
'		. compare files version, if more recent then publish it to share folder (DFS)
'
' V 0.1 | XP | 2017/04/07 | Creation
' + TO DO : DISABLING SILVERLIGHT §
'------------------------------------------------------------------------------------------------------
Version = "0.1"
'------------------------------------------------------------------------------------------------------
' VARIABLES
'------------------------------------------------------------------------------------------------------
DeployServer								= "\\mecaplast.com\info-root\package\ROK"
LogLocation									= "\\mecaplast.com\info-root\log\ROK"
DownloadSourceFromInternet	= "Y"
DistributionServerUpdate		= "Y"
'------------------------------------------------------------------------------------------------------
UseProxy=false
UseSSL=true
UseAdfs=true
IsWindowsAuthenticationEnabled=true
Databasename="MECAPLAST"
'------------------------------------------------------------------------------------------------------
InstallerName								= "Silverlight_x64.exe /q"
Origin											= "https://cloudapp3.rok-solution.com/ClientBin/Bpm.Shell.xap"
ROKConfig										= "cloudapp3.rok-solution.com_443.xml"
'----------------------------------------------------------------
Set wshShell								= CreateObject("WScript.Shell")
Set fso											= CreateObject("Scripting.FileSystemObject")
'----------------------------------------------------------------
strComputer									= "."
strUserProfile							= wshShell.ExpandEnvironmentStrings("%USERPROFILE%")				'C:\Users\xponard
Const HKEY_CLASSES_ROOT			= &H80000000
Const HKEY_CURRENT_USER			= &H80000001
Const HKEY_LOCAL_MACHINE		= &H80000002
Const HKEY_USERS						= &H80000003
Const HKEY_CURRENT_CONFIG		= &H80000005
Const REG_SZ								= 1
Const REG_EXPAND_SZ					= 2
Const REG_BINARY						= 3
Const REG_DWORD							= 4
Const REG_MULTI_SZ					= 7
Const ForAppending					= 8
Const ForReading						= 1
Const ForWriting						= 2
'------------------------------------------------------------------------------------------------------
' MAIN CODE
'------------------------------------------------------------------------------------------------------

DistributionServer	= DeployServer
DeploySource				= DeployServer
LogLocationSource		= LogLocation

WScript.Echo
WScript.Echo
WScript.Echo "----------------------------"
WScript.Echo "   ROK INSTALLATION"
WScript.Echo "----------------------------"

'----------------------------------------------------------------
' download sources
'----------------------------------------------------------------
' SILVERLIGHT
'
DownloadSourceFromInternet	= "N"
' Downloading Silverlight
If DownloadSourceFromInternet = "Y" Then
	WSCript.Echo "   .downloading Silverlight..."
	DownloadSourceFromInternetDir = strUserProfile & "\AppData\Roaming\ROKInstall"
	If Not (fso.FolderExists(DownloadSourceFromInternetDir)) Then fso.CreateFolder(DownloadSourceFromInternetDir)

	strFileURL = "http://download.microsoft.com/download/0/3/E/03EB1393-4F4E-4191-8364-C641FAB20344/50901.00/Silverlight_x64.exe"
	strHDLocation = DownloadSourceFromInternetDir & "\Silverlight_x64.exe"
	Set objXMLHTTP = CreateObject("MSXML2.ServerXMLHTTP")
	objXMLHTTP.open "GET", strFileURL, false
	objXMLHTTP.send()
	If objXMLHTTP.Status = 200 Then
		Set objADOStream = CreateObject("ADODB.Stream")
		objADOStream.Open
		objADOStream.Type = 1
		objADOStream.Write objXMLHTTP.ResponseBody
		objADOStream.Position = 0
		Set objFSO = Createobject("Scripting.FileSystemObject")
		If objFSO.Fileexists(strHDLocation) Then objFSO.DeleteFile strHDLocation
		Set objFSO = Nothing
		objADOStream.SaveToFile strHDLocation
		objADOStream.Close
		Set objADOStream = Nothing
	End if
	Set objXMLHTTP = Nothing
End If

DeploySource=DownloadSourceFromInternetDir
LogLocationSource=DownloadSourceFromInternetDir

'installing Silverlight ??? : NO !
'start /wait %DeploySource%\%InstallerName%
'ECHO    .reporting
'echo %date% %time% Setup ended with error code %errorlevel%. >> %LogLocationSource%\%computername%.txt 

'----------------------------------------------------------------
' download sources
'----------------------------------------------------------------
' ROK CLIENT
'
DownloadSourceFromInternet	= "N"
'Downloading ROK
If DownloadSourceFromInternet = "Y" Then
	WSCript.Echo "   .downloading ROK..."
	DownloadSourceFromInternetDir = strUserProfile & "\AppData\Roaming\ROKInstall"
	If Not (fso.FolderExists(DownloadSourceFromInternetDir)) Then fso.CreateFolder(DownloadSourceFromInternetDir)
	
	strFileURL = "https://cloudapp3.rok-solution.com/ClientBin/Bpm.Shell.xap"
	strHDLocation = DownloadSourceFromInternetDir & "\application.xap"
	Set objXMLHTTP = CreateObject("MSXML2.ServerXMLHTTP")
	objXMLHTTP.open "GET", strFileURL, false
	objXMLHTTP.send()
	If objXMLHTTP.Status = 200 Then
		Set objADOStream = CreateObject("ADODB.Stream")
		objADOStream.Open
		objADOStream.Type = 1
		objADOStream.Write objXMLHTTP.ResponseBody
		objADOStream.Position = 0
		Set objFSO = Createobject("Scripting.FileSystemObject")
		If objFSO.Fileexists(strHDLocation) Then objFSO.DeleteFile strHDLocation
		Set objFSO = Nothing
		objADOStream.SaveToFile strHDLocation
		objADOStream.Close
		Set objADOStream = Nothing
	End if
	Set objXMLHTTP = Nothing
	
End If
DownloadSourceFromInternet	= "Y"
DownloadSourceFromInternetDir = strUserProfile & "\AppData\Roaming\ROKInstall"
'----------------------------------------------------------------
' generating configration file if not exists
'----------------------------------------------------------------
' ROKConfig	= "cloudapp3.rok-solution.com_443.xml"
'
configFilePath = DownloadSourceFromInternetDir & "\" & ROKConfig
If Not (fso.Fileexists(ConfigFilePath)) Then
	WScript.Echo "   .creating config xml file"
	'do nothing
	'fso.DeleteFile ConfigFilePath,true
	Set configFileObject = fso.CreateTextFile(configFilePath,True)
	configFileObject.WriteLine("<?xml version=""1.0"" encoding=""utf-8""?>")
	configFileObject.WriteLine("<Configuration xmlns:i=""http://www.w3.org/2001/XMLSchema-instance"" xmlns=""http://www.opserv.fr/2009/11/"">")
	configFileObject.WriteLine("<DatabaseGroups>")
	configFileObject.WriteLine("<DatabaseGroup>")
	configFileObject.WriteLine("<Name>" & Databasename & "</Name>")
	configFileObject.WriteLine("<IsWindowsAuthenticationEnabled>" & IsWindowsAuthenticationEnabled & "</IsWindowsAuthenticationEnabled>")
	configFileObject.WriteLine("<UseProxy>" & UseProxy & "</UseProxy>")
	configFileObject.WriteLine("<UseSsl>" & UseSSL & "</UseSsl>")
	configFileObject.WriteLine("<UseAdfs>" & UseAdfs & "</UseAdfs>")
	configFileObject.WriteLine("</DatabaseGroup>")
	configFileObject.WriteLine("</DatabaseGroups>")
	configFileObject.WriteLine("<AutomaticClientUpdate>Enabled</AutomaticClientUpdate>")
	configFileObject.WriteLine("</Configuration>")
	configFileObject.Close
End If

'----------------------------------------------------------------
' Analyzing files versions
'----------------------------------------------------------------

'----------------------------------------------------------------
'     getting version / size
'----------------------------------------------------------------
'
'Silverlight
'
currentVersionSilverlight = ""
internetVersionSilverlight = ""
If fso.Fileexists(DeployServer & "\Silverlight_x64.exe") Then
	Set objFile = fso.GetFile(DeployServer & "\Silverlight_x64.exe")
	currentVersionSilverlight = fso.GetFileVersion(DeployServer & "\Silverlight_x64.exe")
End If
If fso.Fileexists(DownloadSourceFromInternetDir & "\Silverlight_x64.exe") Then
	Set objFile = fso.GetFile(DownloadSourceFromInternetDir & "\Silverlight_x64.exe")
	internetVersionSilverlight = fso.GetFileVersion(DownloadSourceFromInternetDir & "\Silverlight_x64.exe")
End If
'
'ROK client
'
currentVersionROKclient = ""
internetVersionROKclient = ""
currentSizeROKclient = ""
internetSizeROKclient = ""
If fso.Fileexists(DeployServer & "\application.xap") Then
	Set objFile = fso.GetFile(DeployServer & "\application.xap")
	currentVersionROKclient = fso.GetFileVersion(DeployServer & "\application.xap")										'version not implemented so far... maybe one day
	currentSizeROKclient = objFile.Size																																'getting file size instead
End If
If fso.Fileexists(DownloadSourceFromInternetDir & "\application.xap") Then
	Set objFile = fso.GetFile(DownloadSourceFromInternetDir & "\application.xap")
	internetVersionROKclient = fso.GetFileVersion(DownloadSourceFromInternetDir & "\application.xap")	'version not implemented so far... maybe one day
	internetSizeROKclient = objFile.Size																															'getting file size instead
End If
'
'Config xml file
'
currentDFSSizeConfigXML = ""
currentSizeConfigXML = ""
If fso.Fileexists(DeployServer & "\" & ROKConfig) Then
	Set objFile = fso.GetFile(DeployServer & "\" & ROKConfig)
	currentDFSSizeConfigXML = objFile.Size
End If
If fso.Fileexists(DownloadSourceFromInternetDir & "\" & ROKConfig) Then
	Set objFile = fso.GetFile(DownloadSourceFromInternetDir & "\" & ROKConfig)
	currentSizeConfigXML = objFile.Size
End If


'----------------------------------------------------------------
'      checking versions
'----------------------------------------------------------------
upgradingSilverlight = NeedUpgradeVersion(currentVersionSilverlight,internetVersionSilverlight)
'WScript.Echo "need to upgrade Silverlight = " & upgradingSilverlight
If internetVersionROKclient = "" Then
	upgradingROKclient = NeedUpgradeVersion(currentSizeROKclient,internetSizeROKclient)					' size checking
Else
	upgradingROKclient = NeedUpgradeVersion(currentVersionROKclient,internetVersionROKclient)		' file version checking (not implemented so far..)
End If
'WScript.Echo "need to upgrade ROK = " & upgradingROKclient
upgradingConfigXML = NeedUpgradeVersion(currentDFSSizeConfigXML,currentSizeConfigXML)
'WScript.Echo "need to upgrade config xml = " & upgradingConfigXML
'----------------------------------------------------------------
'     populating
'----------------------------------------------------------------

DownloadSourceFromInternetDir = strUserProfile & "\AppData\Roaming\ROKInstall"
' Silverlight
If upgradingSilverlight = True Then
	'populate DFS
	fso.CopyFile DownloadSourceFromInternetDir & "\Silverlight_x64.exe",DeployServer & "\Silverlight_x64.exe",True
End If

' ROK client
If upgradingROKclient = True Then
	'populate DFS
	fso.CopyFile DownloadSourceFromInternetDir & "\application.xap",DeployServer & "\application.xap",True
End If

' Config file
If upgradingConfigXML = True Then
	'populate DFS
	fso.CopyFile DownloadSourceFromInternetDir & "\" & ROKConfig,DeployServer & "\" & ROKConfig,True
End If





'------------------------------------------------------------------------------------------------------
' FUNCTIONS
'------------------------------------------------------------------------------------------------------


'--------------------------------------------------------------------
' Retourne la version d'un fichier \\strComputer\Path\FileName
' Utilisé avec GetMcAfee()+GetLotus()
'--------------------------------------------------------------------
Function GetFileVersion(strComputer,Path,FileName)
	Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
	If FileName <> "" Then
		Set colFiles = objWMIService.ExecQuery("Select * From CIM_DataFile Where Name = '"&Path&"\\"&FileName&"'")
	Else
		Set colFiles = objWMIService.ExecQuery("Select * From CIM_DataFile Where Name = '"&Path&"'")
	End If	
	If colFiles.Count = 0 Then
	Else
		For Each objFile in colFiles
			GetFileVersion = objFile.Version
		Next
	End If
End Function


'--------------------------------------------------------------------
' Retourn TRUE if test version if more recent or size different than the current one
'--------------------------------------------------------------------
Function NeedUpgradeVersion (currentVersion, testVersion)
	'WScript.Echo "Compairing : " & currentVersion & " - " & testVersion
	If currentVersion = "" or testVersion = "" Then
		'WScript.Echo "empty found"
		NeedUpgradeVersion = True
		Exit Function
	End If
	
	cV = Split(currentVersion,".")
	tV = Split(testVersion,".")
	
	If (UBound(cV) = 0) OR (UBound(tV) = 0) Then
		'testing size not version - no '.' inside version
		'WScript.Echo " size detection mode"
		If tV(0) <> cV(0) Then
			'Wscript.Echo "   > NEED upgrade - different size"
			NeedUpgradeVersion = True
			Exit Function
		Else
			NeedUpgradeVersion = False
			Exit Function
		End If
	End If
	
	cV_major   = CDbl(cV(0))
	cV_minor   = CDbl(cV(1))
	cV_build   = CDbl(cV(2))
	cV_private = CDbl(cV(3))

	tV_major   = CDbl(tV(0))
	tV_minor   = CDbl(tV(1))
	tV_build   = CDbl(tV(2))
	tV_private = CDbl(tV(3))
	
	'WScript.Echo "major test : " & CStr(tV_major) & " <-> " & CStr(cV_major)
	If tV_major > cV_major Then
		'WScript.Echo "   > major release"
		NeedUpgradeVersion = True
		Exit Function
	ElseIf cV_major > tV_major Then
		'WScript.Echo "   > CURRENT major release"
		NeedUpgradeVersion = False
		Exit Function
	End If
	
	'WScript.Echo "minor test : " & CStr(tV_minor) & " <-> " & CStr(cV_minor)
	If tV_minor > cV_minor Then
		'WScript.Echo "   > minor release"
		NeedUpgradeVersion = True
		Exit Function
	ElseIf cV_minor > tV_minor Then
		'WScript.Echo "   > CURRENT minor release"
		NeedUpgradeVersion = False
		Exit Function
	End If
	
	'WScript.Echo "build test : " & CStr(tV_build) & " <-> " & CStr(cV_build)
	If tV_build > cV_build Then
		'WScript.Echo "   > build release"
		NeedUpgradeVersion = True
		Exit Function
	ElseIf cV_build > tV_build Then
		'WScript.Echo "   > CURRENT build release"
		NeedUpgradeVersion = False
		Exit Function
	End If
	
	'WScript.Echo "private test : " & CStr(tV_private) & " <-> " & CStr(cV_private)
	If tV_private > cV_private Then
		'WScript.Echo "   > private release"
		NeedUpgradeVersion = True
		Exit Function
	ElseIf cV_private > tV_private Then
		'WScript.Echo "   > CURRENT private release"
		NeedUpgradeVersion = False
		Exit Function
	End If
	
	'WScript.Echo " == identical file version"
	NeedUpgradeVersion = False
	'WScript.Echo "EoF"
	
End Function