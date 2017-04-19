'
' ROK Client Installation
'
'		. check presence of existing files
'		. download files from DFS
'		. compare files version, if more recent then publish it to share folder (DFS)
'
' V 0.1 | XP | 2017/04/11 | Creation
'------------------------------------------------------------------------------------------------------
Version = "0.1"
'------------------------------------------------------------------------------------------------------
' VARIABLES
'------------------------------------------------------------------------------------------------------
DeployServer								= "\\mecaplast.com\info-root\package\ROK"
LogLocation									= "\\mecaplast.com\info-root\log\ROK"
DownloadSourceFromInternet	= "N"
DistributionServerUpdate		= "N"
'------------------------------------------------------------------------------------------------------
LogIfAlreadyInstalled				= "Y"
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
Set wshNet									= CreateObject("WScript.Network")
Set wshShell								= CreateObject("WScript.Shell")
Set fso											= CreateObject("Scripting.FileSystemObject")
'----------------------------------------------------------------
strComputer									= "."
strComputerName							= GetComputerName()																					'=48NPHXPONARD
strUserProfile							= wshShell.ExpandEnvironmentStrings("%USERPROFILE%")				'C:\Users\xponard [W7] | C:\Documents and Settings\xponard [XP]
OSArchitecture							= GetOSArchitecture()																				'32-bits or 64-bits
OSVersion										= GetOSVersion																							'5.1.2600 / 6.1.7601 / ...
OSVersionTranslate					= GetOSVersionTranslate(OSVersion)													'WXP / W7 /...
OSType											= GetOSType																									'1:Workstation, 3:Server, 5:DC
strUserNameShort						= LCase(wshNet.Username)																		'=xponard
strDomain										= LCase(wshNet.UserDomain)																	'=mecawin
strNetBIOSDomain						= strDomain																									'=mecawin
strUserNameLong							= strDomain & "\" & strUserNameShort												'mecawin\xponard
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
bDEBUG											= False
bDEBUG_VERBOSE							= False
'----------------------------------------------------------------
'strPrgmFdir								= wshShell.ExpandEnvironmentStrings("%programfiles%")	'C:\Program Files
'strPrgmFdirx86							= wshShell.ExpandEnvironmentStrings("%programfiles(x86)%")	'C:\Program Files (x86)

'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion to get ProgramFiles dir path
strPrgmFdir 								= GetRegistryValue(HKEY_LOCAL_MACHINE,"SOFTWARE\Microsoft\Windows\CurrentVersion","ProgramFilesDir")
strPrgmFdirx86 							= GetRegistryValue(HKEY_LOCAL_MACHINE,"SOFTWARE\Microsoft\Windows\CurrentVersion","ProgramFilesDir (x86)")
If strPrgmFdirx86	= "%programfiles(x86)%" Then strPrgmFdirx86 = ""

'------------------------------------------------------------------------------------------------------
' ARGUMENTS HANDLING
'------------------------------------------------------------------------------------------------------
Set objArgs = WScript.Arguments
If WScript.Arguments.Count>0 Then
	If LCase(objArgs(0)) = "debug" Then bDEBUG = True
	If WScript.Arguments.Count>1 Then
		If LCase(objArgs(1)) = "verbose" Then bDEBUG_VERBOSE=True
	End If
End If


'If bDEBUG Then WScript.Echo " DEBUG MODE : ON"
	'For Each strArg in objArgs
	'	WScript.Echo "arg: " & strArg
	'Next
'End If


'------------------------------------------------------------------------------------------------------
' MAIN CODE
'------------------------------------------------------------------------------------------------------

DistributionServer	= DeployServer
DeploySource				= DeployServer
LogLocationSource		= LogLocation
strLogFileName			= LogLocation & "\" & strComputerName & ".log"											'\\mecaplast.com\info-root\log\ROK\48CTXPONARD.log

If bDEBUG Then
	WScript.Echo
	WScript.Echo
	WScript.Echo "----------------------------"
	WScript.Echo "   ROK INSTALLATION v" & Version
	WScript.Echo "----------------------------"
	WScript.Echo "OS            : " & OSArchitecture
	WScript.Echo "OS Version    : " & OSVersion
	WScript.Echo "OS VersionTran: " & OSVersionTranslate
	WScript.Echo "OS Type       : " & OSType
	WScript.Echo "computerName  : " & strComputerName
	WScript.Echo "userName      : " & strUserNameLong
	WScript.Echo "logIfAlready. : " & LogIfAlreadyInstalled
	WScript.Echo "----------------------------"
End If

If LogIfAlreadyInstalled="Y" Then Call AddToLogFile(strLogFileName,"------------------------------------------------------------------------------")

'----------------------------------------------------------------
' SILVERLIGHT DETECTION
'----------------------------------------------------------------
bSilverlightInstalled		= GetRegistryValue(HKEY_LOCAL_MACHINE,"SOFTWARE\Microsoft\Silverlight","Version")
If bSilverlightInstalled = "" Then
	If bDEBUG Then WScript.Echo "Silverlight is missing. Exit."
	Call AddToLogFile(strLogFileName,"Silverlight is missing.")
	' Exit script because silverlight is missing
	WScript.Quit
Else
	If bDEBUG Then WScript.Echo "Silverlight   : " & bSilverlightInstalled
	If LogIfAlreadyInstalled="Y" Then Call AddToLogFile(strLogFileName,"Silverlight detected - v" & bSilverlightInstalled)
End If
'WScript.Echo strUserProfile & "\AppData\Local\ROK\Configuration\"

'----------------------------------------------------------------
' ROK CONFIGURATION FILE
'----------------------------------------------------------------
'WXP (32/64)
If (OSType=1) And (OSVersionTranslate = "WXP") Then
	'WScript.Echo strUserProfile & "\Local Settings\Application Data\ROK\Configuration\" & ROKConfig
	If fso.FileExists(strUserProfile & "\Local Settings\Application Data\ROK\Configuration\" & ROKConfig) Then
		' xml still present
		If LogIfAlreadyInstalled="Y" Then Call AddToLogFile(strLogFileName,strUserProfile & "\Local Settings\Application Data\ROK\Configuration\" & ROKConfig & " - found.")
		If bDEBUG Then WScript.Echo "ROK config    : " & strUserProfile & "\Local Settings\Application Data\Local\ROK\Configuration\" & ROKConfig
	Else
		'xml not present, copying from DFS
		Call AddToLogFile(strLogFileName,strUserProfile & "\Local Settings\Application Data\ROK\Configuration\" & ROKConfig & " - not found.")
		If Not (fso.FolderExists(strUserProfile & "\Local Settings\Application Data\ROK")) Then fso.CreateFolder(strUserProfile & "\Local Settings\Application Data\ROK")
		If Not (fso.FolderExists(strUserProfile & "\Local Settings\Application Data\ROK\Configuration")) Then fso.CreateFolder(strUserProfile & "\Local Settings\Application Data\ROK\Configuration")
		fso.CopyFile DeployServer & "\" & ROKConfig,strUserProfile & "\Local Settings\Application Data\ROK\Configuration\" & ROKConfig,True
		Call AddToLogFile(strLogFileName,"Copying " & ROKConfig & " from DFS.")
		If bDEBUG Then WScript.Echo "ROK config    : copying from DFS."
	End If

' non WXP -> W7/8/10/.. (32/64) - to prevent execution on server, need to implement OSType check her or GPO WMI filter need to be settle.
Else
	If fso.FileExists(strUserProfile & "\AppData\Local\ROK\Configuration\" & ROKConfig) Then
		' xml still present
		If LogIfAlreadyInstalled="Y" Then Call AddToLogFile(strLogFileName,strUserProfile & "\AppData\Local\ROK\Configuration\" & ROKConfig & " - found.")
		If bDEBUG Then WScript.Echo "ROK config    : " & strUserProfile & "\AppData\Local\ROK\Configuration\" & ROKConfig
	Else
		'xml not present, copying from DFS
		Call AddToLogFile(strLogFileName,strUserProfile & "\AppData\Local\ROK\Configuration\" & ROKConfig & " - not found.")
		If Not (fso.FolderExists(strUserProfile & "\AppData\Local\ROK")) Then fso.CreateFolder(strUserProfile & "\AppData\Local\ROK")
		If Not (fso.FolderExists(strUserProfile & "\AppData\Local\ROK\Configuration")) Then fso.CreateFolder(strUserProfile & "\AppData\Local\ROK\Configuration")
		fso.CopyFile DeployServer & "\" & ROKConfig,strUserProfile & "\AppData\Local\ROK\Configuration\" & ROKConfig,True
		Call AddToLogFile(strLogFileName,"Copying " & ROKConfig & " from DFS.")
		If bDEBUG Then WScript.Echo "ROK config    : copying from DFS."
	End If
End If

'----------------------------------------------------------------
' ROK APPLICATION INSTALLATION
'----------------------------------------------------------------
ROKUninstallKey = SearchRegistryValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Uninstall","DisplayName",REG_SZ,"ROK")
If ROKUninstallKey = "" Then
	'--------------------------------------------
	'Need ROK installation
	If bDEBUG Then WScript.Echo "ROK not installed."
	If fso.FileExists(strPrgmFdir & "\Microsoft Silverlight\sllauncher.exe") Then
		'InstallSoft """" & strPrgmFdir & "\Microsoft Silverlight\sllauncher.exe" & """" & " /install:" & DeployServer & "\application.xap /origin:" & Origin & " /shortcut:desktop+startmenu /overwrite", "application.xap installation v1"	
		If bDEBUG Then WScript.Echo "ROK client    : installing from DFS."
		InstallSoftv2 """" & strPrgmFdir & "\Microsoft Silverlight\sllauncher.exe" & """" & " /install:" & DeployServer & "\application.xap /origin:" & Origin & " /shortcut:desktop+startmenu /overwrite", "application.xap installation v2"	
		Call AddToLogFile(strLogFileName,"Installing ROK application from DFS.")
		WScript.Sleep 20000
		Call AddToLogFile(strLogFileName,"ROK installed : HKCU\" & SearchRegistryValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Uninstall","DisplayName",REG_SZ,"ROK"))
		If bDEBUG Then WScript.Echo "ROK client    > HKCU\" & SearchRegistryValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Uninstall","DisplayName",REG_SZ,"ROK")
	Else
		Call AddToLogFile(strLogFileName,"trying installing ROK failed because sllauncher not found at : " & strPrgmFdir & "\Microsoft Silverlight\sllauncher.exe")
		If bDEBUG Then WScript.Echo "ROK client    : install failed due to missing local Silverlight-sllauncher."
	End If
Else
	'--------------------------------------------
	'No need ROK installation.
	If bDEBUG Then WScript.Echo "ROK client    : True"
	If bDEBUG_VERBOSE Then WScript.Echo "              > HKCU\" & ROKUninstallKey
	If LogIfAlreadyInstalled="Y" Then Call AddToLogFile(strLogFileName,"ROK already installed - HKCU\" & ROKUninstallKey)
End If

If bDEBUG Then WScript.Echo "----------------------------"

'End Of Main Script
'======================================================================================================

'------------------------------------------------------------------------------------------------------
' FUNCTIONS
'------------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------
' Registre, Affiche la clés d'une source
' Utilisé avec GetMcAfee()+GetLotus()+GetBOFC()
'--------------------------------------------------------------------
Function GetRegistryValue(Root,strKeyPath,strKeyName)
	On Error Resume Next
	Set oReg=GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
	oReg.EnumValues Root,strKeyPath,arrValueNames,arrValueTypes

	If strKeyName="" Then
		oReg.GetStringValue Root,strKeyPath,strKeyName,strGetValue
		GetRegistryValue = strGetValue
		Exit Function
	End If
	
	For k=0 To UBound(arrValueNames)
		If arrValueNames(k) = strKeyName Then
			Select Case arrValueTypes(k)
					Case REG_SZ
							'ValueType = "REG_SZ"
							oReg.GetStringValue Root, strKeyPath,arrValueNames(k), strGetValue
							GetRegistryValue = strGetValue
					Case REG_EXPAND_SZ
							'ValueType = "REG_EXPAND_SZ"
							oReg.GetExpandedStringValue Root,strKeyPath,arrValueNames(k),strGetValue
							GetRegistryValue = strGetValue
					Case REG_BINARY
							'ValueType = "REG_BINARY"
							oReg.GetBinaryValue Root,strKeyPath,arrValueNames(k),strGetValue
							For j = lBound(strGetValue) to uBound(strGetValue)
								GetRegistryValue = strGetValue(j)
							Next
					Case REG_DWORD
							'ValueType = "REG_DWORD"
							oReg.GetDWORDValue Root,strKeyPath,arrValueNames(k),strGetValue
							GetRegistryValue = strGetValue
					Case REG_MULTI_SZ
							'ValueType = "REG_MULTI_SZ"
							oReg.GetMultiStringValue Root,strKeyPath,arrValueNames(k),arrValues
							For Each strGetValue In arrValues
								GetRegistryValue = strGetValue
							Next
			End Select
		End If
	Next
End Function
'--------------------------------------------------------------------


'--------------------------------------------------------------------
Sub InstallSoft(strCommand,Label)
	On Error Resume Next
	If bDEBUG Then 
		If Label<>"" Then WScript.Echo " · "&Label & " : " & strCommand
	End If
	wshShell.Run strCommand, 0, True
End Sub
'--------------------------------------------------------------------


'--------------------------------------------------------------------
Sub InstallSoftv2(strCommand,Label)
	On Error Resume Next
	If bDEBUG Then 
		If Label<>"" Then WScript.Echo " · "&Label & " : " & strCommand
	End If
	wshShell.Run WshShell.Environment("PROCESS")("COMSPEC") & " /c """ & strCommand & """", 0
End Sub
'--------------------------------------------------------------------


'--------------------------------------------------------------------
Function GetComputerName()
 	On Error Resume Next
	Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	Set colComputers = objWMIService.ExecQuery ("Select * from Win32_ComputerSystem")
	For Each objComputer in colComputers 
		GetComputerName = objComputer.Name
	Next
End Function
'--------------------------------------------------------------------


'--------------------------------------------------------------------
' Retourne la date/heure
'	GetDateTime(Value)
'		Value:
'		1:DATE			:	DD/MM/YYYY
'		2:TIME			:	hh:mm:ss (return GMT+1 format whatever LocaleTimeZone)
'--------------------------------------------------------------------
Function GetDateTime(Value)
	Select Case Value
		'Date
		Case 1
			strDay = Day(Date)
			If Len(strDay) = 1 Then strDay="0"+CStr(strDay)
			strMonth = Month(Date)
			If Len(strMonth) = 1 Then strMonth="0"+CStr(strMonth)
			strYear = Year(Date)
			'GetDateTime = strDay&"/"&strMonth&"/"&strYear
			'GetDateTime = strMonth&"/"&strDay&"/"&strYear
			GetDateTime = strYear&"/"&strMonth&"/"&strDay
			
		'Time
		Case 2
			strHour = Hour(Time)
			If Len(strHour) = 1 Then strHour="0"+CStr(strHour)
			strMinute	= Minute(Time)
			If Len(strMinute) = 1 Then strMinute="0"+CStr(strMinute)
			strSecond	= Second(Time)
			If Len(strSecond) = 1 Then strSecond="0"+CStr(strSecond)
			Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
			Set colItems = objWMIService.ExecQuery("Select * from Win32_TimeZone")	'Find Bias from WMI (London=0, Paris=60, Istambul=120...)
			For Each objItem in colItems
				strBias = objItem.Bias
			Next
			strDateTime = strHour&":"&strMinute&":"&strSecond
			strBias = strBias-60															'Set GMT+1 (+60) the default reference, so set time -60 to put it the reference
			strDateTime = DateAdd("n",-strBias,strDateTime)		'Remove Bias from current local time
			GetDateTime = strDateTime
	End Select
End Function
'--------------------------------------------------------------------


'--------------------------------------------------------------------
' Add a line to a log file (DFS)
'--------------------------------------------------------------------
Sub AddToLogFile(strFilePath,strMessage)
	On Error Resume Next
	Set configFileObject = fso.OpenTextFile(strFilePath,ForAppending,True)
	configFileObject.WriteLine(GetDateTime(1) & " - " & GetDateTime(2) & " - Client Install - " & OSArchitecture & " - " & strUserNameLong & " - " & strMessage)
	configFileObject.Close
End Sub
'--------------------------------------------------------------------


'--------------------------------------------------------------------
' Registre, recherche la présence d'une clé avec une valeur précise à partir d'une racine (search N-1 + N-2 uniquement)
' Retourne la clé 
' SearchRegistryValue ("HKLM", "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall","DisplayName",REG_SZ,"SAP BusinessObjects Financial Consolidation")
' > SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{38146007-0098-4D02-AAA3-E292D888BB65}\
'--------------------------------------------------------------------
Function SearchRegistryValue(Root,strKeyPath,strKeyName,strKeyType,strKeyValue)
	On Error Resume Next
	Set oReg=GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
	oReg.EnumKey Root,strKeyPath,arrValueNames,arrValueTypes
	' Loop through each key
	For Each sKey In arrValueNames
		' Get all subkeys within the key 'sKey'
		oReg.EnumKey Root, strKeyPath & "\" & sKey, aSubToo
		For Each sKeyToo In aSubToo
			Select Case strKeyType
				Case REG_SZ
						oReg.GetStringValue Root, strKeyPath & "\" & sKey & "\" & sKeyToo, strKeyName, strGetValue
						If strGetValue = strKeyValue Then
							SearchRegistryValue = strKeyPath & "\" & sKey & "\" & sKeyToo
							Exit Function
						End If
				Case REG_EXPAND_SZ
						oReg.GetExpandedStringValue Root,strKeyPath & "\" & sKey & "\" & sKeyToo,strKeyName,strGetValue
						If strGetValue = strKeyValue Then
							SearchRegistryValue = strKeyPath & "\" & sKey & "\" & sKeyToo
							Exit Function
						End If
'				Case REG_BINARY
'						oReg.GetBinaryValue Root,strKeyPath,arrValueNames(k),strGetValue
'						'For j = lBound(strGetValue) to uBound(strGetValue)
'						'	GetRegistryValue = strGetValue(j)
'						'Next
				Case REG_DWORD
						oReg.GetDWORDValue Root, strKeyPath & "\" & sKey & "\" & sKeyToo, strKeyName, strGetValue
						If strGetValue = strKeyValue Then
							SearchRegistryValue = strKeyPath & "\" & sKey & "\" & sKeyToo
							Exit Function
						End If
'				Case REG_MULTI_SZ
'						oReg.GetMultiStringValue Root,strKeyPath,arrValueNames(k),arrValues
'						'For Each strGetValue In arrValues
'						'	GetRegistryValue = strGetValue
'						'Next
			End Select
		Next
		oReg.Nothing
	Next
	SearchRegistryValue = ""
End Function
'--------------------------------------------------------------------


'--------------------------------------------------------------------
' Return '64-bits' or '32-bits' depending on OS Architecture
'--------------------------------------------------------------------
Function GetOSArchitecture()
	On Error Resume Next
	Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
  OSArchitecture = GetRegistryValue(HKEY_LOCAL_MACHINE,"SYSTEM\CurrentControlSet\Control\Session Manager\Environment","PROCESSOR_ARCHITECTURE")
	If OSArchitecture="AMD64" Then GetOSArchitecture = "64-bits"
	If OSArchitecture="x86" Then GetOSArchitecture = "32-bits"
End Function
'--------------------------------------------------------------------


'--------------------------------------------------------------------
' Return OS Version (6.1.7601, ...)
'--------------------------------------------------------------------
Function GetOSVersion()
	On Error Resume Next
	Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
	Set colItems = objWMIService.ExecQuery ("Select * from Win32_OperatingSystem")
	For Each objItem in colItems
		OSVersion = objItem.Version
	Next	
	GetOSVersion = OSVersion
End Function
'--------------------------------------------------------------------


'--------------------------------------------------------------------
' Return OS Version Named (WXP, ...)
'--------------------------------------------------------------------
Function GetOSVersionTranslate(Value)
	On Error Resume Next
	Select Case Value
		Case "5.0.2195" GetOSVersionTranslate="W2000"
		Case "5.1.2600","5.2.3790" GetOSVersionTranslate="WXP"
		Case "6.0.6002" GetOSVersionTranslate="WVista"
		Case "6.1.7600","6.1.7601"	GetOSVersionTranslate="W7"
		Case "6.2.9200" GetOSVersionTranslate="W8"
		Case "6.3.9600" GetOSVersionTranslate="W8.1"
		Case "10.0.10240","10.0.10586","10.0.14393","10.0.15063" GetOSVersionTranslate="W10"
		Case Else GetOSVersionTranslate=Value
	End Select
End Function
'--------------------------------------------------------------------


'--------------------------------------------------------------------
' Return number corresponding to OSType
'--------------------------------------------------------------------
'0 (0x0) Standalone Workstation
'1 (0x1) Member Workstation 
'2 (0x2) Standalone Server
'3 (0x3) Member Server 
'4 (0x4) Backup Domain Controller
'5 (0x5) Primary Domain Controller
Function GetOSType()
	On Error Resume Next
	Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
	Set colItems = objWMIService.ExecQuery ("Select * from Win32_OperatingSystem")
	For Each objItem in colItems
		OSType = objItem.ProductType
	Next	
	GetOSType = OSType
End Function
'--------------------------------------------------------------------

'End Of Script