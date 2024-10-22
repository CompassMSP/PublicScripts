Option Explicit
' CONST section
Const Version = "1.0.1" ' 01.06.2018
' checkMSSQL2 has been added for testing
' added grabLocalMembership 
' getDomainUsers
' getDomainGroups

Const OutputFolder = "Inventory"
Const maxThreadCount = 6
Const maxTargetCPUload = 50
Const maxHostCPUload = 90
Const maxRunningTime = 4320 '3 days = 3 * 24 * 60
Const maxRunningPasses = 3
Const HKCU = &H80000001 'HKEY_CURRENT_USER
Const HKLM = &H80000002 'HKEY_LOCAL_MACHINE
Const HKU  = &H80000003 'HKEY_USERS
Const fileOverwrite = 2
Const fileAppend = 8
Const fileDelimiter = "."

' VARIABLE section
Dim cmdArgs: Set cmdArgs = CreateObject("Scripting.Dictionary")
	' All recognized commandline options
	cmdArgs.Add "/PROCESSLOCAL", False
	cmdArgs.Add "/PROCESSLIST", False
	cmdArgs.Add "/GETLISTFROMAD", False
	cmdArgs.Add "/PROCESSAD", False
	cmdArgs.Add "/PROCESSTARGET:", False
	cmdArgs.Add "/GETDOMAINUSERS", False
	cmdArgs.Add "/GETDOMAINGROUPS", False
Dim domainUserInfoDict: Set domainUserInfoDict = Nothing
Dim oldShell: oldShell = True
Dim objFSO
Dim objDSE
Dim objWMIDNS
Dim Context
Dim objLog
Dim objWMI: objWMI = False
Dim objWMISQL
Dim objRegistry
Dim objShell	
Dim isLocalMode 
Dim isDomainFound
Dim FullPath: FullPath = ""
Dim OutputPath: OutputPath = ""
Dim objToScan: objToScan = "False"
Dim onlyMS: onlyMS = False
Dim domainRole: domainRole = 1
Dim isServer: isServer = False

sub initLog
	On Error Resume Next
	Set objLog = objFSO.OpenTextFile(OutputPath & objToScan & fileDelimiter & "log", fileAppend, True) 	
End Sub

Sub initFSO
	On Error Resume Next
	Err.clear
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	If Err<>0 Then: Call SysLog("FSO",Err.number,Err.description,Err.source): End If
End Sub

Sub initShell
	On Error Resume Next
	Err.clear
	Set objShell = CreateObject("Wscript.Shell")
	If Err<>0 Then: Call SysLog("Shell",Err.number,Err.description,Err.source): End If
End Sub

Sub initWMIDNS
	On Error Resume Next
	objWMIDNS = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & objToScan & "\root\microsoftdns")
	Call LogError("[I] initWMIDNS", 0)
End Sub

Sub initWMISQL
	On Error Resume Next
	Set objWMISQL = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & objToScan & "\root\Microsoft\SqlServer\ComputerManagement")
	Call LogError("[I] initWMISQL: "&"winmgmts:{impersonationLevel=impersonate}!\\" & objToScan & "\root\Microsoft\SqlServer\ComputerManagement", 0)
	If not isObject(objWMISQL) Then
		Set objWMISQL = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & objToScan & "\root\Microsoft\SqlServer\ComputerManagement10")
		Call LogError("[I] initWMISQL: "&"winmgmts:{impersonationLevel=impersonate}!\\" & objToScan & "\root\Microsoft\SqlServer\ComputerManagement10", 0)
		If not isObject(objWMISQL) Then
			Set objWMISQL = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & objToScan & "\root\Microsoft\SqlServer\ComputerManagement11")
			Call LogError("[I] initWMISQL: "&"winmgmts:{impersonationLevel=impersonate}!\\" & objToScan & "\root\Microsoft\SqlServer\ComputerManagement11", 0)
		End If
	End If
End Sub

Sub initDSE
	On Error Resume Next
	Set objDSE = GetObject("LDAP://RootDSE")
	Call LogError("[S] initDSE...getting object", 0)
	Context = objDSE.Get("DefaultNamingContext")
	If Trim(Context) = "" Then
		Call LogError("[-] Error. Unable to identify a domain", 0)
		Wscript.Echo "[-] Error. Unable to identify a domain"
		isDomainFound = False
	Else
		isDomainFound = True
	End If
	Call LogError("[D] initDSE", 0)
End Sub

Sub initWMI
	On Error Resume Next
	Call LogError("[S] initWMI", 0)
	Dim number, description, source
	Set objWMI = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & objToScan & "\root\cimv2")
	number = Err.Number
	description = Err.Description
	source = Err.Source
	Call LogError("[I] objWMI:winmgmts:{impersonationLevel=impersonate}!\\" & objToScan & "\root\cimv2", 1)
	If not isObject(objWMI) Then
		Call SysLog("WMI",number,description,source)
	Else
		Call createFile("online")
		Set objRegistry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & objToScan & "\root\default:StdRegProv")
		number = Err.Number
		description = Err.Description
		source = Err.Source
		Call LogError("[I] objRegistry:winmgmts:{impersonationLevel=impersonate}!\\" & objToScan & "\root\default:StdRegProv", 0)
		If not isObject(objRegistry) Then
			Call SysLog("REGISTRY",number,description,source)
			Set objRegistry = GetObject("winmgmts:{impersonationLevel=impersonate,authenticationLevel=pktPrivacy}!\\" & objToScan & "\root\default:StdRegProv")
			number = Err.Number
			description = Err.Description
			source = Err.Source
			Call LogError("[I] objRegistry:winmgmts:{impersonationLevel=impersonate,authenticationLevel=pktPrivacy}!\\" & objToScan & "\root\default:StdRegProv", 0)
			If not isObject(objRegistry) Then
				Call SysLog("REGISTRY",number,description,source)
			End If
		End If	
	End If
	Call LogError("[D] initWMI", 0)
End Sub

sub closeLog
	On Error Resume Next
	Err.Clear
	objLog.close
end sub

sub LogError(message, errOnly)
	if errOnly <> 0 then
		if typename(Err) = "Object" then
			objLog.WriteLine CSVLineFromArray(array(Now, objToScan, message, ""&Err.Number, ""&Replace(Replace(Replace(Err.Description, vbCrLf, ""),vbCr,""),vbLf,"")))
			Err.Clear
		end if
	else
		objLog.WriteLine CSVLineFromArray(array(Now, objToScan, message, "", ""))
	end if
	
end sub

Function LPad(str, l)
	On Error Resume Next
	LPad = string(l-Len(str)," ") & str
End Function

Function Ping(object)
	On Error Resume Next
	If isLocalMode Then: Ping = True: Exit Function: End If
	With WMIQuery(GetObject("winmgmts:{impersonationLevel=impersonate}"), "Select * from Win32_PingStatus Where address='" & object & "'")
		Call LogError("[I] Pinging", 0)
		If .StatusCode = 0 Then: Ping = True: Else: Ping = False: End If
	End With
End Function

Sub getDomainRole
	On Error Resume Next
	Call LogError("[S] getDomainRole", 0)
	With WMIQuery(objWMI, "Select DomainRole from Win32_ComputerSystem")
		domainRole = .DomainRole
		Call LogError("[I] getDomainRole: " & domainRole, 0)
	End With
	Call LogError("[D] getDomainRole", 0)
End Sub

Function WMIQuery(objWMI, query)
	On Error Resume Next
	Call LogError("[S] WMIQuery", 0)
	Set WMIQuery = Nothing
	If Not IsNull(objWMI) Then
		Dim q,r: Set q = objWMI.ExecQuery(query)
		Call LogError("[I] Query: "&query, 0)
		If isObject(q) Then
			If q.Count > 0 Then: For Each r In q: Set WMIQuery = r: Exit Function: Next: End If
		End If
	' 11010 - request timeout
	End If
	Call LogError("[D] WMIQuery", 0)
End Function

Function filterMS(str)
	On Error Resume Next
	filterMS = True
	Exit Function
	
	'filterMS = False
	'Dim s,ustr: ustr = UCase(Trim(""&str))
	'If ustr="" Then: Exit Function: End If
	'For Each s In Array("MICROSOFT", "VISUAL STUDIO", "SYSTEM CENTER", "EXCHANGE SERVER", "SYSMANSMS", "MS OFFICE", "OFFICE SYSTEM", "VISUAL BASIC", "VISUAL FOXPRO", "SQL SERVER", "SHAREPOINT", "VMWARE", "CITRIX")
	'	If InStr(ustr, s) > 0 Then
	'		filterMS = True
	'		Exit Function
	'	End If
	'Next
End Function

Function getDomainUserInfo(user)
	On Error Resume Next
	Call LogError("[S] Getting domain user info" & user, 0)
	Wscript.Echo "[S] Getting domain user info for " & user & "..."	
	If domainUserInfoDict Is Nothing Then: Set domainUserInfoDict = CreateObject("Scripting.Dictionary"): End If
	If domainUserInfoDict.Exists(user) Then
		getDomainUserInfo = domainUserInfoDict.Item(path)
		Exit Function
	End If
	With GetObject("LDAP://" & user)
		getDomainUserInfo = CSVLineFromArray(Array( _
			""&.sAMAccountName, _
			""&.Description, _
			""&.Class, _
			""&.DisplayName, _
			""&.GivenName, _ 
			""&.SN, _
			""&.PasswordLastChanged, _
			DateDiff("d", IntToDate(.Get("lastLogonTimestamp")), Now), _
			""&.PasswordRequired, _
			""&.userAccountControl, _
			""&.DistinguishedName, _
			""&.WhenCreated, _ 
			""&.WhenChanged _
		))		
	End With
	Call LogError("[I] LDAP Object : LDAP://" & user, 1)
	Call domainUserInfoDict.Add(user, getDomainUserInfo)
	Wscript.Echo "[D] Getting domain user info..."
	Call LogError("[D] Getting domain user info...", 0)
End Function

Sub GetDomainUsers
	On Error Resume Next
	Call LogError("[S] Getting domain users...", 0)
	Wscript.Echo "[S] Getting domain users..."
	Dim File: Set File = objFSO.CreateTextFile(OutputPath & "DomainUsers.txt", True, True)
	' ADODB connection
	Dim objADOConn: Set objADOConn = CreateObject("ADODB.Connection")
	Dim objADOCmd:  Set objADOCmd = CreateObject("ADODB.Command")
	objADOConn.Provider = "ADsDSOObject"
	objADOConn.Open "Active Directory Provider"
	Set objADOCmd.ActiveConnection = objADOConn
	objADOCmd.CommandText = "<LDAP://" & Context & ">;(&(objectCategory=person)(objectClass=user));distinguishedName;subtree"
	objADOCmd.Properties("Page Size") = 1000
	objADOCmd.Properties("Timeout") = 25
	objADOCmd.Properties("Cache Results") = False
	Dim adoRS: Set adoRS = objADOCmd.Execute
	Call LogError("[I] <LDAP://" & Context & ">;(&(objectCategory=person)(objectClass=user));distinguishedName;subtree", 1)
	' Users
	Do Until adoRS.EOF
		File.WriteLine getDomainUserInfo(adoRS.Fields("DistinguishedName").Value)
		adoRS.MoveNext
	Loop
	adoRS.Close
	File.Close
	Wscript.Echo "[D] Getting domain users..."
	Call LogError("[D] Getting domain users...", 0)
End Sub

Function getDomainGroupInfo(group)
	On Error Resume Next
	Call LogError("[S] Getting domain group info for " & group, 0)
	Wscript.Echo "[S] Getting domain group info for " & group	
	With GetObject("LDAP://" & group)
		getDomainGroupInfo = CSVLineFromArray(Array(""&.sAMAccountName, ""&.Description, ""&.Class, ""&.GroupType, ""&.DistinguishedName, ""&.WhenCreated, ""&.WhenChanged))
	End With
	Call LogError("[I] LDAP://" & group, 1)
	Wscript.Echo "[D] Getting domain group info for " & group	
	Call LogError("[D] Getting domain group info for " & group, 0)
End Function

Sub getDomainGroupMembers(group, groups, strGroupInfo, File)
	On Error Resume Next
	Call LogError("[S] Getting domain group members...", 0)
	Wscript.Echo "[S] Getting domain group members..."	
	Dim o, q: Set q = GetObject("LDAP://" & group)
	q.filter("group")
	Call LogError("[I] LDAP object : LDAP://" & group, 1)
	For Each o In q.Members
		With o
			If UCase(.Class) = "GROUP" Then
				Dim ADsPath: ADsPath = Replace(.ADsPath,"LDAP://","")
				If Not groups.Exists(ADsPath) Then
					wscript.echo strGroupInfo
					groups.Add ADsPath, 1
					Call getDomainGroupMembers(ADsPath, groups, strGroupInfo, File)
				End IF
			'ElseIf UCase(.Class) = "USER" Then
			'	File.WriteLine strGroupInfo & getDomainUserInfo(.DistinguishedName)
			End If
		End With
	Next
	Wscript.Echo "[D] Getting domain group members..."	
	Call LogError("[D] Getting domain group members...", 0)
End Sub

Sub getDomainGroups
	On Error Resume Next
	Call LogError("[S] Getting domain groups...", 0)
	Wscript.Echo "[S] Getting domain groups..."	
	Dim File: Set File = objFSO.CreateTextFile(OutputPath & "DomainGroups.txt", True, True)
	' ADODB connection
	Dim objADOConn: Set objADOConn = CreateObject("ADODB.Connection")
	Dim objADOCmd:  Set objADOCmd = CreateObject("ADODB.Command")
	objADOConn.Provider = "ADsDSOObject"
	objADOConn.Open "Active Directory Provider"
	Set objADOCmd.ActiveConnection = objADOConn
	objADOCmd.CommandText = "<LDAP://" & Context & ">;(&(objectCategory=Group));DistinguishedName;subtree"
	objADOCmd.Properties("Page Size") = 1000
	objADOCmd.Properties("Timeout") = 25
	objADOCmd.Properties("Cache Results") = False
	Dim adoRS: Set adoRS = objADOCmd.Execute
	Call LogError("[I] <LDAP://" & Context & ">;(&(objectCategory=Group));DistinguishedName;subtree", 1)
	' Groups
	Dim groups: set groups = CreateObject("Scripting.Dictionary")
	Do Until adoRS.EOF
		Dim strDN: strDN = adoRS.Fields("DistinguishedName").Value
		Dim strGI: strGI = getDomainGroupInfo(strDN)
		groups.RemoveAll
		
		groups.Add strDN, 1
		File.WriteLine strGI
	
		'Call getDomainGroupMembers(strDN, groups, strGI, File)
		adoRS.MoveNext
	Loop
	adoRS.Close
	File.Close
	Wscript.Echo "[D] Getting domain groups..."	
	Call LogError("[D] Getting domain groups...", 0)
End Sub

function grabSharePoint() 
	On Error Resume Next
	grabSharePoint = False
	Call LogError("[S] Checking for SharePoint...", 0)
	Wscript.Echo "[S] Checking for SharePoint..."
	Dim File: Set File = CreateFile("sharepoint")
	Dim tmpPath: tmpPath = "SOFTWARE\Microsoft\Shared Tools\Web Server Extensions"
	Dim subkeys, subKey, guid, role, dsn, keys, key
	Set objRegistry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & objToScan & "\root\default:StdRegProv")
	Call LogError("[I] object: winmgmts:{impersonationLevel=impersonate}!\\" & objToScan & "\root\default:StdRegProv", 1)
	If objRegistry.EnumKey(HKLM,tmpPath,keys) = 0 then
		For Each key In keys
			Call objRegistry.GetStringValue(HKLM,"SOFTWARE\Microsoft\Shared Tools\Web Server Extensions\"&key&"\WSS\","ServerRole", role)
		    Call objRegistry.GetStringValue(HKLM,"SOFTWARE\Microsoft\Shared Tools\Web Server Extensions\"&key&"\Secure\ConfigDB","dsn", dsn)
			If objRegistry.EnumValues (HKLM,tmpPath&"\"&key&"\WSS\InstalledProducts\",subKeys) = 0 Then
				For Each subKey In subkeys
			    	If objRegistry.GetStringValue(HKLM,"SOFTWARE\Microsoft\Shared Tools\Web Server Extensions\"&key&"\WSS\InstalledProducts\",subKey, guid) = 0 Then
			    		File.WriteLine CSVLineFromArray(Array(objToScan, svcName, state, key, ""&role, ""&dsn, ""&subKey, ""&guid))
						grabSharePoint = True		
					End If
				Next
			Else
				File.WriteLine CSVLineFromArray(Array(objToScan, svcName, state, key, ""&role, ""&dsn, "noSUB", "noGUID"))
				grabSharePoint = True
			End If
		Next
	End If	
	File.Close
	If grabSharePoint = False Then
		objFSO.DeleteFile(OutputPath & objToScan & fileDelimiter & "sharepoint")
	End If
	Wscript.Echo "[D] Checking for SharePoint..."
	Call LogError("[D] Checking for SharePoint...", 0)
End Function

Function grabGUID() 
	On Error Resume Next
	Call LogError("[S] Collecting data on installed software...", 0)
	Dim strRegIdentityCode : strRegIdentityCode = Array("SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", _
														"SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\")
	Dim strIdentityCode : strIdentityCode = ""
	Dim arrIdentityCode, strLine
	Dim outArr(37)
	Dim arrRegKey: arrRegKey = Array(_
	   "objToScan", _
	   "reserved", _
	   "strIdentityCode", _ 
	   "DisplayName",_
	   "DisplayVersion",_
	   "Version",_
	   "VersionMajor",_
	   "VersionMinor",_
	   "ReleaseType",_
	   "Publisher",_
	   "ProductID",_
	   "ParentKeyName",_
	   "ParentDisplayName",_
	   "InstallDate",_
	   "UninstallString",_
	   "QuietUninstallString",_
	   "ModifyPath",_
	   "InstallLocation",_
	   "InstallPath",_
	   "InstallSource",_
	   "LogFile",_
	   "DisplayIcon",_
	   "RegistryLocation",_
	   "Size",_
	   "EstimatedSize",_
	   "SystemComponent",_
	   "NoRepair",_
	   "NoModify",_
	   "NoRemove",_
	   "WindowsInstaller",_
	   "TSAware",_
	   "Comments",_
	   "Readme",_
	   "URLInfoAbout",_
	   "URLUpdateInfo",_
	   "HelpLink",_
	   "HelpTelephone",_
	   "Contact")
	Dim deleteFile : deleteFile = True
	Dim File: Set File = CreateFile("guid")
	outArr(0) = objToScan
	outArr(1) = ""
	Dim i: For i = LBound(strRegIdentityCode) to UBound(strRegIdentityCode)
		objRegistry.EnumKey HKLM, strRegIdentityCode(i), arrIdentityCode
		Call LogError("[I] Enum in " & strRegIdentityCode(i), 1)
		For Each strIdentityCode In arrIdentityCode
			outArr(2) = strIdentityCode
			strLine = fncRegKeyValue("HKLM", strRegIdentityCode(i) & "\" & strIdentityCode, arrRegKey, outArr)
			Call LogError("[I] fncRegKeyValue : " & strIdentityCode, 1) 
			If Len(strLine) > 0  And filterMS(strLine) Then
				File.WriteLine CSVLineFromArray(outArr)
				deleteFile = False
			End If
		Next
	Next		
	File.close
	If deleteFile = True Then
		objFSO.DeleteFile(OutputPath & objToScan & fileDelimiter & "guid")
	End If
	grabGUID = True
	Call LogError("[D] Collecting data on installed software...", 0)
End Function

Function fncRegKeyValue(strRegHive, strRegPath, arrRegValue, outArr)
	On Error Resume Next
	Dim strRegKey, strRecord, varRegValue, strFieldName, strFormula, intIndex, strField, bolFilter, strFilter
	strRegKey = strRegHive & "\" & strRegPath & "\"
	strRecord = ""
	Dim j: For j = 3 to 37
		varRegValue = arrRegValue(j)
		If IsArray(varRegValue) Then
			strFieldName = varRegValue(0)
			If strRegHive = "" Then
				strFormula = "strField = strFieldName"
			Else
				strFormula = varRegValue(2)
				If IsArray(varRegValue(1)) Then
					strFormula = Replace(strFormula, "?", "strField")
					For intIndex = UBound(varRegValue(1)) To 0 Step -1
						strFormula = Replace(strFormula, "~" & intIndex, "objShell.RegRead(strRegKey & """ & varRegValue(1)(intIndex) & """)")
					Next
				Else
					strFormula = Replace(strFormula, "~?", "objRegistry.GetExpandedStringValue " & strRegHive & ", strRegPath, """ & varRegValue(1) & """, strField")
					strFormula = Replace(strFormula, "?", "strField")
					strFormula = Replace(strFormula, "~", "objShell.RegRead(strRegKey & """ & varRegValue(1) & """)")
				End If
			End If      
		Else
			strFieldName = varRegValue
			If strRegHive = "" Then
				strFormula = "strField = strFieldName"
			Else
				strFormula = "strField = objShell.RegRead(strRegKey & """ & varRegValue & """)"
			End If      
		End If
		If Left(strFieldName, 1) <> "#" Then
			strField = ""
			Execute strFormula
			Err.Clear
			If Len(strField) > 0 Then
				strField = Replace(strField, vbTab, "<Tab>")
				strField = Replace(strField, vbCrLf, "<CrLf>")
				strField = Replace(strField, vbCr, "<Cr>")
				strField = Replace(strField, vbLf, "<Lf>")
			End If
			Execute "str" & strFieldName & " = strField"
			strRecord = strRecord & vbTab & strField
			outArr(j) = strField
		End If
	Next
	fncRegKeyValue = strRecord
End Function

Function grabBiztalk()
	'http://social.technet.microsoft.com/wiki/cfs-filesystemfile.ashx/__key/communityserver-wikis-components-files/00-00-00-00-05/0743.BizTalk-Version_2C00_-Edition_2C00_-Name.png
	On Error Resume Next
	Call LogError("[S] Checking for BizTalk...", 0)
	Wscript.Echo "[S] Checking for BizTalk..."	
	grabBiztalk = False
	Dim bizVals(10)	
	Set objRegistry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & objToScan & "\root\default:StdRegProv")
	Call LogError("[I] Object: winmgmts:{impersonationLevel=impersonate}!\\" & objToScan & "\root\default:StdRegProv", 1)
	bizVals(0) = objToScan
	Call objRegistry.GetStringValue(HKLM,"SOFTWARE\Microsoft\BizTalk Server\3.0","InstallDate",   bizVals(1))
	Call LogError("[I] Registry...", 1)
	Call objRegistry.GetStringValue(HKLM,"SOFTWARE\Microsoft\BizTalk Server\3.0","InstallPath",   bizVals(2))
	Call objRegistry.GetStringValue(HKLM,"SOFTWARE\Microsoft\BizTalk Server\3.0","InstallTime",   bizVals(3))
	Call objRegistry.GetStringValue(HKLM,"SOFTWARE\Microsoft\BizTalk Server\3.0","ProductCode",   bizVals(4))
	Call objRegistry.GetStringValue(HKLM,"SOFTWARE\Microsoft\BizTalk Server\3.0","ProductCode_R2",bizVals(5))
	Call objRegistry.GetStringValue(HKLM,"SOFTWARE\Microsoft\BizTalk Server\3.0","ProductCode_R3",bizVals(6))
	Call objRegistry.GetStringValue(HKLM,"SOFTWARE\Microsoft\BizTalk Server\3.0","ProductEdition",bizVals(7))
	Call objRegistry.GetStringValue(HKLM,"SOFTWARE\Microsoft\BizTalk Server\3.0","ProductName",   bizVals(8))
	Call objRegistry.GetStringValue(HKLM,"SOFTWARE\Microsoft\BizTalk Server\3.0","ProductVersion",bizVals(9))
	If bizVals(1) <> "" Or bizVals(2) <> "" Or bizVals(3) <> "" Or bizVals(4) <> "" Or bizVals(5) <> "" Or bizVals(6) <> "" Or bizVals(7) <> "" Or bizVals(8) <> "" Or bizVals(9) <> "" Then
		Dim File: Set File = CreateFile("biztalk")
		File.WriteLine CSVLineFromArray(bizVals)
		grabBiztalk = True
	End if 
	File.Close
	Wscript.Echo "[D] Checking for BizTalk..."	
	Call LogError("[D] Checking for BizTalk...", 0)
End Function

Sub grabISAForefrontTMG()
	'todo
	grabISAForefrontTMG = true
End Sub

Sub grabExchange()
	On Error Resume Next
	Call LogError("[S] Collecting data for Exchange servers...", 0)
	Wscript.Echo "[S] Collecting data for Exchange servers..."	
	Dim objAdRootDSE, objRS, varConfigNC, strConnstring, strSQL, objWMIService, colItems, objItem, objServer
	Dim File : Set File = objFSO.CreateTextFile(OutputPath & "EXCHANGESERVERS" , True, True)
	Set objAdRootDSE = GetObject("LDAP://RootDSE")
	Set objRS = CreateObject("adodb.recordset")
	varConfigNC = objAdRootDSE.Get("configurationNamingContext")
	strConnstring = "Provider=ADsDSOObject"
	strSQL = "SELECT * FROM 'LDAP://" & varConfigNC & "' WHERE objectCategory='msExchExchangeServer'"
	objRS.Open strSQL, strConnstring
	Call LogError("[I] strSQL : " & strSQL , 1)
	Do until objRS.eof
		Set objServer = GetObject(objRS.Fields.Item(0))
		Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!" & objServer.CN & "\ROOT\MicrosoftExchangeV2")
		Call LogError("[I] objWMIService : " & "winmgmts:{impersonationLevel=impersonate}!" & objServer.CN & "\ROOT\MicrosoftExchangeV2", 1)
		If isObject(objWMIService) Then
			Set colItems = objWMIService.ExecQuery("Select * from Exchange_Server")
			For Each objItem in colItems
				File.writeline CSVLineFromArray(Array( _
					""&objServer.CN, _
					""&objItem.AdministrativeNote, _
					""&objItem.CreationTime, _
					""&objItem.DN, _
					""&objItem.ExchangeVersion, _
					""&objItem.FQDN, _
					""&objItem.GUID, _
					""&objItem.IsFrontEndServer, _
					""&objItem.LastModificationTime, _
					""&objItem.MessageTrackingEnabled, _
					""&objItem.MessageTrackingLogFileLifetime, _
					""&objItem.MessageTrackingLogFilePath, _
					""&objItem.MonitoringEnabled, _
					""&objItem.MTADataPath, _
					""&objItem.Name, _
					""&objItem.RoutingGroup, _
					""&objItem.SubjectLoggingEnabled, _
					""&objItem.Type _
				)) 
				Call LogError("[I] put : " & objServer.CN, 1)
			Next
		Else
			File.writeline CSVLineFromArray(Array(""&objServer.CN, "NODATA", "NODATA", "NODATA", "NODATA", "NODATA", "NODATA", "NODATA", "NODATA", "NODATA", "NODATA", "NODATA", "NODATA", "NODATA", "NODATA", "NODATA", "NODATA", "NODATA")) 
		End If
		Set colItems = Nothing
		Set objWMIService = Nothing
		Set objServer = Nothing
        objRS.movenext
	Loop
  	objRS.close
	Set objRS = Nothing
	Set objAdRootDSE = Nothing
	File.Close
	Wscript.Echo "[D] Collecting data for Exchange servers..."	
	Call LogError("[D] Collecting data for Exchange servers...", 0)
End Sub

Sub checkMSSQL(instance, state)
	On Error Resume Next
	Call LogError("[S] Checking for MSSQL...", 0)
	Wscript.Echo "[S] Checking for MSSQL..."
	Dim connect: connect = Replace(instance, "MSSQL$", "")
	Dim File : Set File = CreateFile("mssql")
	File.WriteLine CSVLineFromArray(Array(objToScan, state, "Service", connect, "", "", "", "", "", "", "", "", ""))
	' Registry based version/edition detection
	Dim instances, tmpInstance, value, tmpArr(12)
	tmpArr(0) = objToScan
	tmpArr(1) = state
	tmpArr(2) = "registry"
	tmpArr(3) = instance
	Call objRegistry.EnumValues(HKLM, "SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\SQL",instances)
	Call LogError("[I] objRegistry: SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\SQL", 1)
	For Each tmpInstance In instances
		If UCase(tmpInstance) = UCase(connect) Then
			Call objRegistry.GetStringValue(HKLM, "SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\SQL",tmpInstance,value)
			Call LogError("[I] Value: SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\SQL", 1)
			Call objRegistry.GetStringValue(HKLM, "SOFTWARE\Microsoft\Microsoft SQL Server\" & value & "\Setup","Edition",tmpArr(4))
			Call objRegistry.GetStringValue(HKLM, "SOFTWARE\Microsoft\Microsoft SQL Server\" & value & "\Setup","EditionType",tmpArr(5))
			Call objRegistry.GetStringValue(HKLM, "SOFTWARE\Microsoft\Microsoft SQL Server\" & value & "\Setup","Version",tmpArr(6))
			Call objRegistry.GetStringValue(HKLM, "SOFTWARE\Microsoft\Microsoft SQL Server\" & value & "\Setup","PatchLevel",tmpArr(7))
			Call objRegistry.GetStringValue(HKLM, "SOFTWARE\Microsoft\Microsoft SQL Server\" & value & "\Setup","ProductCode",tmpArr(8))
			Call objRegistry.GetStringValue(HKLM, "SOFTWARE\Microsoft\Microsoft SQL Server\" & value & "\Setup","SQLBinRoot",tmpArr(9))
			Call objRegistry.GetStringValue(HKLM, "SOFTWARE\Microsoft\Microsoft SQL Server\" & value & "\Setup","SQLDataRoot",tmpArr(10))
			Call objRegistry.GetStringValue(HKLM, "SOFTWARE\Microsoft\Microsoft SQL Server\" & value & "\Setup","SQLGroup",tmpArr(11))
			Call objRegistry.GetStringValue(HKLM, "SOFTWARE\Microsoft\Microsoft SQL Server\" & value & "\Setup","SQLPath",tmpArr(12))
			File.WriteLine CSVLineFromArray(tmpArr)
		End If
	Next
	Call initWMISQL()
	Call LogError("[I] WMI SQL Initiated", 1)
	Dim i: For i = LBound(tmpArr) to UBound(tmpArr)
		tmpArr(i) = ""
	Next
	Dim o,q: Set q = objWMISQL.ExecQuery("Select * from SqlServiceAdvancedProperty where ServiceName Like '%" & connect & "%'")
	Call LogError("[I] Query: Select * from SqlServiceAdvancedProperty where ServiceName Like '%" & connect & "%'", 1)
	tmpArr(0) = objToScan
	tmpArr(1) = state
	tmpArr(2) = "wmi"
	For Each o In q
		With o
			Select Case UCase(.PropertyName)
				Case "SKUNAME"		: tmpArr(4)  = ""&.PropertyStrValue
				Case "CLUSTERED"	: tmpArr(5)  = ""&.PropertyNumValue
				Case "VERSION"		: tmpArr(6)  = ""&.PropertyStrValue
				Case "SKU"			: tmpArr(7) = ""&.PropertyNumValue
				Case "FILEVERSION"	: tmpArr(8)  = ""&.PropertyStrValue
				Case "INSTALLPATH"	: tmpArr(9)  = ""&.PropertyStrValue
				Case "DATAPATH"		: tmpArr(10) = ""&.PropertyStrValue
				Case "INSTANCEID"	: tmpArr(11) = ""&.PropertyStrValue
				Case "DUMPDIR"		: tmpArr(12) = ""&.PropertyStrValue
				tmpArr(3) = ""&.ServiceName
				Call LogError("[I] PropertyName: "&.PropertyName&" ,ServiceName: "&.ServiceName, 1)
			End Select
		End With
	Next
	If q.Count > 0 Then
		If ""&vals(4) = "" Then: vals(4) = connect: End If
		File.WriteLine CSVLineFromArray(tmpArr)
	End If
	If state = "Running" Then
		connect = objToScan & "\" & connect
		If connectToMSSQL(connect, "", "", instance, File) = False Then
			Call LogError("[I] Failed connect to MSSQL without password : "&connect&" "&instance, 1)
			If connectToMSSQL(connect, "sa", "", instance, File) = False Then
				Call LogError("[I] Failed connect to MSSQL with SA : "&connect&" "&instance, 1)
				File.WriteLine CSVLineFromArray(Array(objToScan, state, "Login", Replace(instance, "MSSQL$", ""), "", "Failed to connect", "", "", "", "", "", "", ""))
			End If
		End If
	End If
	Wscript.Echo "[D] Checking for MSSQL..."
	Call LogError("[D] Checking for MSSQL...", 0)
End Sub

Sub checkMSSQL2(instance, state)
	On Error Resume Next
	Call LogError("[S] Checking for MSSQL2...", 0)
	Wscript.Echo "[S] Checking for MSSQL2..."
	Dim connect: connect = Replace(instance, "MSSQL", "")
	Dim File : Set File = CreateFile("mssql2")
	File.WriteLine CSVLineFromArray(Array(objToScan, state, "Service", connect, "", "", "", "", "", "", "", "", ""))
	' Registry based version/edition detection
	Dim instances, tmpInstance, value, tmpArr(12)
	tmpArr(0) = objToScan
	tmpArr(1) = state
	tmpArr(2) = "registry"
	tmpArr(3) = instance
	Call objRegistry.EnumValues(HKLM, "SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\SQL",instances)
	Call LogError("[I] objRegistry: SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\SQL", 1)
	For Each tmpInstance In instances
		If UCase(tmpInstance) = UCase(connect) Then
			Call objRegistry.GetStringValue(HKLM, "SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\SQL",tmpInstance,value)
			Call LogError("[I] Value: SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\SQL", 1)
			Call objRegistry.GetStringValue(HKLM, "SOFTWARE\Microsoft\Microsoft SQL Server\" & value & "\Setup","Edition",tmpArr(4))
			Call objRegistry.GetStringValue(HKLM, "SOFTWARE\Microsoft\Microsoft SQL Server\" & value & "\Setup","EditionType",tmpArr(5))
			Call objRegistry.GetStringValue(HKLM, "SOFTWARE\Microsoft\Microsoft SQL Server\" & value & "\Setup","Version",tmpArr(6))
			Call objRegistry.GetStringValue(HKLM, "SOFTWARE\Microsoft\Microsoft SQL Server\" & value & "\Setup","PatchLevel",tmpArr(7))
			Call objRegistry.GetStringValue(HKLM, "SOFTWARE\Microsoft\Microsoft SQL Server\" & value & "\Setup","ProductCode",tmpArr(8))
			Call objRegistry.GetStringValue(HKLM, "SOFTWARE\Microsoft\Microsoft SQL Server\" & value & "\Setup","SQLBinRoot",tmpArr(9))
			Call objRegistry.GetStringValue(HKLM, "SOFTWARE\Microsoft\Microsoft SQL Server\" & value & "\Setup","SQLDataRoot",tmpArr(10))
			Call objRegistry.GetStringValue(HKLM, "SOFTWARE\Microsoft\Microsoft SQL Server\" & value & "\Setup","SQLGroup",tmpArr(11))
			Call objRegistry.GetStringValue(HKLM, "SOFTWARE\Microsoft\Microsoft SQL Server\" & value & "\Setup","SQLPath",tmpArr(12))
			File.WriteLine CSVLineFromArray(tmpArr)
		End If
	Next
	Call initWMISQL()
	Call LogError("[I] WMI SQL Initiated", 1)
	Dim i: For i = LBound(tmpArr) to UBound(tmpArr)
		tmpArr(i) = ""
	Next
	Dim o,q: Set q = objWMISQL.ExecQuery("Select * from SqlServiceAdvancedProperty where ServiceName Like '%" & connect & "%'")
	Call LogError("[I] Query: Select * from SqlServiceAdvancedProperty where ServiceName Like '%" & connect & "%'", 1)
	tmpArr(0) = objToScan
	tmpArr(1) = state
	tmpArr(2) = "wmi"
	For Each o In q
		With o
			Select Case UCase(.PropertyName)
				Case "SKUNAME"		: tmpArr(4)  = ""&.PropertyStrValue
				Case "CLUSTERED"	: tmpArr(5)  = ""&.PropertyNumValue
				Case "VERSION"		: tmpArr(6)  = ""&.PropertyStrValue
				Case "SKU"			: tmpArr(7) = ""&.PropertyNumValue
				Case "FILEVERSION"	: tmpArr(8)  = ""&.PropertyStrValue
				Case "INSTALLPATH"	: tmpArr(9)  = ""&.PropertyStrValue
				Case "DATAPATH"		: tmpArr(10) = ""&.PropertyStrValue
				Case "INSTANCEID"	: tmpArr(11) = ""&.PropertyStrValue
				Case "DUMPDIR"		: tmpArr(12) = ""&.PropertyStrValue
				tmpArr(3) = ""&.ServiceName
				Call LogError("[I] PropertyName: "&.PropertyName&" ,ServiceName: "&.ServiceName, 1)
			End Select
		End With
	Next
	If q.Count > 0 Then
		If ""&vals(4) = "" Then: vals(4) = connect: End If
		File.WriteLine CSVLineFromArray(tmpArr)
	End If
	If state = "Running" Then
		connect = objToScan & "\" & connect
		If connectToMSSQL(connect, "", "", instance, File) = False Then
			Call LogError("[I] Failed connect to MSSQL without password : "&connect&" "&instance, 1)
			If connectToMSSQL(connect, "sa", "", instance, File) = False Then
				Call LogError("[I] Failed connect to MSSQL with SA : "&connect&" "&instance, 1)
				File.WriteLine CSVLineFromArray(Array(objToScan, state, "Login", Replace(instance, "MSSQL", ""), "", "Failed to connect", "", "", "", "", "", "", ""))
			End If
		End If
	End If
	Wscript.Echo "[D] Checking for MSSQL2..."
	Call LogError("[D] Checking for MSSQL2...", 0)
End Sub


Function connectToMSSQL(connect, user, password, instance, File)
	On Error Resume Next
	Call LogError("[S] Connecting to MSSQL server...", 0)
	Wscript.Echo "[S] Connecting to MSSQL server..."
	connectToMSSQL = False
	Dim adoConn: Set adoConn = CreateObject("ADODB.Connection")
	Dim adoRS: Set adoRS = CreateObject("ADODB.Recordset")
	' Build connect string depending on credentials
	If user = "" Then
		adoConn.ConnectionString = "Provider=SQLOLEDB;Data Source=" & connect & ";Integrated Security=SSPI;Connect Timeout = 10"
	Else
		If password = "" Then
			adoConn.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & user & ";Data Source=" & connect
		Else
			adoConn.ConnectionString = "Provider=SQLOLEDB.1;Password=" & password & ";Persist Security Info=False;User ID=" & user & ";Data Source=" & connect
		End If
	End If
	' Attempt to connect
	Call adoConn.Open()
	If Err <> 0 Then
		Call LogError("[I] Failed to open connection: " & adoConn.ConnectionString, 1)		
	else
		Dim tmpArr(12) 
		tmpArr(0) = objToScan
		tmpArr(1) = "Running"
		tmpArr(2) = "SQL"
		adoRS.Open "Select SERVERPROPERTY('Edition') AS [Edition],SERVERPROPERTY('InstanceName') AS [InstanceName],@@VERSION AS [Server Information],SERVERPROPERTY('productversion') AS [ProductVersion],SERVERPROPERTY('ProductLevel') AS [ProductLevel], SERVERPROPERTY('ISCLUSTERED') AS ISCLUSTERED", adoConn
		If Err <> 0 Then
			Call LogError("[I] Failed to query for: Select SERVERPROPERTY('Edition') AS [Edition],SERVERPROPERTY('InstanceName') AS [InstanceName],@@VERSION AS [Server Information],SERVERPROPERTY('productversion') AS [ProductVersion],SERVERPROPERTY('ProductLevel') AS [ProductLevel], SERVERPROPERTY('ISCLUSTERED') AS ISCLUSTERED", 1)
		else
			With adoRS.Fields
				tmpArr(4)  = ""&.Item("Edition")
				tmpArr(5)  = ""&.Item("IsClustered")
				tmpArr(6)  = ""&.Item("ProductVersion")
				tmpArr(7) = Trim(Replace(Replace(Replace(""&.Item("Server Information"),chr(13),""),chr(10),""),VbTab," "))
				tmpArr(8)  = ""
				tmpArr(9)  = ""
				tmpArr(10) = ""&.Item("InstanceName")
				tmpArr(11) = instance
				tmpArr(12) = ""&.Item("ProductLevel")
				File.WriteLine CSVLineFromArray(tmpArr)
			End With
			connectToMSSQL = True
		End If
		Call adoConn.Close()
	End If
	Wscript.Echo "[D] Connecting to MSSQL server..."
	Call LogError("[D] Connecting to MSSQL server...", 0)
End Function

Function SysLog(message,id,description,source)
	On Error Resume Next
	Err.Clear
	Dim str : str = CSVLineFromArray(Array(objToScan, message, id, source, description))
	Wscript.Echo "[-] Error connecting to " & objToScan & ". Reason: " & description & "("&id&")"
	If id=462 Then
		CreateFile("offline")
	Else
		CreateFile("error")
	End If
	Dim File: Set File = CreateFile("WMIError")
	File.WriteLine str
	File.Close
End Function

Function CPUload
	On Error Resume Next
	Call LogError("[S] CPULoad", 0)
	Dim o,q: Set q = objWMI.ExecQuery("Select deviceid, loadpercentage from Win32_Processor where DeviceID='CPU0'")
	Call LogError("[I] CPULoad query executed", 1)
	CPUload = 25
	For Each o In q
		CPUload = o.LoadPercentage
		Call LogError("[I] CPULoad = "&o.LoadPercentage, 1)
		If IsNull(CPUload) Then: CPUload = 25: End If
	Next
	Call LogError("[D] CPULoad", 0)
End Function

Function CreateFile(filename)
	On Error Resume Next
	Err.Clear
	objFSO.DeleteFile(OutputPath & objToScan & fileDelimiter & filename)
	Set CreateFile = objFSO.CreateTextFile(OutputPath & objToScan & fileDelimiter & filename, True, True)
End function

Function CSVLineFromArray(arr)
	On Error Resume Next
	Dim quote: quote = chr(34)
	Dim sep:   sep   = ";"
	Dim str
	CSVLineFromArray = ""
	Dim i: For i = LBound(arr) to UBound(arr)
		str = "" & arr(i)
		CSVLineFromArray = CSVLineFromArray & quote & Replace(str, quote, quote & quote) & quote & sep
	Next
End Function

Function getLocalname
	On Error Resume Next
	Wscript.Echo "[S] Getting object name..."
	Dim objNet: Set objNet = Wscript.CreateObject("Wscript.Network")
	If Err = 0 Then: getLocalname = UCase(Trim(objNet.ComputerName)): Else: getLocalname = "" End If
	Wscript.Echo "[D] Getting object name..."
End Function

Function IntToDate(LongInt)
	On Error Resume Next	
	Err.Clear
	Dim lngBiasKey
	Dim glngBias
	lngBiasKey = objShell.RegRead("HKLM\System\CurrentControlSet\Control\TimeZoneInformation\ActiveTimeBias")
	If UCase(TypeName(lngBiasKey)) = "LONG" Then
		glngBias = lngBiasKey
	ElseIf UCase(TypeName(lngBiasKey)) = "VARIANT()" Then
		glngBias = 0
		For k = 0 To UBound(lngBiasKey)
			glngBias = lngBias + (lngBiasKey(k) * 256^k)
		Next
	End If 
	IntToDate = Cdate(#1/1/1601# + (((LongInt.HighPart * (2 ^ 32)) + LongInt.LowPart) / 600000000 - glngBias) / 1440)
End Function

Sub grabObjects
	On Error Resume Next
	Call LogError("[S] Collecting objects from Active Directory...", 0)
	Wscript.Echo "[S] Collecting objects from Active Directory..."
	Dim line, adoRS, counter
	If isDomainFound Then
		' ADODB connection
		Dim objADOConn: Set objADOConn = CreateObject("ADODB.Connection")
		Dim objADOCmd:  Set objADOCmd = CreateObject("ADODB.Command")
		objADOConn.Provider = "ADsDSOObject"
		objADOConn.Open "Active Directory Provider"
		Set objADOCmd.ActiveConnection = objADOConn
		' Domain listing
		Set adoRS = objADOConn.Execute("<GC://" & Context & ">;(objectcategory=domainDNS);name;SubTree")
		Call LogError("[I] Querying for child domains: <GC://" & Context & ">;(objectcategory=domainDNS);name;SubTree", 1)
		line = ""
		Do Until adoRS.EOF
			line = line & CSVLineFromArray(Array(adoRS.Fields("Name").Value))
			adoRS.MoveNext
		Loop
		adoRS.Close
		Dim fileList: Set fileList = objFSO.CreateTextFile(FullPath & "Objectslist.txt", True, True)
		fileList.WriteLine "### Currentdomain=" & Context
		fileList.WriteLine "### Fulldomainlist=" & line
		fileList.WriteLine "### Scanstarttime=" & Now
		' Grab all relevant objects
		objADOCmd.CommandText = "<LDAP://" & Context & ">;(objectCategory=computer);name,createTimeStamp,lastlogonTimeStamp,operatingSystem,operatingSystemVersion,pwdLastSet,whenChanged,whenCreated,distinguishedName,cn,instancetype,lastLogon,logonCount,operatingsystemservicepack;subtree"
		objADOCmd.Properties("Page Size") = 500
		objADOCmd.Properties("Timeout") = 30
		objADOCmd.Properties("Cache Results") = False
		Set adoRS = objADOCmd.Execute
		Call LogError("[I] Querying for objects: <LDAP://" & Context & ">;(objectCategory=computer);name,createTimeStamp,lastlogonTimeStamp,operatingSystem,operatingSystemVersion,pwdLastSet,whenChanged,whenCreated,distinguishedName,cn,instancetype,lastLogon,logonCount,operatingsystemservicepack;subtree", 1)
		Do Until adoRS.EOF
			counter = counter + 1
			If (counter mod 100) = 0 then: Wscript.Echo "[I] Processed " & counter & " objects..." : end if
			Dim tmpName: tmpName= ""&UCase(Trim(adoRS.Fields("Name").Value))
			Call LogError("[I] Empty vals for tmpName", 1)
			Dim tmpcreateTimeStamp: tmpcreateTimeStamp = ""&Trim(adoRS.Fields("createTimeStamp").Value)
			Call LogError("[I] Empty vals for tmpcreateTimeStamp", 1)
			Dim tmplastLogonTimestamp2: tmplastLogonTimestamp2 = IntToDate(adoRS.Fields("lastLogonTimestamp").Value)
			Call LogError("[I] Empty vals for tmplastLogonTimestamp2", 1)
			Dim tmplastLogonTimestamp: tmplastLogonTimestamp = DateDiff("d", IntToDate(adoRS.Fields("lastLogonTimestamp").Value), Now)
			Call LogError("[I] Empty vals for tmplastLogonTimestamp", 1)
			Dim tmpoperatingSystem: tmpoperatingSystem = ""&Trim(adoRS.Fields("operatingSystem").Value)
			Call LogError("[I] Empty vals for tmpoperatingSystem", 1)
			Dim tmpoperatingSystemVersion: tmpoperatingSystemVersion = ""&Trim(adoRS.Fields("operatingSystemVersion").Value)
			Call LogError("[I] Empty vals for tmpoperatingSystemVersion", 1)
			Dim tmppwdLastSet: tmppwdLastSet = DateDiff("d", IntToDate(adoRS.Fields("pwdLastSet").Value), Now)
			Call LogError("[I] Empty vals for tmppwdLastSet", 1)
			Dim tmppwdLastSet2: tmppwdLastSet2 = IntToDate(adoRS.Fields("pwdLastSet").Value)
			Call LogError("[I] Empty vals for tmppwdLastSet2", 1)
			Dim tmpwhenChanged: tmpwhenChanged = ""&Trim(adoRS.Fields("whenChanged").Value)
			Call LogError("[I] Empty vals for tmpwhenChanged", 1)
			Dim tmpwhenCreated: tmpwhenCreated = ""&Trim(adoRS.Fields("whenCreated").Value)
			Call LogError("[I] Empty vals for tmpwhenCreated", 1)
			Dim tmpdistinguishedName: tmpdistinguishedName = ""&Trim(adoRS.Fields("distinguishedName").Value)
			Call LogError("[I] Empty vals for tmpdistinguishedName", 1)
			Dim tmpCN: tmpCN = ""&Trim(adoRS.Fields("cn").Value)
			Call LogError("[I] Empty vals for tmpCN", 1)
			Dim tmpinstancetype: tmpinstancetype = ""&Trim(adoRS.Fields("instancetype").Value)
			Call LogError("[I] Empty vals for tmpinstancetype", 1)
			Dim tmplastLogon
			if (adoRS.Fields("lastLogon").value <> Null) then
				tmplastLogon = IntToDate(adoRS.Fields("lastLogon").Value)
				Call LogError("[I] Empty vals for tmplastLogon", 1)
			else 
				tmplastLogon = "01.01.1900 00:00:00"
			end if
			Dim tmplogonCount: tmplogonCount = ""&adoRS.Fields("logonCount").Value
			Call LogError("[I] Empty vals for tmplogonCount", 1)
			Dim tmpoperatingsystemservicepack: tmpoperatingsystemservicepack = ""&Trim(adoRS.Fields("operatingsystemservicepack").Value)
			Call LogError("[I] Empty vals for tmpoperatingsystemservicepack", 1)
			fileList.WriteLine CSVLineFromArray(Array( _
				tmpName, _
				tmpcreateTimeStamp, _
				tmplastLogonTimestamp, _
				tmpoperatingSystem, _
				tmpoperatingSystemVersion, _
				tmppwdLastSet, _
				tmpwhenChanged, _
				tmpwhenCreated, _
				tmpdistinguishedName, _
				tmplastLogonTimestamp2, _
				tmppwdLastSet2, _
				tmpCN, _
				tmpinstancetype, _
				tmplastLogon, _
				tmplogonCount, _
				tmpoperatingsystemservicepack _
			))
			adoRS.MoveNext
		Loop
		adoRS.Close
		fileList.Close
	End If
	Wscript.Echo "[D] Collecting objects from Active Directory..."
	Call LogError("[D] Collecting objects from Active Directory...", 0)
End Sub

Function LoadList
	On Error Resume Next
	Err.clear
	Wscript.Echo "[S] Loading a list of objects to process..."
	Dim filename: filename = FullPath & "Objectslist.txt"
	Dim arr()
	If Not objFSO.FileExists(filename) Then
		Wscript.Echo "[-] Objectslist.txt not found! Exiting."
		Wscript.Quit
	End If
	Dim File: Set File = objFSO.OpenTextFile(filename, 1, False, -2) ' Possibly force unicode with -1
	Dim line, idx, i: i = 0
	Do While File.AtEndOfStream <> True
		line = Trim(File.ReadLine())
		If Left(line, 3) <> "###" Then
			idx = InStr(line, ";")
			If idx > 1 Then
				line = Left(line, idx-1)
			End If
			line = Trim(Replace(Replace(Replace(line, """", ""), vbCr, ""), vbLf, ""))
			If Len(line) > 0 Then
				ReDim Preserve arr(i)
				arr(i) = UCase(line)
				i = i+1
			End If
		End If
	Loop
	If i > 0 Then
		Wscript.Echo "[I] Loaded " & i & " systems into memory"
		Call objFSO.CopyFile(filename, OutputPath & "#CopyOfObjectslist.txt", True)
	Else
		Wscript.Echo "[-] Objectslist.txt seems to be empty. Exiting"
		Wscript.Quit
	End If
	LoadList = arr
	File.Close
	Wscript.Echo "[D] Loading a list of objects to process..."
End Function

Sub parseCmdArgs
	On Error Resume Next
	Err.Clear
	Dim arg: For Each arg In Wscript.Arguments
		Dim argFound : argFound = False
		Dim key: For Each key In cmdArgs
			If (InStr(UCase(arg),key) > 0) And (Len(arg) > 0) Then
				argFound = True
				cmdArgs(key) = Mid(Trim(arg),Len(key)+1)
				If oldShell And key <> "/PROCESSLOCAL" Then
					Wscript.Echo "[-] Outdated version of Windows Script is used. Only PROCESSLOCAL mode is allowed. Please update WSH."
					Wscript.Quit
				End If
			End If
		Next
		If Not argFound Then
			Wscript.Echo "[-] An invalid argument: " & arg
			Wscript.Quit
		End If
	Next
End Sub

Sub Init
	On Error Resume Next
	Err.clear
	' only cscript is allowed
	If Right(UCase(Wscript.FullName),11)= "Wscript.EXE" Then
		Wscript.Echo "[-] Please use CMD & CSCRIPT.EXE to execute this script."
		Wscript.Quit
	End If
	' let's detect Windows Script version
	If CInt(Replace(UCase(Wscript.Version),".","")) < 56 Then
		oldShell = True
	Else
		oldShell = False
	End If
	Call initFSO()
	Call initShell()
	objToScan = getLocalname()
	
	FullPath = Mid(Wscript.ScriptFullName, 1, InStrRev(Wscript.ScriptFullName, "\"))
	If oldShell Then
		OutputPath = Replace(FullPath & OutputFolder & "\","\\","\")
		If Not objFSO.FolderExists(OutputFolder) Then: objFSO.CreateFolder(OutputFolder): End If
	Else
		objShell.CurrentDirectory = FullPath
		OutputPath = objShell.CurrentDirectory & "\" & OutputFolder & "\"
		If Not objFSO.FolderExists(OutputFolder) Then: objFSO.CreateFolder(OutputFolder): End If
		objShell.CurrentDirectory = OutputFolder
	End If
	
End Sub

Sub processList
	On Error Resume Next
	Err.clear
	If oldShell Then
		Wscript.Echo "[-] This script cannot be run with this Shell version for remote scanning."
		Wscript.Echo "[-] Please update Shell or run the script with a newer OS."
		Wscript.Quit
	End If
	Call initWMI
	Dim o,q,servers: servers = LoadList()
	Dim starttime: starttime = Now
	Dim pass: pass = 1
	Do Until (pass > maxRunningPasses) 
		Dim i: i = 0
		For i=LBound(servers) To UBound(servers)
			If objFSO.FileExists(OutputPath & servers(i) & fileDelimiter & "scanned") Then
				Wscript.Echo "[I] Skipping   [" & LPad(i+1,5) & "/" & UBound(servers)+1 & "] """ & servers(i) & """"
			Else 
				Wscript.Echo "[I] Processing [" & LPad(i+1,5) & "/" & UBound(servers)+1 & "] """ & servers(i) & """"
				Call objShell.Run("CScript.exe //NOLOGO """ & Wscript.ScriptFullName & """ /PROCESSTARGET:" & servers(i), 7, False)
				' Try to save system resources
				If (CLng(i) mod 4) = 2 Then
					' CPU load
					Do While True
						If CPUload() > maxHostCPUload Then
							Wscript.Sleep 3000
						Else
							Exit Do
						End If
					Loop
					Do While True
						Set q = objWMI.ExecQuery("select * from Win32_Process WHERE Name='cscript.exe' AND CommandLine LIKE '%/PROCESSTARGET:%'")
						If q.Count >= maxThreadCount Then
							Wscript.Sleep 1000
						Else
							Exit Do
						End If
					Loop
				End If
			End If
		Next
		' Wait for all threads to finish
		Do While True
			Set q = objWMI.ExecQuery("select * from Win32_Process WHERE Name='cscript.exe' AND CommandLine LIKE '%/PROCESSTARGET:%'")
			If q.Count > 0 Then
				Wscript.Sleep 300
			Else
				Exit Do
			End If
		Loop
		' Check status
		Dim j, iDone, iLoad, iError, iOff, iWork, iAdm: iDone = 0: iLoad = 0: iError = 0: iOff = 0: iWork = 0: iAdm = 0
		For j=LBound(servers) To UBound(servers)
			Dim path: path = OutputPath & servers(j) & fileDelimiter
			If objFSO.FileExists(path & "scanned") Then: iDone=iDone+1: End If
			If objFSO.FileExists(path & "highload") Then: iLoad=iLoad+1: End If
			If objFSO.FileExists(path & "error") Then: iError=iError+1: End If
			If objFSO.FileExists(path & "offline") Then: iOff=iOff+1: End If
			If objFSO.FileExists(path & "noaccess") Then: iAdm=iAdm+1: End If
			If objFSO.FileExists(path & "object") Then: iWork=iWork+1: End If
		Next
		Wscript.Echo VbCrLf
		Wscript.Echo "[I] Status     [" & LPad(iDone,5)  & "/" & UBound(servers)+1 & "] " & "Scanned"
		Wscript.Echo "[I] Status     [" & LPad(iLoad,5)  & "/" & UBound(servers)+1 & "] " & "High load"
		Wscript.Echo "[I] Status     [" & LPad(iError,5) & "/" & UBound(servers)+1 & "] " & "Error"
		Wscript.Echo "[I] Status     [" & LPad(iOff,5)   & "/" & UBound(servers)+1 & "] " & "Offline"
		If iAdm > 0 Then
			Wscript.Echo "[I] Status     [" & LPad(iAdm,5)   & "/" & UBound(servers)+1 & "] " & "NO access"
		End If
		If iWork > 0 Then
			Wscript.Echo "[I] Status     [" & LPad(iWork,5)   & "/" & UBound(servers)+1 & "] " & "Objects"
		End If
		Wscript.Echo VbCrLf
		If iDone+iWork = UBound(servers)-LBound(servers)+1 Then: Exit Do: End If
		Wscript.Sleep 20000
		pass = pass+1
	Loop
	Wscript.Quit
End Sub

Function grabOS
	On Error Resume Next
	Wscript.Echo "[S] Collecting data on operating system..."
	Call LogError("[S] Collecting data on operating system...", 0)
	grabOS = False
	Dim lines: lines = ""
	Dim tmp
	With WMIQuery(objWMI, "Select * from Win32_OperatingSystem")
		objToScan = UCase(objToScan)
		lines = lines & CSVLineFromArray(Array(objToScan, _
			""&.BuildNumber, _
			""&.BuildType, _
			""&.Caption, _
			""&.CSName, _
			""&.Description, _
			""&.Distributed, _
			""&.InstallDate, _
			""&.LastBootUpTime, _
			""&.LocalDateTime, _
			""&.Locale, _
			""&.Manufacturer, _
			""&.Organization, _
			""&.OSArchitecture , _
			""&.OSLanguage, _
			""&.OSProductSuite, _
			""&.OSType , _
			""&.OtherTypeDescription, _
			""&.RegisteredUser, _ 
			""&.SerialNumber, _
			""&.ServicePackMajorVersion, _
			""&.ServicePackMinorVersion, _
			""&.SuiteMask, _ 
			""&.Version _
		)) & vbCrLf
		tmp = .BuildNumber
		Call LogError("[I] Empty vals for BuildNumber", 1)
		tmp = .BuildType
		Call LogError("[I] Empty vals for BuildType", 1)
		tmp = .Caption
		Call LogError("[I] Empty vals for Caption", 1)
		tmp = .CSName
		Call LogError("[I] Empty vals for CSName", 1)
		tmp = .Description
		Call LogError("[I] Empty vals for Description", 1)
		tmp = .Distributed
		Call LogError("[I] Empty vals for Distributed", 1)
		tmp = .InstallDate
		Call LogError("[I] Empty vals InstallDate", 1)
		tmp = .LastBootUpTime
		Call LogError("[I] Empty vals for LastBootUpTime", 1)
		tmp = .LocalDateTime
		Call LogError("[I] Empty vals for LocalDateTime", 1)
		tmp = .Locale
		Call LogError("[I] Empty vals for Locale", 1)
		tmp = .Manufacturer
		Call LogError("[I] Empty vals for Manufacturer", 1)
		tmp = .Organization
		Call LogError("[I] Empty vals for Organization", 1)
		tmp = .OSArchitecture
		Call LogError("[I] Empty vals for OSArchitecture", 1)
		tmp = .OSLanguage
		Call LogError("[I] Empty vals for OSLanguage", 1)
		tmp = .OSProductSuite
		Call LogError("[I] Empty vals for OSProductSuite", 1)
		tmp = .OSType
		Call LogError("[I] Empty vals for OSType", 1)
		tmp = .OtherTypeDescription
		Call LogError("[I] Empty vals for OtherTypeDescription", 1)
		tmp = .RegisteredUser
		Call LogError("[I] Empty vals for RegisteredUser", 1)
		tmp = .SerialNumber
		Call LogError("[I] Empty vals for SerialNumber", 1)
		tmp = .ServicePackMajorVersion
		Call LogError("[I] Empty vals for ServicePackMajorVersion", 1)
		tmp = .ServicePackMinorVersion
		Call LogError("[I] Empty vals for ServicePackMinorVersion", 1)
		tmp = .SuiteMask
		Call LogError("[I] Empty vals for SuiteMask", 1)
		tmp = .Version
		Call LogError("[I] Empty vals for Version", 1)
		If Instr(UCase(.Caption),"SERVER") > 0 Then
			isServer = True
		End If
	End With
	Call LogError("[I] Query: Select * from Win32_OperatingSystem", 1)
	Dim File: Set File = createFile("os")
	File.Write lines
	File.Close
	If Len(lines)>0 Then 
		grabOS = True
	Else
		objFSO.DeleteFile(OutputPath & objToScan & fileDelimiter & "os")
	End if
	Wscript.Echo "[D] Collecting data on operating system..."
	Call LogError("[D] Collecting data on operating system...", 0)
End Function

Function grabLocalUsers
	On Error Resume Next
	grabLocalUsers = False
	Call LogError("[S] Collecting data on local user accounts...", 0)
	Wscript.Echo "[S] Collecting data on local user accounts..."
	If domainRole < 4 Then
		Dim File: Set File = CreateFile("LocalUsers")
		Dim line: Set line = ""
		Dim o,q: Set q = objWMI.ExecQuery("Select * from Win32_UserAccount Where LocalAccount = True and SidType=1")
		Call LogError("[I] Query: Select * from Win32_UserAccount Where LocalAccount = True and SidType=1", 1)
		For Each o In q
		With o
			Dim r: Set r = GetObject("WinNT://" & objToScan & "/" & .Name & ", user")
			Dim strLastLogin: strLastLogin = r.LastLogin
			Call LogError("[I] Empty vals for LastLogin", 1)
			If Err<>0 Then: strLastLogin = "Never": End If
			Call LogError("[I] Object: WinNT://" & objToScan & "/" & .Name & ", user", 1)
			Dim Caption: Caption = .Caption
			If Instr(Caption,"\") > 0 Then: Caption = Split(.Caption,"\")(1): End If
			Dim tmp
			line = CSVLineFromArray(Array( _
				Caption, _
				objToScan, _ 
				""&.AccountType, _ 
				""&.Description,  _
				""&.Disabled, _
				""&.Domain, _
				""&.Lockout, _
				""&.Name, _
				""&.PasswordChangeable, _
				""&.PasswordExpires, _
				""&.PasswordRequired, _
				""&.SID, _
				""&.SIDType, _
				""&.Status, _
				strLastLogin, _
				DateAdd("s", r.PasswordAge * -1, Now), _
				r.PasswordMinimumLength, _
				""&.LocalAccount) _
			)
			tmp = .AccountType
			Call LogError("[I] Empty vals for AccountType", 1)
			tmp = .Description
			Call LogError("[I] Empty vals for Description", 1)
			tmp = .Disabled
			Call LogError("[I] Empty vals for Disabled", 1)
			tmp = .Domain
			Call LogError("[I] Empty vals for Domain", 1)
			tmp = .Lockout
			Call LogError("[I] Empty vals for Lockout", 1)
			tmp = .Name
			Call LogError("[I] Empty vals for Name", 1)
			tmp = .PasswordChangeable
			Call LogError("[I] Empty vals for PasswordChangeable", 1)
			tmp = .PasswordExpires
			Call LogError("[I] Empty vals for PasswordExpires", 1)
			tmp = .PasswordRequired
			Call LogError("[I] Empty vals for PasswordRequired", 1)
			tmp = .SID
			Call LogError("[I] Empty vals for SID", 1)
			tmp = .SIDType
			Call LogError("[I] Empty vals for SIDType", 1)
			tmp = .Status
			Call LogError("[I] Empty vals for Status", 1)
			tmp = .PasswordAge
			Call LogError("[I] Empty vals for PasswordAge", 1)
			tmp = .PasswordMinimumLength
			Call LogError("[I] Empty vals for PasswordMinimumLength", 1)
			tmp = .LocalAccount
			Call LogError("[I] Empty vals for LocalAccount", 1)
			If Len(line) > 0 And grabLocalUsers = False Then
				grabLocalUsers = True
			End If
			File.WriteLine line
		End With
		Next
		File.Close
	Else 
		Wscript.Echo "[-] Skipping as " & objToScan & " is a domain controller..."
		grabLocalUsers = True
	End If
	Wscript.Echo "[D] Collecting data on local user accounts..."
	Call LogError("[D] Collecting data on local user accounts...", 0)
End Function

Function grabLocalGroups
	On Error Resume Next
	grabLocalGroups = False
	Call LogError("[S] Collecting data on local groups...", 0)
	Wscript.Echo "[S] Collecting data on local groups..."
	Dim File: Set File = CreateFile("LocalGroups")
	Dim o,q: Set q = objWMI.ExecQuery("Select * from Win32_Group  Where LocalAccount = True")
	Dim tmp
	For Each o In q
	With o
		Dim str: str = CSVLineFromArray(Array(objToScan, _
			""&.Caption, _
			""&.Description, _
			""&.Domain, _
			""&.Name, _
			""&.SID, _
			""&.SIDType))
		tmp = .Caption
		Call LogError("[I] Empty vals for Caption", 1)
		tmp = .Description
		Call LogError("[I] Empty vals for Description", 1)
		tmp = .Domain
		Call LogError("[I] Empty vals for Domain", 1)
		tmp = .Name
		Call LogError("[I] Empty vals for Name", 1)
		tmp = .SID
		Call LogError("[I] Empty vals for SID", 1)
		tmp = .SIDType
		Call LogError("[I] Empty vals for SIDType", 1)
		Dim r: Set r = GetObject("WinNT://" & objToScan & "/" & .Name)
		Call LogError("[I] Processing object: WinNT://" & objToScan & "/" & .Name, 1)
		For Each p In r
			File.WriteLine str & CSVLineFromArray(Array(Replace(p.adspath,"WinNT://","")))
		Next
		If r.Count = 0 Then
			File.WriteLine str & CSVLineFromArray(Array("Empty"))
		End If
	End With
	Next
	If q.Count = 0 Then
		File.WriteLine CSVLineFromArray(Array(objToScan, "No local groups", "0", "0", "0", "0", "0"))
	End If
	File.Close
	grabLocalGroups = True
	Wscript.Echo "[D] Collecting data on local groups..."
	Call LogError("[D] Collecting data on local groups...", 0)
End Function

Function grabMSI
	On Error Resume Next
	grabMSI = False
	Call LogError("[S] Collecting data on installed software...", 0)
	Wscript.Echo "[S] Collecting data on installed software..."
	Dim File: Set File = CreateFile("MSI")
	' MSIInstaller events
	' ID 1022 Product: %1 - Update '%2' installed successfully. 
	' ID 1033 = Installed software
	' ID 1034 = Removed software
	' ID 1035 = Reconfigured software
	' ID 1036 = Installed update
	' ID 11707 = Installed software (win2003)
	' ID 11724 = Removed software (win2003)
	Dim o,q: Set q = objWMI.ExecQuery("Select * from Win32_NTLogEvent Where LogFile='Application' And SourceName Like 'msiinstaller' AND " & _
								   "((EventCode>=1033 AND EventCode<=1036) OR EventCode=1022 OR EventCode=11707 OR EventCode=11724)")
	Call LogError("[I] ExecQuery: "&"Select * from Win32_NTLogEvent Where LogFile='Application' And SourceName Like 'msiinstaller' AND " & _
								   "((EventCode>=1033 AND EventCode<=1036) OR EventCode=1022 OR EventCode=11707 OR EventCode=11724)", 1)
	For Each o In q
	With o
		Dim strMessage: strMessage = Replace(Replace(Replace("" & ""&.Message, ";", ","), chr(13), ""), chr(10), "")
		Dim tmp
		If filterMS(strMessage) Then
			File.WriteLine CSVLineFromArray(Array(objToScan, _ 
				""&.Category, _
				""&.ComputerName, _
				""&.EventCode, _
				""&.EventIdentifier, _
				""&.EventType, _
				""&.Logfile, _
				strMessage, _ 
				""&.RecordNumber, _
				""&.SourceName, _
				""&.TimeGenerated, _
				""&.TimeWritten, _
				""&.Type, _
				""&.User _
			))
			grabMSI = true
			'tmp = .Category
			'Call LogError("[I] Empty vals for Category", 1)
			'tmp = .ComputerName
			'Call LogError("[I] Empty vals for ComputerName", 1)
			'tmp = .EventCode
			'Call LogError("[I] Empty vals for EventCode", 1)
			'tmp = .EventIdentifier
			'Call LogError("[I] Empty vals for EventIdentifier", 1)
			'tmp = .EventType
			'Call LogError("[I] Empty vals for EventType", 1)
			'tmp = .Logfile
			'Call LogError("[I] Empty vals for Logfile", 1)
			'tmp = .Message
			'Call LogError("[I] Empty vals for Message", 1)
			'tmp = .RecordNumber
			'Call LogError("[I] Empty vals for RecordNumber", 1)
			'tmp = .SourceName
			'Call LogError("[I] Empty vals for SourceName", 1)
			'tmp = .TimeGenerated
			'Call LogError("[I] Empty vals for TimeGenerated", 1)
			'tmp = .TimeWritten
			'Call LogError("[I] Empty vals for TimeWritten", 1)
			'tmp = .Type
			'Call LogError("[I] Empty vals for Type", 1)
			'tmp = .User
			'Call LogError("[I] Empty vals for User", 1)
			err.clear
		End if
	End With
	Next
	File.Close
	Wscript.Echo "[D] Collecting data on installed software..."
	Call LogError("[D] Collecting data on installed software...", 0)
End Function

Function grabLogonsLog
	On Error Resume Next
	grabLogonsLog = False
	Call LogError("[S] Collecting data on logons history...", 0)
	Wscript.Echo "[S] Collecting data on logons history..."
	' Audit events 
	' ID 4624 = Logon
	' ID 4648 = Logon (remote desktop etc)
	' ID 4672 = Special logon
	' ID 4634 = Logoff
	' ID 4647 = Logoff
	Dim o,q: Set q = objWMI.ExecQuery("Select * from Win32_NTLogEvent Where LogFile='Security' AND (EventCode=4624 or EventCode=4628 or EventCode=4672 or EventCode=4634 or EventCode=4647) AND Message LIKE '%WINLOGON.EXE%'")
	If Err <> 0 Then
		Call LogError("[I] ExecQuery: Select * from Win32_NTLogEvent Where LogFile='Security' AND (EventCode=4624 or EventCode=4628 or EventCode=4672 or EventCode=4634 or EventCode=4647) AND Message LIKE '%WINLOGON.EXE%'", 1)
	else
		grabLogonsLog = True
		Dim deleteFile: deleteFile = True
		Dim File: Set File = CreateFile("logonslog")
		Dim tmp
		For Each o In q
		With o
			If IsArray(.InsertionStrings) Then
				If InStr(UCase(""&.Message),"WINLOGON.EXE") > 0 And ""&.InsertionStrings(5)<>"-" And ""&.InsertionStrings(2)<>"-" And Left(""&.InsertionStrings(4),9) = "S-1-5-21-" Then
					File.WriteLine CSVLineFromArray(Array( _
						objToScan, _
						""&.Category, _
						""&.CategoryString, _
						""&.ComputerName, _
						""&.EventCode, _
						""&.EventIdentifier, _
						""&.EventType, _
						""&.RecordNumber, _
						""&.TimeGenerated, _
						""&.TimeWritten, _
						""&.InsertionStrings(1), _
						""&.InsertionStrings(2), _
						""&.InsertionStrings(5), _
						""&.InsertionStrings(6), _
						""&.InsertionStrings(11), _
						""&.InsertionStrings(17) _
					))
					tmp = .Category
					Call LogError("[I] Empty vals for Category", 1)
					tmp = .CategoryString
					Call LogError("[I] Empty vals for CategoryString", 1)
					tmp = .ComputerName
					Call LogError("[I] Empty vals for ComputerName", 1)
					tmp = .EventCode
					Call LogError("[I] Empty vals for EventCode", 1)
					tmp = .EventIdentifier
					Call LogError("[I] Empty vals for EventIdentifier", 1)
					tmp = .EventType
					Call LogError("[I] Empty vals for EventType", 1)
					tmp = .RecordNumber
					Call LogError("[I] Empty vals for RecordNumber", 1)
					tmp = .TimeGenerated
					Call LogError("[I] Empty vals for TimeGenerated", 1)
					tmp = .TimeWritten
					Call LogError("[I] Empty vals for TimeWritten", 1)
					tmp = .InsertionStrings(1)
					Call LogError("[I] Empty vals for InsertionStrings(1)", 1)
					tmp = .InsertionStrings(2)
					Call LogError("[I] Empty vals for InsertionStrings(2)", 1)
					tmp = .InsertionStrings(5)
					Call LogError("[I] Empty vals for InsertionStrings(5)", 1)
					tmp = .InsertionStrings(6)
					Call LogError("[I] Empty vals for InsertionStrings(6)", 1)
					tmp = .InsertionStrings(11)
					Call LogError("[I] Empty vals for InsertionStrings(11)", 1)
					tmp = .InsertionStrings(17)
					Call LogError("[I] Empty vals for InsertionStrings(17)", 1)
					deleteFile = False
				End If
			End If
		End With
		Next
		File.Close
		If deleteFile Then
			objFSO.DeleteFile(OutputPath & objToScan & fileDelimiter & "logonslog")
		End If 
	End If
	Wscript.Echo "[D] Collecting data on logons history..."
	Call LogError("[D] Collecting data on logons history...", 0)
End Function

Function grabLogons
	On Error Resume Next
	grabLogons = False
	Call LogError("[S] Collecting data on logons...", 0)
	Wscript.Echo "[S] Collecting data on logons..."
	Dim File: Set File = CreateFile("logons")
	Dim profiles: Call objRegistry.GetStringValue(HKLM, "SOFTWARE\MICROSOFT\WINDOWS NT\CURRENTVERSION\PROFILELIST", "PROFILESDIRECTORY", profiles)
	Call LogError("[I] Registry: "&"SOFTWARE\MICROSOFT\WINDOWS NT\CURRENTVERSION\PROFILELIST", 1)
	Dim path: path = Replace(Replace(Right(profiles,Len(profiles)-2)  & "\","\","\\"),"\\\","\\")
	Dim drive: drive = Left(profiles,2)
	' Skip network locations
	If Not (InStr(profiles, "\\") > 0) Then
		Dim o,q: Set q = objWMI.ExecQuery("Select * from Win32_Directory Where Hidden=False And Path='" & path & "' And drive='" & drive & "'")
		Call LogError("[I] ExecQuery: "&"Select * from Win32_Directory Where Hidden=False And Path='" & path & "' And drive='" & drive & "'", 1)
		Dim tmp
		For Each o In q
		With o
			If UCase(.FileName) <> "ALL USERS" Then
				File.WriteLine CSVLineFromArray(Array( _
					objToScan, _ 
					""&.CreationDate, _
					""&.CSName, _
					""&.FileName, _
					""&.InstallDate, _ 
					""&.LastAccessed, _
					""&.LastModified, _
					""&.Name _
				))
				tmp = .CreationDate
				Call LogError("[I] Empty vals for CreationDate", 1)
				tmp = .CSName
				Call LogError("[I] Empty vals for CSName", 1)
				tmp = .FileName
				Call LogError("[I] Empty vals for FileName", 1)
				tmp = .InstallDate
				Call LogError("[I] Empty vals for InstallDate", 1)
				tmp = .LastAccessed
				Call LogError("[I] Empty vals for LastAccessed", 1)
				tmp = .LastModified
				Call LogError("[I] Empty vals for LastModified", 1)
				tmp = .Name
				Call LogError("[I] Empty vals for Name", 1)
				grabLogons = true			
			End If
		End With
		Next
	End If
	File.Close
	Call LogError("[D] Collecting data on logons...", 0)
	Wscript.Echo "[D] Collecting data on logons..."
End Function

Function grabProcesses
	On Error Resume Next
	grabProcesses = False
	Wscript.Echo "[S] Collecting data on system processes..."
	Call LogError("[S] Collecting data on system processes...", 0)
	Dim File: Set File = CreateFile("processes")
	Dim o,q: Set q = objWMI.ExecQuery("Select * from Win32_Process")
	Call LogError("[I] ExecQuery: "&"Select * from Win32_Process", 1)
	Dim tmp
	For Each o In q
		With o
			File.WriteLine CSVLineFromArray(Array( _
				objToScan, _ 
				""&.Caption, _
				""&.CreationDate, _
				""&.CSName, _
				""&.Description, _
				""&.ExecutablePath, _
				""&.Name, _
				""&.WindowsVersion _
			))
			tmp = .Caption
			Call LogError("[I] Empty vals for Caption", 1)
			tmp = .CreationDate
			Call LogError("[I] Empty vals for CreationDate", 1)
			tmp = .CSName
			Call LogError("[I] Empty vals for CSName", 1)
			tmp = .Description
			Call LogError("[I] Empty vals for Description", 1)
			tmp = .ExecutablePath
			Call LogError("[I] Empty vals for ExecutablePath", 1)
			tmp = .Name
			Call LogError("[I] Empty vals for Name", 1)
			tmp = .WindowsVersion
			Call LogError("[I] Empty vals for WindowsVersion", 1)
			grabProcesses = true
		End With
	Next
	File.Close
	Call LogError("[D] Collecting data on system processes...", 0)
	Wscript.Echo "[D] Collecting data on system processes..."
End Function

Function grabServices
	On Error Resume Next
	Call LogError("[S] Collecting data on services...", 0)
	Wscript.Echo "[S] Collecting data on services..."
	grabServices = False
	Dim File: Set File = CreateFile("services")
	Dim o,q: Set q = objWMI.ExecQuery("Select * from Win32_Service")
	Call LogError("[I] ExecQuery Select * from Win32_Service", 1)
	Dim tmp
	For Each o In q
		With o
			If filterMS(.Caption) Or filterMS(.DisplayName) Then
				File.WriteLine CSVLineFromArray(Array( _
					objToScan, _
					""&.Caption, _
					""&.DisplayName, _
					""&.Name, _
					""&.PathName, _
					""&.Started, _
					""&.StartMode, _
					""&.StartName, _
					""&.State, _
					""&.SystemName _
				))
				grabServices = True
				tmp = .Caption
				Call LogError("[I] Empty vals for Caption", 1)
				tmp = .DisplayName
				Call LogError("[I] Empty vals for DisplayName", 1)
				tmp = .Name
				Call LogError("[I] Empty vals for Name", 1)
				tmp = .PathName
				Call LogError("[I] Empty vals for PathName", 1)
				tmp = .Started
				Call LogError("[I] Empty vals for Started", 1)
				tmp = .StartMode
				Call LogError("[I] Empty vals for StartMode", 1)
				tmp = .StartName
				Call LogError("[I] Empty vals for StartName", 1)
				tmp = .State
				Call LogError("[I] Empty vals for State", 1)
				tmp = .SystemName
				Call LogError("[I] Empty vals for SystemName", 1)
				If InStr(UCase(.Name), "MSSQL$") > 0 Then
					Call checkMSSQL(UCase(""&.Name), ""&.State)
				End If
				If InStr(UCase(.Name), "MSSQL") > 0 Then
					Call checkMSSQL2(UCase(""&.Name), ""&.State)
				End If
			End If
		End With
	Next
	File.Close
	Wscript.Echo "[D] Collecting data on services..."
	Call LogError("[D] Collecting data on services...", 0)
End Function

Function grabNetwork
	On Error Resume Next
	Call LogError("[S] Collecting network settings...", 0)
	Wscript.Echo "[S] Collecting network settings..."
	grabNetwork = False
	Dim File: Set File = CreateFile("network")
	Dim o,q: Set q = objWMI.ExecQuery("Select * from Win32_NetworkAdapterConfiguration Where IPEnabled=TRUE")
	Call LogError("[I] ExecQuery Select * from Win32_NetworkAdapterConfiguration Where IPEnabled=TRUE", 1)
	Dim tmp
	For Each o In q
	With o
		If Not IsNull(.IPAddress) Then
			Dim i, c: c = 0 
			For i=LBound(.IPAddress) To UBound(.IPAddress)
				File.WriteLine CSVLineFromArray(Array( _ 
					objToScan, _ 
					c, _
					""&.DHCPLeaseExpires, _
					""&.DHCPLeaseObtained, _
					""&.DNSHostName , _
					""&.IPAddress(i), _
					""&.MACAddress(i) _
				))
				c = c + 1
				grabNetwork = True
				tmp = .DHCPLeaseExpires
				Call LogError("[I] Empty vals for DHCPLeaseExpires", 1)
				tmp = .DHCPLeaseObtained
				Call LogError("[I] Empty vals for DHCPLeaseObtained", 1)
				tmp = .DNSHostName
				Call LogError("[I] Empty vals for DNSHostName", 1)
				tmp = .IPAddress(i)
				Call LogError("[I] Empty vals for IPAddress("&i&")", 1)
				tmp = .MACAddress(i)
				Call LogError("[I] Empty vals for MACAddress("&i&")", 1)
			Next
		End If
	End With
	Next
	File.Close
	Wscript.Echo "[D] Collecting network settings..."
	Call LogError("[D] Collecting network settings...", 0)
End Function

Function grabCPU
	On Error Resume Next
	Call LogError("[S] Collecting data on CPUs...", 0)
	Wscript.Echo "[S] Collecting data on CPUs..."
	grabCPU = False
	Dim arr(4), c: c = 0
	Dim File: Set File = CreateFile("cpu")
	Dim o,q: Set q = objWMI.ExecQuery("Select * from Win32_Environment Where name like '%processor%'")
	Call LogError("[I] ExecQuery Select * from Win32_Environment Where name like '%processor%'", 1)
	Dim tmp
	For Each o In q
	With o
		Select Case .Name 
			Case "NUMBER_OF_PROCESSORS":   arr(0) = ""&.VariableValue
			Case "PROCESSOR_ARCHITECTURE": arr(1) = ""&.VariableValue
			Case "PROCESSOR_IDENTIFIER":   arr(2) = ""&.VariableValue
			Case "PROCESSOR_REVISION":     arr(3) = ""&.VariableValue
		End Select
		Call LogError("[I] Empty value for "&.Name, 1)
	End With
	Next
	' CPU name from registry
	Call objRegistry.GetStringValue(HKLM, "HARDWARE\DESCRIPTION\System\CentralProcessor\0", "ProcessorNameString", arr(4))
	Call LogError("[I] registry HARDWARE\DESCRIPTION\System\CentralProcessor\0", 1)
	' String to be prepended processor information
	Dim tmpStr: tmpStr = """" & objToScan & """;" & CSVLineFromArray(arr)
	File.WriteLine tmpStr
	Set q = objWMI.ExecQuery("Select * from Win32_Processor")
	Call LogError("[I] ExecQuery Select * from Win32_Processor", 1)
	For Each o In q
	With o
		' CPU information
		File.WriteLine CSVLineFromArray(Array(c, _
			""&.AddressWidth, _
			""&.Architecture, _
			""&.Availability, _
			""&.Caption, _
			""&.CpuStatus, _
			""&.CreationClassName, _
			""&.DataWidth, _
			""&.Description, _
			""&.DeviceID, _
			""&.Family, _
			""&.Level, _
			""&.Manufacturer, _
			""&.Name, _
			""&.NumberOfCores, _
			""&.NumberOfLogicalProcessors, _
			""&.ProcessorId, _
			""&.ProcessorType, _
			""&.Revision, _
			""&.Role, _
			""&.StatusInfo, _
			""&.SystemName _
		))
		tmp = .AddressWidth
		Call LogError("[I] Empty vals for AddressWidth", 1)
		tmp = .Architecture
		Call LogError("[I] Empty vals for Architecture", 1)
		tmp = .Availability
		Call LogError("[I] Empty vals for Availability", 1)
		tmp = .Caption
		Call LogError("[I] Empty vals for Caption", 1)
		tmp = .CpuStatus
		Call LogError("[I] Empty vals for CpuStatus", 1)
		tmp = .CreationClassName
		Call LogError("[I] Empty vals for CreationClassName", 1)
		tmp = .DataWidth
		Call LogError("[I] Empty vals for DataWidth", 1)
		tmp = .Description
		Call LogError("[I] Empty vals for Description", 1)
		tmp = .DeviceID
		Call LogError("[I] Empty vals for DeviceID", 1)
		tmp = .Family
		Call LogError("[I] Empty vals for Family", 1)
		tmp = .Level
		Call LogError("[I] Empty vals for Level", 1)
		tmp = .Manufacturer
		Call LogError("[I] Empty vals for Manufacturer", 1)
		tmp = .Name
		Call LogError("[I] Empty vals for Name", 1)
		tmp = .NumberOfCores
		Call LogError("[I] Empty vals for NumberOfCores", 1)
		tmp = .NumberOfLogicalProcessors
		Call LogError("[I] Empty vals for NumberOfLogicalProcessors", 1)
		tmp = .ProcessorId
		Call LogError("[I] Empty vals for ProcessorId", 1)
		tmp = .ProcessorType
		Call LogError("[I] Empty vals for ProcessorType", 1)
		tmp = .Revision
		Call LogError("[I] Empty vals for Revision", 1)
		tmp = .Role
		Call LogError("[I] Empty vals for Role", 1)
		tmp = .StatusInfo
		Call LogError("[I] Empty vals for StatusInfo", 1)
		tmp = .SystemName
		Call LogError("[I] Empty vals for SystemName", 1)
		grabCPU = True
		c = c + 1
	End With
	Next
	File.Close
	Wscript.Echo "[D] Collecting data on CPUs..."
	Call LogError("[D] Collecting data on CPUs...", 0)
End Function

Function grabHardware
	On Error Resume Next
	Call LogError("[S] Collecting hardware information...", 0)
	Wscript.Echo "[S] Collecting hardware information..."
	grabHardware = False
	Dim File: Set File = CreateFile("hw")
	Dim line: line = ""
	Dim tmpArr(27) 
	tmpArr(0) = objToScan
	With WMIQuery(objWMI, "Select * from Win32_ComputerSystem")
		tmpArr(1)  = ""&.Caption
		Call LogError("[I] Empty vals for Caption", 1)
		tmpArr(2)  = ""&.DNSHostName
		Call LogError("[I] Empty vals for DNSHostName", 1)
		tmpArr(3)  = ""&.DomainRole
		Call LogError("[I] Empty vals for DomainRole", 1)
		tmpArr(4)  = ""&.Manufacturer
		Call LogError("[I] Empty vals for Manufacturer", 1)
		tmpArr(5)  = ""&.Model
		Call LogError("[I] Empty vals for Model", 1)
		tmpArr(6)  = ""&.Name
		Call LogError("[I] Empty vals for Name", 1)
		tmpArr(7)  = ""&.NetworkServerModeEnabled
		Call LogError("[I] Empty vals for NetworkServerModeEnabled", 1)
		tmpArr(8)  = ""&.NumberOfLogicalProcessors
		Call LogError("[I] Empty vals for NumberOfLogicalProcessors", 1)
		tmpArr(9)  = ""&.NumberOfProcessors
		Call LogError("[I] Empty vals for NumberOfProcessors", 1)
		tmpArr(10) = ""&.PartOfDomain
		Call LogError("[I] Empty vals for PartOfDomain", 1)
		tmpArr(11) = ""&.PCSystemType
		Call LogError("[I] Empty vals for PCSystemType", 1)
		tmpArr(12) = ""&.TotalPhysicalMemory
		Call LogError("[I] Empty vals for TotalPhysicalMemory", 1)
	End With
	With WMIQuery(objWMI, "Select * from Win32_BIOS")
		tmpArr(13) = ""&.Caption
		Call LogError("[I] Empty vals for Caption", 1)
		tmpArr(14) = ""&.Description
		Call LogError("[I] Empty vals for Description", 1)
		tmpArr(15) = ""&.Manufacturer
		Call LogError("[I] Empty vals for Manufacturer", 1)
		tmpArr(16) = ""&.Name
		Call LogError("[I] Empty vals for Name", 1)
		tmpArr(17) = ""&.SerialNumber
		Call LogError("[I] Empty vals for SerialNumber", 1)
		tmpArr(18) = ""&.Version
		Call LogError("[I] Empty vals for Version", 1)
	End With
	With WMIQuery(objWMI, "Select * from Win32_BaseBoard")
		tmpArr(19) = ""&.Caption
		Call LogError("[I] Empty vals for Caption", 1)
		tmpArr(20) = ""&.Manufacturer
		Call LogError("[I] Empty vals for Manufacturer", 1)
		tmpArr(21) = ""&.Product
		Call LogError("[I] Empty vals for Product", 1)
		tmpArr(22) = ""&.SerialNumber
		Call LogError("[I] Empty vals for SerialNumber", 1)
	End With
	With WMIQuery(objWMI, "Select * from Win32_SystemEnclosure")
		line = line & CSV(Array(.SerialNumber, .Model, .Name, .Manufacturer))
		tmpArr(23) = ""&.Manufacturer
		Call LogError("[I] Empty vals for Manufacturer", 1)
		tmpArr(24) = ""&.SerialNumber
		Call LogError("[I] Empty vals for SerialNumber", 1)
		tmpArr(25) = ""&.Name
		Call LogError("[I] Empty vals for Name", 1)
		tmpArr(26) = ""&.Model
		Call LogError("[I] Empty vals for Model", 1)
	End With
	File.WriteLine CSVLineFromArray(tmpArr)
	If tmpArr.count > 0 Then
		grabHardware = true
	End if
	File.Close
	Wscript.Echo "[D] Collecting hardware information..."
	Call LogError("[D] Collecting hardware information...", 0)
End Function

Function grabSoftware
	On Error Resume Next
	Call LogError("[S] Collecting information about installed software from the registry...", 0)
	Wscript.Echo "[S] Collecting information about installed software from the registry..."
	grabSoftware = False
	Dim File: Set File = CreateFile("sw")
	Dim keys: keys = Array("SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\", _
						   "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\")
	Dim key: For Each key In keys 
		Dim subKey, subKeys
		Call objRegistry.EnumKey(HKLM, key, subKeys)
		Call LogError("[I] Registry: "&key, 1)
		If IsArray(subKeys) Then
			grabSoftware = true
			For Each subKey In subKeys
				Dim vals(6)
				If objRegistry.GetStringValue(HKLM, key & subKey, "DisplayName", vals(0)) <> 0 Then
					Call objRegistry.GetStringValue(HKLM, key & subKey, "QuietDisplayName", vals(0))
					Call LogError("[I] Registry: "&key&subKey, 1)
				End If
				If Trim(vals(0)) <> "" Then
					Call objRegistry.GetStringValue (HKLM, key & subKey, "InstallDate", vals(1))
					Call objRegistry.GetDWORDValue  (HKLM, key & subKey, "VersionMajor", vals(2))
					Call objRegistry.GetDWORDValue  (HKLM, key & subKey, "VersionMinor", vals(3))
					Call objRegistry.GetStringValue (HKLM, key & subKey, "DisplayVersion", vals(4))
					Call objRegistry.GetStringValue (HKLM, key & subKey, "Publisher", vals(5))
					' Version
					vals(2) = Trim(vals(2)) & "." & Trim(vals(3))
					If vals(2) = "." Then: vals(2) = "": End If
					If filterMS(vals(0)) Or filterMS(vals(5)) Then
						File.WriteLine CSVLineFromArray(Array(objToScan, vals(0), vals(1), vals(2), vals(4), vals(5)))
					End If
				End If
			Next
		End If
		Dim subUser, subUsers
		Call objregistry.EnumKey(HKU, "", subUsers)
		For Each subUser In subUsers
			Call objRegistry.EnumKey(HKU, subUser & "\" & key, subKeys)
			If IsArray(subKeys) Then
				grabSoftware = True
				For Each subKey In subKeys
					If objRegistry.GetStringValue(HKU, subUser & "\" & key & subKey, "DisplayName", vals(0)) <> 0 Then
						Call objRegistry.GetStringValue(HKU, subUser & "\" & key & subKey, "QuietDisplayName", vals(0))
						Call LogError("[I] Registry: "&subUser & "\" & key & subKey, 1)
					End If
					If Trim(vals(0)) <> "" Then
						Call objRegistry.GetStringValue (HKU, subUser & "\" & key & subKey, "InstallDate", vals(1))
						Call objRegistry.GetDWORDValue  (HKU, subUser & "\" & key & subKey, "VersionMajor", vals(2))
						Call objRegistry.GetDWORDValue  (HKU, subUser & "\" & key & subKey, "VersionMinor", vals(3))
						Call objRegistry.GetStringValue (HKU, subUser & "\" & key & subKey, "DisplayVersion", vals(4))
						Call objRegistry.GetStringValue (HKU, subUser & "\" & key & subKey, "Publisher", vals(5))
						' Version
						vals(2) = Trim(vals(2)) & "." & Trim(vals(3))
						If vals(2) = "." Then: vals(2) = "": End If
						If filterMS(vals(0)) Or filterMS(vals(5)) Then
							File.WriteLine CSVLineFromArray(Array(objToScan, vals(0), vals(1), vals(2), vals(4), vals(5)))
						End If
					End If
				Next
			End If
		Next
	Next
	File.Close
	Call LogError("[D] Collecting information about installed software from the registry...", 0)
	Wscript.Echo "[D] Collecting information about installed software from the registry..."
End Function

Function grabLocalMembership
	On Error Resume Next
	
	Call LogError("[S] Collecting information about local groups & users...", 0)
	Wscript.Echo "[S] Collecting information about local groups & users..."
	grabLocalMembership = False
	Dim File: Set File = CreateFile("membership")
	Dim colGroups: Set colGroups = GetObject("WinNT://" & objToScan & "")
	colGroups.Filter = Array("Group")
	Dim strTestString: strTestString = "/" & objToScan & "/"
	Dim objGroup, membersCount, objUser
	For Each objGroup In colGroups
		membersCount = 0
		For Each objUser in objGroup.Members
			membersCount = membersCount + 1
			If InStr(objUser.AdsPath, strTestString) Then
				File.WriteLine CSVLineFromArray(Array(objToScan, objGroup.Name, objUser.Name, "Local", objUser.AdsPath))
			Else
				File.WriteLine CSVLineFromArray(Array(objToScan, objGroup.Name, objUser.Name, "Domain", objUser.AdsPath))
			End If
		Next
		if (membersCount = 0) then
			File.WriteLine CSVLineFromArray(Array(objToScan, objGroup.Name, "Empty", "Empty", "Empty"))
		end if
	Next
	grabLocalMembership = True
	File.Close
	Call LogError("[D] Collecting information about local groups & users...", 0)
	Wscript.Echo "[D] Collecting information about local groups & users..."
End Function

Function grabDNS
	On Error Resume Next
	Call LogError("[S] Collecting DNS information...", 0)
	Wscript.Echo "[S] Collecting DNS information..."
	grabDNS = True
	Call initWMIDNS()
	If objWMIDNS Is Nothing Then: Exit Function: End If
	Dim File: Set File = CreateFile("dns")
	Dim o,q: Set q = objWMIDNS.ExecQuery("Select * From MicrosoftDNS_AType")
	Call LogError("[I] ExecQuery Select * From MicrosoftDNS_AType", 1)
	Dim tmp
	For Each o In q
	With o
		File.WriteLine CSVLineFromArray(Array(objToScan, ""&.OwnerName, "", ""&.DomainName))
		tmp = .OwnerName
		Call LogError("[I] Empty vals for OwnerName", 1)
		tmp = .DomainName
		Call LogError("[I] Empty vals for DomainName", 1)
	End With
	Next
	Set q = objWMIDNS.ExecQuery("Select * From MicrosoftDNS_CNAMEType")
	Call LogError("[I] ExecQuery Select * From MicrosoftDNS_CNAMEType", 1)
	For Each o In q
	With o
		File.WriteLine CSVLineFromArray(Array(objToScan, ""&.OwnerName, ""&.PrimaryName, ""&.DomainName))
		tmp = .OwnerName
		Call LogError("[I] Empty vals for OwnerName", 1)
		tmp = .PrimaryName
		Call LogError("[I] Empty vals for PrimaryName", 1)
		tmp = .DomainName
		Call LogError("[I] Empty vals for DomainName", 1)
	End With
	Next
	File.Close
	Wscript.Echo "[D] Collecting DNS information..."
	Call LogError("[D] Collecting DNS information...", 0)
End Function

Sub scanObject
	On Error Resume Next
	Err.clear
	Call initLog()

	If objFSO.FileExists(OutputPath & objToScan & fileDelimiter & "scanned") Then
		Call LogError("[*] " & objToScan & " has been scanned. Skipping.", 0)
		Wscript.Echo "[*] " & objToScan & " has been scanned. Skipping."
		Wscript.Quit
	End If
	Dim isOk: isOk =  True
	Call LogError("[S] Deleting files from the previous scan.", 0)
	Call objFSO.DeleteFile(OutputPath & objToScan & fileDelimiter & "offline", true)
	Call LogError("[I] offline", 0)
	Call objFSO.DeleteFile(OutputPath & objToScan & fileDelimiter & "error", true)
	Call LogError("[I] error", 0)
	Call objFSO.DeleteFile(OutputPath & objToScan & fileDelimiter & "noaccess", true)
	Call LogError("[I] noaccess", 0)
	Call objFSO.DeleteFile(OutputPath & objToScan & fileDelimiter & "workstation", true)
	Call LogError("[I] workstation", 0)
	Call objFSO.DeleteFile(OutputPath & objToScan & fileDelimiter & "highload", true)
	Call LogError("[I] highload", 0)
	Call objFSO.DeleteFile(OutputPath & objToScan & fileDelimiter & "WMIError", true)
	Call LogError("[I] wmierror", 0)
	Call objFSO.DeleteFile(OutputPath & objToScan & fileDelimiter & "cpu", true)
	Call LogError("[I] cpu", 0)
	Call objFSO.DeleteFile(OutputPath & objToScan & fileDelimiter & "hw", true)
	Call LogError("[I] hw", 0)
	Call objFSO.DeleteFile(OutputPath & objToScan & fileDelimiter & "localgroups", true)
	Call LogError("[I] localgroups", 0)
	Call objFSO.DeleteFile(OutputPath & objToScan & fileDelimiter & "localusers", true)
	Call LogError("[I] localusers", 0)
	Call objFSO.DeleteFile(OutputPath & objToScan & fileDelimiter & "logons", true)
	Call LogError("[I] logons", 0)
	Call objFSO.DeleteFile(OutputPath & objToScan & fileDelimiter & "logonslog", true)
	Call LogError("[I] logonslog", 0)
	Call objFSO.DeleteFile(OutputPath & objToScan & fileDelimiter & "mssql", true)
	Call LogError("[I] mssql", 0)
	Call objFSO.DeleteFile(OutputPath & objToScan & fileDelimiter & "network", true)
	Call LogError("[I] network", 0)
	Call objFSO.DeleteFile(OutputPath & objToScan & fileDelimiter & "online", true)
	Call LogError("[I] online", 0)
	Call objFSO.DeleteFile(OutputPath & objToScan & fileDelimiter & "os", true)
	Call LogError("[I] os", 0)
	Call objFSO.DeleteFile(OutputPath & objToScan & fileDelimiter & "processes", true)
	Call LogError("[I] processes", 0)
	Call objFSO.DeleteFile(OutputPath & objToScan & fileDelimiter & "scanned", true)
	Call LogError("[I] scanned", 0)
	Call objFSO.DeleteFile(OutputPath & objToScan & fileDelimiter & "services", true)
	Call LogError("[I] services", 0)
	Call objFSO.DeleteFile(OutputPath & objToScan & fileDelimiter & "sw", true)
	Call LogError("[I] sw", 0)
	Call objFSO.DeleteFile(OutputPath & objToScan & fileDelimiter & "msi", true)
	Call LogError("[I] msi", 0)
	Call objFSO.DeleteFile(OutputPath & objToScan & fileDelimiter & "guid", true)
	Call LogError("[I] GUID", 0)
	Call objFSO.DeleteFile(OutputPath & objToScan & fileDelimiter & "membership", true)
	Call LogError("[I] membership", 0)
	
	Call LogError("[D] Deleting files from the previous scan.", 0)
	Call initWMI()
	If TypeName(objWMI) = "Boolean" Then
		isOk = False
	End If
	If isOk And Ping(objToScan) Then
		Dim onlineFile: Set onlineFile = CreateFile("online")
		onlineFile.WriteLine Now
		onlineFile.close
	End If
	' Local OS
	If isOk Then
		Call getDomainRole()
		Dim load: Set load = CPUload() 
		If load > maxTargetCPUload Then
			Call CreateFile("highload")
			Wscript.Echo "[-] " & objToScan & " is overloaded at the moment with " & load & "%. Skipping..."
			Wscript.Quit
		End If

		
		call grabLocalMembership()
		Call grabOS()
		Call grabCPU()
		Call grabSoftware()
		Call grabHardware()
		Call grabProcesses()
		Call grabServices()
		Call grabNetwork()
		Call grabMSI()
		Call grabGUID()
		Call grabLocalUsers()
		Call grabLocalGroups()
		Call grabLogons()
		Call grabLogonsLog()
		Call grabBiztalk()
		Call grabISAForefrontTMG()
		Call grabSharePoint()
	End If
	If isOk Then
		Dim File: Set File = CreateFile("scanned")
		File.WriteLine Now
	End If
	Wscript.Echo "[+] Done for " & objToScan
	Call closeLog
	Wscript.Quit
End Sub

Sub Process
	On Error Resume Next
	' By default the script reads Objectslist.txt file and goes through it otherwise shows menu to select usage mode
	
	If Wscript.Arguments.Count = 0 Then
		Dim choice
		Do
			Wscript.Echo
			Wscript.Echo "[+] Please select usage mode and press ENTER"
			Wscript.Echo "    1  - to scan local system"
			Wscript.Echo "    2  - to scan each system from Objectslist.txt file"
			Wscript.Echo "    3  - to create list of systems based on Active Directory"
			Wscript.Echo "    4  - to create and process list of systems"
			Wscript.Echo "    5  - to scan specific object"
			Wscript.Echo "    6  - to extract Domain Users"
			Wscript.Echo "    7  - to extract Domain Groups"
			
			choice = UCase(Wscript.StdIn.ReadLine)
		Loop While (choice <> "1" and choice <> "2" and choice <> "3" and choice <> "4" and choice <> "5" and choice <> "6" and choice <> "7")
		If choice = "1" Then 
			cmdArgs("/PROCESSLOCAL") = ""
			objToScan = getLocalName 
			isLocalMode = True
			Call scanObject()
		ElseIf choice = "2" Then 
			cmdArgs("/PROCESSLIST") = true 
		ElseIf choice = "3" Then  
			cmdArgs("/GETLISTFROMAD") = true 
		ElseIf choice = "4" Then 
			cmdArgs("/PROCESSAD") = true
		ElseIf choice = "5" Then 
			Wscript.Echo "[+] Please specify object name:"	
			cmdArgs("/PROCESSTARGET:") = UCase(Wscript.StdIn.ReadLine)
		ElseIf choice = "6" Then 
			cmdArgs("/GETDOMAINUSERS") = true 
		ElseIf choice = "7" Then 
			cmdArgs("/GETDOMAINGROUPS") = true
		End If
	End If
	If cmdArgs("/PROCESSTARGET:") <> False And objToScan <> False Then
		objToScan = UCase(cmdArgs("/PROCESSTARGET:"))
		isLocalMode = False
		Call scanObject()
		Wscript.Quit
	End If
	If cmdArgs("/GETLISTFROMAD") = True Then
		Call initDSE()
		Call initLog()
		Call grabObjects()
		Wscript.Quit
	End If
	If cmdArgs("/GETDOMAINUSERS") = True Then
		Call initDSE()
		Call initLog()
		Call GetDomainUsers()
		Wscript.Quit
	End If
	If cmdArgs("/GETDOMAINGROUPS") = True Then
		Call initDSE()
		Call initLog()
		Call GetDomainGroups()
		Wscript.Quit
	End If
	If cmdArgs("/PROCESSAD") = True Then
		Call initDSE()
		Call initLog()
		Call grabObjects()
		Call grabExchange()
		Call processList()
		Wscript.Quit
	End If
	If cmdArgs("/PROCESSLIST") = True then 
		Call initDSE()
		Call initLog()
		Call grabExchange()
		Call processList()
		Wscript.Quit
	End If
End Sub

Sub Start
	On Error Resume Next
	Err.Clear
	Wscript.Echo "[I] Starting up..."
	Call Init()
	Call parseCmdArgs()
	Call Process()
End Sub

Call Start()