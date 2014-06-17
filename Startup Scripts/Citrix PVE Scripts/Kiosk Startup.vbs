' -------------------------
' | Setup PVS Environment |
' -------------------------
'
'
' 1) Logs out the Information regarding the machine
' 2) Setup the system to autologon to an account that is generated using the machine name - KIOSK Systems
'
'
' Error Handling or Lack Thereof
'On Error Resume Next

' Get Required Variables
Set wshShell = WScript.CreateObject( "WScript.Shell" )
Set objFSO = CreateObject("Scripting.FileSystemObject")
strComputerName = wshShell.ExpandEnvironmentStrings( "%COMPUTERNAME%" )

' Setup Logging
Set fLog = objFSO.OpenTextFile("\\xxxxx\log$\" & strComputerName & "-log.txt", 8, True)

' Tell the log we booted up
WriteLog "Workstation Bootup"
WriteLog "Workstation In OU - " & GetPCOU

' Setup the Autologon
SetAutoLogon()

' Set the Autologon Information for PC
Sub SetAutoLogon
	' Set Workstation Autologon
	strWksUsername = "domain\u_" & LCase(strComputerName)
	strWksPassword = LCase(strComputerName)
	
	keyDefaultUsername = "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\DefaultUserName"
	WshShell.RegWrite keyDefaultUsername,strWksUsername,"REG_SZ"
	
	keyDefaultPassword = "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\DefaultPassword"
	WshShell.RegWrite keyDefaultPassword,strWksPassword,"REG_SZ"
	
	WriteLog "Processed Autologon details with Username: " & strWksUsername
	
End Sub

' Output Information To Log File
Sub WriteLog(strMessage)
	fLog.WriteLine Now & " : " & strMessage
End Sub

' Get's the OU the PC Belongs To From AD
Function GetPCOU
	Set objSysInfo = CreateObject("ADSystemInfo")
	strComputer = objSysInfo.ComputerName

	Set objComputer = GetObject("LDAP://" & strComputer)

	arrOUs = Split(objComputer.Parent, ",")
	arrMainOU = Split(arrOUs(0), "=")

	GetPCOU = arrMainOU(1)
End Function