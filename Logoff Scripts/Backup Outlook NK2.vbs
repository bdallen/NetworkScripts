'This script will copy the users NK2 file to a file server.

On Error Resume Next

Const OverwriteExisting = TRUE
Const OverWriteFiles = TRUE
strComputer = "."
strDestFolder = "\\xxxxxxxx\users$\Backups\"  'Do Not Remove the trailing backslash! 
strDestFolder2 = ""
strDestFolder3 = ""
strLocalFolder = "\Application Data\Microsoft\Outlook\"
strReplace = ""
strUserName = ""

'Call Functions
UserCheck()
ProfileCheck()
FolderCheck()
CopyNK2()
'WScript.Echo "Done..."
WScript.Quit



Function UserCheck()
	Dim objNetwork

	Set objNetwork = CreateObject("WScript.Network")
	strUserName = objNetwork.Username
End Function



Function ProfileCheck()
	Set objFSO = CreateObject("Scripting.FileSystemObject")

	If objFSO.FolderExists(strDestFolder & strUserName) Then
		'WScript.Echo "Folder does exist."
		'WScript.Echo strDestFolder & strUserName & "\" & "NK2"
	Else
		'WScript.Echo "Folder does not exist."
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		Set objFolder = objFSO.CreateFolder(strDestFolder & "\" & strUserName)
			'WScript.Echo strDestFolder & strUserName
	End If
End Function



Function FolderCheck()
	Set objFSO = CreateObject("Scripting.FileSystemObject")

	If objFSO.FolderExists(strDestFolder & strUserName & "\" & "NK2") Then
		'WScript.Echo "Folder does exist."
		'WScript.Echo strDestFolder & strUserName & "\" & "NK2"
		strDestFolder2 = strDestFolder & strUserName & "\" & "NK2"
	Else
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		Set objFolder = objFSO.CreateFolder(strDestFolder & "\" & strUserName & "\" & "NK2")
			strDestFolder2 = strDestFolder & strUserName & "\" & "NK2"
	End If
End Function



Sub CopyNK2()
	Set objShell = CreateObject("Wscript.Shell")
		strProfile = objShell.ExpandEnvironmentStrings("%userprofile%") 
		'Wscript.Echo strProfile 
	Set objFSO = CreateObject("Scripting.FileSystemObject")
		objFSO.CopyFile strProfile & strLocalFolder & "*.nk2" , strDestFolder2 , OverwriteExisting
End Sub