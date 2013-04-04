' *********************************************
' * Create User Home Directory on First Logon *
' *********************************************
' This VBS Script should run at users logon to create their home directory
' if it does not aready exist.
'
' 04/04/2013 - B.Allen - Initial Write
'
' On Error Resume Next

Dim objNet,objFSO, objFolder, strDirectory, UserName

' Create the netwrk object and get the username
Set objNet = CreateObject("WScript.Network")
UserName = objNet.UserName

' Change me when needing to point to another location
strDirectory = "\\xxx\users$\" & UserName

' Create the File System Object
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Note If..Exists. Then, Else ... End If construction
If objFSO.FolderExists(strDirectory) Then
	WScript.Quit
ELSE
	Set objFolder = objFSO.CreateFolder(strDirectory)
	Set objNet = Nothing 'Destroy the Object to free the Memory
End If

WScript.Quit
Set objNet = Nothing 'Destroy the Object to free the Memory