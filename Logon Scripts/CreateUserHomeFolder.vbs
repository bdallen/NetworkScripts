' *********************************************
' * Create User Home Directory on First Logon *
' *********************************************
' This VBS Script should run at users logon to create their home directory
' if it does not aready exist.
'
' 04/04/2013 - B.Allen - Initial Write
' 04/04/2013 - B.Allen - Made it create and check on local profile to see if it has been processed already (Used for Slow Links)
'						 Creates an empty folder in the user profile directory called .hdirtoken
'
'On Error Resume Next

Dim objNet,objFSO, objFolder, strDirectory, UserName, strToken, strProfile, objToken, objFToken, wshShell

' Get the Userprofile Environment Variable
Set wshShell = CreateObject( "WScript.Shell" )
strProfile = wshShell.ExpandEnvironmentStrings( "%USERPROFILE%" )

' HomeDirectory Creation Token
strToken = strProfile & "\.hdirtoken"

' See if this has been run already, if so stop processing
Set objToken = CreateObject("Scripting.FileSystemObject")
If objToken.FolderExists(strToken) Then
	WScript.Quit
End If

' Create the netwrk object and get the username
Set objNet = CreateObject("WScript.Network")
UserName = objNet.UserName

' Change me when needing to point to another location
strDirectory = "\\xxxxxxx\users$\" & UserName

' Create the File System Object
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Note If..Exists. Then, Else ... End If construction
If objFSO.FolderExists(strDirectory) Then
	WScript.Quit
ELSE
	Set objFolder = objFSO.CreateFolder(strDirectory)
	Set objFToken = objToken.CreateFolder(strToken)
	Set objNet = Nothing 'Destroy the Object to free the Memory
End If

WScript.Quit
Set objNet = Nothing 'Destroy the Object to free the Memory