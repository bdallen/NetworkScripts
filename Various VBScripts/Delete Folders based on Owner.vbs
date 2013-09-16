' ---------------------------------
' | Clean Sentrian Updates Script |
' ---------------------------------
' 16/09/2013 - BA - Initial Write

CleanSentrianUpdates "c:\"

Sub CleanSentrianUpdates(folderspec)
    Dim fs, f, f1, fc, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFolder(folderspec)
    Set fc = f.SubFolders
    For Each f1 in fc
		' Sanity Checks
		Select Case f1
		Case "WINDOWS"
		Case "Program Files"
		Case "Program Files (x86)"
		Case Else
			If folderOwner("C:\" & f1.name) = "sentrianadmin" Then
			s = s & f1.name
			s = s &  vbCrLf
			fs.DeleteFolder("C:\" & f1.name)
			End If
		End Select
    Next
    'MsgBox s
End Sub

Function folderOwner(sFolder)
 On Error Resume Next
  Dim oShell, oFSO, oExec, sText, iStart, iEnd
  Set oShell = CreateObject("WScript.Shell")
  Set oFSO = CreateObject("Scripting.FileSystemObject")
  If oFSO.FolderExists(sFolder) Then
   Set oExec = oShell.Exec("cmd /c dir /q " & sFolder & " | find /I ""<DIR>""")
   Do While oExec.StdOut.AtEndOfStream <> True
    sText = oExec.StdOut.ReadAll
   Loop
   iEnd = InStr(sText,"."&vbCrLf)
   iStart = InStrRev(sText,">",InStr(sText,"."&vbCrLf)) + 1
   If iEnd>iStart Then
    folderOwner = Trim(Mid(sText,iStart,iEnd-iStart))
	user = Split(folderOwner, "\", -1, 1)
	folderOwner = user(1)
   Else
    folderOwner = "Unable to determine"
   End If
  Else
   folderOwner = "Folder not found"
  End If
 End Function