Dim oFSO : Set oFSO = CreateObject("Scripting.FileSystemObject")
Dim oShell : Set oShell = createobject("wscript.shell")
SetLocale "en-us"  ' Do not remove!

sWinDir = oFSO.GetSpecialFolder(0)
sPrograms = oShell.SpecialFolders("AllUsersPrograms")

iOsVer = GetOsVersionNumber

'================================================================
' Uninstall games from Windows XP
'================================================================
If iOsVer = 5.1 Then
' WinXP
If (oFSO.FolderExists(sPrograms & "\games")) Then
  ' Create file for uninstalling games
  Set f = oFSO.CreateTextFile(sWinDir & "\inf\wmdtocm.txt", True)
  f.WriteLine("[Components]")
  f.WriteLine("freecell=off")
  f.WriteLine("hearts=off")
  f.WriteLine("minesweeper=off")
  f.WriteLine("msnexplr=off")
  f.WriteLine("pinball=off")
  f.WriteLine("solitaire=off")
  f.WriteLine("spider=off")
  f.WriteLine("zonegames=off")
  f.Close
  oShell.Run "sysocmgr.exe /i:%windir%\inf\sysoc.inf" _
           & " /u:""%windir%\inf\wmdtocm.txt"" /q /r", 0, True

  oShell.Run "%Comspec% /C RD /S /Q " _
            & Chr(34) & sPrograms & "\games" & Chr(34), 0, True
End If
End If

'================================================================
' Uninstall games from Windows 2000
'================================================================

If iOsVer = 5 Then
' Win2k
If (oFSO.FolderExists(sPrograms & "\accessories\games")) Then
  ' Create file for uninstalling games
  Set f = oFSO.CreateTextFile(sWinDir & "\inf\wmdtocm.txt", True)
  f.WriteLine("[Components]")
  f.WriteLine("freecell=off")
  f.WriteLine("minesweeper=off")
  f.WriteLine("pinball=off")
  f.WriteLine("solitaire=off")
  f.Close
  oShell.Run "sysocmgr.exe /i:%windir%\inf\sysoc.inf" _
           & " /u:""%windir%\inf\wmdtocm.txt"" /q /r", 0, True
End If
End If



Function GetOsVersionNumber()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Determines OS by reading reg val & comparing to known values
' OS version number returned as number of type double:
'    Windows 2k:   5
'    Windows XP:   5.1
'    Windows Server 2003: 5.2
'    Windows x:  >5.2

' Note: Decimal point returned is based on the Locale setting
' of the computer, so it might be returned as 5,1 as well.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  Dim sOStype, sOSversion

  On Error Resume Next
  sOStype = oShell.RegRead(_
    "HKLM\SYSTEM\CurrentControlSet\Control\ProductOptions\ProductType")
  If Err.Number<>0 Then
    ' Hex(Err.Number)="80070002"
    ' - Could not find this key, OS must be Win9x
    Err.Clear

    sOStype = oShell.RegRead(_
      "HKLM\SOFTWARE\Microsoft\Windows" & _
      "\CurrentVersion\VersionNumber")

    Select Case sOStype
      Case "4.00.950"
        sOSversion = 1   ' Windows 95A
      Case "4.00.1111"
        Dim sSubVersion
        sSubVersion = oShell.RegRead(_
          "HKLM\SOFTWARE\Microsoft\Windows" & _
          "\CurrentVersion\SubVersionNumber")
        Select Case sSubVersion
          Case " B"
            sOSversion = 1   ' Windows 95B
          Case " C"
            sOSversion = 1   ' Windows 95C
          Case Else
            sOSversion = 1   ' Unknown Windows 95
        End Select
      Case "4.03.1214"
        sOSversion = 1   ' Windows 95B
      Case "4.10.1998"
        sOSversion = 2   ' Windows 98
      Case "4.10.2222"
        sOSversion = 2   ' Windows 98SE
      Case "4.90.3000"
        sOSversion = 3   ' Windows Me
      Case Else
        sOSversion = 1   ' Unknown W9x/Me
    End Select
  Else  ' OS is NT based
    sOSversion = oShell.RegRead(_
      "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\CurrentVersion")
    If Err.Number<>0 Then
      GetOsVersion = "Unknown NTx"
      ' Could not determine NT version
      Exit Function  ' >>>
    End If
  End If

  ' Setting Locale to "en-us" to be indifferent to country settings.
  ' CDbl might err else
  SetLocale "en-us"
  GetOsVersionNumber = CDbl(sOSversion)
End Function 