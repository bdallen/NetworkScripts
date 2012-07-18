' Setup Branch Computers
' ----------------------
' 18-07-2012 : BAllen : Initial Write

' Environment Varialbes
set oShell = WScript.CreateObject( "WScript.Shell" )
CompName = oShell.ExpandEnvironmentStrings("%COMPUTERNAME%")

' File Operation Constants
Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8

' Port Map File
Const PMAPFILE = "\\xxx\PrinterPorts.txt"
Const SETPRINTER = "\\xxx\setprinter.exe"
Const XP32DRIVERS = "\\xxx\wxp-32\"
Const WIN7DRIVERS = "\\xxx\w7-64\"


' Objects
Set WSHNetwork = WScript.CreateObject("WScript.Network")
Set objWMIService = GetObject("winmgmts:\\" & CompName & "\root\cimv2")
Set objPrinter = objWMIService.Get("Win32_Printer").SpawnInstance_
Set objPrinterCfg = objWMIService.Get("Win32_PrinterConfiguration").SpawnInstance_
Set objNewPort = objWMIService.Get("Win32_TCPIPPrinterPort").SpawnInstance_
Set FileSys = CreateObject("Scripting.FileSystemObject")
Set colOperatingSystems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")

' Support taking a command argument if not prompt for a location
location = Left(CompName,5)

SetupPrinters

MsgBox "Printer Setup Complete"

Sub SetupPrinters()
	' Read in the Printer Mapping File
	strPrinters = FileSys.OpenTextFile(PMAPFILE,ForReading).ReadAll
	arrLines = Split(strPrinters,vbCrLf)

	' Loop through each line of the Port Mapping Script
	For Each strLine in arrLines

		' Split and check the location of the printer
		arrPrinter = Split(strLine,":")
		
		If arrPrinter(0) = location Then
			createPort arrPrinter(1),arrPrinter(2)
			
			' Create the Printer Name
			SingleSided = arrPrinter(3) & "_SS"
			DoubleSided = arrPrinter(3) & "_DS"
			
			If GetOS = "6.1.7601" Then			' Windows 7
			
				DriverINF = WIN7DRIVERS & arrPrinter(4)
				CreatePrinterXP arrPrinter(1), DriverINF, SingleSided, arrPrinter(5)
				If arrPrinter(7) = "Y" Then
					SetA4(SingleSided)
				End If
				If arrPrinter(6) = "Y" Then
					CreatePrinterXP arrPrinter(1), DriverINF, DoubleSided, arrPrinter(5)
					SetDuplex(DoubleSided)
					If arrPrinter(7) = "Y" Then
						SetA4(DoubleSided)
					End If
				End If
			
			End If
			If GetOS = "5.1.2600" Then		' Windows XP

				DriverINF = XP32DRIVERS & arrPrinter(4)
				CreatePrinterXP arrPrinter(1), DriverINF, SingleSided, arrPrinter(5)
				If arrPrinter(7) = "Y" Then
					SetA4(SingleSided)
				End If
				If arrPrinter(6) = "Y" Then
					CreatePrinterXP arrPrinter(1), DriverINF, DoubleSided, arrPrinter(5)
					SetDuplex(DoubleSided)
					If arrPrinter(7) = "Y" Then
						SetA4(DoubleSided)
					End If
				End If
			End If
		End If
	Next
End Sub

Sub CreatePrinter7(PortName, DriverINF, PrinterName, DriverName)
	oShell.run "RUNDLL32 PRINTUI.DLL,PrintUIEntry /if /b """ & PrinterName & """    /f """ & DriverINF & """  /r """ & PortName & """ /m """ & DriverName & """"  ,,true
End Sub

Sub CreatePrinterXP(PortName, DriverINF, PrinterName, DriverName)
	oShell.run "RUNDLL32 PRINTUI.DLL,PrintUIEntry /if /b """ & PrinterName & """    /f """ & DriverINF & """  /r """ & PortName & """ /m """ & DriverName & """"  ,,true
End Sub

Sub createPort (name, ip)
    objNewPort.Name = name
    objNewPort.Protocol = 1
    objNewPort.HostAddress = ip
    objNewPort.SNMPEnabled = False
    objNewPort.Put_
End Sub

sub SetDuplex(name)
	Set objScriptExec = oShell.Exec(SETPRINTER & " " & Chr(34) & name & Chr(34) & " 8 " & Chr(34) & "pdevmode=dmDuplex=2,dmCollate=1,dmFormName=A4,dmFields=|duplex collate FormName" & Chr(34))
end sub

Sub SetA4(name)
	Set objScriptExec = oShell.Exec(SETPRINTER & " " & Chr(34) & name & Chr(34) & " 8 " & Chr(34) & "pdevmode=dmFormName=A4,dmFields=|FormName" & Chr(34))
end sub


Function GetOS()
	For Each objOperatingSystem in colOperatingSystems
		os = objOperatingSystem.Version
		GetOS = os
	Next
End Function