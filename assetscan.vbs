Option Explicit

'On Error Resume Next

'***** DECLARATIONS*****************************
Const wbemFlagReturnImmediately = &H10
Const wbemFlagForwardOnly = &H20
Const ForReading = 1
Const ForWriting = 2
Const DEV_ID = 0
Const FSYS = 1
Const DSIZE = 2
Const FSPACE = 3
Const USPACE = 4
Const TITLE = "AssetScan"

Dim fso, f, fsox, fx, objXL, wmiPath, strNoPing, strMBProduct
Dim computerIndex, wscr, adsi, intbutton, strStart, Cshell, strNoConnect
Dim inputFile, outputFile, objKill, strAction, strComplete, strManufact
Dim strPC, intRow, strFilter, RowNum, strCompName, strVideo, strFSB
Dim strDEV_ID, strFSYS, strDSIZE, strFSPACE, strUSPACE, strSD
Dim strRAM, strVir, strPage, strOS, strSP, strProdID, strStatic, struser
Dim strNIC, strIP, strMask, strGate, strMAC, strProc, strSpeed, strHostName
Dim pathlength, Scriptpath, strDocName, currentIP, i, test2, strSubIP, strFooter1, strFooter2
Dim strDomain, strRole, strMake, strModel, strSerial, strBIOSrev, strNICmodel(4), strDateInstalled

outputFile = "IP_table.txt"

Set fsox = CreateObject("Scripting.FileSystemObject")
Set fx = fsox.OpenTextFile(outputFile, ForWriting, True)

' What is the name of the output file?
strDocName = InputBox("What would you like to name the output file?", TITLE)

' Create IP list to scan
Call IPCREATE

'Get Script Location
pathlength = Len(Wscript.ScriptFullName) - Len(Wscript.ScriptName)
Scriptpath = Mid(Wscript.ScriptFullName, 1, pathlength)

Set adsi = CreateObject("ADSystemInfo")
Set wscr = CreateObject("WScript.Network")

inputFile = "IP_table.txt"
outputFile = "NA_IP.txt"

Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.OpenTextFile(inputFile, ForReading, True)
Set fsox = CreateObject("Scripting.FileSystemObject")
Set fx = fsox.OpenTextFile(outputFile, ForWriting, True)
Set Cshell = CreateObject("WScript.Shell")
computerIndex = 1

'*****[ FUNCTIONS ]*******************************

Function Ask(strAction)
   intbutton = MsgBox(strAction, vbQuestion + vbYesNo, TITLE)
   Ask = intbutton = vbNo
End Function

Function IsConnectible(sHost, iPings, iTO)

Const OpenAsASCII = 0
Const FailIfNotExist = 0
Const ForReading = 1
Dim oShell, oFSO, sTempFile, fFile

If iPings = "" Then iPings = 2
If iTO = "" Then iTO = 750

Set oShell = CreateObject("WScript.Shell")
Set oFSO = CreateObject("Scripting.FileSystemObject")

sTempFile = oFSO.GetSpecialFolder(2).ShortPath & "\" & oFSO.GetTempName

oShell.run "%comspec% /c ping.exe -n " & iPings & " -w " & iTO & " " & sHost & ">" & sTempFile, 0, True
Set fFile = oFSO.OpenTextFile(sTempFile, ForReading, FailIfNotExist, OpenAsASCII)

Select Case InStr(fFile.ReadAll, "TTL=")
Case 0 IsConnectible = False
Case Else IsConnectible = True
End Select

fFile.Close
oFSO.DeleteFile (sTempFile)

End Function


'*****[ MAIN SCRIPT ]*****************************

If Ask("Run AssetScan now?") Then
Wscript.Quit
Else
strStart = "Inventory run started: " & Date & " at " & Time
End If

Call BuildXLS
Call Connect
Call Footer

objXL.ActiveWorkbook.SaveAs Scriptpath & strDocName & "-AssetScan.xls"
MsgBox "Your inventory run is complete!", vbInformation + vbOKOnly, TITLE

'*****[ SUB ROUTINES ]****************************

'*** Subroutine create ip table
Sub IPCREATE()

   currentIP = GetIP()

   Dim Seps(2)
   Seps(0) = "."
   Seps(1) = "."
   test2 = Tokenize(currentIP, Seps)

   strSubIP = test2(0) & "." & test2(1) & "." & test2(2) & "."
   strSubIP = InputBox("Enter Subnet to Scan - ie: 192.168.5. Press <enter> to Scan Local Subnet", TITLE, strSubIP)
    On Error Resume Next
        intStartingAddress = InputBox("Start at :", "Scanning Subnet: " & strSubIP, 61)
        intEndingAddress = InputBox ("End at :", "Scanning Subnet: "&strSubIP&intStartingAddress, 254)

    For i = intStartingAddress To intEndingAddress
        strComputer = strSubIP & i
        fx.WriteLine (strSubIP & i)
    Next

End Sub

Function Tokenize(ByVal TokenString, ByRef TokenSeparators())

   Dim NumWords, a()
   NumWords = 0
   
   Dim NumSeps
   NumSeps = UBound(TokenSeparators)
   
   Do
      Dim SepIndex, SepPosition
      SepPosition = 0
      SepIndex = -1
      
      For i = 0 To NumSeps - 1
      
         ' Find location of separator in the string
         Dim pos
         pos = InStr(TokenString, TokenSeparators(i))
         
         ' Is the separator present, and is it closest to the beginning of the string?
         If pos > 0 And ((SepPosition = 0) Or (pos < SepPosition)) Then
            SepPosition = pos
            SepIndex = i
         End If
         
      Next

      ' Did we find any separators?
      If SepIndex < 0 Then

         ' None found - so the token is the remaining string
         ReDim Preserve a(NumWords + 1)
         a(NumWords) = TokenString
         
      Else

         ' Found a token - pull out the substring
         Dim substr
         substr = Trim(Left(TokenString, SepPosition - 1))
   
         ' Add the token to the list
         ReDim Preserve a(NumWords + 1)
         a(NumWords) = substr
      
         ' Cutoff the token we just found
         Dim TrimPosition
         TrimPosition = SepPosition + Len(TokenSeparators(SepIndex))
         TokenString = Trim(Mid(TokenString, TrimPosition))
                  
      End If
      
      NumWords = NumWords + 1
   Loop While (SepIndex >= 0)
   
   Tokenize = a
   
End Function


Function GetIP()
  Dim ws: Set ws = CreateObject("WScript.Shell")
  Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
  Dim TmpFile: TmpFile = fso.GetSpecialFolder(2) & "/ip.txt"
  Dim ThisLine, IP
  If ws.Environment("SYSTEM")("OS") = "" Then
    ws.run "winipcfg /batch " & TmpFile, 0, True
  Else
    ws.run "%comspec% /c ipconfig > " & TmpFile, 0, True
  End If
  With fso.GetFile(TmpFile).OpenAsTextStream
    Do While Not .AtEndOfStream
      ThisLine = .ReadLine
      If InStr(ThisLine, "Address") <> 0 Then IP = Mid(ThisLine, InStr(ThisLine, ":") + 2)
    Loop
    .Close
  End With
  'WinXP (NT? 2K?) leaves a carriage return at the end of line
  If IP <> "" Then
    If Asc(Right(IP, 1)) = 13 Then IP = Left(IP, Len(IP) - 1)
  End If
  GetIP = IP
  fso.GetFile(TmpFile).Delete
  Set fso = Nothing
  Set ws = Nothing
End Function

Function TranslateDomainRole(ByVal roleID)
   Dim a

   Select Case roleID
      Case 0
         a = "Standalone Workstation"
      Case 1
         a = "Member Workstation"
      Case 2
         a = "Standalone Server"
      Case 3
         a = "Member Server"
      Case 4
         a = "Backup Domain Controller"
      Case 5
         a = "Primary Domain Controller"
   End Select
   TranslateDomainRole = a
End Function

'*********************************************************
Sub Connect()
    Do While f.AtEndOfLine <> True
        strPC = f.ReadLine
        If strPC <> "" Then
            If Not IsConnectible(strPC, "", "") Then
                strNoPing = "Couldn't ping " & strPC
                'Call MsgNoPing()
                Call Error
            Else
                On Error Resume Next
                Set oWMI = GetObject("winmgmts:{impersonationLevel=impersonate}!//" & strPC & "/root/cimv2")

                If Err.Number <> 0 Then

                strNoConnect = "Couldn't connect to " & strPC
                'Call MsgNoConnect()
                Call Error

               Else
                 
                  'Get IP Address
                  strCompName = UCase(strPC)
               
                  'Get Hostname
                  Set HostName = oWMI.ExecQuery("select DNSHostName from Win32_NetworkAdapterConfiguration where IPEnabled=TRUE")
                  For Each Host In HostName
                     strHostName = Host.DNSHostName
                  Next

                  'Get Domain and Role
                  Set colItems = oWMI.ExecQuery("SELECT * FROM Win32_ComputerSystem", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

                  For Each objItem In colItems
                     strDomain = objItem.Domain
                     strRole = TranslateDomainRole(objItem.DomainRole)
                  Next

                  'Get Make, Model, Serial Number
                  Set colItems = oWMI.ExecQuery("SELECT * FROM Win32_ComputerSystemProduct", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

                  For Each objItem In colItems
                     strSerial = objItem.IdentifyingNumber
                     strModel = objItem.Name
                     strMake = objItem.Vendor
                  Next

                  'Get RAM (Total)
                  Set MemorySet = oWMI.ExecQuery("select TotalPhysicalMemory, " & "TotalVirtualMemory, TotalPageFileSpace from " & "Win32_LogicalMemoryConfiguration")
                  For Each Memory In MemorySet
                     strRAM = FormatNumber(Memory.TotalPhysicalMemory / 1024, 1) & " Mb"
                  Next

                  'Get Operating System and Service Pack Info
                  Set OSSet = oWMI.ExecQuery("select Caption, CSDVersion, SerialNumber " & "from Win32_OperatingSystem")
                  For Each OS In OSSet
                     strOS = OS.Caption
                     strSP = OS.CSDVersion
                  Next

                  'Get BIOS Revision
                  Set colSettings = oWMI.ExecQuery("Select * from Win32_BIOS")
                  For Each objBIOS In colSettings
                     strBIOSrev = objBIOS.Version
                  Next
                  
                  'Get Processor Type
                  Set ProSet = oWMI.ExecQuery("select Name, MaxClockSpeed from Win32_Processor")
                  For Each Pro In ProSet
                     strProc = Pro.Name
                     strSpeed = Pro.MaxClockSpeed & " MHZ"
                  Next

                  'Get Logged in user
                  Set loggeduser = oWMI.ExecQuery("select UserName from Win32_ComputerSystem")
                  For Each logged In loggeduser
                     struser = logged.UserName
                  Next

                  'Get NIC Model 'ISOLATE PRIMARY NIC INFO
                  Set colSettings = oWMI.ExecQuery("Select * from Win32_NetworkAdapter")
                  i = 1
                  For Each Objcomputer In colSettings
                     If Objcomputer.AdapterType = "Ethernet 802.3" Then
                        strNICmodel(i - 1) = strMsg & "Interface[" & i & "]: " & Objcomputer.Name
                        i = i + 1
                     End If
                  Next

                  'Get Subnet Mask, MAC Address, Default Gateway
                  Set IPConfigSet = oWMI.ExecQuery("select ServiceName, IPAddress, " & "IPSubnet, DefaultIPGateway, MACAddress from " & "Win32_NetworkAdapterConfiguration where IPEnabled=TRUE")
                  Count = 0
                  For Each IPConfig In IPConfigSet
                     Count = Count + 1
                  Next
                  ReDim sName(Count - 1)
                  ReDim sIP(Count - 1)
                  ReDim sMask(Count - 1)
                  ReDim sGate(Count - 1)
                  ReDim sMAC(Count - 1)
                  Count = 0
                  For Each IPConfig In IPConfigSet
                     sName(Count) = IPConfig.ServiceName(0)
                     strNIC = sName(Count)
                     sIP(Count) = IPConfig.IPAddress(0)
                     strIP = sIP(Count)
                     sMask(Count) = IPConfig.IPSubnet(0)
                     strMask = sMask(Count)
                     sGate(Count) = IPConfig.DefaultIPGateway(0)
                     strGate = sGate(Count)
                     sMAC(Count) = IPConfig.MACAddress(0)
                     strMAC = sMAC(Count)
                     Count = Count + 1
                  Next

                  'Date Installed
                  Set colSettings = oWMI.ExecQuery("Select * from Win32_OperatingSystem")
                  For Each Objcomputer In colSettings
                     strDateInstalled = Objcomputer.InstallDate
                  Next

                  'EXTRA LOOP to call Add lines
                  Set DiskSet = oWMI.ExecQuery("select DeviceID, FileSystem, Size, FreeSpace " & "from Win32_LogicalDisk where DriveType = '3'")

                  ReDim strDisk(RowNum, 4)
                    For Each Disk In DiskSet
                        Call AddLineToXLS(strCompName, strHostName, strDomain, strRole, strMake, _
                        strModel, strSerial, strRAM, strOS, strSP, strBIOSrev, strProc, strSpeed, struser, _
                        strMask, strGate, strMAC, strDateInstalled, strNICmodel)
                  Next
                  
               End If
            End If
        End If
    Loop
End Sub

'*** Subroutine to Build XLS ***
Sub BuildXLS()

intRow = 1
Set objXL = Wscript.CreateObject("Excel.Application")
objXL.Visible = True
objXL.WorkBooks.Add
objXL.Sheets("Sheet1").Select()
objXL.Sheets("Sheet1").Name = " AssetScan Inventory"

'** Set Row Height
objXL.Rows(1).RowHeight = 25

'** Set Column widths
objXL.Columns(1).ColumnWidth = 9
objXL.Columns(2).ColumnWidth = 14
objXL.Columns(3).ColumnWidth = 7
objXL.Columns(4).ColumnWidth = 17
objXL.Columns(5).ColumnWidth = 16
objXL.Columns(6).ColumnWidth = 10
objXL.Columns(7).ColumnWidth = 15
objXL.Columns(8).ColumnWidth = 7
objXL.Columns(9).ColumnWidth = 26
objXL.Columns(10).ColumnWidth = 12
objXL.Columns(11).ColumnWidth = 14
objXL.Columns(12).ColumnWidth = 24
objXL.Columns(13).ColumnWidth = 15
objXL.Columns(14).ColumnWidth = 19
objXL.Columns(15).ColumnWidth = 11
objXL.Columns(16).ColumnWidth = 11
objXL.Columns(17).ColumnWidth = 14
objXL.Columns(18).ColumnWidth = 22
objXL.Columns(19).ColumnWidth = 37
objXL.Columns(20).ColumnWidth = 35
objXL.Columns(21).ColumnWidth = 35
objXL.Columns(22).ColumnWidth = 35
objXL.Columns(23).ColumnWidth = 35

'*** Set Cell Format for Column Titles ***
objXL.Range("A1:Z1").Select
objXL.Selection.Font.Bold = True
objXL.Selection.Font.Size = 8
objXL.Selection.Interior.ColorIndex = 11
objXL.Selection.Interior.Pattern = 1 'xlSolid
objXL.Selection.Font.ColorIndex = 2
objXL.Selection.WrapText = True
objXL.Columns("A:Z").Select
objXL.Selection.HorizontalAlignment = 3 'xlCenter


'*** Set Column Titles ***
Dim arrNicTitle(4)
arrNicTitle(0) = "NIC #1 Model"
arrNicTitle(1) = "NIC #2 Model"
arrNicTitle(2) = "NIC #3 Model"
arrNicTitle(3) = "NIC #4 Model"
arrNicTitle(4) = "NIC #5 Model"

' 15,16,17
Call AddLineToXLS("IP Address", "Hostname", "Domain", "Role", "Make", "Model", "Serial Number", _
"RAM", "Operating System", "Service Pack", "BIOS Revision", "Processor Type", "Processor Speed", _
"Logged in user", "Subnet Mask", "Default Gateway", "MAC Address", "Date Installed", arrNicTitle)

End Sub

'*** Subroutine Add Lines to XLS ***
objXL.Columns("A:AA").Select
objXL.Selection.HorizontalAlignment = 3 'xlCenter
objXL.Selection.Font.Size = 8

Sub AddLineToXLS(strCompName, strHostName, strDomain, strRole, strMake, strModel, strSerial, strRAM, _
strOS, strSP, strBIOSrev, strProc, strSpeed, struser, strMask, strGate, strMAC, strDateInstalled, ByRef strNICmodel)
    objXL.Cells(intRow, 1).Value = strCompName
    objXL.Cells(intRow, 2).Value = strHostName
    objXL.Cells(intRow, 3).Value = strDomain
    objXL.Cells(intRow, 4).Value = strRole
    objXL.Cells(intRow, 5).Value = strMake
    objXL.Cells(intRow, 6).Value = strModel
    objXL.Cells(intRow, 7).Value = strSerial
    objXL.Cells(intRow, 8).Value = strRAM
    objXL.Cells(intRow, 9).Value = strOS
    objXL.Cells(intRow, 10).Value = strSP
    objXL.Cells(intRow, 11).Value = strBIOSrev
    objXL.Cells(intRow, 12).Value = strProc
    objXL.Cells(intRow, 13).Value = strSpeed
    objXL.Cells(intRow, 14).Value = struser
    objXL.Cells(intRow, 15).Value = strMask
    objXL.Cells(intRow, 16).Value = strGate
    objXL.Cells(intRow, 17).Value = strMAC
    objXL.Cells(intRow, 18).Value = strDateInstalled
    objXL.Cells(intRow, 19).Value = strNICmodel(0)
    objXL.Cells(intRow, 20).Value = strNICmodel(1)
    objXL.Cells(intRow, 21).Value = strNICmodel(2)
    objXL.Cells(intRow, 22).Value = strNICmodel(3)
    objXL.Cells(intRow, 23).Value = strNICmodel(4)
    intRow = intRow + 1
    objXL.Cells(1, 1).Select
End Sub

'*** Subroutine Add Lines to XLS for Disk Info. ***
'objXL.Columns("A:AA").Select
'objXL.Selection.HorizontalAlignment = 3 'xlCenter
'objXL.Selection.Font.Size = 8

Sub AddLineToDisk(strDEV_ID, strFSYS, strDSIZE, strFSPACE, strUSPACE)
    objXL.Cells(intRow, 11).Value = strDEV_ID
    objXL.Cells(intRow, 12).Value = strFSYS
    objXL.Cells(intRow, 13).Value = strDSIZE
    objXL.Cells(intRow, 14).Value = strFSPACE
    objXL.Cells(intRow, 15).Value = strUSPACE
    intRow = intRow + 1
    objXL.Cells(1, 1).Select
End Sub

'*** Sub to add footer when speadsheet is complete ***
Sub Footer()
   strFooter1 = "Inventory AssetScan"
   strFooter2 = "Script was created by Sean Kelly and is free for personal/small business use"
   strComplete = "Inventory run completed at: " & Date & " at " & Time

   intRow = intRow + 4

   '** Set Cell Format for Row
   objXL.Cells(intRow, 4).Select
   objXL.Selection.Font.ColorIndex = 1
   objXL.Selection.Font.Size = 8
   objXL.Selection.Font.Bold = False
   objXL.Selection.HorizontalAlignment = 2 'xlRight
   objXL.Cells(intRow, 4).Value = strFooter1

   intRow = intRow + 1

   '** Set Cell Format for Row
   objXL.Cells(intRow, 4).Select
   objXL.Selection.Font.ColorIndex = 1
   objXL.Selection.Font.Size = 8
   objXL.Selection.Font.Bold = False
   objXL.Selection.HorizontalAlignment = 2 'xlRight
   objXL.Cells(intRow, 4).Value = strFooter2

   intRow = intRow + 1

   '** Set Cell Format for Row
   objXL.Cells(intRow, 4).Select
   objXL.Selection.Font.ColorIndex = 1
   objXL.Selection.Font.Size = 8
   objXL.Selection.Font.Bold = False
   objXL.Selection.HorizontalAlignment = 2 'xlRight
   objXL.Cells(intRow, 4).Value = strStart

   intRow = intRow + 1

   '** Set Cell Format for Row
   objXL.Cells(intRow, 4).Select
   objXL.Selection.Font.ColorIndex = 1
   objXL.Selection.Font.Size = 8
   objXL.Selection.Font.Bold = False
   objXL.Selection.HorizontalAlignment = 2 'xlRight
   objXL.Cells(intRow, 4).Value = strComplete

   intRow = intRow + 1

End Sub

'*** ErrorHandler ***
Sub Error()

fx.WriteLine (strPC)

End Sub
