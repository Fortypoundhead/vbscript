' List Terminal Services Client Settings

strComputer = "."
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_TSClientSetting")

For Each objItem in colItems
    Wscript.Echo "Audio mapping: " & objItem.AudioMapping
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Clipboard mapping: " & objItem.ClipboardMapping
    Wscript.Echo "Color depth: " & objItem.ColorDepth
    Wscript.Echo "Color depth policy: " & objItem.ColorDepthPolicy
    Wscript.Echo "COM port mapping: " & objItem.COMPortMapping
    Wscript.Echo "Connect client drives at logon: " & objItem.ConnectClientDrivesAtLogon
    Wscript.Echo "Connection policy: " & objItem.ConnectionPolicy
    Wscript.Echo "Connect printer at logon: " & objItem.ConnectPrinterAtLogon
    Wscript.Echo "Default to client printer: " & objItem.DefaultToClientPrinter
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Drive mapping: " & objItem.DriveMapping
    Wscript.Echo "LPT port mapping: " & objItem.LPTPortMapping
    Wscript.Echo "Setting ID: " & objItem.SettingID
    Wscript.Echo "Terminal name: " & objItem.TerminalName
    Wscript.Echo "Windows printer mapping: " & objItem.WindowsPrinterMapping
    Wscript.Echo
Next