' List Available Disk Space

Const HARD_DISK = 3

strComputer = "."
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colDisks = objWMIService.ExecQuery("Select * from Win32_LogicalDisk Where DriveType = " & HARD_DISK & "")

For Each objDisk in colDisks
    Wscript.Echo "DeviceID: "& vbTab &  objDisk.DeviceID       
    Wscript.Echo "Free Disk Space: "& vbTab & objDisk.FreeSpace
Next