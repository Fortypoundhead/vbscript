'local machine. can be remote, if your account has the rights

strComputer = "."

' Connect to the WMI Service

Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

' Fetch all details from Win32_computersystem

Set colComputerSystem = objWMIService.ExecQuery ("Select * from Win32_computersystem")
Set colBIOS = objWMIService.ExecQuery ("Select * from Win32_BIOS")

' Look through all values, and make variables for manufacturer and model

For each objComputerSystem in colComputerSystem
    
	GetComputerManufacturer = objComputerSystem.Manufacturer
    GetComputerModel = objComputerSystem.Model

Next

'uncomment to test retrieved strings
'Wscript.echo "The system you are on is a " & GetComputerManufacturer & " " & GetComputerModel

If GetComputerModel="HP ProBook 640 G1" then 

	' True, do something
	
	wscript.echo "true"
	
else
	
	'false, do something else, or nothing
	
	wscript.echo "false"
	
end if
	