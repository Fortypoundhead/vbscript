' Create a Custom Event Log
 
Const NO_VALUE = Empty

Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.RegWrite "HKLM\System\CurrentControlSet\Services\EventLog\DWScripts\", NO_VALUE