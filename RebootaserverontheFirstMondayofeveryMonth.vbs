' Reboot a server on the First Monday of every Month

Set WshShell = WScript.CreateObject( "WScript.Shell" )
 
DateofMonth = Day(Date) 
DayofWeek  = UCASE(WeekDayName(WeekDay(Date)))
 
If (DateofMonth < "8") AND (DayofWeek = UCASE("Monday")) Then
 Command = "E:\EM\site\bin\Reboot.bat"
 WshShell.Run (Command)
End If
 
Set WshShell = Nothing

