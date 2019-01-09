' Monitor Physical Disk Drive Performance

strComputer = "."
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
set objRefresher = CreateObject("WbemScripting.SWbemRefresher")
Set colDisks = objRefresher.AddEnum (objWMIService, "win32_perfformatteddata_perfdisk_physicaldisk").objectSet
objRefresher.Refresh

For i = 1 to 100
    For Each objDisk in colDisks
        Wscript.Echo "Average Disk Bytes Per Read: " & vbTab & objDisk.AvgDiskBytesPerRead
        Wscript.Echo "Average Disk Bytes Per Transfer: " & vbTab & objDisk.AvgDiskBytesPerTransfer
        Wscript.Echo "Average Disk Bytes Per Write: " & vbTab & objDisk.AvgDiskBytesPerWrite
        Wscript.Echo "Average Disk Queue Length: " & vbTab & objDisk.AvgDiskQueueLength
        Wscript.Echo "Average Disk Read Queue Length: " & vbTab & objDisk.AvgDiskReadQueueLength
        Wscript.Echo "Average Disk Seconds Per Read: " & vbTab & objDisk.AvgDiskSecPerRead
        Wscript.Echo "Average Disk Seconds Per Transfer: " & vbTab & objDisk.AvgDiskSecPerTransfer      
        Wscript.Echo "Average Disk Seconds Per Write: " & vbTab & objDisk.AvgDiskSecPerWrite      
        Wscript.Echo "Average Disk Write Queue Length: " & vbTab & objDisk.AvgDiskWriteQueueLength      
        Wscript.Echo "Current Disk Queue Length: " & vbTab & objDisk.CurrentDiskQueueLength
        Wscript.Echo "Disk Bytes Per Second: " & vbTab & objDisk.DiskBytesPerSec     
        Wscript.Echo "Disk Read Bytes Per Second: " & vbTab & objDisk.DiskReadBytesPerSec
        Wscript.Echo "Disk Reads Per Second: " & vbTab & objDisk.DiskReadsPerSec
        Wscript.Echo "Disk Transfers Per Second: " & vbTab & objDisk.DiskTransfersPerSec
        Wscript.Echo "Disk Write Bytes Per Second: " & vbTab & objDisk.DiskWriteBytesPerSec
        Wscript.Echo "Disk Writes Per Second: " & vbTab & objDisk.DiskWritesPerSec
        Wscript.Echo "Name: " & vbTab &  objDisk.Name
        Wscript.Echo "Percent Disk Read Time: " & vbTab & objDisk.PercentDiskReadTime
        Wscript.Echo "Percent Disk Time: " & vbTab & objDisk.PercentDiskTime     
        Wscript.Echo "Percent Disk Write Time: " & vbTab & objDisk.PercentDiskWriteTime       
        Wscript.Echo "Percent Idle Time: " & vbTab & objDisk.PercentIdleTime     
        Wscript.Echo "Split IO Per Second: " & vbTab & objDisk.SplitIOPerSec       
        Wscript.Sleep 2000
        objRefresher.Refresh
    Next
Next