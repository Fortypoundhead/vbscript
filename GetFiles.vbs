Set objFSO = CreateObject("Scripting.FileSystemObject")
objStartFolder = "C:\scripts"

Set objFolder = objFSO.GetFolder(objStartFolder)

Set colFiles = objFolder.Files

'For Each objFile in colFiles

'    Wscript.Echo objFolder.Path & "\" & objFile.Name

'Next

ShowSubfolders objFSO.GetFolder(objStartFolder)

Sub ShowSubFolders(Folder)

    For Each Subfolder in Folder.SubFolders
        
		Wscript.Echo Subfolder.Path
        Set objFolder = objFSO.GetFolder(Subfolder.Path)
        Set colFiles = objFolder.Files
        
		For Each objFile in colFiles
		
            Wscript.Echo objFolder & "\" & objFile.Name
			
        Next
 
        ShowSubFolders Subfolder
		
    Next
	
End Sub