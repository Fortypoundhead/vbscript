' Convert SID to Account Name

On Error Resume Next

wscript.echo getSID

Private Function getSID()
	' Get SID from user
	Const POPUP_TITLE = "User To SID Conversion"
	SID = InputBox("Enter SID",POPUP_TITLE)	
		server = "."
		Set objWMIService = GetObject("winmgmts:\\" & server & "\root\cimv2")

		Set objAccount = objWMIService.Get("Win32_SID.SID='" & SID & "'")
		strUser = objAccount.AccountName
		strDomain = objAccount.ReferencedDomainName
		If Err.Number <> 0 Then
        getSID = Err.Description
        Err.Clear
    Else
				getSID = "User: " & vbtab & UCase(strUser) & vbcrlf & "Domain: " & vbtab & UCase(strDomain)
    End If
End Function
' =======================================