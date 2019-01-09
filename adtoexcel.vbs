'*******************************************************************************
'***
'*** TITLE: 		ADToExcel.vbs 
'***
'*** DATE: 			2010-08-02
'*** AUTHOR: 		FortyPoundHead
'***
'*** COMPANY: 		redacted
'*** DEPARTMENT: 	Information Services 
'***					\ Infrastructure 
'***					\ System Administrators
'***
'*** PURPOSE: 		Dump Active Directory user objects to Excel Spreadsheet
'***
'*** USAGE:			Not hard.  Double-click the .vbs file, and the script will
'***				open up a copy of Excel, and populate the spreadsheet with
'***				user information from Active Directory.  The script will
'***				take a few minutes to run, and you will be given a pop-up
'***				message when it completes.
'***
'***				Elevated rights are not required, since AD can be read by 
'***				all in the Domain Users or Authenticated Users groups.
'***
'*******************************************************************************

'*******************************************************************************
'*** 
'*** MODIFICATIONS:
'*** 
'*** 2013-11-08 - Added traps for account status, added field for raw status code
'***
'*******************************************************************************

'*******************************************************************************
'***
'*** Set aside variables
'***
'*******************************************************************************

Dim objWriteBook
Dim ObjExcel 
Dim x, intSEAC 

'*******************************************************************************
'***
'*** Setup Active directory connectivity
'***
'*******************************************************************************

Set objRoot = GetObject("LDAP://RootDSE") 										' grab the root
strDNC = objRoot.Get("DefaultNamingContext") 									' where are we starting
Set objDomain = GetObject("LDAP://" & strDNC) 									' Bind to the top of the Domain using LDAP using ROotDSE 

'*******************************************************************************
'***
'*** Do the job.  First, set up spreadsheet, then enumerate AD objects
'***
'*******************************************************************************

Call ExcelSetup("Sheet1") 														' Sub to make Excel Document 
x = 1 
Call EnumerateObjects(objDomain) 

MsgBox "Done" 																	' inform user of script completion

Sub EnumerateObjects(objDomain) 

	On Error Resume Next 														' nope - we are not stopping on errors. Bad, I know. Back up off me.
																				
	Dim SecondaryEmail(20) 														' Secondary email address storage

	For Each objMember In objDomain 											' iterate through the entire collection 

		If ObjMember.Class = "user" Then 										' if not User object, move on. 
			x = x +1 															' counter used to increment the cells in Excel 

			objWriteBook.Cells(x, 1).Value = objMember.Class 
			
			'*******************************************************************
			'***
			'*** Grab values from objMember
			'*** 
			'*******************************************************************
			
			SamAccountName = ObjMember.samAccountName 
			Cn = ObjMember.CN 
			FirstName = objMember.GivenName 
			LastName = objMember.sn 
			initials = objMember.initials 
			Descrip = objMember.description 
			Office = objMember.physicalDeliveryOfficeName 
			Telephone = objMember.telephonenumber 
			EmailAddr = objMember.mail 
			WebPage = objMember.wwwHomePage 
			Addr1 = objMember.streetAddress 
			City = objMember.l 
			State = objMember.st 
			ZipCode = objMember.postalCode 
			Title = ObjMember.Title 
			Department = objMember.Department 
			Company = objMember.Company 
			Manager = ObjMember.Manager 
			Profile = objMember.profilePath 
			LoginScript = objMember.scriptpath 
			HomeDirectory = ObjMember.HomeDirectory 
			HomeDrive = ObjMember.homeDrive 
			AdsPath = Objmember.Adspath 
			LastLogin = objMember.LastLogin
			GUA = objMember.ExtensionAttribute10
			AccountExpires = objMember.AccountExpirationDate
			LastBadPassword = objMember.badPasswordTime
			BadPasswordCount = objMember.badPwdCount
			LogonCount = objMember.LogonCount
			pwdLastSet = objMember.PasswordLastChanged
			WhenCreated = objMember.WhenCreated
			whenChanged = objMember.whenChanged
			UACStatus = objMember.userAccountControl
			msTSAllowLogon = objMember.msTSAllowLogon
			Contractor=objMember.Contractor
			ILMManaged=objMember.extensionAttribute11
			
			'*******************************************************************
			'***
			'*** Determine account status, make the display human-readable
			'***
			'*******************************************************************
			
			Select Case UACStatus
				
				Case 64
					UserAccountControl="Cannot Change Password"
				
				Case 512
					
					UserAccountControl="Normal Account"
					
				Case 514
					
					UserAccountControl="Disabled"
					
				Case 544
					UserAccountControl="Enabled, Password Not Required"
					
				Case 546
					UserAccountControl="Disabled, Password Not Required"
				
				case 8192
					UserAccountControl="Server Trust Account"
				
				case 2048
					UserAccountControl="Interdomain Trust Account"
					
				Case 4096
					UserAccountControl="Workstation Trust Account"
				
				Case 16777216 
					UserAccountControl="Trusted for Delegation"
				
				Case 33554432 
					UserAccountControl="No Authentication Required"
					
				Case 66048
					UserAccountControl="Enabled, Password Never Expires"

				Case 66050
					UserAccountControl="Disabled"
					
				Case 66080	
					UserAccountControl="Enabled, Password Doesn't Expire & Not Required"
					
				Case 66082	
					UserAccountControl="Disabled, Password Doesn't Expire & Not Required"

				Case 65536
					UserAccountControl="Password Never Expires"
					
				Case 262656	
					UserAccountControl="Enabled, Smartcard Required"
					
				Case 262658	
					UserAccountControl="Disabled, Smartcard Required"
					
				Case 262688	
					UserAccountControl="Enabled, Smartcard Required, Password Not Required"
					
				Case 262690	
					UserAccountControl="Disabled, Smartcard Required, Password Not Required"
					
				Case 328192	
					UserAccountControl="Enabled, Smartcard Required, Password Doesn't Expire"
					
				Case 328194	
					UserAccountControl="Disabled, Smartcard Required, Password Doesn't Expire"
					
				Case 328224	
					UserAccountControl="Enabled, Smartcard Required, Password Doesn't Expire & Not Required"
					
				Case 328226	
					UserAccountControl="Disabled, Smartcard Required, Password Doesn't Expire & Not Required"
					
				Case Else
					
					UserAccountCountrol=UACStatus
					
			End Select
					
			
			intSEAC = 1 														' Counter for array of Secondary email addresses 
			
			For each email in ObjMember.proxyAddresses 
			
				If Left (email,5) = "SMTP:" Then 
				
					Primary = Mid (email,6) 									' if SMTP is all caps, then it's the Primary 
				
				ElseIf Left (email,5) = "smtp:" Then 
				
					SecondaryEmail(intSEAC) = Mid (email,6) 					' load the list of 2ndary SMTP emails into Array. 
					intSEAC = intSEAC + 1 
				
				End If 
				
			Next 
			
			'*******************************************************************
			'***
			'*** Send the values to Excel
			'***
			'*******************************************************************
			
			objWriteBook.Cells(x, 2).Value = SamAccountName 
			objWriteBook.Cells(x, 3).Value = CN 
			objWriteBook.Cells(x, 4).Value = FirstName 
			objWriteBook.Cells(x, 5).Value = LastName 
			objWriteBook.Cells(x, 6).Value = Initials 
			objWriteBook.Cells(x, 7).Value = Descrip 
			objWriteBook.Cells(x, 8).Value = Office 
			objWriteBook.Cells(x, 9).Value = Telephone 
			objWriteBook.Cells(x, 10).Value = EmailAddr
			objWriteBook.Cells(x, 11).Value = WebPage 
			objWriteBook.Cells(x, 12).Value = Addr1 
			objWriteBook.Cells(x, 13).Value = City 
			objWriteBook.Cells(x, 14).Value = State 
			objWriteBook.Cells(x, 15).Value = ZipCode 
			objWriteBook.Cells(x, 16).Value = Title 
			objWriteBook.Cells(x, 17).Value = Department 
			objWriteBook.Cells(x, 18).Value = Company 
			objWriteBook.Cells(x, 19).Value = Manager 
			objWriteBook.Cells(x, 20).Value = Profile 
			objWriteBook.Cells(x, 21).Value = LoginScript 
			objWriteBook.Cells(x, 22).Value = HomeDirectory 
			objWriteBook.Cells(x, 23).Value = HomeDrive 
			objWriteBook.Cells(x, 24).Value = Adspath 
			objWriteBook.Cells(x, 25).Value = LastLogin 
			objWriteBook.Cells(x,26).Value = GUA
			objWriteBook.Cells(x,27).Value = AccountExpires
			objWriteBook.Cells(x,28).Value = LastBadPassword
			objWriteBook.Cells(x,29).Value = BadPasswordCount
			objWriteBook.Cells(x,30).Value = LogonCount
			objWriteBook.Cells(x,31).Value = pwdLastSet
			objWriteBook.Cells(x,32).Value = WhenCreated
			objWriteBook.Cells(x,33).Value = whenChanged
			objWriteBook.Cells(x,34).Value = userAccountControl
			objWriteBook.Cells(x,35).Value = UACStatus
			objWriteBook.Cells(x,36).Value = msTSAllowLogon
			objWriteBook.Cells(x,37).Value = Primary
			objWriteBook.Cells(x,38).value = Contractor
			objWriteBook.Cells(x,39).value = ILMManaged
			
			
			'*******************************************************************
			'***
			'*** Write out the Array for the Secondary email addresses. 
			'***
			'*******************************************************************
			
			For intWSEA = 1 To 20 
				
				objWriteBook.Cells(x,39+intWSEA).Value = SecondaryEmail(intWSEA) 
			
			Next 
			
			'*******************************************************************
			'***
			'*** Reset the values to nothingness. Prevents stale data from being
			'*** introduced into the next retrieved record.
			'***
			'*******************************************************************

			SamAccountName = "-" 
			Cn = "-" 
			FirstName = "-" 
			LastName = "-" 
			initials = "-" 
			Descrip = "-" 
			Office = "-" 
			Telephone = "-" 
			EmailAddr = "-" 
			WebPage = "-" 
			Addr1 = "-" 
			City = "-" 
			State = "-" 
			ZipCode = "-" 
			Title = "-" 
			Department = "-" 
			Company = "-" 
			Manager = "-" 
			Profile = "-" 
			LoginScript = "-" 
			HomeDirectory = "-" 
			HomeDrive = "-" 
			Primary = "-" 
			GUA = "-"
			AccountExpires = "-"
			LastBadPassword = "-"
			BadPasswordCount = "-"
			LogonCount = "-"
			pwdLastSet = "-"
			WhenCreated = "-"
			whenChanged = "-"
			userAccountControl = "UACStatus Not Present in AD"
			RawStatus="0"
			msTSAllowLogon = "-"
			Contractor="-"
			ILMManaged="-"
			
			'***
			'*** Clear out the secondary email address array
			'*** 
			
			For intWSEA = 1 To 20 
				
				SecondaryEmail(intWSEA) = "" 
			
			Next 
			
		End If 

		'***********************************************************************
		'***
		'*** If we run into an OU along the way, call myself again to iterate 
		'*** through the sub-OU
		'***
		'***********************************************************************

		If objMember.Class = "organizationalUnit" or OBjMember.Class = "container" Then 
			
			EnumerateObjects (objMember) 
		
		End If 

	Next
	
End Sub 

Sub ExcelSetup(shtName) 
	
	'***************************************************************************
	'***
	'*** Create the Excel spreadsheet, adding headings at the top.
	'***
	'***************************************************************************
	
	Set objExcel = CreateObject("Excel.Application") 
	Set objWriteBook = objExcel.Workbooks.Add 
	Set objWriteBook = objExcel.ActiveWorkbook.Worksheets(shtName) 
	
	objWriteBook.Name = "Active Directory Users" 								' name the sheet 
	objWriteBook.Activate 
	
	objExcel.Visible = True 
	objWriteBook.Cells(1, 1).Value = "ObjClass"
	objWriteBook.Cells(1, 2).Value = "SamAccountName" 
	objWriteBook.Cells(1, 3).Value = "CN" 
	objWriteBook.Cells(1, 4).Value = "FirstName" 
	objWriteBook.Cells(1, 5).Value = "LastName" 
	objWriteBook.Cells(1, 6).Value = "Initials" 
	objWriteBook.Cells(1, 7).Value = "Description" 
	objWriteBook.Cells(1, 8).Value = "Office" 
	objWriteBook.Cells(1, 9).Value = "Telephone" 
	objWriteBook.Cells(1, 10).Value = "Email" 
	objWriteBook.Cells(1, 11).Value = "WebPage" 
	objWriteBook.Cells(1, 12).Value = "Addr1" 
	objWriteBook.Cells(1, 13).Value = "City" 
	objWriteBook.Cells(1, 14).Value = "State" 
	objWriteBook.Cells(1, 15).Value = "ZipCode" 
	objWriteBook.Cells(1, 16).Value = "Title" 
	objWriteBook.Cells(1, 17).Value = "Department" 
	objWriteBook.Cells(1, 18).Value = "Company" 
	objWriteBook.Cells(1, 19).Value = "Manager" 
	objWriteBook.Cells(1, 20).Value = "Profile" 
	objWriteBook.Cells(1, 21).Value = "LoginScript" 
	objWriteBook.Cells(1, 22).Value = "HomeDirectory" 
	objWriteBook.Cells(1, 23).Value = "HomeDrive" 
	objWriteBook.Cells(1, 24).Value = "Adspath" 
	objWriteBook.Cells(1, 25).Value = "LastLogin" 
	objWriteBook.Cells(1, 26).Value = "GUA"
	objWriteBook.Cells(1, 27).Value = "Account Expires" 
	objWriteBook.Cells(1, 28).Value = "Last Bad Password" 
	objWriteBook.Cells(1, 29).Value = "Bad Password Count" 
	objWriteBook.Cells(1, 30).Value = "Logon Count"
	objWriteBook.Cells(1, 31).Value = "Password Last Set"
	objWriteBook.Cells(1, 32).Value = "When Created"
	objWriteBook.Cells(1, 33).Value = "When Changed"
	objWriteBook.Cells(1, 34).Value = "Status"
	objWriteBook.Cells(1, 35).Value = "Raw Status Code"
	objWriteBook.Cells(1, 36).Value = "Allowed TS?"
	objWriteBook.Cells(1, 37).Value = "Primary SMTP" 
	objWriteBook.Cells(1, 38).Value = "Contractor"
	objWriteBook.Cells(1, 39).Value = "ILM Managed ?"
	objWriteBook.Cells(1, 40).Value = "Secondary Email Addresses"
End Sub 