'On Error Resume Next
'Option Explicit
Dim objCommand, objConnection, strQuery, strDBPath, objRecordset
Dim objAccessConnection, objAccessRecordset, objItem
Const ADS_UF_ACCOUNTDISABLE = 2
Const adLockOptimistic 	= 3

' Connect to AD.
Set objCommand = CreateObject("ADODB.Command")
Set objConnection = CreateObject("ADODB.Connection")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"
objCommand.ActiveConnection = objConnection
' ******* BEGIN CALLOUT A *******
'strQuery = "LDAP://DC=Corp,DC=domain,DC=com;(objectCategory=person);adsPath;subtree" 
strQuery = "LDAP://RootDSE;(objectCategory=person);adsPath;subtree" 
' ******* END CALLOUT A *******
objCommand.CommandText = strQuery

Set objRecordset = objCommand.Execute

' Open the Access database.
Set objAccessConnection = CreateObject("ADODB.Connection")
Set objAccessRecordset = CreateObject("ADODB.Recordset")
' ******* BEGIN CALLOUT B *******
strDBPath = "c:\mydirectory\ADUsers.mdb"
' ******* END CALLOUT B *******
objAccessConnection.open ("DRIVER=Microsoft Access Driver (*.mdb);DBQ=" & strDBPath)
Set objAccessRecordset.ActiveConnection = objAccessConnection
objAccessRecordset.LockType = adLockOptimistic

' Delete all the rows from the database so it can be repopulated.
objAccessRecordset.source = "DELETE * FROM ADUsers"
objAccessRecordset.Open
objAccessRecordset.source = "Select * FROM ADUsers"
objAccessRecordset.Open

' ******* BEGIN CALLOUT C *******
' Loop through all the user accounts and insert their details into the database.
Do Until objRecordset.EOF
  Set objItem = GetObject(objRecordset.Fields ("ADsPath"))
  objAccessRecordset.AddNew
  objAccessRecordset.Fields("DisplayName") = objItem.displayName
  objAccessRecordset.Fields("UserID") = objItem.sAMAccountName
  objAccessRecordset.Fields("EmailAddress") = objItem.mail
    If objitem.userAccountControl And ADS_UF_ACCOUNTDISABLE Then
      objAccessRecordset.Fields("UserDisabled") = True
    Else
      objAccessRecordset.Fields("UserDisabled") = False
    End If
    objAccessRecordset.Update
  objRecordSet.MoveNext
Loop
' ******* END CALLOUT C *******
