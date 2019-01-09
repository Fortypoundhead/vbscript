' Modify User Account Address Attributes

Set objUser = GetObject("LDAP://cn=UserName,ou=OU,dc=fabrikam,dc=com") 
 
objUser.Put "streetAddress", "Building 43" & _
    VbCrLf & "One Microsoft Way"
objUser.Put "l", "Redmond"
objUser.Put "st", "Washington"
objUser.Put "postalCode", "98053"
objUser.Put "c", "US"
objUser.Put "postOfficeBox", "2222"
objUser.SetInfo

