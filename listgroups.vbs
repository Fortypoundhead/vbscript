Option Explicit 
Dim objRootLDAP, objGroup, objUser, objOU, objmemberOf 
Dim strOU, strUser, strDNSDomain, strLDAP, strList

' Commands to bind to AD and extract domain name
Set objRootLDAP = GetObject("LDAP://RootDSE")
strDNSDomain = objRootLDAP.Get("DefaultNamingContext")

' Build the LDAP DN from strUser, strOU and strDNSDomain
strUser ="cn=dsison,"
strOU ="CN=migration,CN=MigLRW,CN=LRWUsers,"
strLDAP ="LDAP://" & strUser & strOU & strDNSDomain

wscript.echo strLDAP
wscript.End

Set objUser = GetObject(strLDAP)

 ' Heart of the script, extract a list of Groups from memberOf 
objmemberOf  = objUser.GetEx("memberOf")
For Each objGroup in objmemberOf 
   strList = strList & objGroup & vbcr
Next

WScript.Echo "Groups for " & strUser & vbCr & strList

WScript.Quit

' End of Sample User memberOf  VBScript
