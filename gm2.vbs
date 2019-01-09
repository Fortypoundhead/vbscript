Option Explicit

Dim objRootDSE, strDNSDomain, adoConnection
Dim strBase, strFilter, strAttributes, strQuery, adoRecordset
Dim strGroup, objList

Const adOpenStatic = 3
Const adLockOptimistic = 3
Const adUseClient = 3

' Dictionary object to track duplicates.
Set objList = CreateObject("Scripting.Dictionary")
objList.CompareMode = vbTextCompare

' Specify DN of group to enumerate.
strGroup = "cn=domain admins,dc=trueblueinc,dc=com"

' Add to dictionary object.
objList.Add strGroup, True

' Determine DNS domain name.
Set objRootDSE = GetObject("LDAP://RootDSE")
strDNSDomain = objRootDSE.Get("defaultNamingContext")

' Use ADO to search Active Directory.
Set adoConnection = CreateObject("ADODB.Connection")
adoConnection.Provider = "ADsDSOObject"
adoConnection.Open "Active Directory Provider"

Set adoRecordset = CreateObject("ADODB.Recordset")
Set adoRecordset.ActiveConnection = adoConnection
adoRecordset.CursorLocation = adUseClient
adoRecordset.CursorType = adOpenStatic
adoRecordset.LockType = adLockOptimistic

' Search entire domain.
strBase = "<LDAP://" & strDNSDomain & ">"

' Filter on all group objects.
strFilter = "(objectCategory=group)"

' Comma delimited list of attribute values to retrieve.
strAttributes = "distinguishedName,member"

' Construct the LDAP query.
strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"

' Run the query.
adoRecordset.Source = strQuery
adoRecordset.Open

' Disconnect the recordset.
Set adoRecordset.ActiveConnection = Nothing
adoConnection.Close

' Enumerate members of the group.
Call EnumMembers(adoRecordset, strGroup, "")

' Clean up.
adoRecordset.Close

Sub EnumMembers(adoDiscRS, strGroupDN, strOffset)
    ' Subroutine to filter disconnected recordset to enumerate
    ' members of strGroupDN. The object reference objList must
    ' have global scope. strOffset shows how groups are nested.

    Dim arrMembers, strMember

    ' Filter the recordset on the group to be enumerated.
    adoDiscRS.Filter = "distinguishedName='" & strGroupDN & "'"

    ' If strGroupDN is not a group, recordset will be empty.
    If adoDiscRS.EOF Then
        Exit Sub
    End If

    ' Retrieve direct members of strGroupDN.
    adoDiscRS.MoveFirst
    Do Until adoDiscRS.EOF
        arrMembers = adoDiscRS.Fields("member").Value
        adoDiscRS.MoveNext
    Loop

    If Not IsNull(arrMembers) Then
        ' Enumerate direct members of strGroupDN.
        For Each strMember In arrMembers
            ' Check if this member seen before. This avoids infinite
            ' loop if group nesting is circular.
            If (objList.Exists(strMember) = True) Then
                Wscript.Echo strOffset & strMember & " (Duplicate)"
            Else
                Wscript.Echo strOffset & strMember
                ' Add to dictionary object.
                objList.Add strMember, True
                ' Call method recursively to reveal nested membership.
                Call EnumMembers(adoDiscRS, strMember, strOffset & "  ")
            End If
        Next
    End If
End Sub
