option explicit

Dim objArgs, strGroupDN
set objArgs = WScript.Arguments
if objArgs.Count <> 1 then
   Dim objRootDSE
   set objRootDSE = GetObject("LDAP://RootDSE")
   'Corp.Trueblueinc.com/Migration/MigLRW/LRWGroups/Security Groups/Auditdb Access
   strGroupDN = "cn=Group,ou=OrgUnit," & objRootDSE.Get("defaultNamingContext")
else
   strGroupDN = objArgs.Item(0)
end if

Dim dicSeenGroupMember
set dicSeenGroupMember = CreateObject("Scripting.Dictionary")
Wscript.Echo "Members of " & strGroupDN & ":"
DisplayMembers "LDAP://" & strGroupDN, " ", dicSeenGroupMember

Function DisplayMembers (strGroupADsPath, strSpaces, dicSeenGroupMember)

   Dim objGroup, objMember
   set objGroup = GetObject(strGroupADsPath)
   for each objMember In objGroup.Members

      Wscript.Echo strSpaces & objMember.Get("distinguishedname")
      if objMember.Class = "group" then

         if dicSeenGroupMember.Exists(objMember.ADsPath) then
            Wscript.Echo strSpaces & "   ^ already seen group member " & _
                                     "(stopping to avoid loop)"
         else
            dicSeenGroupMember.Add objMember.ADsPath, 1
            DisplayMembers objMember.ADsPath, strSpaces & "  ", _
                           dicSeenGroupMember
         end if

      end if

   next
End Function

