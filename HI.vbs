Set objSysInfo = CreateObject("ADSystemInfo")
strUser = objSysInfo.UserName
Set objUser = GetObject("LDAP://" & strUser)
strFullName = objUser.Get("displayName")
MsgBox ("hai.. " & strFullName & " Please check your desktop for SLD EXCEL FOLDER GOOD LUCK......!!")