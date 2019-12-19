Option Explicit
On Error Resume Next

Dim objNetwork, mappedDrives, objShell
Dim objFSO
Dim WshShell
Dim ObjUser, ObjRootDSE, ObjConn, ObjRS
Dim GroupCollection, ObjGroup
Dim StrUserName, StrDomName, StrSQL
Dim grouplistD
Dim objADSysInf

'Auslesen der notwendigen Umgebungsvariablen
Set ObjRootDSE = GetObject("LDAP://RootDSE")
StrDomName = Trim(ObjRootDSE.Get("DefaultNamingContext"))
Set ObjRootDSE = Nothing

Set objNetwork = CreateObject("WScript.Network") 
Set mappedDrives = objNetwork.EnumNetworkDrives
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set WshShell = WScript.CreateObject("WScript.Shell") 

Set objADSysInf = CreateObject("ADSystemInfo")

' *****************************************************
'This function returns a particular environment variable's value.
' for example, if you use EnvString("username"), it would return
' the value of %username%.
Function EnvString(variable)
set objShell = WScript.CreateObject( "WScript.Shell" )
    variable = "%" & variable & "%"
    EnvString = objShell.ExpandEnvironmentStrings(variable)
End Function
' *****************************************************

'******************************************************************************************
'                              Check User Groups:
'******************************************************************************************

StrUserName = EnvString("username")
StrSQL = "Select ADsPath From 'LDAP://" & StrDomName & "' Where ObjectCategory = 'User' AND SAMAccountName = '" & StrUserName & "'"

Set ObjConn = CreateObject("ADODB.Connection")
ObjConn.Provider = "ADsDSOObject":	ObjConn.Open "Active Directory Provider"
Set ObjRS = CreateObject("ADODB.Recordset")
ObjRS.Open StrSQL, ObjConn

' *****************************************************
'This function checks to see if the passed group name contains the current
' user as a member. Returns True or False
Function MemberOf(groupName)
    If IsEmpty(groupListD) then
        Set groupListD = CreateObject("Scripting.Dictionary")

        If Not ObjRS.EOF Then
            ObjRS.MoveLast:	ObjRS.MoveFirst
            Set ObjUser = GetObject (Trim(ObjRS.Fields("ADsPath").Value))
            Set GroupCollection = ObjUser.Groups
            For Each ObjGroup In GroupCollection
                groupListD.Add Trim(ObjGroup.CN), "-"
            Next
            Set ObjGroup = Nothing:	Set GroupCollection = Nothing:	Set ObjUser = Nothing
        End If
        ObjRS.Close:	Set ObjRS = Nothing
        ObjConn.Close:	Set ObjConn = Nothing
    End if
    MemberOf = CBool(groupListD.Exists(groupName))
End Function
' *****************************************************

'******************************************************************************************
'                                      Disconnect Shares:
'******************************************************************************************
Dim x
For x = 0 to mappedDrives.Count-1 Step 2  
    objNetwork.RemoveNetworkDrive mappedDrives.Item(x), True, True
Next 

'******************************************************************************************
'                                      Little Break
'******************************************************************************************

WScript.Sleep 1000 'delay is in milliseconds

'******************************************************************************************
'                                      Connect Shares:
'******************************************************************************************

If MemberOf("xyz") Then 'check group membership
    objNetwork.MapNetworkDrive "X:","\\srv\share$", True 'connect share
End if
