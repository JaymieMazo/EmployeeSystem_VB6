VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clschecking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Function getUserlevelid(ByVal userlevel As String) As Integer
Dim strqry As String
Dim rs As New ADODB.Recordset

strqry = "select Userlevelid from userlevels where Userlevel ='" & userlevel & "'"
Call Connect
rs.Open strqry, cn, adOpenDynamic, adLockReadOnly
rs.MoveLast
 getUserlevelid = rs.Fields(0).Value
 Call Disconnect
End Function

Public Function checkuser() As Boolean
Dim rsUser As New ADODB.Recordset
Dim strUser As String

strUser = "select * from users where  userid ='" & LoggedinEmployeeid & "' and  UserRights=1"
Call Connect

rsUser.Open strUser, cn, adOpenDynamic, adLockReadOnly
If rsUser.RecordCount = 0 Then

checkuser = False
Else
checkuser = True
End If
Call Disconnect
End Function

