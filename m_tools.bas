Attribute VB_Name = "mTools"
Option Explicit
Public userid As String
Public pword As String
Public LoggedinEmployeeid As String
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Const SIZEOF_INT64 As Long = 8
Public Type Int64 'LowPart must be the first one in LittleEndian systems
    'required part
    LowPart  As Long
    HighPart As Long

    'optional part
    SignBit As Byte 'define this as long if you want to get minimum CPU access time.
End Type
'with the SignBit you can emulate both Int64 and UInt64 without changing the real sign bit in HighPart.
'but if you want to change it you can access it like this  mySign = (myVar.HighPart And &H80000000)
'or turn on the sign bit using  myVar.HighPart = (myVar.HighPart Or &H80000000)

Public Function CInt64(ByVal vCur As Currency) As Int64
    vCur = (CCur(vCur) * 0.0001@)
    Call CopyMemory(CInt64, vCur, SIZEOF_INT64)
End Function

Sub main()
On Error GoTo FetchErr
If App.PrevInstance = True Then MsgBox ("System Already Open"), vbOKOnly, "Message": Exit Sub
frmSplash.Show
DoEvents


Sleep 2000
frmLogin.Show

Unload frmSplash
Exit Sub
FetchErr:
MsgBox Err.Number & " : " & Err.Description, vbCritical, "Error"

End Sub

Public Sub Refresh_employees(strfilter As String, IsLoad As Boolean)
Dim rs As New ADODB.Recordset
Dim a As Long
Dim b As Integer
Dim strDesignation As String
Dim rsDesignation As New ADODB.Recordset

Dim totalrec As Double

Call Connect
rs.Open strfilter, cn, adOpenDynamic, adLockReadOnly

With frm_m_employees.msh_emp
    .Cols = 12
    .Rows = rs.RecordCount + 1
    .TextMatrix(0, 0) = "Employee Code"
    .TextMatrix(0, 1) = "Employee Name"
    .TextMatrix(0, 2) = "Department"
    .TextMatrix(0, 3) = "Section "
    .TextMatrix(0, 4) = "Team "
    .TextMatrix(0, 5) = "Designation  "
    .TextMatrix(0, 6) = "Gender "
    .TextMatrix(0, 7) = "Retired Date "
    .TextMatrix(0, 8) = "Regular Date "
    .TextMatrix(0, 9) = "Contract Status "
    .TextMatrix(0, 10) = "Date Birth "
    .TextMatrix(0, 11) = "Updated Date"
    
    totalrec = rs.RecordCount
    
    Do While Not rs.EOF
            For a = 1 To totalrec
                For b = 0 To .Cols - 1
                    .TextMatrix(a, 0) = rs.Fields(0).Value
                    .TextMatrix(a, 1) = rs.Fields(1).Value
                    .TextMatrix(a, 2) = rs.Fields(2).Value
                    .TextMatrix(a, 3) = rs.Fields(3).Value
                    .TextMatrix(a, 4) = rs.Fields(4).Value
                    .TextMatrix(a, 5) = rs.Fields(5).Value
                    .TextMatrix(a, 6) = rs.Fields(6).Value
                    .TextMatrix(a, 7) = rs.Fields(7).Value
                    .TextMatrix(a, 8) = rs.Fields(8).Value
                    .TextMatrix(a, 9) = rs.Fields(9).Value
                    .TextMatrix(a, 10) = rs.Fields(10).Value
                      .TextMatrix(a, 11) = rs.Fields(11).Value
                Next
            rs.MoveNext
            Next
    Loop
    
            .ColAlignment(0) = flexAlignCenterTop
            .WordWrap = True
            .ColWidth(0) = 1200
            .ColWidth(1) = 3000
             .ColWidth(2) = 1500
               .ColWidth(3) = 1500
            .RowHeight(0) = 500
            
            
            With frm_m_employees
            If IsLoad = True Then
                .lbltotalRec = rs.RecordCount
                .cbodept.Visible = False
                .cboSections.Visible = False
                .txtEmpCode.Visible = True
                .lblFind.Visible = True
                .cboDesignation.Clear
                strDesignation = "Select DesignationName from Designations"
                rsDesignation.Open strDesignation, cn, adOpenDynamic, adLockReadOnly
                Do While Not rsDesignation.EOF
                        .cboDesignation.AddItem rsDesignation.Fields(0).Value
                        rsDesignation.MoveNext
                Loop
                End If
              End With
End With
Call Disconnect
End Sub

Public Function loadusers(strfilter As String) As Object
Dim rs_Users As New ADODB.Recordset
Dim counterA As Integer
Dim counterB As Integer


Call Connect
rs_Users.Open strfilter, cn, adOpenDynamic, adLockReadOnly

With frm_m_user.msh_users
  If rs_Users.EOF Then
  MsgBox "No record Found", vbInformation, "Information"
  .Clear
  End If
  
            .Cols = 6
                .Rows = rs_Users.RecordCount + 1
                .FixedRows = 0
                .TextMatrix(0, 0) = "Username"
                .TextMatrix(0, 1) = "Name"
                .TextMatrix(0, 2) = "Password"
                .TextMatrix(0, 3) = "User Level"
                .TextMatrix(0, 4) = "Updated Date"
                .TextMatrix(0, 5) = "Updated By"
           
                
                
                
     Do While Not rs_Users.EOF
                For counterA = 1 To rs_Users.RecordCount
                        For counterB = 0 To .Cols - 1
                            .TextMatrix(counterA, 0) = rs_Users(0).Value
                            .TextMatrix(counterA, 1) = rs_Users(1).Value
                            .TextMatrix(counterA, 2) = rs_Users(2).Value
                            .TextMatrix(counterA, 3) = rs_Users(3).Value
                            .TextMatrix(counterA, 4) = rs_Users(4).Value
                            
                            .TextMatrix(counterA, 5) = rs_Users(5).Value
                            .ColWidth(1) = 3000
                            .ColWidth(4) = 2500
                            .ColWidth(5) = 2500
                              .ColAlignment(counterB) = flexAlignLeftCenter
                                       
                 
                        Next

                rs_Users.MoveNext
                Next

    Loop
    
            End With
       Set loadusers = rs_Users
    
End Function




