Attribute VB_Name = "mTools"
Option Explicit
Public userid As String
Public pword As String

Private Declare Sub Sleep Lib "KERNEL32" (ByVal dwMilliseconds As Long)

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
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim a As Integer
Dim b As Integer
Dim strDesignation As String
Dim rsDesignation As New ADODB.Recordset

With cn
        .CursorLocation = adUseClient
        .ConnectionString = "Provider=SQLOLEDB;Data Source=SD_SQL_TRAINING ;Initial Catalog=Jai;UID=sa ; PWD=81at84"
        .Open
End With

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


    Do While Not rs.EOF
            For a = 1 To rs.RecordCount
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

End Sub

