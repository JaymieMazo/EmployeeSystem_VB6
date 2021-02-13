VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm_m_user 
   BackColor       =   &H00400000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Master-> Users"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   9735
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00FFFF80&
      Caption         =   "&Add"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5040
      Width           =   1455
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh_users 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   7011
      _Version        =   393216
      BackColor       =   4194304
      ForeColor       =   16777215
      FixedCols       =   0
      BackColorBkg    =   4194304
      GridColorUnpopulated=   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Users"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "frm_m_user"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
frmadduser.Show
End Sub

Private Sub Form_Load()
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim strqry As String
Dim rowcount As Integer
Dim counterA As Integer
Dim counterB As Integer



strqry = " Select userid ,c.EmployeeName,  password , b.userlevel, a.updatedate , d.EmployeeName from users a" & _
        " inner join userlevels b on a.userrights=b.userlevelid" & _
        " INNER JOIN  Employees c  ON c.EmployeeCode=a.UserID " & _
        " INNER JOIN  Employees d ON  d.EmployeeCode=a.UpdatedBy"

With cn
.CursorLocation = adUseClient
.ConnectionString = "Provider=SQLOLEDB;Data Source=SD_SQL_TRAINING ; Initial Catalog=Jai;UID=sa ;PWD=81at84"
.Open
End With

rs.Open strqry, cn, adOpenDynamic, adLockReadOnly

Do While Not rs.EOF
If rs.EOF Then MsgBox "No record Found", vbInformation, "Information": Exit Sub


    With msh_users
            .Cols = 6
            .Rows = rs.RecordCount + 1
            .TextMatrix(0, 0) = "Username"
            .TextMatrix(0, 1) = "Name"
            .TextMatrix(0, 2) = "Password"
            .TextMatrix(0, 3) = "User Level"
            .TextMatrix(0, 4) = "Updated Date"
            .TextMatrix(0, 5) = "Updated By"
           
             

            For counterA = 1 To rs.RecordCount
                    For counterB = 0 To .Cols - 1
                        .TextMatrix(counterA, 0) = rs(0).Value
                        .TextMatrix(counterA, 1) = rs(1).Value
                        .TextMatrix(counterA, 2) = rs(2).Value
                        .TextMatrix(counterA, 3) = rs(3).Value
                        .TextMatrix(counterA, 4) = rs(4).Value
                        .TextMatrix(counterA, 5) = rs(5).Value
                
                        .ColWidth(1) = 2000
                        .ColWidth(4) = 2000
                        .ColWidth(5) = 2000
                    Next
                    rs.MoveNext
            Next
    End With
    
    
    
Loop
        
        

End Sub
