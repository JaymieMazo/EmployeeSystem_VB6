VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm_m_employees 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Master->Employees"
   ClientHeight    =   12570
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   17475
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   12570
   ScaleWidth      =   17475
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Height          =   11295
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   17175
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00FFFF80&
         Caption         =   "&Delete"
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
         Left            =   15840
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton cmdEdit 
         BackColor       =   &H00FFFF80&
         Caption         =   "&Edit"
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
         Left            =   14520
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   840
         Width           =   1215
      End
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
         Left            =   13200
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   840
         Width           =   1215
      End
      Begin VB.ComboBox cboTeams 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frm_m_employees.frx":0000
         Left            =   9240
         List            =   "frm_m_employees.frx":0002
         TabIndex        =   14
         Text            =   "[select teams]"
         Top             =   960
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00FFFF80&
         Caption         =   "&Clear"
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
         Left            =   10560
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdFind 
         BackColor       =   &H00FFFF80&
         Caption         =   "&Find"
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
         Left            =   9240
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   360
         Width           =   1215
      End
      Begin VB.ComboBox cboDesignation 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frm_m_employees.frx":0004
         Left            =   1920
         List            =   "frm_m_employees.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   960
         Width           =   3135
      End
      Begin VB.TextBox txtEmpCode 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5880
         TabIndex        =   8
         Top             =   360
         Width           =   3135
      End
      Begin VB.ComboBox cboSearchby 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frm_m_employees.frx":0008
         Left            =   1920
         List            =   "frm_m_employees.frx":001B
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   360
         Width           =   3135
      End
      Begin VB.ComboBox cbodept 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frm_m_employees.frx":0065
         Left            =   5160
         List            =   "frm_m_employees.frx":0067
         TabIndex        =   6
         Text            =   "[select department]"
         Top             =   360
         Width           =   3975
      End
      Begin VB.ComboBox cboSections 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frm_m_employees.frx":0069
         Left            =   5160
         List            =   "frm_m_employees.frx":006B
         TabIndex        =   5
         Text            =   "[select section]"
         Top             =   960
         Visible         =   0   'False
         Width           =   3975
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh_emp 
         Height          =   9375
         Left            =   120
         TabIndex        =   4
         Top             =   1560
         Width           =   16935
         _ExtentX        =   29871
         _ExtentY        =   16536
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         FixedCols       =   0
         BackColorFixed  =   8421376
         ForeColorFixed  =   -2147483637
         BackColorBkg    =   16777215
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
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label5 
         BackColor       =   &H00404040&
         Caption         =   "Designation: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   360
         TabIndex        =   18
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label lblFind 
         BackColor       =   &H00404040&
         Caption         =   "Find:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   5160
         TabIndex        =   11
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00404040&
         Caption         =   "Filter by:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   840
         TabIndex        =   9
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Records:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   12000
      Width           =   1935
   End
   Begin VB.Label lbltotalRec 
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Top             =   12000
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Employees"
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
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frm_m_employees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim IsEmpcode As Boolean
Dim IsEmpname As Boolean
Dim IsDesignation As Boolean
Dim IsDep As Boolean
Dim isSec As Boolean
Dim isTeam As Boolean
Dim stremployees As String

Private Sub cbodept_Click()
        cboTeams.Visible = False
        Call load_dept(False)
        cboSections.Visible = True
        IsDep = True
        isSec = False
        isTeam = False
End Sub


Private Sub cboDesignation_Click()
IsDesignation = True
End Sub

Private Sub cboSearchby_Click()
Dim intlist As Integer

intlist = cboSearchby.ListIndex
IsDep = False
isSec = False
isTeam = False
IsEmpcode = False
IsEmpname = False
txtEmpCode = ""
    If intlist = 2 Then
    cbodept.Text = "[select department]"
        Call load_dept(True)
        cbodept.Visible = True
        txtEmpCode = ""
        txtEmpCode.Visible = False
        lblFind.Visible = False
    Else
    cbodept.ListIndex = -1
        txtEmpCode.Enabled = True
        cbodept.Visible = False
        txtEmpCode.Visible = True
        lblFind.Visible = True
        cboSections.Visible = False
    End If
End Sub
Public Sub load_dept(isSec As Boolean)
    Dim rs As New ADODB.Recordset
    Dim strdep As String
    
        If isSec = True Then
            strdep = "Select * from departments"
        Else
        cboSections.Clear
        cboSections.Text = "[select section]"
            strdep = "SELECT a.SectionName  FROM Sections a " & _
                    " INNER JOIN  DepartmentSectionRelations b ON a.SectionCode=b.SectionCode " & _
                    " INNER JOIN Departments   c ON c.DepartmentCode  =b.DepartmentCode" & _
                    " WHERE  c.Departmentname='" & cbodept.List(cbodept.ListIndex) & "' "

        End If
        
    Call Connect
    rs.Open strdep, cn, adOpenDynamic, adLockReadOnly
    
    Do While Not rs.EOF
        If isSec = True Then
            cbodept.AddItem rs.Fields(1).Value
        Else
            cboSections.AddItem rs.Fields(0).Value
        End If
        rs.MoveNext
    Loop
    
    Call Disconnect
End Sub
Private Sub cboSections_Click()
cboTeams.Text = "[select Teams]"
isSec = True
Call Load_Teams
End Sub



Private Sub cboTeams_Click()
isTeam = True
End Sub

Private Sub cmdClear_Click()
Call ClearSearch
Call Refresh_employees(stremployees, True)
End Sub

Private Sub cmdFind_Click()
Dim strfind As String
Dim strfilter As String
Dim rs As New ADODB.Recordset
Dim cn1 As New ADODB.Connection




strfind = ""
'MsgBox "dep-sec-teams-desig-empcode-empname: " & IsDep & isSec & isTeam & IsDesignation & IsEmpcode & IsEmpname


If IsDep <> False Or isSec <> False Or isTeam <> False Or IsDesignation <> False Or IsEmpcode <> False Or IsEmpname <> False Then
     strfind = " where  "
        If IsDep <> False Then
        strfind = Trim(strfind) & "  b.DepartmentName    = '" & cbodept.List(cbodept.ListIndex) & "' AND "
            If isSec <> False Then
                strfind = Trim(strfind) & " c.SectionName   ='" & cboSections.List(cboSections.ListIndex) & "' AND "
            End If
            
            If isTeam <> False Then
                 strfind = Trim(strfind) & " e.TeamName   ='" & cboTeams.List(cboTeams.ListIndex) & "' AND"
            End If
            
        End If
        
        If IsDesignation <> False Then
                   strfind = Trim(strfind) & " d.DesignationName   ='" & cboDesignation.List(cboDesignation.ListIndex) & "' AND"
        End If
        
        If IsEmpcode <> False Then
                 strfind = Trim(strfind) & "  a.EmployeeCode  like '%" & txtEmpCode & "%' AND"
        ElseIf IsEmpname <> False Then
                 strfind = Trim(strfind) & "  a.Employeename  like '%" & txtEmpCode & "%' AND "
        End If
                strfilter = Trim(stremployees) & " " & Trim(strfind)
        strfilter = StrReverse(Mid(StrReverse(Trim(strfilter)), 4, Len(Trim(strfilter)) - 3))

Call Refresh_employees(strfilter, False)
Else
MsgBox "Please input something first ", vbCritical, "Oops!"
End If


End Sub

Private Sub Form_Load()
'stremployees = " SELECT  a.EmployeeCode , a.EmployeeName  , b.DepartmentName  , c.SectionName  , " & _
'"e.TeamName   , d.DesignationName,CASE WHEN a.Gender  IS NULL THEN '' ELSE  a.gender END Gender ," & _
' " CASE WHEN a.RetiredDate      IS NULL THEN ' ' ELSE   convert(VARCHAR(30) , a.RetiredDate     , 111)" & _
' " END  [Retired Date],  CASE WHEN a.RegularHiredDate   IS NULL  THEN '' ELSE    convert(VARCHAR(30) " & _
' " , a.RegularHiredDate , 111) END [Regular Date]  , a.ContractStatus   ,CASE WHEN a.DateBirth IS NULL " & _
' " THEN '' ELSE   convert(VARCHAR(10),a.DateBirth , 111 ) END Birthday , a.UpdatedDate " & _
' " from Employees a " & _
'" INNER JOIN Departments b  ON b.DepartmentCode  =a.DepartmentCode " & _
'" INNER JOIN Sections  c ON c.SectionCode      =a.SectionCode " & _
'" INNER JOIN  Designations     d ON d.DesignationCode =a.DesignationCode" & _
'" LEFT JOIN Teams     e ON e.TeamCode     =a.TeamCode "

stremployees = " SELECT  a.EmployeeCode , a.EmployeeName  , b.DepartmentName  , c.SectionName  ," & _
                " CASE WHEN e.TeamName  IS NULL  THEN '' ELSE E.TeamName  END TeamName  , " & _
                " d.DesignationName,CASE WHEN a.Gender  IS NULL THEN '' ELSE  a.gender END Gender , " & _
                " CASE WHEN a.RetiredDate      IS NULL THEN ' '  ELSE   convert(VARCHAR(30) , a.RetiredDate   " & _
                " , 111) END  [Retired Date], CASE WHEN a.RegularHiredDate   IS NULL  THEN '' ELSE  " & _
                "convert(VARCHAR(30) , a.RegularHiredDate , 111) END [Regular Date] , a.ContractStatus   , " & _
                " CASE WHEN a.DateBirth IS NULL THEN '' ELSE   convert(VARCHAR(10),a.DateBirth , 111 )  " & _
                " END Birthday , a.UpdatedDate  from Employees a " & _
                " INNER JOIN Departments b  ON b.DepartmentCode  =a.DepartmentCode " & _
                " INNER JOIN Sections  c ON c.SectionCode      =a.SectionCode " & _
                " INNER JOIN  Designations     d ON d.DesignationCode =a.DesignationCode" & _
                " LEFT JOIN Teams     e ON e.TeamCode     =a.TeamCode "
Call Refresh_employees(stremployees, True)

End Sub

Private Sub txtEmpCode_Change()
    If Trim(txtEmpCode.Text) <> "" Then
         If cboSearchby.ListIndex = 0 Then
            IsEmpcode = True
            IsEmpname = False
        Else
            IsEmpname = True
            IsEmpcode = False
        End If
        IsDep = False
        isSec = False
        isTeam = False
    End If
End Sub
Public Sub Load_Teams()
Dim rs As New ADODB.Recordset
Dim strTeams As String
strTeams = " SELECT b.TeamName FROM  SectionTeamRelations a " & _
            " INNER JOIN Teams b ON a.TeamCode     =b.TeamCode" & _
            " INNER    JOIN Sections       c ON a.sectioncode     =c.sectioncode " & _
            " WHERE   sectionname='" & cboSections.List(cboSections.ListIndex) & "' "
Call Connect
rs.Open strTeams, cn, adOpenDynamic, adLockReadOnly

If rs.RecordCount = 0 Then
Else

cboTeams.Clear
cboTeams.Visible = True
cboTeams.Text = "[seSlect teams]"

End If

Do While Not rs.EOF
    cboTeams.AddItem rs.Fields(0).Value
    rs.MoveNext
Loop
End Sub


Public Sub ClearSearch()
txtEmpCode = ""
cboSections.Clear
cbodept.Clear
cboTeams.Clear
cboTeams.Visible = False
cboDesignation.ListIndex = -1
cboSearchby.ListIndex = -1
cbodept.ListIndex = -1
cboSections.ListIndex = -1
cboTeams.ListIndex = -1
IsDep = False
isSec = False
isTeam = False
IsDesignation = False
IsEmpcode = False
IsEmpname = False
End Sub
