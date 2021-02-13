VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm_select_users 
   Caption         =   "Select Users"
   ClientHeight    =   7935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6165
   LinkTopic       =   "Form1"
   ScaleHeight     =   7935
   ScaleWidth      =   6165
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSearch 
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
      Height          =   375
      Left            =   4920
      TabIndex        =   4
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   7320
      Width           =   1095
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   6135
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   10821
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.TextBox txtSearch 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   240
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Search: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frm_select_users"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSelect_Click()
sel_empid = "test"
End Sub

Private Sub Form_Load()
Dim strUsers1 As String
strUsers1 = ""
Call loadusers(strUsers1)

stremployees = " SELECT  a.EmployeeCode , a.EmployeeName  , b.DepartmentName  , c.SectionName  , " & _
"e.TeamName   , d.DesignationName,CASE WHEN a.Gender  IS NULL THEN '' ELSE  a.gender END Gender ," & _
 " CASE WHEN a.RetiredDate      IS NULL THEN ' ' ELSE   convert(VARCHAR(30) , a.RetiredDate     , 111)" & _
 " END  [Retired Date],  CASE WHEN a.RegularHiredDate   IS NULL  THEN '' ELSE    convert(VARCHAR(30) " & _
 " , a.RegularHiredDate , 111) END [Regular Date]  , a.ContractStatus   ,CASE WHEN a.DateBirth IS NULL " & _
 " THEN '' ELSE   convert(VARCHAR(10),a.DateBirth , 111 ) END Birthday , a.UpdatedDate " & _
 " from Employees a " & _
" INNER JOIN Departments b  ON b.DepartmentCode  =a.DepartmentCode " & _
" INNER JOIN Sections  c ON c.SectionCode      =a.SectionCode " & _
" INNER JOIN  Designations     d ON d.DesignationCode =a.DesignationCode" & _
" INNER JOIN Teams     e ON e.TeamCode     =a.TeamCode "

Call Refresh_employees(stremployees, True)
End Sub

