VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm_m_user 
   BackColor       =   &H00400000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Master-> Users"
   ClientHeight    =   7845
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12645
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   12645
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
      Height          =   375
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1680
      Width           =   1300
   End
   Begin VB.CommandButton Command1 
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
      Height          =   375
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1680
      Width           =   1300
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
      Height          =   375
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1680
      Width           =   1300
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00400000&
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
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   12495
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
         Height          =   375
         Left            =   10800
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   600
         Width           =   1300
      End
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H00FFFF80&
         Caption         =   "&Search"
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
         Left            =   9360
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   600
         Width           =   1300
      End
      Begin VB.TextBox txtuname 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5880
         TabIndex        =   5
         Top             =   600
         Width           =   2895
      End
      Begin VB.TextBox txtuid 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   3
         Top             =   600
         Width           =   2895
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "User Name: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   4560
         TabIndex        =   4
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "User ID: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   1215
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh_users 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   2160
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   9763
      _Version        =   393216
      BackColor       =   4194304
      ForeColor       =   16777215
      FixedCols       =   0
      BackColorBkg    =   4194304
      GridColorUnpopulated=   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "frm_m_user"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdAdd_Click()
Dim checkuser As New clschecking
Dim isAdmin As Boolean

isAdmin = checkuser.checkuser
If isAdmin = True Then
frmadduser.Show
Else
MsgBox "You are not authorized to add new user. Pls Contact Administrator. Thank you!", vbCritical, "Warning"
End If
End Sub

Private Sub cmdClear_Click()
txtuid = ""
txtUname = ""
Call All_users
End Sub

Private Sub cmdSearch_Click()
Dim strUsers As String
Dim rsSearch As New ADODB.Recordset

If Trim(txtuid) <> "" Or Trim(txtUname) <> "" Then
        strUsers = " Select userid ,c.EmployeeName,  password , b.userlevel, a.updatedate , d.EmployeeName from users a" & _
        " inner  join userlevels b on a.userrights=b.userlevelid" & _
        " inner JOIN  Employees c  ON c.EmployeeCode=a.UserID " & _
        " inner JOIN  Employees d ON  d.EmployeeCode=a.UpdatedBy"

        If Trim(txtuid) <> "" And Trim(txtUname) <> "" Then
          strUsers = strUsers & "  WHERE  userid like '%" & txtuid & "%' and   c.employeename like '%" & txtUname & "%'"
        Else
        
                If Trim(txtuid) <> "" Then
                           strUsers = strUsers & " WHERE userid like '%" & txtuid & "%' "
                ElseIf Trim(txtUname) <> "" Then
                             strUsers = strUsers & " WHERE  c.employeename like '%" & txtUname & "%'"
                End If
         
        End If
        
    
Set rsSearch = loadusers(strUsers)
Call Disconnect
Else
MsgBox "Please input filter first!", vbCritical, "Oops!"
End If
End Sub

Private Sub Form_Load()
Call All_users
End Sub

Public Sub All_users()
Dim counterA As Integer
Dim counterB As Integer
Dim rsUsers  As New ADODB.Recordset
Dim strqry As String

strqry = " Select userid ,c.EmployeeName,  password , b.userlevel, a.updatedate , d.EmployeeName from users a" & _
        " inner  join userlevels b on a.userrights=b.userlevelid" & _
        " inner JOIN  Employees c  ON c.EmployeeCode=a.UserID " & _
        " inner JOIN  Employees d ON  d.EmployeeCode=a.UpdatedBy"
        
        
Set rsUsers = loadusers(strqry)
Call Disconnect
End Sub

