VERSION 5.00
Begin VB.Form frmadduser 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Creating New User"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   4665
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
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2400
      Width           =   1215
   End
   Begin VB.ComboBox cboUlvl 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1680
      TabIndex        =   4
      Top             =   1800
      Width           =   2655
   End
   Begin VB.TextBox txtPword 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   3
      Top             =   1080
      Width           =   2655
   End
   Begin VB.TextBox txtUname 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label lbluserlevel 
      BackStyle       =   0  'Transparent
      Caption         =   "User Level: "
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
      Left            =   240
      TabIndex        =   5
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
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
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label lbluname 
      BackStyle       =   0  'Transparent
      Caption         =   "Username: "
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
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "frmadduser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sel_empid As String
Public str_ulevelid As New clschecking



Private Sub cmdAdd_Click()
Dim stradd As String
If Trim(txtUname) <> "" And Trim(txtPword) <> "" And cboUlvl.ListIndex <> -1 Then
stradd = "insert into users(userid  , password , userrights , updatedby) " & _
        "values ( '" & txtUname & "', '" & txtPword & "', " & _
        str_ulevelid.getUserlevelid(cboUlvl.List(cboUlvl.ListIndex)) & " , '" & LoggedinEmployeeid & "')"
   Call Connect
   rs.Open stradd, cn, adOpenDynamic, adLockReadOnly
   MsgBox "New user Successfully addded.", vbInformation, "Information"
   Call Disconnect
Else
MsgBox "Please complete all input fields", vbCritical, "Oops!"
End If
End Sub



Private Sub Form_Load()
Dim rs As New ADODB.Recordset
Dim strUserlevel As String

strUserlevel = "select * from userlevels where deleteddate is null"
Call Connect
    rs.Open strUserlevel, cn, adOpenDynamic, adLockReadOnly
    Do While Not rs.EOF
            cboUlvl.AddItem rs(1).Value
            rs.MoveNext
    Loop
Call Disconnect
End Sub




Private Sub txtUname_DblClick()
frm_select_users.Show
End Sub
