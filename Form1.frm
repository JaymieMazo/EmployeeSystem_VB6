VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form test 
   Caption         =   "Form1"
   ClientHeight    =   6075
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8820
   LinkTopic       =   "Form1"
   ScaleHeight     =   6075
   ScaleWidth      =   8820
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   5160
      TabIndex        =   1
      Top             =   480
      Width           =   1455
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh 
      Height          =   3135
      Left            =   1920
      TabIndex        =   0
      Top             =   1800
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   5530
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "test"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim str As String


str = "select * from users"

With cn
.ConnectionString = "Provider=SQLOLEDB;Data Source=SD_SQL_TRAINING; Initial Catalog=Jai; UID=sa ; PWD=81at84"
.CursorLocation = adUseClient
.ConnectionString = str
End With

 rs.Open str, cn, adOpenDynamic, adLockReadOnly






With msh
            

End With






End Sub
