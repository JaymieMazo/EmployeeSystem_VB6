VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00800000&
   BorderStyle     =   0  'None
   ClientHeight    =   1230
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9375
   FillColor       =   &H00400000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1230
   ScaleWidth      =   9375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmSplash.frx":0000
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   120
      Width           =   8175
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Copyright © July 2018  HRD-SMD-SD, All Rights Reserved"
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   840
      Width           =   4335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Employee Management System"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   8175
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   120
      Picture         =   "frmSplash.frx":00A3
      Stretch         =   -1  'True
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()
Unload Me
End Sub
