Attribute VB_Name = "mconnection"
Option Explicit
Public cn As New ADODB.Connection
Public rs As New ADODB.Recordset
Public strconnection As String



'Public Const strConnectionString As String = "Provider=SQLOLEDB.1;Password=81at84;Persist Security Info=True;User ID=sa;Initial Catalog=Jai;Data Source=SD_SQL_TRAINING_"

Public Sub Connect()
strconnection = "Provider=SQLOLEDB;Data Source=SD_SQL_TRAINING ; Initial Catalog=Jai;UID=sa ; PWD=81at84"
       cn.CursorLocation = adUseClient
        cn.ConnectionString = strconnection
        cn.Open

 End Sub


Public Sub Disconnect()
cn.Close
Set cn = Nothing


End Sub
