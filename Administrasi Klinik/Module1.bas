Attribute VB_Name = "Module1"
Public Connect As ADODB.Connection
Public RsAdmin As ADODB.Recordset

Public lokasidb As String

Public Sub connectdb()
Set Connect = New ADODB.Connection
Set RsAdmin = New ADODB.Recordset

lokasidb = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\dbLogin.mdb"
Connect.Open lokasidb
End Sub
