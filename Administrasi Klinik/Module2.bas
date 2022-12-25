Attribute VB_Name = "Module2"
Option Explicit
Public con As ADODB.Connection
Public Rs As ADODB.Recordset
Public RsLpp As ADODB.Recordset

Sub Connects()
Set con = New ADODB.Connection
Set RsLpp = New ADODB.Recordset
con.Provider = "Microsoft.jet.oledb.4.0"
con.Open App.Path & "\Laporan Penjualan & Pelayanan.mdb"
End Sub
