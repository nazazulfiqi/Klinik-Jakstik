VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form5 
   ClientHeight    =   11280
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   20250
   ControlBox      =   0   'False
   LinkTopic       =   "Form5"
   ScaleHeight     =   11280
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   18480
      Top             =   4320
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Laporan Penjualan & Pelayanan.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Laporan Penjualan & Pelayanan.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Obatan"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Masukkan Data Diri"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7095
      Left            =   2160
      TabIndex        =   15
      Top             =   2160
      Width           =   5415
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         TabIndex        =   21
         Top             =   480
         Width           =   3135
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         TabIndex        =   20
         Top             =   1560
         Width           =   3135
      End
      Begin VB.TextBox Text4 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         TabIndex        =   19
         Top             =   4080
         Width           =   3135
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   1920
         TabIndex        =   18
         Top             =   5160
         Width           =   3135
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Laki-laki"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         TabIndex        =   17
         Top             =   2400
         Width           =   2415
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Perempuan"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   16
         Top             =   3120
         Width           =   2775
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Pasien"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   26
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis Kelamin"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   25
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Usia"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   24
         Top             =   4080
         Width           =   1335
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   23
         Top             =   5400
         Width           =   1455
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Pasien"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   22
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Kode Pembayaran"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   13800
      TabIndex        =   11
      Top             =   5040
      Visible         =   0   'False
      Width           =   4455
      Begin VB.Label Label22 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2160
         TabIndex        =   33
         Top             =   2640
         Width           =   1815
      End
      Begin VB.Label Label20 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2160
         TabIndex        =   29
         Top             =   4200
         Width           =   1815
      End
      Begin VB.Label Label19 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2160
         TabIndex        =   28
         Top             =   3480
         Width           =   1815
      End
      Begin VB.Label Label18 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2160
         TabIndex        =   27
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   3975
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Masukkan"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11040
      TabIndex        =   10
      Top             =   9480
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Obat-obatan"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7095
      Left            =   7800
      TabIndex        =   1
      Top             =   2160
      Width           =   4815
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   32
         Top             =   6240
         Width           =   3735
      End
      Begin VB.TextBox Text9 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   14
         Top             =   5160
         Width           =   3735
      End
      Begin VB.TextBox Text8 
         DataSource      =   "Adodc2"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   9
         Top             =   4080
         Width           =   3735
      End
      Begin VB.TextBox Text7 
         DataSource      =   "Adodc2"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   7
         Top             =   3000
         Width           =   3735
      End
      Begin VB.TextBox Text6 
         DataSource      =   "Adodc2"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   5
         Top             =   1920
         Width           =   3735
      End
      Begin VB.ComboBox Combo1 
         DataField       =   "Nama_obat"
         DataSource      =   "Adodc2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   600
         TabIndex        =   3
         Top             =   1080
         Width           =   3735
      End
      Begin VB.Label Label1 
         Caption         =   "Total Harga"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   31
         Top             =   5880
         Width           =   1215
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Beli"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   13
         Top             =   4800
         Width           =   1455
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Harga"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   8
         Top             =   3720
         Width           =   1095
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Kedaluwarsa"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   6
         Top             =   2640
         Width           =   2055
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Stock Obat"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   4
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Masukkan Obat Yang Ingin Dibeli"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   2
         Top             =   600
         Width           =   3375
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Kembali"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9360
      TabIndex        =   0
      Top             =   9480
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   11760
      Left            =   0
      Picture         =   "Form5.frx":0000
      ScaleHeight     =   11700
      ScaleWidth      =   20670
      TabIndex        =   30
      Top             =   0
      Width           =   20730
      Begin VB.CommandButton Command4 
         Caption         =   "Cetak"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   13800
         TabIndex        =   34
         Top             =   6360
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Silahkan Catat/Foto Kode Pembayaran,Lalu Klik Cetak."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   15480
         TabIndex        =   35
         Top             =   6360
         Visible         =   0   'False
         Width           =   2415
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub abis()
Label12.Caption = ""
Label11.Caption = ""
Command4.Visible = False
Frame3.Visible = False

End Sub
Private Sub Command1_Click()
pesan = MsgBox("Keluar dari Membeli Obat?", vbQuestion + vbYesNo, "Pertanyaan")
If pesan = vbYes Then
Form5.Hide
Form3.Show
Else
Form5.SetFocus
End If
End Sub


Sub bersih()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Combo1.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Option1.Value = False
Option2.Value = False

End Sub

Sub pilih()
Set RsLpp = New ADODB.Recordset
RsLpp.Open "select * from Stock_obat order by Nama_obat ASC", _
con, adOpenDynamic, _
adLockBatchOptimistic

Do While Not RsLpp.BOF And Not RsLpp.EOF
If Not RsLpp.EOF Then
Combo1.AddItem RsLpp.Fields(0)
RsLpp.MoveNext
End If
Loop

End Sub

Private Sub Combo1_Click()
Set RsLpp = New ADODB.Recordset
RsLpp.LockType = adLockBatchOptimistic
RsLpp.CursorType = adOpenDynamic
RsLpp.Open "select * from Stock_obat", con, , , adCmdText
RsLpp.Filter = "Nama_obat= '" & Combo1 & "'"
If Not RsLpp.EOF Then
Text6.Text = RsLpp.Fields(1)
Text7.Text = RsLpp.Fields(2)
Text8.Text = RsLpp.Fields(3)
End If

End Sub

Private Sub Command3_Click()
If Text1.Text = "" Or Text2.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text8.Text = "" Or Text9.Text = "" Or Combo1.Text = "" Then
MsgBox "Semua Data Harus Diisi1", vbCritical, "Perhatian"

Else

If Val(Text9.Text) > Val(Text6.Text) Then
MsgBox ("Stock Tidak Memenuhi")
Text9.Text = ""
Else
Text6.Text = Val(Text6.Text) - Val(Text9.Text)
con.Execute "update Stock_obat set Stock= '" & Text6.Text & "' where Nama_obat = '" & Combo1.Text & "'"
MsgBox "Berhasil", vbOKOnly, "Pemberitahuan"
Label12.Caption = Left(Text2.Text, 1) + Mid(Text1.Text, 6, 3) + Left(Text8.Text, 3)
Frame3.Visible = True
Form7.Adodc1.Recordset.AddNew
Form7.Adodc1.Recordset!Kode_Pasien = Text1.Text
Form7.Adodc1.Recordset!Nama = Text2.Text
If Option1.Value = True Then
Form7.Adodc1.Recordset!Jenis_Kelamin = Option1.Caption
Else
Form7.Adodc1.Recordset!Jenis_Kelamin = Option2.Caption
End If
Form7.Adodc1.Recordset!Usia = Text4.Text
Form7.Adodc1.Recordset!Alamat = Text5.Text
Form7.Adodc1.Recordset!Jenis_Obat = Combo1.Text
Form7.Adodc1.Recordset!Jumlah = Text9.Text
Form7.Adodc1.Recordset!Biaya = Text3.Text
Form7.Adodc1.Recordset!Kode_Pembayaran = Label12.Caption
Form7.Adodc1.Recordset.Update
Form7.DataGrid1.Refresh
Command4.Visible = True
Label11.Visible = True
Call bersih
End If

End If

End Sub

Private Sub Command4_Click()
Form11.Label2.Caption = Form5.Label12.Caption
Form5.Hide
Form11.Show

End Sub

Private Sub Form_activate()
AutoNumber

End Sub

Private Sub Form_Load()
Connects
pilih
bersih
End Sub

Private Sub AutoNumber()
Connects
RsLpp.Open ("SELECT * FROM Obatan WHERE Kode_pasien in(select max(Kode_pasien) from Obatan)order by Kode_pasien desc"), con
RsLpp.Requery
    Dim Urut As String * 6
    Dim Hitung As Long
    With RsLpp
        If .EOF Then
            Urut = "00001"
            Text1 = "KJO + Urut"
        Else
            Hitung = Right(!Kode_Pasien, 6) + 1
            Urut = Right("0000" & Hitung, 6)
        End If
        Text1 = "KJO" + Urut
    End With
End Sub

Private Sub Label21_Click()

End Sub

Private Sub Option1_Click()
If Option1 = True Then
Option1 = True
End If
End Sub

Private Sub Option2_Click()
If Option2 = True Then
Option2 = True
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then
  KeyAscii = 8
ElseIf KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
End If
End Sub

Private Sub Text9_Change()
Text3.Text = Val(Text8.Text) * Val(Text9.Text)
End Sub
