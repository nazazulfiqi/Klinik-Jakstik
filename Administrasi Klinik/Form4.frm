VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form4 
   ClientHeight    =   11280
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   20250
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   ScaleHeight     =   11280
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   15000
      Top             =   1680
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1085
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
      RecordSource    =   "Periksa"
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
      Left            =   14280
      TabIndex        =   17
      Top             =   4560
      Visible         =   0   'False
      Width           =   4455
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
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   4215
      End
   End
   Begin VB.CommandButton Command2 
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
      Left            =   11520
      TabIndex        =   16
      Top             =   8520
      Width           =   1335
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
      Left            =   9840
      TabIndex        =   15
      Top             =   8520
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Periksa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   7800
      TabIndex        =   10
      Top             =   2400
      Width           =   5055
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   495
         Left            =   600
         Top             =   4560
         Visible         =   0   'False
         Width           =   3735
         _ExtentX        =   6588
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
         RecordSource    =   "dokter"
         Caption         =   "Adodc2"
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
      Begin VB.TextBox Text6 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """Rp""#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
         DataSource      =   "Adodc2"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   14
         Top             =   2880
         Width           =   3975
      End
      Begin VB.ComboBox Combo1 
         DataField       =   "spesialis"
         DataSource      =   "Adodc2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "Form4.frx":0000
         Left            =   360
         List            =   "Form4.frx":0002
         TabIndex        =   12
         Top             =   1440
         Width           =   4215
      End
      Begin VB.Label Label9 
         Caption         =   "Harga"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   13
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "Keluhan yang ingin diperiksa"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   11
         Top             =   720
         Width           =   2895
      End
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
      Height          =   5655
      Left            =   2040
      TabIndex        =   0
      Top             =   2400
      Width           =   5415
      Begin VB.OptionButton Option2 
         Caption         =   "Perempuan"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   19
         Top             =   2400
         Width           =   2775
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Laki-laki"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         TabIndex        =   18
         Top             =   1920
         Width           =   2415
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   2040
         TabIndex        =   9
         Top             =   3960
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
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         TabIndex        =   8
         Top             =   2880
         Width           =   3135
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         TabIndex        =   7
         Top             =   1200
         Width           =   3135
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         TabIndex        =   6
         Top             =   480
         Width           =   3135
      End
      Begin VB.Label Label7 
         Caption         =   "Kode Pasien"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   5
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Alamat"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   4
         Top             =   4200
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Usia"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Jenis Kelamin"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Nama"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   1320
         Width           =   975
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   11760
      Left            =   0
      Picture         =   "Form4.frx":0004
      ScaleHeight     =   11700
      ScaleWidth      =   22950
      TabIndex        =   20
      Top             =   0
      Width           =   23010
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
         Left            =   14280
         TabIndex        =   21
         Top             =   6000
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
         Left            =   15840
         TabIndex        =   22
         Top             =   6000
         Visible         =   0   'False
         Width           =   2415
      End
   End
End
Attribute VB_Name = "Form4"
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
Private Sub Combo1_Click()
Set RsLpp = New ADODB.Recordset
RsLpp.LockType = adLockBatchOptimistic
RsLpp.CursorType = adOpenDynamic
RsLpp.Open "select * from dokter", con, , , adCmdText
RsLpp.Filter = "spesialis= '" & Combo1 & "'"
If Not RsLpp.EOF Then
Text6.Text = RsLpp.Fields(1)
End If

End Sub

Sub pilih()
Set RsLpp = New ADODB.Recordset
RsLpp.Open "select * from dokter order by spesialis ASC", _
con, adOpenDynamic, _
adLockBatchOptimistic

Do While Not RsLpp.BOF And Not RsLpp.EOF
If Not RsLpp.EOF Then
Combo1.AddItem RsLpp.Fields(0)
RsLpp.MoveNext
End If
Loop

End Sub

Sub bersih()
Text1.Text = ""
Text2.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Combo1.Text = ""
Option1.Value = False
Option2.Value = False

End Sub


Private Sub Command1_Click()
pesan = MsgBox("Keluar dari Pelayanan Periksa?", vbQuestion + vbYesNo, "Pertanyaan")
If pesan = vbYes Then
Form4.Hide
Form3.Show
Else
Form4.SetFocus
End If
End Sub

Private Sub Command2_Click()
If Text1.Text = "" Or Text2.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Combo1.Text = "" Then
MsgBox "Semua Data Harus Diisi1", vbCritical, "Perhatian"

Else

MsgBox "Berhasil", vbOKOnly, "Pemberitahuan"
Label12.Caption = Left(Text2.Text, 1) + Mid(Text1.Text, 6, 3) + Mid(Text6.Text, 3, 2)
Frame3.Visible = True

Form6.Adodc1.Recordset.AddNew
Form6.Adodc1.Recordset!Kode_Pasien = Text1.Text
Form6.Adodc1.Recordset!Nama = Text2.Text
If Option1.Value = True Then
    Form6.Adodc1.Recordset!Jenis_Kelamin = Option1.Caption
    Else
    Form6.Adodc1.Recordset!Jenis_Kelamin = Option2.Caption
    End If
Form6.Adodc1.Recordset!Usia = Text4.Text
Form6.Adodc1.Recordset!Alamat = Text5.Text
Form6.Adodc1.Recordset!Keluhan = Combo1.Text
Form6.Adodc1.Recordset!Biaya = Text6.Text
Form6.Adodc1.Recordset!Kode_Pembayaran = Label12.Caption
Form6.Adodc1.Recordset.Update
Form6.DataGrid1.Refresh
Command4.Visible = True
Label11.Visible = True
Call bersih

End If

End Sub

Private Sub Command3_Click()
pesan = MsgBox("Melakukan Transaksi Lagi?", vbQuestion + vbYesNo, "Pertanyaan")
If pesan = vbYes Then
Call bersih
Label10.Caption = ""
Label11.Caption = ""
Label12.Caption = ""
Label14.Caption = ""
Frame3.Visible = False
Command3.Visible = False
Else
Call bersih
Label10.Caption = ""
Label11.Caption = ""
Label12.Caption = ""
Label14.Caption = ""
Frame3.Visible = False
Command3.Visible = False
Form4.Hide
Form3.Show
End If

End Sub

Private Sub Command4_Click()
Form12.Label2.Caption = Form4.Label12.Caption
Form4.Hide
Form12.Show

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
RsLpp.Open ("SELECT * FROM Periksa WHERE Kode_pasien in(select max(Kode_pasien) from Periksa)order by Kode_pasien desc"), con
RsLpp.Requery
    Dim Urut As String * 6
    Dim Hitung As Long
    With RsLpp
        If .EOF Then
            Urut = "00001"
            Text1 = "KJP" + Urut
        Else
            Hitung = Right(!Kode_Pasien, 6) + 1
            Urut = Right("0000" & Hitung, 6)
        End If
        Text1 = "KJP" + Urut
    End With
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
