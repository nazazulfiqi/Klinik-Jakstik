VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   ClientHeight    =   11280
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   20250
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   11280
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.CheckBox Check1 
      Caption         =   "Perlihatkan Kata Sandi"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   5
      Top             =   6240
      Width           =   3015
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Daftar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      TabIndex        =   4
      Top             =   8040
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   4440
      Top             =   3240
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=dbLogin.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=dbLogin.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "dbLogin"
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
   Begin VB.CommandButton Command2 
      Caption         =   "Keluar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5760
      TabIndex        =   3
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Masuk"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4080
      TabIndex        =   2
      Top             =   6840
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   4080
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   5520
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      TabIndex        =   0
      Top             =   4560
      Width           =   3015
   End
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   11760
      Left            =   0
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   11700
      ScaleWidth      =   20670
      TabIndex        =   6
      Top             =   0
      Width           =   20730
      Begin VB.Timer Timer2 
         Interval        =   1
         Left            =   18960
         Top             =   120
      End
      Begin VB.Label Label5 
         Caption         =   "Tidak Memiliki Akun? Silahkan Daftar."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4080
         TabIndex        =   10
         Top             =   7680
         Width           =   2655
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "Digital-7"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   615
         Left            =   15120
         TabIndex        =   9
         Top             =   240
         Width           =   3615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Kata Sandi"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2040
         TabIndex        =   8
         Top             =   5640
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Pengguna"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1200
         TabIndex        =   7
         Top             =   4560
         Width           =   2535
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim query As String

Private Sub Check1_Click()
If Check1.Value = 1 Then
Text2.PasswordChar = ""
Else
Text2.PasswordChar = "*"
End If
End Sub

Private Sub Command1_Click()
Call connectdb
If Text1.Text = "" Then
MsgBox "Mohon Isi Nama Pengguna!", vbCritical, "Perhatian"
Text1.SetFocus
ElseIf Text2.Text = "" Then
MsgBox "Mohon Isi Kata Sandi!", vbCritical, "Perhatian"
Text2.SetFocus
End If
query = "select * from dblogin where nama_pengguna='" & Text1.Text & "' and kata_sandi='" & Text2.Text & "'"
RsAdmin.Open (query), Connect
If Not RsAdmin.EOF Then
    If RsAdmin!nama_pengguna = "admin" Then
    Unload Me
    Form2.Show
    MsgBox "Berhasil Masuk!", vbInformation, "Berhasil"
    Else
    Unload Me
    Form3.Show
    MsgBox "Berhasil Masuk!", vbInformation, "Berhasil"
    End If
    Else
    MsgBox "Nama Pengguna atau Kata Sandi Salah!", vbExclamation, "Gagal"
    Text1.Text = ""
    Text2.Text = ""
    Text1.SetFocus
    End If
End Sub

Private Sub Command2_Click()
pesan = MsgBox("Keluar Program", vbQuestion + vbYesNo, "Pertanyaan")
If pesan = vbYes Then
End
Else
Form1.SetFocus
End If
End Sub

Private Sub Command3_Click()
Form1.Hide
Form9.Show
End Sub

Private Sub Timer1_Timer()
Label3.Caption = Right(Label3.Caption, Len(Label3.Caption) - 1) & Left(Label3.Caption, 1)
End Sub

Private Sub Timer2_Timer()
Label4.Caption = Time
End Sub
