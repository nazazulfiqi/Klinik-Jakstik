VERSION 5.00
Begin VB.Form Form2 
   ClientHeight    =   11280
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20250
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   11280
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      Caption         =   "Beli Obat"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10320
      TabIndex        =   3
      Top             =   4560
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Masuk"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9120
      TabIndex        =   2
      Top             =   6720
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Periksa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8160
      TabIndex        =   1
      Top             =   4560
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
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
      Left            =   18120
      TabIndex        =   0
      Top             =   10320
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   11760
      Left            =   -480
      Picture         =   "Form2.frx":0000
      ScaleHeight     =   11700
      ScaleWidth      =   20670
      TabIndex        =   4
      Top             =   0
      Width           =   20730
      Begin VB.Shape Shape2 
         BorderColor     =   &H00000000&
         FillColor       =   &H00FFFFFF&
         Height          =   2055
         Left            =   4800
         Top             =   5880
         Width           =   11535
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000000&
         Height          =   2295
         Index           =   0
         Left            =   4800
         Top             =   3480
         Width           =   11535
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "MENU PELAYANAN DAN PENJUALAN"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   855
         Left            =   5880
         TabIndex        =   6
         Top             =   6000
         Width           =   9495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "DATABASE PELAYANAN DAN PENJUALAN"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   5160
         TabIndex        =   5
         Top             =   3720
         Width           =   10695
      End
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Height          =   2295
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   11535
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
pesan = MsgBox("Kembali ke Masuk Akun?", vbQuestion + vbYesNo, "Pertanyaan")
If pesan = vbYes Then
Form2.Hide
Form1.Show
Else
Form2.SetFocus
End If
End Sub

Private Sub Command2_Click()
Form2.Hide
Form6.Show
End Sub

Private Sub Command3_Click()
Form2.Hide
Form3.Show

Form3.Command4.Visible = True
End Sub

Private Sub Command4_Click()
Form2.Hide
Form7.Show
End Sub

