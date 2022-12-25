VERSION 5.00
Begin VB.Form Form3 
   ClientHeight    =   11280
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   20250
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   ScaleHeight     =   11280
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      Caption         =   "Kembali"
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
      Left            =   16680
      TabIndex        =   3
      Top             =   10080
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
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
      Left            =   18480
      TabIndex        =   2
      Top             =   10080
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "MASUK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9000
      TabIndex        =   1
      Top             =   7200
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "MASUK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9000
      TabIndex        =   0
      Top             =   5160
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      Height          =   11760
      Left            =   0
      Picture         =   "Form3.frx":0000
      ScaleHeight     =   11700
      ScaleWidth      =   20670
      TabIndex        =   4
      Top             =   0
      Width           =   20730
      Begin VB.Timer Timer2 
         Interval        =   1
         Left            =   12600
         Top             =   3840
      End
      Begin VB.Timer Timer1 
         Interval        =   150
         Left            =   12600
         Top             =   3360
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "Digital-7"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   375
         Left            =   8280
         TabIndex        =   8
         Top             =   3840
         Width           =   3735
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "SELAMAT DATANG!   "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   735
         Left            =   6960
         TabIndex        =   7
         Top             =   3360
         Width           =   6255
      End
      Begin VB.Shape Shape2 
         Height          =   1935
         Left            =   6960
         Top             =   6360
         Width           =   6135
      End
      Begin VB.Shape Shape1 
         Height          =   1935
         Left            =   6960
         Top             =   4320
         Width           =   6135
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "BELI OBAT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   8160
         TabIndex        =   6
         Top             =   6480
         Width           =   3615
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "PERIKSA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   7920
         TabIndex        =   5
         Top             =   4440
         Width           =   4095
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form3.Hide
Form4.Show
End Sub

Private Sub Command2_Click()
Form3.Hide
Form5.Show
End Sub

Private Sub Command3_Click()
pesan = MsgBox("Kembali ke Masuk Akun?", vbQuestion + vbYesNo, "Pertanyaan")
If pesan = vbYes Then
Form3.Hide
Form1.Show
Else
Form3.SetFocus
End If
End Sub

Private Sub Command4_Click()
Form3.Hide
Form2.Show
End Sub

Private Sub Timer1_Timer()
Label3.Caption = Right(Label3.Caption, Len(Label3.Caption) - 1) & Left(Label3.Caption, 1)
End Sub

Private Sub Timer2_Timer()
Label4.Caption = Time
End Sub
