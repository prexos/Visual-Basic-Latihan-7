VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6480
      TabIndex        =   7
      Top             =   5520
      Width           =   4095
   End
   Begin VB.TextBox Text6 
      Height          =   360
      Left            =   6480
      TabIndex        =   6
      Top             =   4800
      Width           =   2775
   End
   Begin VB.TextBox Text5 
      Height          =   360
      Left            =   6480
      TabIndex        =   5
      Top             =   4200
      Width           =   2775
   End
   Begin VB.TextBox Text4 
      Height          =   360
      Left            =   6480
      TabIndex        =   4
      Top             =   3480
      Width           =   2775
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Left            =   6480
      TabIndex        =   3
      Top             =   2880
      Width           =   2775
   End
   Begin VB.TextBox Text3 
      Height          =   360
      Left            =   6480
      TabIndex        =   2
      Top             =   2280
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Height          =   735
      Left            =   6480
      TabIndex        =   1
      Top             =   1200
      Width           =   4095
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Left            =   6480
      TabIndex        =   0
      Top             =   600
      Width           =   2895
   End
   Begin VB.Label Label8 
      Caption         =   "Nomor Sertifikat"
      Height          =   255
      Left            =   4800
      TabIndex        =   15
      Top             =   5520
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "3 Huruf Pertama Program Dipilih"
      Height          =   495
      Left            =   4800
      TabIndex        =   14
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "3 Huruf Terakhir Nomor Induk"
      Height          =   495
      Left            =   4800
      TabIndex        =   13
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Program Dipilih"
      Height          =   255
      Left            =   4800
      TabIndex        =   12
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Pilihan Program"
      Height          =   255
      Left            =   4800
      TabIndex        =   11
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Tahun Masuk"
      Height          =   255
      Left            =   4800
      TabIndex        =   10
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Nama Siswa"
      Height          =   615
      Left            =   4800
      TabIndex        =   9
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Nomor Induk"
      Height          =   255
      Left            =   4800
      TabIndex        =   8
      Top             =   600
      Width           =   1455
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
Text4.Text = Combo1.Text
Text4.Enabled = False

Text6.Text = Strings.Left(Text4.Text, 3)
Text7.Text = Text5.Text + "/" + UCase(Text6.Text) + "/" + "POL" + "/" + Text3.Text
End Sub

Private Sub Form_Activate()
 Combo1.AddItem "Akuntansi"
 Combo1.AddItem "Informatika"
 Combo1.AddItem "Bisnis"
 Combo1.AddItem "Antropologi"
 Combo1.AddItem "Dokter"
End Sub

Private Sub Text1_LostFocus()
Text5.Text = Strings.Right(Text1.Text, 3)
End Sub
