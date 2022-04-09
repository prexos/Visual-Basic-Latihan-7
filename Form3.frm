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
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtpropil 
      Height          =   375
      Left            =   6960
      TabIndex        =   14
      Top             =   3000
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generate!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6360
      TabIndex        =   13
      Top             =   6840
      Width           =   2295
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Left            =   6960
      TabIndex        =   11
      Text            =   "Combo1"
      Top             =   2400
      Width           =   2655
   End
   Begin VB.TextBox Text3 
      Height          =   600
      Left            =   6960
      TabIndex        =   10
      Top             =   3720
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Left            =   6960
      TabIndex        =   3
      Top             =   1200
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      Height          =   360
      Left            =   6960
      TabIndex        =   2
      Top             =   1800
      Width           =   2655
   End
   Begin VB.TextBox Text4 
      Height          =   600
      Left            =   6960
      TabIndex        =   1
      Top             =   4560
      Width           =   2655
   End
   Begin VB.TextBox Text5 
      Height          =   360
      Left            =   6960
      TabIndex        =   0
      Top             =   5400
      Width           =   2655
   End
   Begin VB.Label Label7 
      Caption         =   "Program Pilihan"
      Height          =   375
      Left            =   5040
      TabIndex        =   12
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Nomor Induk"
      Height          =   255
      Left            =   5040
      TabIndex        =   9
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Nama Siswa"
      Height          =   255
      Left            =   5040
      TabIndex        =   8
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Pilihan Program"
      Height          =   255
      Left            =   5040
      TabIndex        =   7
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "3 Huruf Pertama Nomor Induk"
      Height          =   495
      Left            =   5040
      TabIndex        =   6
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "3 Huruf Pertama Program Dipilih"
      Height          =   495
      Left            =   5040
      TabIndex        =   5
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "Nomor Sertifikat"
      Height          =   255
      Left            =   5040
      TabIndex        =   4
      Top             =   5400
      Width           =   1815
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
txtpropil.Text = Combo1.Text
txtpropil.Enabled = False
End Sub

Private Sub Command1_Click()
Text3.Text = Strings.Left(Text1.Text, 3)
Text4.Text = Strings.Left(txtpropil.Text, 4)
Text5.Text = Text3.Text + "/" + UCase(Text4.Text) + "/" + "POL" + "/" + Format(Now, "yyyy")
End Sub

Private Sub Form_Activate()
 Combo1.AddItem "Akuntansi"
 Combo1.AddItem "Administrasi Bisnis"
 Combo1.AddItem "Informatika"
End Sub
