VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
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
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text6 
      Height          =   360
      Left            =   6840
      TabIndex        =   10
      Text            =   "Text6"
      Top             =   3960
      Width           =   2655
   End
   Begin VB.TextBox Text5 
      Height          =   360
      Left            =   6840
      TabIndex        =   4
      Text            =   "Text5"
      Top             =   3360
      Width           =   2655
   End
   Begin VB.TextBox Text4 
      Height          =   360
      Left            =   6840
      TabIndex        =   3
      Text            =   "Text4"
      Top             =   2760
      Width           =   2655
   End
   Begin VB.TextBox Text3 
      Height          =   360
      Left            =   6840
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   2160
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      Height          =   360
      Left            =   6840
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   1560
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Left            =   6840
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label Label6 
      Caption         =   "Nama Hari"
      Height          =   255
      Left            =   4920
      TabIndex        =   11
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "Tanggal Gabungan"
      Height          =   255
      Left            =   4920
      TabIndex        =   9
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "Tahun Saja"
      Height          =   255
      Left            =   4920
      TabIndex        =   8
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Bulan Saja"
      Height          =   255
      Left            =   4920
      TabIndex        =   7
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Tanggal Saja"
      Height          =   255
      Left            =   4920
      TabIndex        =   6
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Tanggal Sekarang"
      Height          =   255
      Left            =   4920
      TabIndex        =   5
      Top             =   960
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
 Text1.Text = Format(Now, "dd-MM-yyyy")
 Text2.Text = Format(Now, "dd")
 Text3.Text = Format(Now, "MM")
 Text4.Text = Format(Now, "yyyy")
 Text5.Text = Text2 + Text3 + Text4
 Text6.Text = Format(Now, "dddd")
End Sub
