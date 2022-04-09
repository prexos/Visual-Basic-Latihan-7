VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
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
   LinkTopic       =   "Form2"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   360
      Left            =   7560
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1800
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      Height          =   360
      Left            =   7560
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   2400
      Width           =   2655
   End
   Begin VB.TextBox Text3 
      Height          =   360
      Left            =   7560
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   3000
      Width           =   2655
   End
   Begin VB.TextBox Text4 
      Height          =   360
      Left            =   7560
      TabIndex        =   1
      Text            =   "Text4"
      Top             =   3600
      Width           =   2655
   End
   Begin VB.TextBox Text5 
      Height          =   360
      Left            =   7560
      TabIndex        =   0
      Text            =   "Text5"
      Top             =   4200
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Jam Sekarang"
      Height          =   255
      Left            =   5640
      TabIndex        =   9
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Jam Saja"
      Height          =   255
      Left            =   5640
      TabIndex        =   8
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Menit Saja"
      Height          =   255
      Left            =   5640
      TabIndex        =   7
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "Detik Saja"
      Height          =   255
      Left            =   5640
      TabIndex        =   6
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "Jam dan Tanggal"
      Height          =   255
      Left            =   5640
      TabIndex        =   5
      Top             =   4200
      Width           =   1695
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Text1.Text = Format(Now, "hh:mm:ss")
Text2.Text = Format(Now, "hh")
Text3.Text = Format(Now, "mm")
Text4.Text = Format(Now, "ss")
Text5.Text = Format(Now)
End Sub
