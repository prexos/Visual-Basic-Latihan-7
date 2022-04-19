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
   Begin VB.TextBox Text5 
      Height          =   360
      Left            =   7080
      TabIndex        =   4
      Top             =   3360
      Width           =   2775
   End
   Begin VB.TextBox Text4 
      Height          =   360
      Left            =   7080
      TabIndex        =   3
      Top             =   2760
      Width           =   2775
   End
   Begin VB.TextBox Text3 
      Height          =   360
      Left            =   7080
      TabIndex        =   2
      Top             =   2160
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Height          =   360
      Left            =   7080
      TabIndex        =   1
      Top             =   1560
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Left            =   7080
      TabIndex        =   0
      Top             =   960
      Width           =   2775
   End
   Begin VB.Label Label5 
      Caption         =   "Jam dan Tanggal"
      Height          =   255
      Left            =   5160
      TabIndex        =   9
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Detik Saja"
      Height          =   255
      Left            =   5160
      TabIndex        =   8
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Menit Saja"
      Height          =   255
      Left            =   5160
      TabIndex        =   7
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Jam Saja"
      Height          =   255
      Left            =   5160
      TabIndex        =   6
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Jam Sekarang"
      Height          =   255
      Left            =   5160
      TabIndex        =   5
      Top             =   960
      Width           =   1695
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text1_LostFocus()
Text1.Text = Format(Now, "hh:mm:dd")
Text2.SetFocus
End Sub

Private Sub Text2_LostFocus()
Text2.Text = Format(Now, "hh")
Text3.SetFocus
End Sub

Private Sub Text3_LostFocus()
Text3.Text = Format(Now, "mm")
Text4.SetFocus
End Sub

Private Sub Text4_LostFocus()
Text4.Text = Format(Now, "ss")
Text5.SetFocus
End Sub

Private Sub Text5_LostFocus()
Text5.Text = Format(Now, "hh:mm:ss DDDD, MMMM YYYY")
Text1.SetFocus
End Sub
