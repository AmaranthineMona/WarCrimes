VERSION 5.00
Begin VB.Form frmmain 
   Caption         =   "Menu"
   ClientHeight    =   10500
   ClientLeft      =   -11505
   ClientTop       =   -945
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   Picture         =   "Main Menu.frx":0000
   ScaleHeight     =   10500
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   8280
      Top             =   1800
   End
   Begin VB.CommandButton cmdend 
      BackColor       =   &H00C00000&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4320
      Width           =   2535
   End
   Begin VB.CommandButton cmdhigh 
      BackColor       =   &H0080C0FF&
      Caption         =   "How to play"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3000
      Width           =   2535
   End
   Begin VB.CommandButton cmdplay 
      BackColor       =   &H000000FF&
      Caption         =   "Play"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Label lbltitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Hangman - Super Mario version"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   33.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1215
      Left            =   3960
      TabIndex        =   3
      Top             =   120
      Width           =   10575
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdend_Click()
End
End Sub

Private Sub cmdhigh_Click()
frmopt.Show
End Sub

Private Sub cmdplay_Click()
frmdif.Show
frmmain.Hide
End Sub


Private Sub Timer1_Timer()
Static x As Integer
x = x + 1
If x = 1 Then
    lbltitle.ForeColor = &HC0&
End If
If x = 2 Then
    lbltitle.ForeColor = &H80C0FF
End If
If x = 3 Then
    lbltitle.ForeColor = &HC00000
End If
If x = 4 Then
    x = 0
End If
End Sub
