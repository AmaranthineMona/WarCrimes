VERSION 5.00
Begin VB.Form frmwin 
   BackColor       =   &H00FFFFFF&
   Caption         =   "EP"
   ClientHeight    =   9630
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10470
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   9630
   ScaleWidth      =   10470
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2040
      Top             =   2400
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Back to Menu"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3360
      TabIndex        =   0
      Top             =   1320
      Width           =   3495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   855
      Left            =   6600
      TabIndex        =   2
      Top             =   3720
      Width           =   3495
   End
   Begin VB.Image imgjump 
      Height          =   1530
      Left            =   4560
      Picture         =   "Win.frx":0000
      Top             =   2760
      Width           =   1530
   End
   Begin VB.Image Image1 
      Height          =   3600
      Left            =   2040
      Picture         =   "Win.frx":0C94
      Top             =   6360
      Width           =   6450
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "You win"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   4200
      TabIndex        =   1
      Top             =   0
      Width           =   1725
   End
End
Attribute VB_Name = "frmwin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmwin.Hide
frmmain.Show
End Sub

Private Sub Form_Load()
x = 0

End Sub

Private Sub Timer1_Timer()
Static x As Integer
x = x + 1
If x = 1 Then
    imgjump.Top = 4800
End If
If x = 3 Then
    imgjump.Top = 3360
End If
If x = 5 Then
    imgjump.Top = 2760
End If
If x = 7 Then
    imgjump.Top = 3360
End If
If x = 9 Then
    x = 0
End If
End Sub
