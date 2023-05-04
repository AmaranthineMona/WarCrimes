VERSION 5.00
Begin VB.Form frmdif 
   Caption         =   "Difficulty"
   ClientHeight    =   8730
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9615
   LinkTopic       =   "Form1"
   ScaleHeight     =   8730
   ScaleWidth      =   9615
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdhard 
      Caption         =   "Hard"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   6480
      TabIndex        =   2
      Top             =   7440
      Width           =   2895
   End
   Begin VB.CommandButton cmdeasy 
      Caption         =   "Easy"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   1
      Top             =   7440
      Width           =   2895
   End
   Begin VB.CommandButton cmdmed 
      Caption         =   "Medium"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3360
      TabIndex        =   0
      Top             =   7440
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Choose your difficulty"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   1080
      TabIndex        =   3
      Top             =   0
      Width           =   7605
   End
   Begin VB.Image Image1 
      Height          =   8775
      Left            =   0
      Picture         =   "frmdif.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9615
   End
End
Attribute VB_Name = "frmdif"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdeasy_Click()
frmhang.Show
frmdif.Hide
mil = 180
jkl = 1
End Sub

Private Sub cmdhard_Click()
frmhang.Show
frmdif.Hide
mil = 60
jkl = 3
End Sub

Private Sub cmdmed_Click()
frmhang.Show
frmdif.Hide
mil = 90
jkl = 2
End Sub
