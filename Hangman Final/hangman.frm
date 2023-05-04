VERSION 5.00
Begin VB.Form frmhang 
   Caption         =   "Hang man"
   ClientHeight    =   12315
   ClientLeft      =   660
   ClientTop       =   1740
   ClientWidth     =   15285
   FillStyle       =   2  'Horizontal Line
   LinkTopic       =   "Form1"
   Picture         =   "hangman.frx":0000
   ScaleHeight     =   12315
   ScaleWidth      =   15285
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   75
      Left            =   13080
      Top             =   2400
   End
   Begin VB.Timer timershell2 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1920
      Top             =   6120
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1440
      Top             =   6120
   End
   Begin VB.Timer lblmi 
      Interval        =   1000
      Left            =   12480
      Top             =   10920
   End
   Begin VB.Timer timerjump 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   11280
      Top             =   10920
   End
   Begin VB.Timer timershell 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2400
      Top             =   6120
   End
   Begin VB.Image imgi3 
      Height          =   1800
      Left            =   12000
      Picture         =   "hangman.frx":33B042
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   1080
   End
   Begin VB.Image imgy2 
      Height          =   1695
      Left            =   13560
      Picture         =   "hangman.frx":34091E
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   1410
   End
   Begin VB.Image Image3 
      Height          =   1575
      Left            =   13200
      Picture         =   "hangman.frx":34142B
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   780
   End
   Begin VB.Label lbld8 
      BackColor       =   &H00000000&
      Height          =   255
      Left            =   11040
      TabIndex        =   45
      Top             =   4080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lbld7 
      BackColor       =   &H00000000&
      Height          =   255
      Left            =   10200
      TabIndex        =   44
      Top             =   4080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lbld6 
      BackColor       =   &H00000000&
      Height          =   255
      Left            =   9360
      TabIndex        =   43
      Top             =   4080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lbld5 
      BackColor       =   &H00000000&
      Height          =   255
      Left            =   8520
      TabIndex        =   42
      Top             =   4080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lbld4 
      BackColor       =   &H00000000&
      Height          =   255
      Left            =   7680
      TabIndex        =   41
      Top             =   4080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lbld3 
      BackColor       =   &H00000000&
      Height          =   255
      Left            =   6840
      TabIndex        =   40
      Top             =   4080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lbld2 
      BackColor       =   &H00000000&
      Height          =   255
      Left            =   6000
      TabIndex        =   39
      Top             =   4080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lbld1 
      BackColor       =   &H00000000&
      Height          =   255
      Left            =   5160
      TabIndex        =   38
      Top             =   4080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lbl5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8400
      TabIndex        =   37
      Top             =   3240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lbl6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9240
      TabIndex        =   36
      Top             =   3240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lbl7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10080
      TabIndex        =   35
      Top             =   3240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lbl8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10920
      TabIndex        =   34
      Top             =   3240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lbl4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7560
      TabIndex        =   33
      Top             =   3240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lbl3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6720
      TabIndex        =   32
      Top             =   3240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lbl2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5880
      TabIndex        =   31
      Top             =   3240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5040
      TabIndex        =   30
      Top             =   3240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblnum 
      Caption         =   "Label2"
      Height          =   375
      Left            =   6000
      TabIndex        =   29
      Top             =   240
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Label lblword 
      Caption         =   "Label2"
      Height          =   615
      Left            =   7080
      TabIndex        =   28
      Top             =   600
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Image imgy1 
      Height          =   1170
      Left            =   13920
      Picture         =   "hangman.frx":3426C0
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   1050
   End
   Begin VB.Image imgl1 
      Height          =   1830
      Left            =   12120
      Picture         =   "hangman.frx":34283F
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   840
   End
   Begin VB.Image img5 
      Height          =   540
      Left            =   2280
      Picture         =   "hangman.frx":343C29
      Top             =   5160
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image img9 
      Height          =   540
      Left            =   4200
      Picture         =   "hangman.frx":343E67
      Top             =   5160
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image img7 
      Height          =   540
      Left            =   3240
      Picture         =   "hangman.frx":3440A5
      Top             =   5160
      Width           =   540
   End
   Begin VB.Image img3 
      Height          =   540
      Left            =   1320
      Picture         =   "hangman.frx":3442E3
      Top             =   5160
      Width           =   540
   End
   Begin VB.Image img1 
      Height          =   540
      Left            =   240
      Picture         =   "hangman.frx":344521
      Top             =   5160
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image img6 
      Height          =   540
      Left            =   2280
      Picture         =   "hangman.frx":34475F
      Top             =   5280
      Width           =   540
   End
   Begin VB.Image img10 
      Height          =   540
      Left            =   4200
      Picture         =   "hangman.frx":34499D
      Top             =   5280
      Width           =   540
   End
   Begin VB.Image img8 
      Height          =   540
      Left            =   3240
      Picture         =   "hangman.frx":344BDB
      Top             =   5280
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image img4 
      Height          =   540
      Left            =   1320
      Picture         =   "hangman.frx":344E19
      Top             =   5280
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image img2 
      Height          =   540
      Left            =   240
      Picture         =   "hangman.frx":345057
      Top             =   5280
      Width           =   540
   End
   Begin VB.Label lblmil 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   555
      Left            =   7935
      TabIndex        =   27
      Top             =   5760
      Width           =   165
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Timer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   555
      Left            =   7290
      TabIndex        =   26
      Top             =   9240
      Width           =   1365
   End
   Begin VB.Image Image1 
      Height          =   2775
      Left            =   7320
      Picture         =   "hangman.frx":345295
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   1335
   End
   Begin VB.Image imgshell 
      Appearance      =   0  'Flat
      Height          =   720
      Left            =   2760
      Picture         =   "hangman.frx":34545E
      Stretch         =   -1  'True
      Top             =   10560
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Image Image4 
      Height          =   3255
      Left            =   360
      Picture         =   "hangman.frx":345908
      Stretch         =   -1  'True
      Top             =   8040
      Width           =   2655
   End
   Begin VB.Label lblt 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "T"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   9240
      TabIndex        =   25
      Top             =   11640
      Width           =   450
   End
   Begin VB.Label lblq 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Q"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   7800
      TabIndex        =   24
      Top             =   11640
      Width           =   450
   End
   Begin VB.Label lblw 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   10680
      TabIndex        =   23
      Top             =   11640
      Width           =   570
   End
   Begin VB.Label lblp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "P"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   7320
      TabIndex        =   22
      Top             =   11640
      Width           =   450
   End
   Begin VB.Label lblr 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   8280
      TabIndex        =   21
      Top             =   11640
      Width           =   450
   End
   Begin VB.Label lbly 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Y"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   11880
      TabIndex        =   20
      Top             =   11640
      Width           =   450
   End
   Begin VB.Label lblo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   6840
      TabIndex        =   19
      Top             =   11640
      Width           =   450
   End
   Begin VB.Label lblz 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Z"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   12360
      TabIndex        =   18
      Top             =   11640
      Width           =   450
   End
   Begin VB.Label lbls 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   8760
      TabIndex        =   17
      Top             =   11640
      Width           =   450
   End
   Begin VB.Label lblx 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   11280
      TabIndex        =   16
      Top             =   11640
      Width           =   450
   End
   Begin VB.Label lblu 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "U"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   9720
      TabIndex        =   15
      Top             =   11640
      Width           =   450
   End
   Begin VB.Label lblv 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "V"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   10200
      TabIndex        =   14
      Top             =   11640
      Width           =   450
   End
   Begin VB.Label lbln 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "N"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   6360
      TabIndex        =   13
      Top             =   11640
      Width           =   450
   End
   Begin VB.Label lblg 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "G"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   3120
      TabIndex        =   12
      Top             =   11640
      Width           =   450
   End
   Begin VB.Label lbld 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   1560
      TabIndex        =   11
      Top             =   11640
      Width           =   450
   End
   Begin VB.Label lblj 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "J"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   4320
      TabIndex        =   10
      Top             =   11640
      Width           =   450
   End
   Begin VB.Label lblc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   1080
      TabIndex        =   9
      Top             =   11640
      Width           =   450
   End
   Begin VB.Label lble 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   2040
      TabIndex        =   8
      Top             =   11640
      Width           =   510
   End
   Begin VB.Label lbll 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   5280
      TabIndex        =   7
      Top             =   11640
      Width           =   450
   End
   Begin VB.Label lblb 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   480
      TabIndex        =   6
      Top             =   11640
      Width           =   450
   End
   Begin VB.Label lblm 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   5760
      TabIndex        =   5
      Top             =   11640
      Width           =   570
   End
   Begin VB.Label lblf 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   2640
      TabIndex        =   4
      Top             =   11640
      Width           =   390
   End
   Begin VB.Label lblk 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "K"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   4800
      TabIndex        =   3
      Top             =   11640
      Width           =   450
   End
   Begin VB.Label lblh 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "H"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   3600
      TabIndex        =   2
      Top             =   11640
      Width           =   510
   End
   Begin VB.Label lbli 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   3960
      TabIndex        =   1
      Top             =   11640
      Width           =   450
   End
   Begin VB.Label lbla 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   11640
      Width           =   450
   End
   Begin VB.Image Image6 
      Appearance      =   0  'Flat
      Height          =   720
      Left            =   0
      Picture         =   "hangman.frx":3463FE
      Stretch         =   -1  'True
      Top             =   10560
      Width           =   840
   End
   Begin VB.Image Image7 
      Appearance      =   0  'Flat
      Height          =   720
      Left            =   360
      Picture         =   "hangman.frx":3468A8
      Stretch         =   -1  'True
      Top             =   9960
      Width           =   840
   End
   Begin VB.Image Image8 
      Appearance      =   0  'Flat
      Height          =   720
      Left            =   960
      Picture         =   "hangman.frx":346D52
      Stretch         =   -1  'True
      Top             =   10560
      Width           =   840
   End
   Begin VB.Image imgshell2 
      Appearance      =   0  'Flat
      Height          =   720
      Left            =   2760
      Picture         =   "hangman.frx":3471FC
      Stretch         =   -1  'True
      Top             =   10560
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Image imghurt 
      Height          =   1920
      Left            =   12600
      Picture         =   "hangman.frx":3476A6
      Stretch         =   -1  'True
      Top             =   9480
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Image imgmario 
      Height          =   1620
      Left            =   13080
      Picture         =   "hangman.frx":347BCB
      Stretch         =   -1  'True
      Top             =   9720
      Width           =   960
   End
   Begin VB.Image imgl2 
      Height          =   1710
      Left            =   11760
      Picture         =   "hangman.frx":347FDE
      Stretch         =   -1  'True
      Top             =   4080
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Menu mnufile 
      Caption         =   "File"
      Begin VB.Menu mnunew 
         Caption         =   "New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuseperator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnumenu 
         Caption         =   "Main Menu"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuquit 
         Caption         =   "Quit"
         Shortcut        =   ^Q
      End
   End
End
Attribute VB_Name = "frmhang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim s As Long
Dim a As Long
Dim c As Integer
Dim e As Long
Dim test As String
Dim arrwords(10) As String
Dim word As String
Dim num As Integer
Dim t As Integer
Dim b As Integer

Private Sub cmds2_Click()
timershell2.Enabled = True
e = 1.5
End Sub

Private Sub cmdtest_Click()
Randomize

    For asdf = 0 To 0
        arrwords(asdf) = Int(11 * Rnd)
        lblnum = arrwords(asdf)
            
        
        
    Next asdf
lbl1 = ""
lbl2 = ""
lbl3 = ""
lbl4 = ""
lbl5 = ""
lbl6 = ""
lbl7 = ""
lbl8 = ""
lbl9 = ""
lbl10 = ""
If lblnum = 0 Then
    lblword = "ROFL"
    lbl1.Caption = "R"
    lbl2.Caption = "O"
    lbl3.Caption = "F"
    lbl4.Caption = "L"
End If
If lblnum = 1 Then
    lblword = "SHYGUY"
    lbl1.Caption = "S"
    lbl2.Caption = "H"
    lbl3.Caption = "Y"
    lbl4.Caption = "G"
    lbl5.Caption = "U"
    lbl6.Caption = "Y"
End If
If lblnum = 2 Then
    lblword = "GOOMBA"
    lbl1.Caption = "G"
    lbl2.Caption = "O"
    lbl3.Caption = "O"
    lbl4.Caption = "M"
    lbl5.Caption = "B"
    lbl6.Caption = "A"
End If
If lblnum = 3 Then
    lblword = "MUSHROOM"
    lbl1.Caption = "M"
    lbl2.Caption = "U"
    lbl3.Caption = "S"
    lbl4.Caption = "H"
    lbl5.Caption = "R"
    lbl6.Caption = "O"
    lbl7.Caption = "O"
    lbl8.Caption = "M"
End If
If lblnum = 4 Then
    lblword = "MARIO"
    lbl1.Caption = "M"
    lbl2.Caption = "A"
    lbl3.Caption = "R"
    lbl4.Caption = "I"
    lbl5.Caption = "O"
End If
If lblnum = 5 Then
    lblword = "LUIGI"
    lbl1.Caption = "L"
    lbl2.Caption = "U"
    lbl3.Caption = "I"
    lbl4.Caption = "G"
    lbl5.Caption = "I"
End If
If lblnum = 6 Then
    lblword = "BOWSER"
    lbl1.Caption = "B"
    lbl2.Caption = "O"
    lbl3.Caption = "W"
    lbl4.Caption = "S"
    lbl5.Caption = "E"
    lbl6.Caption = "R"
End If
If lblnum = 7 Then
    lblword = "PEACH"
    lbl1.Caption = "P"
    lbl2.Caption = "E"
    lbl3.Caption = "A"
    lbl4.Caption = "C"
    lbl5.Caption = "H"
End If
If lblnum = 8 Then
    lblword = "YOSHI"
    lbl1.Caption = "Y"
    lbl2.Caption = "O"
    lbl3.Caption = "S"
    lbl4.Caption = "H"
    lbl5.Caption = "I"
End If
If lblnum = 9 Then
    lblword = "TOAD"
    lbl1.Caption = "T"
    lbl2.Caption = "O"
    lbl3.Caption = "A"
    lbl4.Caption = "D"
End If
If lblnum = 10 Then
    lblword = "KOOPA"
    lbl1.Caption = "K"
    lbl2.Caption = "O"
    lbl3.Caption = "O"
    lbl4.Caption = "P"
    lbl5.Caption = "A"
End If
End Sub

Private Sub Command1_Click()
init = Trim(UCase(InputBox("Enter name:", "Highscore", "")))
End Sub

Private Sub Command2_Click()
high = Trim(UCase(InputBox("Enter num:", "Highscore", "")))
End Sub

Private Sub Command3_Click()
timershell.Enabled = True
s = 1.5
End Sub

Private Sub Command4_Click()
Dim x As Integer
x = x + 1
lbllives = x

End Sub

Private Sub Form_Activate()
    lbld1.Visible = False
    lbld2.Visible = False
    lbld3.Visible = False
    lbld4.Visible = False
    lbld5.Visible = False
    lbld6.Visible = False
    lbld7.Visible = False
    lbld8.Visible = False
lbl1.Visible = False
lbl2.Visible = False
lbl3.Visible = False
lbl4.Visible = False
lbl5.Visible = False
lbl6.Visible = False
lbl7.Visible = False
lbl8.Visible = False
lbla.Enabled = 1
lblb.Enabled = 1
lblc.Enabled = 1
lbld.Enabled = 1
lble.Enabled = 1
lblf.Enabled = 1
lblg.Enabled = 1
lblh.Enabled = 1
lbli.Enabled = 1
lblj.Enabled = 1
lblk.Enabled = 1
lbll.Enabled = 1
lblm.Enabled = 1
lbln.Enabled = 1
lblo.Enabled = 1
lblp.Enabled = 1
lblq.Enabled = 1
lblr.Enabled = 1
lbls.Enabled = 1
lblt.Enabled = 1
lblu.Enabled = 1
lblv.Enabled = 1
lblw.Enabled = 1
lblx.Enabled = 1
lbly.Enabled = 1
lblz.Enabled = 1
lbl1 = ""
lbl2 = ""
lbl3 = ""
lbl4 = ""
lbl5 = ""
lbl6 = ""
lbl7 = ""
lbl8 = ""
lbl9 = ""
lbl10 = ""
Randomize
e = 0
Dim asdf As Integer
For asdf = 0 To 0
    arrwords(asdf) = Int(10 * Rnd)
    lblnum.Caption = arrwords(asdf)
Next asdf
lbl1 = ""
lbl2 = ""
lbl3 = ""
lbl4 = ""
lbl5 = ""
lbl6 = ""
lbl7 = ""
lbl8 = ""
lbl9 = ""
lbl10 = ""
If lblnum = 0 Then
    lblword = "ROFL"
    lbl1.Caption = "R"
    lbl2.Caption = "O"
    lbl3.Caption = "F"
    lbl4.Caption = "L"
    lbld1.Visible = True
    lbld2.Visible = True
    lbld3.Visible = True
    lbld4.Visible = True
End If
If lblnum = 1 Then
    lblword = "SHYGUY"
    lbl1.Caption = "S"
    lbl2.Caption = "H"
    lbl3.Caption = "Y"
    lbl4.Caption = "G"
    lbl5.Caption = "U"
    lbl6.Caption = "Y"
    lbld1.Visible = True
    lbld2.Visible = True
    lbld3.Visible = True
    lbld4.Visible = True
    lbld5.Visible = True
    lbld6.Visible = True
End If
If lblnum = 2 Then
    lblword = "GOOMBA"
    lbl1.Caption = "G"
    lbl2.Caption = "O"
    lbl3.Caption = "O"
    lbl4.Caption = "M"
    lbl5.Caption = "B"
    lbl6.Caption = "A"
    lbld1.Visible = True
    lbld2.Visible = True
    lbld3.Visible = True
    lbld4.Visible = True
    lbld5.Visible = True
    lbld6.Visible = True
End If
If lblnum = 3 Then
    lblword = "MUSHROOM"
    lbl1.Caption = "M"
    lbl2.Caption = "U"
    lbl3.Caption = "S"
    lbl4.Caption = "H"
    lbl5.Caption = "R"
    lbl6.Caption = "O"
    lbl7.Caption = "O"
    lbl8.Caption = "M"
    lbld1.Visible = True
    lbld2.Visible = True
    lbld3.Visible = True
    lbld4.Visible = True
    lbld5.Visible = True
    lbld6.Visible = True
    lbld7.Visible = True
    lbld8.Visible = True
End If
If lblnum = 4 Then
    lblword = "MARIO"
    lbl1.Caption = "M"
    lbl2.Caption = "A"
    lbl3.Caption = "R"
    lbl4.Caption = "I"
    lbl5.Caption = "O"
    lbld1.Visible = True
    lbld2.Visible = True
    lbld3.Visible = True
    lbld4.Visible = True
    lbld5.Visible = True
End If
If lblnum = 5 Then
    lblword = "LUIGI"
    lbl1.Caption = "L"
    lbl2.Caption = "U"
    lbl3.Caption = "I"
    lbl4.Caption = "G"
    lbl5.Caption = "I"
    lbld1.Visible = True
    lbld2.Visible = True
    lbld3.Visible = True
    lbld4.Visible = True
    lbld5.Visible = True
End If
If lblnum = 6 Then
    lblword = "BOWSER"
    lbl1.Caption = "B"
    lbl2.Caption = "O"
    lbl3.Caption = "W"
    lbl4.Caption = "S"
    lbl5.Caption = "E"
    lbl6.Caption = "R"
    lbld1.Visible = True
    lbld2.Visible = True
    lbld3.Visible = True
    lbld4.Visible = True
    lbld5.Visible = True
    lbld6.Visible = True
End If
If lblnum = 7 Then
    lblword = "PEACH"
    lbl1.Caption = "P"
    lbl2.Caption = "E"
    lbl3.Caption = "A"
    lbl4.Caption = "C"
    lbl5.Caption = "H"
    lbld1.Visible = True
    lbld2.Visible = True
    lbld3.Visible = True
    lbld4.Visible = True
    lbld5.Visible = True
End If
If lblnum = 8 Then
    lblword = "YOSHI"
    lbl1.Caption = "Y"
    lbl2.Caption = "O"
    lbl3.Caption = "S"
    lbl4.Caption = "H"
    lbl5.Caption = "I"
    lbld1.Visible = True
    lbld2.Visible = True
    lbld3.Visible = True
    lbld4.Visible = True
    lbld5.Visible = True
End If
If lblnum = 9 Then
    lblword = "TOAD"
    lbl1.Caption = "T"
    lbl2.Caption = "O"
    lbl3.Caption = "A"
    lbl4.Caption = "D"
    lbld1.Visible = True
    lbld2.Visible = True
    lbld3.Visible = True
    lbld4.Visible = True
End If
If lblnum = 10 Then
    lblword = "KOOPA"
    lbl1.Caption = "K"
    lbl2.Caption = "O"
    lbl3.Caption = "O"
    lbl4.Caption = "P"
    lbl5.Caption = "A"
    lbld1.Visible = True
    lbld2.Visible = True
    lbld3.Visible = True
    lbld4.Visible = True
    lbld5.Visible = True
End If
numb = lblnum
high = 0
topp = 0
s = 0
topp = 0
b = 0
a = 0
lblmil = mil
x = 10
z = 0
lblmi.Enabled = True
End Sub

Private Sub Form_Deactivate()
lblmi.Enabled = False
If asdf = 0 Then
    qwerty = 0
End If
If asdf = 1 Then
    qwerty = 1
End If
If asdf = 2 Then
    qwerty = 2
End If
If asdf = 3 Then
    qwerty = 3
End If
If asdf = 4 Then
    qwerty = 4
End If
If asdf = 5 Then
    qwerty = 5
End If
If asdf = 6 Then
    qwerty = 6
End If
If asdf = 7 Then
    qwerty = 7
End If
If asdf = 8 Then
    qwerty = 8
End If
If asdf = 9 Then
    qwerty = 9
End If
If asdf = 10 Then
    qwerty = 10
End If
End Sub


Private Sub Form_Load()
x = 10
End Sub

Private Sub lbla_Click()

lbla.Enabled = False
If lblnum = 0 Then
    timershell2.Enabled = True
    e = 1.5
mil = mil - 10
End If
If lblnum = 1 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 2 Then
    timershell.Enabled = True
    s = 1.5
    lbl6 = "A"
    lbl6.Visible = True

End If
If lblnum = 3 Then
    timershell2.Enabled = True
    e = 1.5
mil = mil - 10
End If
If lblnum = 4 Then
    timershell.Enabled = True
    s = 1.5
    lbl2.Visible = True
    lbl2 = "A"
    lifes = lifes + 10
    
End If
If lblnum = 5 Then
    timershell2.Enabled = True
    e = 1.5
mil = mil - 10
End If
If lblnum = 6 Then
    timershell2.Enabled = True
    e = 1.5
mil = mil - 10
End If
If lblnum = 7 Then
    timershell.Enabled = True
    s = 1.5
    lbl3.Visible = True
    lbl3 = "A"
    lifes = lifes + 10
End If
If lblnum = 8 Then
    timershell2.Enabled = True
    e = 1.5
mil = mil - 10
End If
If lblnum = 9 Then
    timershell.Enabled = True
    s = 1.5
    lbl3.Visible = True
    lbl3 = "A"

End If
If lblnum = 10 Then
    timershell.Enabled = True
    s = 1.5
    lbl5.Visible = True
    lbl5 = "A"

End If
If lblnum = 0 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 1 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 2 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 3 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True And lbl7.Visible = True And lbl8.Visible = True Then
        frmwin.Show
        frmhang.Hide
        
    End If
End If
If lblnum = 4 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 5 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 6 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 7 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 8 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 9 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 10 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
End Sub

Private Sub lblb_Click()
lblb.Enabled = False
If lblnum = 0 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 1 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 2 Then
    timershell.Enabled = True
    s = 1.5
        lbl5.Visible = True
    lbl5 = "B"
End If
If lblnum = 3 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 4 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 5 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 6 Then
    timershell.Enabled = True
    s = 1.5
        lbl1.Visible = True
    lbl1 = "B"
End If
If lblnum = 7 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 8 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 9 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 10 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 0 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 1 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 2 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 3 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True And lbl7.Visible = True And lbl8.Visible = True Then
        frmwin.Show
        frmhang.Hide
        
    End If
End If
If lblnum = 4 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 5 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 6 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 7 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 8 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 9 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 10 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
End Sub

Private Sub lblc_Click()
lblc.Enabled = False
If lblnum = 0 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 1 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 2 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 3 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 4 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 5 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 6 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 7 Then
    timershell.Enabled = True
    s = 1.5
    lbl4.Visible = True
    lbl4 = "C"
End If
If lblnum = 8 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 9 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 10 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 0 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 1 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 2 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 3 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True And lbl7.Visible = True And lbl8.Visible = True Then
        frmwin.Show
        frmhang.Hide
        
    End If
End If
If lblnum = 4 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 5 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 6 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 7 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 8 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 9 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 10 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
End Sub

Private Sub lbld_Click()
lbld.Enabled = False
If lblnum = 0 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 1 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 2 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 3 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 4 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 5 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 6 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 7 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 8 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 9 Then
    timershell.Enabled = True
    s = 1.5
    lbl4.Visible = True
    lbl4 = "D"
End If
If lblnum = 10 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 0 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 1 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 2 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 3 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True And lbl7.Visible = True And lbl8.Visible = True Then
        frmwin.Show
        frmhang.Hide
        
    End If
End If
If lblnum = 4 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 5 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 6 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 7 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 8 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 9 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 10 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
End Sub

Private Sub lble_Click()
lble.Enabled = False
If lblnum = 0 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 1 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 2 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 3 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 4 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 5 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 6 Then
    timershell.Enabled = True
    s = 1.5
    lbl5.Visible = True
    lbl5 = "E"
End If
If lblnum = 7 Then
    timershell.Enabled = True
    s = 1.5
        lbl2.Visible = True
    lbl2 = "E"
End If
If lblnum = 8 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 9 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 10 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 0 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 1 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 2 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 3 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True And lbl7.Visible = True And lbl8.Visible = True Then
        frmwin.Show
        frmhang.Hide
        
    End If
End If
If lblnum = 4 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 5 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 6 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 7 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 8 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 9 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 10 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
End Sub
Private Sub lblf_Click()
lblf.Enabled = False
If lblnum = 0 Then
    timershell.Enabled = True
    s = 1.5
        lbl3.Visible = True
    lbl3 = "F"
End If
If lblnum = 1 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 2 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 3 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 4 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 5 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 6 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 7 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 8 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 9 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 10 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 0 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 1 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 2 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 3 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True And lbl7.Visible = True And lbl8.Visible = True Then
        frmwin.Show
        frmhang.Hide
        
    End If
End If
If lblnum = 4 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 5 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 6 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 7 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 8 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 9 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 10 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
End Sub

Private Sub lblg_Click()
lblg.Enabled = False
If lblnum = 0 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 1 Then
    timershell.Enabled = True
    s = 1.5
        lbl4.Visible = True
    lbl4 = "G"
End If
If lblnum = 2 Then
    timershell.Enabled = True
    s = 1.5
        lbl1.Visible = True
    lbl1 = "G"
End If
If lblnum = 3 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 4 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 5 Then
    timershell.Enabled = True
    s = 1.5
    lbl4.Visible = True
    lbl4 = "G"
End If
If lblnum = 6 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 7 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 8 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 9 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 10 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 0 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 1 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 2 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 3 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True And lbl7.Visible = True And lbl8.Visible = True Then
        frmwin.Show
        frmhang.Hide
        
    End If
End If
If lblnum = 4 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 5 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 6 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 7 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 8 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 9 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 10 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
End Sub

Private Sub lblh_Click()
lblh.Enabled = False
If lblnum = 0 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 1 Then
    timershell.Enabled = True
    s = 1.5
    lbl2.Visible = True
    lbl2 = "H"
End If
If lblnum = 2 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 3 Then
    timershell.Enabled = True
    s = 1.5
    lbl4.Visible = True
    lbl4 = "H"
End If
If lblnum = 4 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 5 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 6 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 7 Then
    timershell.Enabled = True
    s = 1.5
    lbl5.Visible = True
    lbl5 = "H"
End If
If lblnum = 8 Then
    timershell.Enabled = True
    s = 1.5
    lbl4.Visible = True
    lbl4 = "H"
End If
If lblnum = 9 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 10 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 0 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 1 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 2 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 3 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True And lbl7.Visible = True And lbl8.Visible = True Then
        frmwin.Show
        frmhang.Hide
        
    End If
End If
If lblnum = 4 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 5 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 6 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 7 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 8 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 9 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 10 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
End Sub

Private Sub lbli_Click()
lbli.Enabled = False
If lblnum = 0 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 1 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 2 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 3 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 4 Then
    timershell.Enabled = True
    s = 1.5
    lbl4.Visible = True
    lbl4 = "I"
End If
If lblnum = 5 Then
    timershell.Enabled = True
    s = 1.5
    lbl3.Visible = True
    lbl3 = "I"
    lbl5.Visible = True
    lbl5 = "I"
End If
If lblnum = 6 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 7 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 8 Then
    timershell.Enabled = True
    s = 1.5
    lbl5.Visible = True
    lbl5 = "I"
End If
If lblnum = 9 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 10 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 0 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 1 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 2 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 3 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True And lbl7.Visible = True And lbl8.Visible = True Then
        frmwin.Show
        frmhang.Hide
        
    End If
End If
If lblnum = 4 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 5 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 6 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 7 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 8 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 9 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 10 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
End Sub

Private Sub lblj_Click()
lblj.Enabled = False
If lblnum = 0 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 1 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 2 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 3 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 4 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 5 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 6 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 7 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 8 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 9 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 10 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 0 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 1 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 2 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 3 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True And lbl7.Visible = True And lbl8.Visible = True Then
        frmwin.Show
        frmhang.Hide
        
    End If
End If
If lblnum = 4 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 5 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 6 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 7 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 8 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 9 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 10 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
End Sub

Private Sub lblk_Click()
lblk.Enabled = False
If lblnum = 0 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 1 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 2 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 3 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 4 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 5 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 6 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 7 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 8 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 9 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 10 Then
    timershell.Enabled = True
    s = 1.5
    lbl1.Visible = True
    lbl1 = "K"
End If
If lblnum = 0 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 1 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 2 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 3 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True And lbl7.Visible = True And lbl8.Visible = True Then
        frmwin.Show
        frmhang.Hide
        
    End If
End If
If lblnum = 4 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 5 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 6 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 7 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 8 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 9 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 10 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
End Sub

Private Sub lbll_Click()
lbll.Enabled = False
If lblnum = 0 Then
    timershell.Enabled = True
    s = 1.5
    lbl4.Visible = True
    lbl4 = "L"
End If
If lblnum = 1 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 2 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 3 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 4 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 5 Then
    timershell.Enabled = True
    s = 1.5
    lbl1.Visible = True
    lbl1 = "L"
End If
If lblnum = 6 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 7 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 8 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 9 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 10 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 0 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 1 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 2 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 3 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True And lbl7.Visible = True And lbl8.Visible = True Then
        frmwin.Show
        frmhang.Hide
        
    End If
End If
If lblnum = 4 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 5 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 6 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 7 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 8 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 9 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 10 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
End Sub

Private Sub lblm_Click()
lblm.Enabled = False
If lblnum = 0 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 1 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 2 Then
    timershell.Enabled = True
    s = 1.5
    lbl4.Visible = True
    lbl4 = "M"
End If
If lblnum = 3 Then
    timershell.Enabled = True
    s = 1.5
    lbl1.Visible = True
    lbl1 = "M"
    lbl8.Visible = True
    lbl8 = "M"
End If
If lblnum = 4 Then
    timershell.Enabled = True
    s = 1.5
    lbl1.Visible = True
    lbl1 = "M"
End If
If lblnum = 5 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 6 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 7 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 8 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 9 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 10 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 0 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 1 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 2 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 3 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True And lbl7.Visible = True And lbl8.Visible = True Then
        frmwin.Show
        frmhang.Hide
        
    End If
End If
If lblnum = 4 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 5 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 6 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 7 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 8 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 9 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 10 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
End Sub

Private Sub lblmi_Timer()
Randomize
lblmil = mil
mil = mil - 1
If mil < 1 Then
    lblmil = 0
    lbl1.Visible = True
    lbl2.Visible = True
    lbl3.Visible = True
    lbl4.Visible = True
    lbl5.Visible = True
    lbl6.Visible = True
    lbl7.Visible = True
    lbl8.Visible = True
    If lblnum = 0 Then
    lblword = "ROFL"
    lbl1.Caption = "R"
    lbl2.Caption = "O"
    lbl3.Caption = "F"
    lbl4.Caption = "L"
    lbld1.Visible = True
    lbld2.Visible = True
    lbld3.Visible = True
    lbld4.Visible = True
End If
If lblnum = 1 Then
    lblword = "SHYGUY"
    lbl1.Caption = "S"
    lbl2.Caption = "H"
    lbl3.Caption = "Y"
    lbl4.Caption = "G"
    lbl5.Caption = "U"
    lbl6.Caption = "Y"
    lbld1.Visible = True
    lbld2.Visible = True
    lbld3.Visible = True
    lbld4.Visible = True
    lbld5.Visible = True
    lbld6.Visible = True
End If
If lblnum = 2 Then
    lblword = "GOOMBA"
    lbl1.Caption = "G"
    lbl2.Caption = "O"
    lbl3.Caption = "O"
    lbl4.Caption = "M"
    lbl5.Caption = "B"
    lbl6.Caption = "A"
    lbld1.Visible = True
    lbld2.Visible = True
    lbld3.Visible = True
    lbld4.Visible = True
    lbld5.Visible = True
    lbld6.Visible = True
End If
If lblnum = 3 Then
    lblword = "MUSHROOM"
    lbl1.Caption = "M"
    lbl2.Caption = "U"
    lbl3.Caption = "S"
    lbl4.Caption = "H"
    lbl5.Caption = "R"
    lbl6.Caption = "O"
    lbl7.Caption = "O"
    lbl8.Caption = "M"
    lbld1.Visible = True
    lbld2.Visible = True
    lbld3.Visible = True
    lbld4.Visible = True
    lbld5.Visible = True
    lbld6.Visible = True
    lbld7.Visible = True
    lbld8.Visible = True
End If
If lblnum = 4 Then
    lblword = "MARIO"
    lbl1.Caption = "M"
    lbl2.Caption = "A"
    lbl3.Caption = "R"
    lbl4.Caption = "I"
    lbl5.Caption = "O"
    lbld1.Visible = True
    lbld2.Visible = True
    lbld3.Visible = True
    lbld4.Visible = True
    lbld5.Visible = True
End If
If lblnum = 5 Then
    lblword = "LUIGI"
    lbl1.Caption = "L"
    lbl2.Caption = "U"
    lbl3.Caption = "I"
    lbl4.Caption = "G"
    lbl5.Caption = "I"
    lbld1.Visible = True
    lbld2.Visible = True
    lbld3.Visible = True
    lbld4.Visible = True
    lbld5.Visible = True
End If
If lblnum = 6 Then
    lblword = "BOWSER"
    lbl1.Caption = "B"
    lbl2.Caption = "O"
    lbl3.Caption = "W"
    lbl4.Caption = "S"
    lbl5.Caption = "E"
    lbl6.Caption = "R"
    lbld1.Visible = True
    lbld2.Visible = True
    lbld3.Visible = True
    lbld4.Visible = True
    lbld5.Visible = True
    lbld6.Visible = True
End If
If lblnum = 7 Then
    lblword = "PEACH"
    lbl1.Caption = "P"
    lbl2.Caption = "E"
    lbl3.Caption = "A"
    lbl4.Caption = "C"
    lbl5.Caption = "H"

End If
If lblnum = 8 Then
    lblword = "YOSHI"
    lbl1.Caption = "Y"
    lbl2.Caption = "O"
    lbl3.Caption = "S"
    lbl4.Caption = "H"
    lbl5.Caption = "I"
    lbld1.Visible = True
    lbld2.Visible = True
    lbld3.Visible = True
    lbld4.Visible = True
    lbld5.Visible = True
End If
If lblnum = 9 Then
    lblword = "TOAD"
    lbl1.Caption = "T"
    lbl2.Caption = "O"
    lbl3.Caption = "A"
    lbl4.Caption = "D"
    lbld1.Visible = True
    lbld2.Visible = True
    lbld3.Visible = True
    lbld4.Visible = True
End If
If lblnum = 10 Then
    lblword = "KOOPA"
    lbl1.Caption = "K"
    lbl2.Caption = "O"
    lbl3.Caption = "O"
    lbl4.Caption = "P"
    lbl5.Caption = "A"
    lbld1.Visible = True
    lbld2.Visible = True
    lbld3.Visible = True
    lbld4.Visible = True
    lbld5.Visible = True
End If
    test = MsgBox("Retry???", vbYesNo, "Game Over")
    If test = vbYes Then
        frmdif.Show
        frmhang.Hide
    End If
    If test = vbNo Then
        frmmain.Show
        frmhang.Hide
    End If
        lbl1.Visible = False
lbl2.Visible = False
lbl3.Visible = False
lbl4.Visible = False
lbl5.Visible = False
lbl6.Visible = False
lbl7.Visible = False
lbl8.Visible = False
lbl1 = ""
lbl2 = ""
lbl3 = ""
lbl4 = ""
lbl5 = ""
lbl6 = ""
lbl7 = ""
lbl8 = ""
lbl9 = ""
lbl10 = ""
lbla.Enabled = 1
lblb.Enabled = 1
lblc.Enabled = 1
lbld.Enabled = 1
lble.Enabled = 1
lblf.Enabled = 1
lblg.Enabled = 1
lblh.Enabled = 1
lbli.Enabled = 1
lblj.Enabled = 1
lblk.Enabled = 1
lbll.Enabled = 1
lblm.Enabled = 1
lbln.Enabled = 1
lblo.Enabled = 1
lblp.Enabled = 1
lblq.Enabled = 1
lblr.Enabled = 1
lbls.Enabled = 1
lblt.Enabled = 1
lblu.Enabled = 1
lblv.Enabled = 1
lblw.Enabled = 1
lblx.Enabled = 1
lbly.Enabled = 1
lblz.Enabled = 1
    End If


End Sub

Private Sub lbln_Click()
lbln.Enabled = False
If lblnum = 0 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 1 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 2 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 3 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 4 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 5 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 6 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 7 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 8 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 9 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 10 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 0 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 1 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 2 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 3 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True And lbl7.Visible = True And lbl8.Visible = True Then
        frmwin.Show
        frmhang.Hide
        
    End If
End If
If lblnum = 4 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 5 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 6 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 7 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 8 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 9 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 10 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
End Sub

Private Sub lblo_Click()
lblo.Enabled = False
If lblnum = 0 Then
    timershell.Enabled = True
    s = 1.5
    lbl2.Visible = True
    lbl2 = "O"
End If
If lblnum = 1 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 2 Then
    timershell.Enabled = True
    s = 1.5
    lbl2.Visible = True
    lbl2 = "O"
    lbl3.Visible = True
    lbl3 = "O"
End If
If lblnum = 3 Then
    timershell.Enabled = True
    s = 1.5
    lbl6.Visible = True
    lbl6 = "O"
    lbl7.Visible = True
    lbl7 = "O"
End If
If lblnum = 4 Then
    timershell.Enabled = True
    s = 1.5
    lbl5.Visible = True
    lbl5 = "O"
End If
If lblnum = 5 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 6 Then
    timershell.Enabled = True
    s = 1.5
    lbl2.Visible = True
    lbl2 = "O"
End If
If lblnum = 7 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 8 Then
    timershell.Enabled = True
    s = 1.5
    lbl2.Visible = True
    lbl2 = "O"
End If
If lblnum = 9 Then
    timershell.Enabled = True
    s = 1.5
    lbl2.Visible = True
    lbl2 = "O"
End If
If lblnum = 10 Then
    timershell.Enabled = True
    s = 1.5
    lbl2.Visible = True
    lbl2 = "O"
    lbl3.Visible = True
    lbl3 = "O"
End If
If lblnum = 0 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 1 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 2 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 3 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True And lbl7.Visible = True And lbl8.Visible = True Then
        frmwin.Show
        frmhang.Hide
        
    End If
End If
If lblnum = 4 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 5 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 6 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 7 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 8 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 9 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 10 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
End Sub

Private Sub lblp_Click()
lblp.Enabled = False
If lblnum = 0 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 1 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 2 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 3 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 4 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 5 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 6 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 7 Then
    timershell.Enabled = True
    s = 1.5
    lbl1.Visible = True
    lbl1 = "P"
End If
If lblnum = 8 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 9 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 10 Then
    timershell.Enabled = True
    s = 1.5
    lbl4.Visible = True
    lbl4 = "P"
End If
If lblnum = 0 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 1 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 2 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 3 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True And lbl7.Visible = True And lbl8.Visible = True Then
        frmwin.Show
        frmhang.Hide
        
    End If
End If
If lblnum = 4 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 5 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 6 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 7 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 8 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 9 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 10 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
End Sub

Private Sub lblq_Click()
mil = mil - 10
lblq.Enabled = False
If lblnum = 0 Then
    timershell2.Enabled = True
    e = 1.5
End If
If lblnum = 1 Then
    timershell2.Enabled = True
    e = 1.5
End If
If lblnum = 2 Then
    timershell2.Enabled = True
    e = 1.5
End If
If lblnum = 3 Then
    timershell2.Enabled = True
    e = 1.5
End If
If lblnum = 4 Then
    timershell2.Enabled = True
    e = 1.5
End If
If lblnum = 5 Then
    timershell2.Enabled = True
    e = 1.5
End If
If lblnum = 6 Then
    timershell2.Enabled = True
    e = 1.5
End If
If lblnum = 7 Then
    timershell2.Enabled = True
    e = 1.5
End If
If lblnum = 8 Then
    timershell2.Enabled = True
    e = 1.5
End If
If lblnum = 9 Then
    timershell2.Enabled = True
    e = 1.5
End If
If lblnum = 10 Then
    timershell2.Enabled = True
    e = 1.5
End If
If lblnum = 0 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 1 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 2 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 3 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True And lbl7.Visible = True And lbl8.Visible = True Then
        frmwin.Show
        frmhang.Hide
        
    End If
End If
If lblnum = 4 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 5 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 6 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 7 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 8 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 9 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 10 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
End Sub

Private Sub lblr_Click()
lblr.Enabled = False
If lblnum = 0 Then
    timershell.Enabled = True
    s = 1.5
    lbl1.Visible = True
    lbl1 = "R"
End If
If lblnum = 1 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 2 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 3 Then
    timershell.Enabled = True
    s = 1.5
    lbl5.Visible = True
    lbl5 = "R"
End If
If lblnum = 4 Then
    timershell.Enabled = True
    s = 1.5
    lbl3.Visible = True
    lbl3 = "R"
End If
If lblnum = 5 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 6 Then
    timershell.Enabled = True
    s = 1.5
    lbl6.Visible = True
    lbl6 = "R"
End If
If lblnum = 7 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 8 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 9 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 10 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 0 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 1 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 2 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 3 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True And lbl7.Visible = True And lbl8.Visible = True Then
        frmwin.Show
        frmhang.Hide
        
    End If
End If
If lblnum = 4 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 5 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 6 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 7 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 8 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 9 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 10 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
End Sub

Private Sub lbls_Click()
lbls.Enabled = False
If lblnum = 0 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 1 Then
    timershell.Enabled = True
    s = 1.5
    lbl1.Visible = True
    lbl1 = "S"
    
End If
If lblnum = 2 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 3 Then
    timershell.Enabled = True
    s = 1.5
    lbl3.Visible = True
    lbl3 = "S"
End If
If lblnum = 4 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 5 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 6 Then
    timershell.Enabled = True
    s = 1.5
    lbl4.Visible = True
    lbl4 = "S"
End If
If lblnum = 7 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 8 Then
    timershell.Enabled = True
    s = 1.5
    lbl3.Visible = True
    lbl3 = "S"
End If
If lblnum = 9 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 10 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 0 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 1 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 2 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 3 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True And lbl7.Visible = True And lbl8.Visible = True Then
        frmwin.Show
        frmhang.Hide
        
    End If
End If
If lblnum = 4 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 5 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 6 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 7 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 8 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 9 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 10 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
End Sub

Private Sub lblt_Click()
lblt.Enabled = False
If lblnum = 0 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 1 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 2 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 3 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 4 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 5 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 6 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 7 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 8 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 9 Then
    timershell.Enabled = True
    s = 1.5
    lbl1.Visible = True
    lbl1 = "T"
End If
If lblnum = 10 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 0 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 1 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 2 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 3 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True And lbl7.Visible = True And lbl8.Visible = True Then
        frmwin.Show
        frmhang.Hide
        
    End If
End If
If lblnum = 4 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 5 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 6 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 7 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 8 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 9 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 10 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
End Sub

Private Sub lblu_Click()
lblu.Enabled = False
If lblnum = 0 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 1 Then
    timershell.Enabled = True
    s = 1.5
    lbl5.Visible = True
    lbl5 = "U"
End If
If lblnum = 2 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 3 Then
    timershell.Enabled = True
    s = 1.5
    lbl2.Visible = True
    lbl2 = "U"
End If
If lblnum = 4 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 5 Then
    timershell.Enabled = True
    s = 1.5
    lbl2.Visible = True
    lbl2 = "U"
End If
If lblnum = 6 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 7 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 8 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 9 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 10 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 0 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 1 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 2 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 3 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True And lbl7.Visible = True And lbl8.Visible = True Then
        frmwin.Show
        frmhang.Hide
        
    End If
End If
If lblnum = 4 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 5 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 6 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 7 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 8 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 9 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 10 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
End Sub

Private Sub lblv_Click()
mil = mil - 10
lblv.Enabled = False
If lblnum = 0 Then
    timershell2.Enabled = True
    e = 1.5
End If
If lblnum = 1 Then
    timershell2.Enabled = True
    e = 1.5
End If
If lblnum = 2 Then
    timershell2.Enabled = True
    e = 1.5
End If
If lblnum = 3 Then
    timershell2.Enabled = True
    e = 1.5
End If
If lblnum = 4 Then
    timershell2.Enabled = True
    e = 1.5
End If
If lblnum = 5 Then
    timershell2.Enabled = True
    e = 1.5
End If
If lblnum = 6 Then
    timershell2.Enabled = True
    e = 1.5
End If
If lblnum = 7 Then
    timershell2.Enabled = True
    e = 1.5
End If
If lblnum = 8 Then
    timershell2.Enabled = True
    e = 1.5
End If
If lblnum = 9 Then
    timershell2.Enabled = True
    e = 1.5
End If
If lblnum = 10 Then
    timershell2.Enabled = True
    e = 1.5
End If
If lblnum = 0 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 1 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 2 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 3 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True And lbl7.Visible = True And lbl8.Visible = True Then
        frmwin.Show
        frmhang.Hide
        
    End If
End If
If lblnum = 4 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 5 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 6 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 7 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 8 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 9 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 10 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
End Sub

Private Sub lblw_Click()
mil = mil - 10
lblw.Enabled = False
If lblnum = 0 Then
    timershell2.Enabled = True
    e = 1.5
End If
If lblnum = 1 Then
    timershell2.Enabled = True
    e = 1.5
End If
If lblnum = 2 Then
    timershell2.Enabled = True
    e = 1.5
End If
If lblnum = 3 Then
    timershell2.Enabled = True
    e = 1.5
End If
If lblnum = 4 Then
    timershell2.Enabled = True
    e = 1.5
End If
If lblnum = 5 Then
    timershell2.Enabled = True
    e = 1.5
End If
If lblnum = 6 Then
    timershell.Enabled = True
    s = 1.5
    lbl3.Visible = True
    lbl3 = "W"
    mil = mil + 10
End If
If lblnum = 7 Then
    timershell2.Enabled = True
    e = 1.5
End If
If lblnum = 8 Then
    timershell2.Enabled = True
    e = 1.5
End If
If lblnum = 9 Then
    timershell2.Enabled = True
    e = 1.5
End If
If lblnum = 10 Then
    timershell2.Enabled = True
    e = 1.5
End If
If lblnum = 0 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 1 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 2 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 3 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True And lbl7.Visible = True And lbl8.Visible = True Then
        frmwin.Show
        frmhang.Hide
        
    End If
End If
If lblnum = 4 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 5 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 6 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 7 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 8 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 9 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 10 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
End Sub

Private Sub lblx_Click()
mil = mil - 10
lblx.Enabled = False
If lblnum = 0 Then
    timershell2.Enabled = True
    e = 1.5
End If
If lblnum = 1 Then
    timershell2.Enabled = True
    e = 1.5
End If
If lblnum = 2 Then
    timershell2.Enabled = True
    e = 1.5
End If
If lblnum = 3 Then
    timershell2.Enabled = True
    e = 1.5
End If
If lblnum = 4 Then
    timershell2.Enabled = True
    e = 1.5
End If
If lblnum = 5 Then
    timershell2.Enabled = True
    e = 1.5
End If
If lblnum = 6 Then
    timershell2.Enabled = True
    e = 1.5
End If
If lblnum = 7 Then
    timershell2.Enabled = True
    e = 1.5
End If
If lblnum = 8 Then
    timershell2.Enabled = True
    e = 1.5
End If
If lblnum = 9 Then
    timershell2.Enabled = True
    e = 1.5
End If
If lblnum = 10 Then
    timershell2.Enabled = True
    e = 1.5
End If
If lblnum = 0 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 1 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 2 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 3 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True And lbl7.Visible = True And lbl8.Visible = True Then
        frmwin.Show
        frmhang.Hide
        
    End If
End If
If lblnum = 4 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 5 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 6 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 7 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 8 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 9 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 10 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
End Sub

Private Sub lbly_Click()
lbly.Enabled = False
If lblnum = 0 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 1 Then
    timershell.Enabled = True
    s = 1.5
    lbl3.Visible = True
    lbl3 = "Y"
    lbl6.Visible = True
    lbl6 = "Y"
End If
If lblnum = 2 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 3 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 4 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 5 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 6 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 7 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 8 Then
    timershell.Enabled = True
    s = 1.5
    lbl1.Visible = True
    lbl1 = "Y"
End If
If lblnum = 9 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 10 Then
    timershell2.Enabled = True
    e = 1.5
    mil = mil - 10
End If
If lblnum = 0 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 1 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 2 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 3 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True And lbl7.Visible = True And lbl8.Visible = True Then
        frmwin.Show
        frmhang.Hide
        
    End If
End If
If lblnum = 4 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 5 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 6 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 7 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 8 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 9 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 10 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
End Sub

Private Sub lblz_Click()
mil = mil - 10
lblz.Enabled = False
If lblnum = 0 Then
    timershell2.Enabled = True
    e = 1.5
End If
If lblnum = 1 Then
    timershell2.Enabled = True
    e = 1.5
End If
If lblnum = 2 Then
    timershell2.Enabled = True
    e = 1.5
End If
If lblnum = 3 Then
    timershell2.Enabled = True
    e = 1.5
End If
If lblnum = 4 Then
    timershell2.Enabled = True
    e = 1.5
End If
If lblnum = 5 Then
    timershell2.Enabled = True
    e = 1.5
End If
If lblnum = 6 Then
    timershell2.Enabled = True
    e = 1.5
End If
If lblnum = 7 Then
    timershell2.Enabled = True
    e = 1.5
End If
If lblnum = 8 Then
    timershell2.Enabled = True
    e = 1.5
End If
If lblnum = 9 Then
    timershell2.Enabled = True
    e = 1.5
End If
If lblnum = 10 Then
    timershell2.Enabled = True
    e = 1.5
End If
If lblnum = 0 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 1 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 2 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 3 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True And lbl7.Visible = True And lbl8.Visible = True Then
        frmwin.Show
        frmhang.Hide
        
    End If
End If
If lblnum = 4 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 5 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 6 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True And lbl6.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 7 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 8 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 9 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
If lblnum = 10 Then
    If lbl1.Visible = True And lbl2.Visible = True And lbl3.Visible = True And lbl4.Visible = True And lbl5.Visible = True Then
        frmwin.Show
        frmhang.Hide
    End If
End If
End Sub

Private Sub mnuhigh_Click()
frmhigh.Show
End Sub

Private Sub mnumenu_Click()
frmmain.Show
frmhang.Hide
End Sub

Private Sub mnunew_Click()

Randomize
Dim asdf As Integer
test = MsgBox("Retry???", vbYesNo, "New")
If test = vbYes Then
    mil = 180
    lblmil = 180
    lbl1.Visible = False
    lbld1.Visible = False
    lbld2.Visible = False
    lbld3.Visible = False
    lbld4.Visible = False
    lbld5.Visible = False
    lbld6.Visible = False
    lbld7.Visible = False
    lbld8.Visible = False
lbl2.Visible = False
lbl3.Visible = False
lbl4.Visible = False
lbl5.Visible = False
lbl6.Visible = False
lbl7.Visible = False
lbl8.Visible = False
lbl1 = ""
lbl2 = ""
lbl3 = ""
lbl4 = ""
lbl5 = ""
lbl6 = ""
lbl7 = ""
lbl8 = ""
lbl9 = ""
lbl10 = ""
    For asdf = 0 To 0
        arrwords(asdf) = Int(10 * Rnd)
        lblnum = arrwords(asdf)
        
    Next asdf
End If

If lblnum = 0 Then
    lblword = "ROFL"
    lbl1.Caption = "R"
    lbl2.Caption = "O"
    lbl3.Caption = "F"
    lbl4.Caption = "L"
    lbld1.Visible = True
    lbld2.Visible = True
    lbld3.Visible = True
    lbld4.Visible = True
End If
If lblnum = 1 Then
    lblword = "SHYGUY"
    lbl1.Caption = "S"
    lbl2.Caption = "H"
    lbl3.Caption = "Y"
    lbl4.Caption = "G"
    lbl5.Caption = "U"
    lbl6.Caption = "Y"
    lbld1.Visible = True
    lbld2.Visible = True
    lbld3.Visible = True
    lbld4.Visible = True
    lbld5.Visible = True
    lbld6.Visible = True
End If
If lblnum = 2 Then
    lblword = "GOOMBA"
    lbl1.Caption = "G"
    lbl2.Caption = "O"
    lbl3.Caption = "O"
    lbl4.Caption = "M"
    lbl5.Caption = "B"
    lbl6.Caption = "A"
    lbld1.Visible = True
    lbld2.Visible = True
    lbld3.Visible = True
    lbld4.Visible = True
    lbld5.Visible = True
    lbld6.Visible = True
End If
If lblnum = 3 Then
    lblword = "MUSHROOM"
    lbl1.Caption = "M"
    lbl2.Caption = "U"
    lbl3.Caption = "S"
    lbl4.Caption = "H"
    lbl5.Caption = "R"
    lbl6.Caption = "O"
    lbl7.Caption = "O"
    lbl8.Caption = "M"
    lbld1.Visible = True
    lbld2.Visible = True
    lbld3.Visible = True
    lbld4.Visible = True
    lbld5.Visible = True
    lbld6.Visible = True
    lbld7.Visible = True
    lbld8.Visible = True
End If
If lblnum = 4 Then
    lblword = "MARIO"
    lbl1.Caption = "M"
    lbl2.Caption = "A"
    lbl3.Caption = "R"
    lbl4.Caption = "I"
    lbl5.Caption = "O"
    lbld1.Visible = True
    lbld2.Visible = True
    lbld3.Visible = True
    lbld4.Visible = True
    lbld5.Visible = True
End If
If lblnum = 5 Then
    lblword = "LUIGI"
    lbl1.Caption = "L"
    lbl2.Caption = "U"
    lbl3.Caption = "I"
    lbl4.Caption = "G"
    lbl5.Caption = "I"
    lbld1.Visible = True
    lbld2.Visible = True
    lbld3.Visible = True
    lbld4.Visible = True
    lbld5.Visible = True
End If
If lblnum = 6 Then
    lblword = "BOWSER"
    lbl1.Caption = "B"
    lbl2.Caption = "O"
    lbl3.Caption = "W"
    lbl4.Caption = "S"
    lbl5.Caption = "E"
    lbl6.Caption = "R"
    lbld1.Visible = True
    lbld2.Visible = True
    lbld3.Visible = True
    lbld4.Visible = True
    lbld5.Visible = True
    lbld6.Visible = True
End If
If lblnum = 7 Then
    lblword = "PEACH"
    lbl1.Caption = "P"
    lbl2.Caption = "E"
    lbl3.Caption = "A"
    lbl4.Caption = "C"
    lbl5.Caption = "H"
    lbld1.Visible = True
    lbld2.Visible = True
    lbld3.Visible = True
    lbld4.Visible = True
    lbld5.Visible = True
End If
If lblnum = 8 Then
    lblword = "YOSHI"
    lbl1.Caption = "Y"
    lbl2.Caption = "O"
    lbl3.Caption = "S"
    lbl4.Caption = "H"
    lbl5.Caption = "I"
    lbld1.Visible = True
    lbld2.Visible = True
    lbld3.Visible = True
    lbld4.Visible = True
    lbld5.Visible = True
End If
If lblnum = 9 Then
    lblword = "TOAD"
    lbl1.Caption = "T"
    lbl2.Caption = "O"
    lbl3.Caption = "A"
    lbl4.Caption = "D"
    lbld1.Visible = True
    lbld2.Visible = True
    lbld3.Visible = True
    lbld4.Visible = True
End If
If lblnum = 10 Then
    lblword = "KOOPA"
    lbl1.Caption = "K"
    lbl2.Caption = "O"
    lbl3.Caption = "O"
    lbl4.Caption = "P"
    lbl5.Caption = "A"
    lbld1.Visible = True
    lbld2.Visible = True
    lbld3.Visible = True
    lbld4.Visible = True
    lbld5.Visible = True
End If
If test = vbNo Then
End If
jump = 0
mil = 180
lblmil = 180
lbla.Enabled = 1
lblb.Enabled = 1
lblc.Enabled = 1
lbld.Enabled = 1
lble.Enabled = 1
lblf.Enabled = 1
lblg.Enabled = 1
lblh.Enabled = 1
lbli.Enabled = 1
lblj.Enabled = 1
lblk.Enabled = 1
lbll.Enabled = 1
lblm.Enabled = 1
lbln.Enabled = 1
lblo.Enabled = 1
lblp.Enabled = 1
lblq.Enabled = 1
lblr.Enabled = 1
lbls.Enabled = 1
lblt.Enabled = 1
lblu.Enabled = 1
lblv.Enabled = 1
lblw.Enabled = 1
lblx.Enabled = 1
lbly.Enabled = 1
lblz.Enabled = 1
End Sub

Private Sub mnuquit_Click()
End
End Sub

Private Sub Timer1_Timer()
Static z As Integer
z = z + 1
If z = 1 Then
    img1.Visible = False
    img3.Visible = True
    img5.Visible = False
    img7.Visible = True
    img9.Visible = False
    img2.Visible = True
    img4.Visible = False
    img6.Visible = True
    img8.Visible = False
    img10.Visible = True
End If
If z = 2 Then
    img1.Visible = True
    img3.Visible = False
    img5.Visible = True
    img7.Visible = False
    img9.Visible = True
    img2.Visible = False
    img4.Visible = True
    img6.Visible = False
    img8.Visible = True
    img10.Visible = False
End If
If z = 3 Then
    z = 0
End If
End Sub

Private Sub Timer2_Timer()
Static count As Integer
count = count + 1
If count = 1 Then
    imgl1.Visible = True
    imgl2.Visible = False
    imgi3.Visible = False
    imgy1.Visible = True
    imgy2.Visible = False
End If
If count = 5 Then
    imgl1.Visible = False
    imgl2.Visible = True
    imgi3.Visible = False
    imgy1.Visible = False
    imgy2.Visible = True
End If
If count = 9 Then
    imgl1.Visible = False
    imgl2.Visible = False
    imgi3.Visible = True
    imgy1.Visible = True
    imgy2.Visible = False
End If
If count = 13 Then
    count = 0
End If
End Sub

Private Sub timerjump_Timer()
Static jump As Integer
jump = jump + 1
If jump = 1 Then
    imgmario.Top = 7800
End If
If jump = 4 Then
    imgmario.Top = 9720
End If
If jump = 7 Then
    timerjump.Enabled = False
    jump = jump + 1
End If
If jump = 8 Then
    jump = 0
End If
End Sub

Private Sub timershell_Timer()
imgshell.Move imgshell.Left + s * 50
imgshell.Visible = True
If imgshell.Left > 15360 Then
    imgshell.Left = 2760
    imgshell.Visible = False
    timershell.Enabled = False
    imgshell.Visible = False
End If
If imgshell.Left > 11280 Then
    timerjump.Enabled = True
End If

End Sub

Private Sub timershell2_Timer()
imgshell2.Visible = True

imgshell2.Move imgshell2.Left + e * 50
If imgshell2.Left > 15360 Then
    imgshell2.Left = 2760
    timershell2.Enabled = False
    imghurt.Visible = False
    imgmario.Visible = True
    imgshell2.Visible = False
End If
If imgshell2.Left > 12480 Then
    imghurt.Visible = True
    imgmario.Visible = False
End If
End Sub
