VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BackColor       =   &H80000016&
   Caption         =   "SPapak ver.10.01"
   ClientHeight    =   7680
   ClientLeft      =   105
   ClientTop       =   675
   ClientWidth     =   11880
   Icon            =   "SPapak.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "SPapak.frx":030A
   ScaleHeight     =   7680
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command14 
      Caption         =   "DESIGN PERIMETERS"
      Height          =   375
      Left            =   8040
      TabIndex        =   126
      Top             =   6120
      Width           =   3255
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Show ult. moments"
      Enabled         =   0   'False
      Height          =   375
      Left            =   9720
      TabIndex        =   111
      Top             =   7080
      Width           =   1575
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10680
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Show Loads"
      Enabled         =   0   'False
      Height          =   375
      Left            =   9720
      TabIndex        =   100
      Top             =   6600
      Width           =   1605
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Calc. coef."
      Enabled         =   0   'False
      Height          =   375
      Left            =   8040
      TabIndex        =   95
      Top             =   7080
      Width           =   1605
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Simpan data !"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8040
      TabIndex        =   94
      Top             =   6600
      Width           =   1575
   End
   Begin VB.CommandButton Command8 
      Caption         =   "LUKIS"
      Height          =   375
      Left            =   6120
      TabIndex        =   92
      Top             =   7080
      Width           =   1575
   End
   Begin VB.CommandButton Command7 
      Caption         =   "BEBAN HIDUP"
      Height          =   375
      Left            =   6120
      TabIndex        =   91
      Top             =   6600
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      Caption         =   "BEBAN KEMASAN"
      Height          =   375
      Left            =   4440
      TabIndex        =   90
      Top             =   6600
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "TEBAL PAPAK"
      Height          =   375
      Left            =   2760
      TabIndex        =   89
      Top             =   6600
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "UKUR RASUK (H)"
      Height          =   375
      Left            =   4440
      TabIndex        =   88
      Top             =   7080
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "UKUR RASUK (B)"
      Height          =   375
      Left            =   2760
      TabIndex        =   87
      Top             =   7080
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "UKUR RENTANG"
      Height          =   375
      Left            =   1080
      TabIndex        =   86
      Top             =   6600
      Width           =   1575
   End
   Begin VB.TextBox Text29 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   11270
      TabIndex        =   85
      Text            =   "Text29"
      Top             =   330
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   11270
      TabIndex        =   84
      Text            =   "Text28"
      Top             =   1990
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text27 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7560
      TabIndex        =   83
      Text            =   "Text27"
      Top             =   3550
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text26 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5640
      TabIndex        =   82
      Text            =   "Text26"
      Top             =   5130
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text25 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10800
      TabIndex        =   81
      Text            =   "Text25"
      Top             =   3120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text24 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8880
      TabIndex        =   80
      Text            =   "Text24"
      Top             =   3120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text23 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6960
      TabIndex        =   79
      Text            =   "Text23"
      Top             =   4680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text22 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5040
      TabIndex        =   78
      Text            =   "Text22"
      Top             =   6240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text21 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3000
      TabIndex        =   77
      Text            =   "Text21"
      Top             =   6240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text20 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   840
      TabIndex        =   76
      Text            =   "Text20"
      Top             =   6240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text19 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9360
      TabIndex        =   75
      Text            =   "Text19"
      Top             =   4200
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text18 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3600
      TabIndex        =   74
      Text            =   "Text18"
      Top             =   4320
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text17 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1560
      TabIndex        =   73
      Text            =   "Text17"
      Top             =   4320
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text16 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5520
      TabIndex        =   72
      Text            =   "Text16"
      Top             =   2760
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text15 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3600
      TabIndex        =   71
      Text            =   "Text15"
      Top             =   2760
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text14 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1560
      TabIndex        =   70
      Text            =   "Text14"
      Top             =   2760
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text13 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9360
      TabIndex        =   69
      Text            =   "Text13"
      Top             =   1080
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text12 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7440
      TabIndex        =   68
      Text            =   "Text12"
      Top             =   1080
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text11 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5520
      TabIndex        =   67
      Text            =   "Text11"
      Top             =   1080
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text10 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3600
      TabIndex        =   66
      Text            =   "Text10"
      Top             =   1080
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text9 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1560
      TabIndex        =   65
      Text            =   "Text9"
      Top             =   1080
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      MaxLength       =   6
      TabIndex        =   56
      Text            =   "10000"
      Top             =   4680
      Width           =   735
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      MaxLength       =   6
      TabIndex        =   55
      Text            =   "20000"
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "LABEL GRID"
      Height          =   375
      Left            =   1080
      TabIndex        =   38
      Top             =   7080
      Width           =   1575
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      MaxLength       =   6
      TabIndex        =   33
      Text            =   "30000"
      Top             =   1440
      Width           =   735
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   9840
      MaxLength       =   6
      TabIndex        =   32
      Text            =   "40000"
      Top             =   450
      Width           =   735
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7800
      MaxLength       =   6
      TabIndex        =   31
      Text            =   "80000"
      Top             =   450
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5880
      MaxLength       =   6
      TabIndex        =   30
      Text            =   "70000"
      Top             =   450
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3960
      MaxLength       =   6
      TabIndex        =   29
      Text            =   "60000"
      Top             =   450
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1920
      MaxLength       =   6
      TabIndex        =   28
      Text            =   "40000"
      Top             =   450
      Width           =   735
   End
   Begin VB.Label Label87 
      Caption         =   "Minimun Ast = 0.13% -> 0.24%bh."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   7320
      TabIndex        =   128
      Top             =   0
      Width           =   3975
   End
   Begin VB.Line Line2 
      X1              =   6450
      X2              =   6500
      Y1              =   50
      Y2              =   240
   End
   Begin VB.Label Label86 
      Caption         =   "Modification factor,  m.f. > 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   127
      Top             =   0
      Width           =   3135
   End
   Begin VB.Label Label85 
      Alignment       =   2  'Center
      Caption         =   "email : sohaimi@jkr.gov.my"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   8040
      TabIndex        =   125
      Top             =   7440
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label Label84 
      Alignment       =   2  'Center
      Caption         =   "PAPAK MS1195"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   1200
      TabIndex        =   124
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label83 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "83"
      Height          =   255
      Left            =   11400
      TabIndex        =   123
      Top             =   720
      Width           =   375
   End
   Begin VB.Label Label82 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "82"
      Height          =   255
      Left            =   11400
      TabIndex        =   122
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label Label81 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "81"
      Height          =   255
      Left            =   7680
      TabIndex        =   121
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label Label80 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "80"
      Height          =   255
      Left            =   5760
      TabIndex        =   120
      Top             =   5520
      Width           =   375
   End
   Begin VB.Label Label79 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "79"
      Height          =   255
      Left            =   10920
      TabIndex        =   119
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label Label78 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "78"
      Height          =   255
      Left            =   9000
      TabIndex        =   118
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label Label77 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "77"
      Height          =   255
      Left            =   7080
      TabIndex        =   117
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label Label76 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "76"
      Height          =   255
      Left            =   5160
      TabIndex        =   116
      Top             =   5880
      Width           =   375
   End
   Begin VB.Label Label75 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "75"
      Height          =   255
      Left            =   3120
      TabIndex        =   115
      Top             =   5880
      Width           =   375
   End
   Begin VB.Label Label74 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "74"
      Height          =   255
      Left            =   960
      TabIndex        =   114
      Top             =   5880
      Width           =   375
   End
   Begin VB.Label Label73 
      BackStyle       =   0  'Transparent
      Caption         =   "73"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   495
      Left            =   10080
      TabIndex        =   113
      Top             =   5040
      Width           =   975
   End
   Begin VB.Label Label72 
      BackStyle       =   0  'Transparent
      Caption         =   "72"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   495
      Left            =   4320
      TabIndex        =   112
      Top             =   5040
      Width           =   975
   End
   Begin VB.Line Line23 
      BorderColor     =   &H80000009&
      X1              =   11040
      X2              =   11160
      Y1              =   480
      Y2              =   720
   End
   Begin VB.Shape Shape42 
      FillStyle       =   7  'Diagonal Cross
      Height          =   135
      Left            =   11040
      Top             =   720
      Width           =   135
   End
   Begin VB.Shape Shape41 
      FillStyle       =   7  'Diagonal Cross
      Height          =   135
      Left            =   11040
      Top             =   2400
      Width           =   135
   End
   Begin VB.Shape Shape40 
      BorderColor     =   &H000040C0&
      BorderStyle     =   3  'Dot
      Height          =   1575
      Left            =   9240
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label71 
      BackStyle       =   0  'Transparent
      Caption         =   "71"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   495
      Left            =   2280
      TabIndex        =   110
      Top             =   5040
      Width           =   975
   End
   Begin VB.Label Label70 
      BackStyle       =   0  'Transparent
      Caption         =   "70"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   495
      Left            =   6240
      TabIndex        =   109
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label Label69 
      BackStyle       =   0  'Transparent
      Caption         =   "69"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   495
      Left            =   4320
      TabIndex        =   108
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label Label68 
      BackStyle       =   0  'Transparent
      Caption         =   "68"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   495
      Left            =   2280
      TabIndex        =   107
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label Label67 
      BackStyle       =   0  'Transparent
      Caption         =   "67"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   495
      Left            =   10080
      TabIndex        =   106
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label66 
      BackStyle       =   0  'Transparent
      Caption         =   "66"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   495
      Left            =   8040
      TabIndex        =   105
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label65 
      BackStyle       =   0  'Transparent
      Caption         =   "65"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   495
      Left            =   6240
      TabIndex        =   104
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label64 
      BackStyle       =   0  'Transparent
      Caption         =   "64"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   495
      Left            =   4320
      TabIndex        =   103
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label63 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "63"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   495
      Left            =   2280
      TabIndex        =   102
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label62 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "62"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   135
      Left            =   9480
      TabIndex        =   101
      Top             =   4800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Shape Shape39 
      FillColor       =   &H80000000&
      Height          =   495
      Left            =   11250
      Shape           =   3  'Circle
      Top             =   600
      Width           =   615
   End
   Begin VB.Shape Shape38 
      FillColor       =   &H80000000&
      Height          =   495
      Left            =   11280
      Shape           =   3  'Circle
      Top             =   2280
      Width           =   615
   End
   Begin VB.Shape Shape37 
      FillColor       =   &H80000000&
      Height          =   495
      Left            =   7560
      Shape           =   3  'Circle
      Top             =   3840
      Width           =   615
   End
   Begin VB.Shape Shape36 
      FillColor       =   &H80000000&
      Height          =   495
      Left            =   5640
      Shape           =   3  'Circle
      Top             =   5400
      Width           =   615
   End
   Begin VB.Shape Shape35 
      FillColor       =   &H80000000&
      Height          =   495
      Left            =   10800
      Shape           =   3  'Circle
      Top             =   2640
      Width           =   615
   End
   Begin VB.Shape Shape34 
      FillColor       =   &H80000000&
      Height          =   495
      Left            =   8880
      Shape           =   3  'Circle
      Top             =   2640
      Width           =   615
   End
   Begin VB.Shape Shape33 
      FillColor       =   &H80000000&
      Height          =   495
      Left            =   6960
      Shape           =   3  'Circle
      Top             =   4200
      Width           =   615
   End
   Begin VB.Shape Shape32 
      FillColor       =   &H80000000&
      Height          =   495
      Left            =   5040
      Shape           =   3  'Circle
      Top             =   5760
      Width           =   615
   End
   Begin VB.Shape Shape31 
      FillColor       =   &H80000000&
      Height          =   495
      Left            =   3000
      Shape           =   3  'Circle
      Top             =   5760
      Width           =   615
   End
   Begin VB.Shape Shape30 
      FillColor       =   &H80000000&
      Height          =   495
      Left            =   840
      Shape           =   3  'Circle
      Top             =   5760
      Width           =   615
   End
   Begin VB.Label Label61 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "61"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   135
      Left            =   9840
      TabIndex        =   99
      Top             =   4560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label60 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "60"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   135
      Left            =   3600
      TabIndex        =   98
      Top             =   5400
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label59 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "59"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   135
      Left            =   3600
      TabIndex        =   97
      Top             =   4920
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label58 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "58"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   135
      Left            =   3600
      TabIndex        =   96
      Top             =   4080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label57 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "57"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   165
      Left            =   4560
      TabIndex        =   93
      Top             =   4680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label56 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "56"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   135
      Left            =   4080
      TabIndex        =   64
      Top             =   4680
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Shape Shape29 
      BorderColor     =   &H80000014&
      BorderWidth     =   3
      Height          =   405
      Left            =   1080
      Top             =   30
      Width           =   2295
   End
   Begin VB.Label Label55 
      BackStyle       =   0  'Transparent
      Caption         =   "55"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   135
      Left            =   3360
      TabIndex        =   63
      Top             =   4680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label54 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "54"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   135
      Left            =   1560
      TabIndex        =   62
      Top             =   5400
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label53 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "53"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   135
      Left            =   1560
      TabIndex        =   61
      Top             =   4920
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label52 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "52"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   135
      Left            =   1560
      TabIndex        =   60
      Top             =   4080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label51 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "51"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   135
      Left            =   2520
      TabIndex        =   59
      Top             =   4680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label50 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "50"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   135
      Left            =   1920
      TabIndex        =   58
      Top             =   4680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label49 
      BackStyle       =   0  'Transparent
      Caption         =   "49"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   135
      Left            =   1200
      TabIndex        =   57
      Top             =   4680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape Shape28 
      FillColor       =   &H00008000&
      FillStyle       =   7  'Diagonal Cross
      Height          =   135
      Left            =   1080
      Top             =   3960
      Width           =   135
   End
   Begin VB.Shape Shape27 
      FillColor       =   &H00008000&
      FillStyle       =   7  'Diagonal Cross
      Height          =   135
      Left            =   5280
      Top             =   720
      Width           =   135
   End
   Begin VB.Shape Shape26 
      FillColor       =   &H00008000&
      FillStyle       =   7  'Diagonal Cross
      Height          =   135
      Left            =   1080
      Top             =   720
      Width           =   135
   End
   Begin VB.Shape Shape25 
      FillColor       =   &H00008000&
      FillStyle       =   7  'Diagonal Cross
      Height          =   135
      Left            =   7200
      Top             =   3960
      Width           =   135
   End
   Begin VB.Shape Shape24 
      FillColor       =   &H00008000&
      FillStyle       =   7  'Diagonal Cross
      Height          =   135
      Left            =   5280
      Top             =   5520
      Width           =   135
   End
   Begin VB.Shape Shape23 
      FillColor       =   &H00008000&
      FillStyle       =   7  'Diagonal Cross
      Height          =   135
      Left            =   3240
      Top             =   5520
      Width           =   135
   End
   Begin VB.Shape Shape22 
      FillColor       =   &H00008000&
      FillStyle       =   7  'Diagonal Cross
      Height          =   135
      Left            =   5280
      Top             =   3960
      Width           =   135
   End
   Begin VB.Shape Shape21 
      FillColor       =   &H00008000&
      FillStyle       =   7  'Diagonal Cross
      Height          =   135
      Left            =   7200
      Top             =   2400
      Width           =   135
   End
   Begin VB.Shape Shape20 
      FillColor       =   &H00008000&
      FillStyle       =   7  'Diagonal Cross
      Height          =   135
      Left            =   9120
      Top             =   2400
      Width           =   135
   End
   Begin VB.Shape Shape19 
      FillColor       =   &H00008000&
      FillStyle       =   7  'Diagonal Cross
      Height          =   135
      Left            =   1080
      Top             =   5520
      Width           =   135
   End
   Begin VB.Shape Shape18 
      FillStyle       =   7  'Diagonal Cross
      Height          =   135
      Left            =   9120
      Top             =   3960
      Width           =   135
   End
   Begin VB.Shape Shape17 
      FillStyle       =   7  'Diagonal Cross
      Height          =   135
      Left            =   11040
      Top             =   3960
      Width           =   135
   End
   Begin VB.Shape Shape16 
      FillStyle       =   7  'Diagonal Cross
      Height          =   135
      Left            =   11040
      Top             =   5520
      Width           =   135
   End
   Begin VB.Shape Shape15 
      FillStyle       =   7  'Diagonal Cross
      Height          =   135
      Left            =   9120
      Top             =   5520
      Width           =   135
   End
   Begin VB.Shape Shape14 
      FillColor       =   &H00008000&
      FillStyle       =   7  'Diagonal Cross
      Height          =   135
      Left            =   9120
      Top             =   720
      Width           =   135
   End
   Begin VB.Shape Shape13 
      FillColor       =   &H00008000&
      FillStyle       =   7  'Diagonal Cross
      Height          =   135
      Left            =   3240
      Top             =   2400
      Width           =   135
   End
   Begin VB.Line Line22 
      X1              =   1080
      X2              =   11040
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line21 
      X1              =   1080
      X2              =   1080
      Y1              =   720
      Y2              =   5640
   End
   Begin VB.Line Line19 
      X1              =   5400
      X2              =   5400
      Y1              =   4080
      Y2              =   5520
   End
   Begin VB.Line Line18 
      X1              =   7320
      X2              =   7320
      Y1              =   2520
      Y2              =   3960
   End
   Begin VB.Line Line17 
      X1              =   11160
      X2              =   11160
      Y1              =   840
      Y2              =   2400
   End
   Begin VB.Shape Shape12 
      BorderColor     =   &H000040C0&
      BorderStyle     =   3  'Dot
      FillColor       =   &H00C0C0C0&
      Height          =   1455
      Left            =   9240
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Line Line16 
      X1              =   7320
      X2              =   11040
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line15 
      X1              =   5400
      X2              =   7320
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Line Line4 
      X1              =   3480
      X2              =   5400
      Y1              =   5640
      Y2              =   5640
   End
   Begin VB.Line Line3 
      X1              =   1200
      X2              =   3480
      Y1              =   5640
      Y2              =   5640
   End
   Begin VB.Label Label48 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "48"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   135
      Left            =   5760
      TabIndex        =   54
      Top             =   3840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label47 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "47"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   135
      Left            =   5760
      TabIndex        =   53
      Top             =   3360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label46 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "46"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   135
      Left            =   5760
      TabIndex        =   52
      Top             =   2520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label45 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "45"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   135
      Left            =   6600
      TabIndex        =   51
      Top             =   3120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label44 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "44"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   135
      Left            =   6120
      TabIndex        =   50
      Top             =   3120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label43 
      BackStyle       =   0  'Transparent
      Caption         =   "43"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   135
      Left            =   5400
      TabIndex        =   49
      Top             =   3120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label42 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "42"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   135
      Left            =   3600
      TabIndex        =   48
      Top             =   3840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label41 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "41"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   135
      Left            =   3600
      TabIndex        =   47
      Top             =   3360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label40 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "40"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   135
      Left            =   3600
      TabIndex        =   46
      Top             =   2520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label39 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "39"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   135
      Left            =   4680
      TabIndex        =   45
      Top             =   3120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label38 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "38"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   135
      Left            =   4080
      TabIndex        =   44
      Top             =   3120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label37 
      BackStyle       =   0  'Transparent
      Caption         =   "37"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   135
      Left            =   3360
      TabIndex        =   43
      Top             =   3120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000009&
      X1              =   9120
      X2              =   9240
      Y1              =   480
      Y2              =   720
   End
   Begin VB.Shape Shape9 
      BorderColor     =   &H000040C0&
      BorderStyle     =   3  'Dot
      FillColor       =   &H00C0C0C0&
      Height          =   1575
      Left            =   7320
      Top             =   840
      Width           =   1815
   End
   Begin VB.Shape Shape10 
      Height          =   1695
      Left            =   9120
      Top             =   3960
      Width           =   2055
   End
   Begin VB.Label Label36 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "36"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   135
      Left            =   1560
      TabIndex        =   42
      Top             =   3840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label35 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "35"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   135
      Left            =   1560
      TabIndex        =   41
      Top             =   3360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label34 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "34"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   135
      Left            =   1560
      TabIndex        =   40
      Top             =   2520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label33 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "33"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   135
      Left            =   2640
      TabIndex        =   39
      Top             =   3120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label32 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "32"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   135
      Left            =   1920
      TabIndex        =   37
      Top             =   3120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label31 
      BackStyle       =   0  'Transparent
      Caption         =   "31"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   135
      Left            =   1200
      TabIndex        =   36
      Top             =   3120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label30 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "30"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   135
      Left            =   9600
      TabIndex        =   35
      Top             =   2280
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label29 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "29"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   135
      Left            =   9600
      TabIndex        =   34
      Top             =   1680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Line Line14 
      BorderColor     =   &H80000009&
      X1              =   1080
      X2              =   840
      Y1              =   5520
      Y2              =   5640
   End
   Begin VB.Line Line13 
      BorderColor     =   &H80000009&
      X1              =   1080
      X2              =   840
      Y1              =   3960
      Y2              =   4080
   End
   Begin VB.Line Line12 
      BorderColor     =   &H80000009&
      X1              =   1080
      X2              =   840
      Y1              =   2400
      Y2              =   2520
   End
   Begin VB.Line Line11 
      BorderColor     =   &H80000009&
      X1              =   1080
      X2              =   840
      Y1              =   720
      Y2              =   840
   End
   Begin VB.Line Line10 
      BorderColor     =   &H80000009&
      X1              =   7200
      X2              =   7320
      Y1              =   480
      Y2              =   720
   End
   Begin VB.Line Line9 
      BorderColor     =   &H80000009&
      X1              =   5280
      X2              =   5400
      Y1              =   480
      Y2              =   720
   End
   Begin VB.Line Line8 
      BorderColor     =   &H80000009&
      X1              =   3240
      X2              =   3360
      Y1              =   480
      Y2              =   720
   End
   Begin VB.Line Line7 
      BorderColor     =   &H80000009&
      X1              =   1080
      X2              =   1200
      Y1              =   480
      Y2              =   720
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000009&
      X1              =   960
      X2              =   960
      Y1              =   720
      Y2              =   5760
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000009&
      X1              =   840
      X2              =   11280
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Shape Shape8 
      BorderStyle     =   3  'Dot
      FillColor       =   &H00C0C0C0&
      Height          =   1455
      Left            =   5400
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Shape Shape7 
      BorderStyle     =   3  'Dot
      FillColor       =   &H00C0C0C0&
      Height          =   1575
      Left            =   5400
      Top             =   840
      Width           =   1815
   End
   Begin VB.Shape Shape6 
      BorderStyle     =   3  'Dot
      FillColor       =   &H00C0C0C0&
      Height          =   1455
      Left            =   3360
      Top             =   4080
      Width           =   1935
   End
   Begin VB.Shape Shape5 
      BorderStyle     =   3  'Dot
      FillColor       =   &H00C0C0C0&
      Height          =   1455
      Left            =   1200
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H000040C0&
      BorderStyle     =   3  'Dot
      FillColor       =   &H00C0C0C0&
      Height          =   1455
      Left            =   3360
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H000040C0&
      BorderStyle     =   3  'Dot
      FillColor       =   &H00C0C0C0&
      Height          =   1455
      Left            =   1200
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000040C0&
      BorderStyle     =   3  'Dot
      FillColor       =   &H00C0C0C0&
      Height          =   1575
      Left            =   3360
      Top             =   840
      Width           =   1935
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000040C0&
      BorderStyle     =   3  'Dot
      FillColor       =   &H00C0C0C0&
      Height          =   1575
      Left            =   1200
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "28"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   135
      Left            =   9600
      TabIndex        =   27
      Top             =   840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label27 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "27"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   135
      Left            =   10440
      TabIndex        =   26
      Top             =   1440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "26"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   255
      Left            =   9840
      TabIndex        =   25
      Top             =   1440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "25"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   135
      Left            =   9240
      TabIndex        =   24
      Top             =   1440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "24"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   135
      Left            =   7680
      TabIndex        =   23
      Top             =   2280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "23"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   135
      Left            =   7680
      TabIndex        =   22
      Top             =   1680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "22"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   135
      Left            =   7680
      TabIndex        =   21
      Top             =   840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label21 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "21"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   135
      Left            =   8520
      TabIndex        =   20
      Top             =   1440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "20"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   135
      Left            =   8040
      TabIndex        =   19
      Top             =   1440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "19"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   135
      Left            =   7320
      TabIndex        =   18
      Top             =   1440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "18"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   135
      Left            =   5760
      TabIndex        =   17
      Top             =   2280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "17"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   135
      Left            =   5760
      TabIndex        =   16
      Top             =   1680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "16"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   135
      Left            =   5760
      TabIndex        =   15
      Top             =   840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "15"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   135
      Left            =   6600
      TabIndex        =   14
      Top             =   1440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "14"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   135
      Left            =   6120
      TabIndex        =   13
      Top             =   1440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "13"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   135
      Left            =   5400
      TabIndex        =   12
      Top             =   1440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   135
      Left            =   3600
      TabIndex        =   11
      Top             =   2280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "11"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   135
      Left            =   3600
      TabIndex        =   10
      Top             =   1680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   135
      Left            =   3600
      TabIndex        =   9
      Top             =   840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   135
      Left            =   4680
      TabIndex        =   8
      Top             =   1440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   135
      Left            =   4080
      TabIndex        =   7
      Top             =   1440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   135
      Left            =   3360
      TabIndex        =   6
      Top             =   1440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   135
      Left            =   1560
      TabIndex        =   5
      Top             =   2280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   135
      Left            =   1560
      TabIndex        =   4
      Top             =   1680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   135
      Left            =   1560
      TabIndex        =   3
      Top             =   840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   135
      Left            =   2520
      TabIndex        =   2
      Top             =   1440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   135
      Left            =   1680
      TabIndex        =   1
      Top             =   1440
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   135
      Left            =   1200
      TabIndex        =   0
      Top             =   1440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&Klik di sini"
      Begin VB.Menu mnuFileOpenDwg 
         Caption         =   "&OpenDWG"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''THIS COMPUTER PROGRAM (SOFTWARE) IS DEDICATED TO MY LATE FATHER            '''
'''WAN MOHAMED BIN WAN AWANG (1938 - 15.02.02), SEMOGA DIRAHMATI ALLAH.       '''
'''THE PROGRAM IS AN INPUT TEMPLATE FOR STRUCTURAL REINFORCED CONCRETE SLAB,  '''
'''TO INTERFACE INTO AUTOCAD ENVIRONMENT WITH PRE-DESIGNED ALGORITHMS THAT    '''
'''ENABLE TO CONVERT THE INPUT DATA AUTOMATICALLY INTO STRUCT. R.C. DRAWINGS. '''
'''CREATED IN 2001 BY : WAN SOHAIMI BIN WAN MOHAMED.                          '''
'''(LATEST REVISION FEB 2002)- [butiran papak sahaja]                         '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''Permission is hereby granted, free of charge, to any person obtaining a copy  '''
'''of this software and associated documentation files (the "Software"), to deal '''
'''in the Software without restriction, including without limitation the rights  '''
'''to use, copy, modify, merge, publish, distribute, sub-license, and/or sell    '''
'''copies of the Software, and to permit persons to whom the Software is         '''
'''furnished to do so, subject to the following conditions: nil                  '''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
Option Explicit

Dim Bh1, Bh2, Bh3, Bh4, Bh5, Bh6, Bh7, Bh8, Bh9, Bh10, Bh11 As String
Dim Bk1, Bk2, Bk3, Bk4, Bk5, Bk6, Bk7, Bk8, Bk9, Bk10, Bk11 As String
Dim LG1, LG2, LG3, LG4, LG5, LG6, LG7, LG8, LG9, LG10 As String
Dim RB1, RB2, RB3, RB4, RB5, RB6, RB7, RB8, RB9, RB10 As String
Dim RH1, RH2, RH3, RH4, RH5, RH6, RH7, RH8, RH9, RH10 As String
Dim In1, In2, In3, In4, In5, In6, In7, In8 As String
Dim TP1, TP2, TP3, TP4, TP5, TP6, TP7, TP8, TP9, TP10, TP11 As String

Dim Coef1, Coef2, Coef3, Coef4, Coef5, Coef6, Lx1 As Double
Dim Coef7, Coef8, Coef9, Coef10, Coef11, Coef12, Lx2 As Double
Dim Coef13, Coef14, Coef15, Coef16, Coef17, Coef18, Lx3 As Double
Dim Coef19, Coef20, Coef21, Coef22, Coef23, Coef24, Lx4 As Double
Dim Coef25, Coef26, Coef27, Coef28, Coef29, Coef30, Lx5 As Double

Dim Coef31, Coef32, Coef33, Coef34, Coef35, Coef36, Lx6 As Double
Dim Coef37, Coef38, Coef39, Coef40, Coef41, Coef42, Lx7 As Double
Dim Coef43, Coef44, Coef45, Coef46, Coef47, Coef48, Lx8 As Double
Dim Coef49, Coef50, Coef51, Coef52, Coef53, Coef54, Lx9 As Double
Dim Coef55, Coef56, Coef57, Coef58, Coef59, Coef60, Lx10 As Double
Dim Coef61, Coef62, Lx11 As Double

Dim BHidupNew As New BHidup
Dim BKemasNew As New BKemas
Dim LabelNew As New Label
Dim Rasuk_BNew As New Rasuk_B
Dim Rasuk_HNew As New Rasuk_H
Dim RentangNew As New Rentang
Dim TPapakNew As New TPapak
Dim DsgPeriNew As New DsgPerimeters

Dim DFCU, DFYV, DDIA, DCVR As Double
Dim DBMK As Integer

Public Xinsertion, Yinsertion As Double
Public DataFile, NamaFolder As String
Public dwgName As String
Public acadApp As Object
Public acadDoc As Object
Public moSpace As Object
Public paSpace As Object


''''''''''''''''''''''''''''
Private Sub StartAutoCAD()
Form1.Picture = LoadPicture("C:\autodraf\icon\statacad.ico")
  
On Error Resume Next

Set acadApp = GetObject(, "AutoCAD.Application")

If Err Then
   Set acadApp = CreateObject("AutoCAD.Application")
'   MsgBox "Set acadApp = CreateObject", , "Check setting:"
    Err.Clear
            If Err Then
              MsgBox Err.Description
              Exit Sub
            End If
End If


''''Set OSNAP mode for duration of the VB program.
'sysVarName = "OSMODE"
'sysVarData = acadDoc.GetVariable(sysVarName)
'osMode = CInt(sysVarData)
'acadDoc.SetVariable sysVarName, 1

''''Set SDI mode for duration of the VB program.
'sysVarName = "SDI"
'sysVarData = acadDoc.GetVariable(sysVarName)
'sdiMode = CInt(sysVarData)
'acadDoc.SetVariable sysVarName, 1

Set acadDoc = acadApp.ActiveDocument

If acadDoc.FullName <> dwgName Then
   acadDoc.Open dwgName
End If

acadApp.Visible = True
acadApp.Top = 0
acadApp.Left = 0
acadApp.Width = 1000
acadApp.Height = 900
Form1.Picture = LoadPicture("C:\autodraf\icon\ukad4.ico")

End Sub

'''''''''''''''''''''''''''''''''''''''''''''

Private Sub PaneL1()
Dim LyOnLx As Double

LyOnLx = Val(In1) / Val(In6)
If Val(In1) < Val(In6) Then
   LyOnLx = 1 / LyOnLx
   End If
If LyOnLx > 2 Then
   LyOnLx = 2
   End If
   
 If Val(In1) >= Val(In6) Then
    Label1.Caption = 0
    Label2.Caption = 0.034
    Label3.Caption = 0.045
    Label4.Caption = 0
    Label5.Caption = "0" & Str(Int(1000 * TwoAdjEDiscoSag(LyOnLx)) / 1000)
    Label6.Caption = "0" & Str(Int(1000 * TwoAdjEDiscoHog(LyOnLx)) / 1000)
    Label63.Caption = "v" & Str(Int(10 * LyOnLx) / 10)
    Coef1 = Val(Label1.Caption)
    Coef2 = Val(Label2.Caption)
    Coef3 = Val(Label3.Caption)
    Coef4 = Val(Label4.Caption)
    Coef5 = Val(Label5.Caption)
    Coef6 = Val(Label6.Caption)
    Lx1 = Val(In6)
          Else
    Label1.Caption = 0
    Label2.Caption = "0" & Str(Int(1000 * TwoAdjEDiscoSag(LyOnLx)) / 1000)
    Label3.Caption = "0" & Str(Int(1000 * TwoAdjEDiscoHog(LyOnLx)) / 1000)
    Label4.Caption = 0
    Label5.Caption = 0.034
    Label6.Caption = 0.045
    Label63.Caption = ">" & Str(Int(10 * LyOnLx) / 10)
    Coef1 = Val(Label1.Caption)
    Coef2 = Val(Label2.Caption)
    Coef3 = Val(Label3.Caption)
    Coef4 = Val(Label4.Caption)
    Coef5 = Val(Label5.Caption)
    Coef6 = Val(Label6.Caption)
    Lx1 = Val(In1)
    
        End If
End Sub

Private Sub PaneL2()
Dim LyOnLx As Double

LyOnLx = Val(In2) / Val(In6)
If Val(In2) < Val(In6) Then
   LyOnLx = 1 / LyOnLx
   End If
If LyOnLx > 2 Then
   LyOnLx = 2
   End If
   
 If Val(In2) >= Val(In6) Then
    Label7.Caption = 0.037
    Label8.Caption = 0.028
    Label9.Caption = 0.037
    Label10.Caption = 0
    Label11.Caption = "0" & Str(Int(1000 * OneLongEDiscoSag(LyOnLx)) / 1000)
    Label12.Caption = "0" & Str(Int(1000 * OneLongEDiscoHog(LyOnLx)) / 1000)
    Label64.Caption = "v" & Str(Int(10 * LyOnLx) / 10)
    Coef7 = Val(Label7.Caption)
    Coef8 = Val(Label8.Caption)
    Coef9 = Val(Label9.Caption)
    Coef10 = Val(Label10.Caption)
    Coef11 = Val(Label11.Caption)
    Coef12 = Val(Label12.Caption)
    Lx2 = Val(In6)
           Else
    Label7.Caption = "0" & Str(Int(1000 * OneShortEDiscoHog(LyOnLx)) / 1000)
    Label8.Caption = "0" & Str(Int(1000 * OneShortEDiscoSag(LyOnLx)) / 1000)
    Label9.Caption = "0" & Str(Int(1000 * OneShortEDiscoHog(LyOnLx)) / 1000)
    Label10.Caption = 0
    Label11.Caption = 0.028
    Label12.Caption = 0.037
    Label64.Caption = ">" & Str(Int(10 * LyOnLx) / 10)
    Coef7 = Val(Label7.Caption)
    Coef8 = Val(Label8.Caption)
    Coef9 = Val(Label9.Caption)
    Coef10 = Val(Label10.Caption)
    Coef11 = Val(Label11.Caption)
    Coef12 = Val(Label12.Caption)
    Lx2 = Val(In2)
        End If
End Sub
Private Sub PaneL3()
Dim LyOnLx As Double

LyOnLx = Val(In3) / Val(In6)
If Val(In3) < Val(In6) Then
   LyOnLx = 1 / LyOnLx
   End If
If LyOnLx > 2 Then
   LyOnLx = 2
   End If
   
 If Val(In3) >= Val(In6) Then
    Label13.Caption = 0.037
    Label14.Caption = 0.028
    Label15.Caption = 0.037
    Label16.Caption = 0
    Label17.Caption = "0" & Str(Int(1000 * OneLongEDiscoSag(LyOnLx)) / 1000)
    Label18.Caption = "0" & Str(Int(1000 * OneLongEDiscoHog(LyOnLx)) / 1000)
    Label65.Caption = "v" & Str(Int(10 * LyOnLx) / 10)
    Coef13 = Val(Label13.Caption)
    Coef14 = Val(Label14.Caption)
    Coef15 = Val(Label15.Caption)
    Coef16 = Val(Label16.Caption)
    Coef17 = Val(Label17.Caption)
    Coef18 = Val(Label18.Caption)
    Lx3 = Val(In6)
           Else
    Label13.Caption = "0" & Str(Int(1000 * OneShortEDiscoHog(LyOnLx)) / 1000)
    Label14.Caption = "0" & Str(Int(1000 * OneShortEDiscoSag(LyOnLx)) / 1000)
    Label15.Caption = "0" & Str(Int(1000 * OneShortEDiscoHog(LyOnLx)) / 1000)
    Label16.Caption = 0
    Label17.Caption = 0.028
    Label18.Caption = 0.037
    Label65.Caption = ">" & Str(Int(10 * LyOnLx) / 10)
    Coef13 = Val(Label13.Caption)
    Coef14 = Val(Label14.Caption)
    Coef15 = Val(Label15.Caption)
    Coef16 = Val(Label16.Caption)
    Coef17 = Val(Label17.Caption)
    Coef18 = Val(Label18.Caption)
    Lx3 = Val(In3)
        End If
End Sub
Private Sub PaneL4()
Dim LyOnLx As Double

LyOnLx = Val(In4) / Val(In6)
If Val(In4) < Val(In6) Then
   LyOnLx = 1 / LyOnLx
   End If
If LyOnLx > 2 Then
   LyOnLx = 2
   End If
   
 If Val(In4) >= Val(In6) Then
    Label19.Caption = 0.045
    Label20.Caption = 0.034
    Label21.Caption = 0.045
    Label22.Caption = 0
    Label23.Caption = "0" & Str(Int(1000 * OneShortEDiscoHog(LyOnLx)) / 1000)
    Label24.Caption = 0
    Label66.Caption = "v" & Str(Int(10 * LyOnLx) / 10)
    Coef19 = Val(Label19.Caption)
    Coef20 = Val(Label20.Caption)
    Coef21 = Val(Label21.Caption)
    Coef22 = Val(Label22.Caption)
    Coef23 = Val(Label23.Caption)
    Coef24 = Val(Label24.Caption)
    Lx4 = Val(In6)
          Else
    Label19.Caption = "0" & Str(Int(1000 * TwoShortEDiscoHog(LyOnLx)) / 1000)
    Label20.Caption = "0" & Str(Int(1000 * TwoShortEDiscoSag(LyOnLx)) / 1000)
    Label21.Caption = "0" & Str(Int(1000 * TwoShortEDiscoHog(LyOnLx)) / 1000)
    Label22.Caption = 0
    Label23.Caption = 0.034
    Label24.Caption = 0
    Label66.Caption = ">" & Str(Int(10 * LyOnLx) / 10)
    Coef19 = Val(Label19.Caption)
    Coef20 = Val(Label20.Caption)
    Coef21 = Val(Label21.Caption)
    Coef22 = Val(Label22.Caption)
    Coef23 = Val(Label23.Caption)
    Coef24 = Val(Label24.Caption)
    Lx4 = Val(In4)
        End If
End Sub

Private Sub PaneL5()
Dim LyOnLx As Double

LyOnLx = Val(In5) / Val(In6)
If Val(In5) < Val(In6) Then
   LyOnLx = 1 / LyOnLx
   End If
If LyOnLx > 2 Then
   LyOnLx = 2
   End If
   
 If Val(In5) >= Val(In6) Then
    Label25.Caption = 0.058
    Label26.Caption = 0.044
    Label27.Caption = 0
    Label28.Caption = 0
    Label29.Caption = "0" & Str(Int(1000 * ThreeEShortDiscoSag(LyOnLx)) / 1000)
    Label30.Caption = 0
    Label67.Caption = "v" & Str(Int(10 * LyOnLx) / 10)
    Coef25 = Val(Label25.Caption)
    Coef26 = Val(Label26.Caption)
    Coef27 = Val(Label27.Caption)
    Coef28 = Val(Label28.Caption)
    Coef29 = Val(Label29.Caption)
    Coef30 = Val(Label30.Caption)
    Lx5 = Val(In6)
           Else
    Label25.Caption = "0" & Str(Int(1000 * ThreeELongDiscoHog(LyOnLx)) / 1000)
    Label26.Caption = "0" & Str(Int(1000 * ThreeELongDiscoSag(LyOnLx)) / 1000)
    Label27.Caption = 0
    Label28.Caption = 0
    Label29.Caption = 0.044
    Label30.Caption = 0
    Label67.Caption = ">" & Str(Int(10 * LyOnLx) / 10)
    Coef25 = Val(Label25.Caption)
    Coef26 = Val(Label26.Caption)
    Coef27 = Val(Label27.Caption)
    Coef28 = Val(Label28.Caption)
    Coef29 = Val(Label29.Caption)
    Coef30 = Val(Label30.Caption)
    Lx5 = Val(In5)
        End If
End Sub

Private Sub PaneL6()
Dim LyOnLx As Double

LyOnLx = Val(In1) / Val(In7)
If Val(In1) < Val(In7) Then
   LyOnLx = 1 / LyOnLx
   End If
If LyOnLx > 2 Then
   LyOnLx = 2
   End If
   
 If Val(In1) >= Val(In7) Then
    Label31.Caption = 0
    Label32.Caption = 0.028
    Label33.Caption = 0.037
    Label34.Caption = "0" & Str(Int(1000 * OneLongEDiscoHog(LyOnLx)) / 1000)
    Label35.Caption = "0" & Str(Int(1000 * OneLongEDiscoSag(LyOnLx)) / 1000)
    Label36.Caption = "0" & Str(Int(1000 * OneLongEDiscoHog(LyOnLx)) / 1000)
    Label68.Caption = "v" & Str(Int(10 * LyOnLx) / 10)
    Coef31 = Val(Label31.Caption)
    Coef32 = Val(Label32.Caption)
    Coef33 = Val(Label33.Caption)
    Coef34 = Val(Label34.Caption)
    Coef35 = Val(Label35.Caption)
    Coef36 = Val(Label36.Caption)
    Lx6 = Val(In7)
           Else
    Label31.Caption = 0
    Label32.Caption = "0" & Str(Int(1000 * OneShortEDiscoSag(LyOnLx)) / 1000)
    Label33.Caption = "0" & Str(Int(1000 * OneShortEDiscoHog(LyOnLx)) / 1000)
    Label34.Caption = 0.037
    Label35.Caption = 0.028
    Label36.Caption = 0.037
    Label68.Caption = ">" & Str(Int(10 * LyOnLx) / 10)
    Coef31 = Val(Label31.Caption)
    Coef32 = Val(Label32.Caption)
    Coef33 = Val(Label33.Caption)
    Coef34 = Val(Label34.Caption)
    Coef35 = Val(Label35.Caption)
    Coef36 = Val(Label36.Caption)
    Lx6 = Val(In1)
        End If
End Sub

Private Sub PaneL7()
Dim LyOnLx As Double

LyOnLx = Val(In2) / Val(In7)
If Val(In2) < Val(In7) Then
   LyOnLx = 1 / LyOnLx
   End If
If LyOnLx > 2 Then
   LyOnLx = 2
   End If
   
 If Val(In2) >= Val(In7) Then
    Label37.Caption = 0.032
    Label38.Caption = 0.024
    Label39.Caption = 0.032
    Label40.Caption = "0" & Str(Int(1000 * FourEContinHog(LyOnLx)) / 1000)
    Label41.Caption = "0" & Str(Int(1000 * FourEContinSag(LyOnLx)) / 1000)
    Label42.Caption = "0" & Str(Int(1000 * FourEContinHog(LyOnLx)) / 1000)
    Label69.Caption = "v" & Str(Int(10 * LyOnLx) / 10)
    Coef37 = Val(Label37.Caption)
    Coef38 = Val(Label38.Caption)
    Coef39 = Val(Label39.Caption)
    Coef40 = Val(Label40.Caption)
    Coef41 = Val(Label41.Caption)
    Coef42 = Val(Label42.Caption)
    Lx7 = Val(In7)
           Else
    Label37.Caption = "0" & Str(Int(1000 * FourEContinHog(LyOnLx)) / 1000)
    Label38.Caption = "0" & Str(Int(1000 * FourEContinSag(LyOnLx)) / 1000)
    Label39.Caption = "0" & Str(Int(1000 * FourEContinHog(LyOnLx)) / 1000)
    Label40.Caption = 0.032
    Label41.Caption = 0.024
    Label42.Caption = 0.032
    Label69.Caption = ">" & Str(Int(10 * LyOnLx) / 10)
    Coef37 = Val(Label37.Caption)
    Coef38 = Val(Label38.Caption)
    Coef39 = Val(Label39.Caption)
    Coef40 = Val(Label40.Caption)
    Coef41 = Val(Label41.Caption)
    Coef42 = Val(Label42.Caption)
    Lx7 = Val(In2)
        End If
End Sub

Private Sub PaneL8()
Dim LyOnLx As Double

LyOnLx = Val(In3) / Val(In7)
If Val(In3) < Val(In7) Then
   LyOnLx = 1 / LyOnLx
   End If
If LyOnLx > 2 Then
   LyOnLx = 2
   End If
   
 If Val(In3) >= Val(In7) Then
    Label43.Caption = 0.045
    Label44.Caption = 0.034
    Label45.Caption = 0
    Label46.Caption = "0" & Str(Int(1000 * TwoAdjEDiscoHog(LyOnLx)) / 1000)
    Label47.Caption = "0" & Str(Int(1000 * TwoAdjEDiscoSag(LyOnLx)) / 1000)
    Label48.Caption = 0
    Label70.Caption = "v" & Str(Int(10 * LyOnLx) / 10)
    Coef43 = Val(Label43.Caption)
    Coef44 = Val(Label44.Caption)
    Coef45 = Val(Label45.Caption)
    Coef46 = Val(Label46.Caption)
    Coef47 = Val(Label47.Caption)
    Coef48 = Val(Label48.Caption)
    Lx8 = Val(In7)
           Else
    Label43.Caption = "0" & Str(Int(1000 * TwoAdjEDiscoHog(LyOnLx)) / 1000)
    Label44.Caption = "0" & Str(Int(1000 * TwoAdjEDiscoSag(LyOnLx)) / 1000)
    Label45.Caption = 0
    Label46.Caption = 0.045
    Label47.Caption = 0.034
    Label48.Caption = 0
    Label70.Caption = ">" & Str(Int(10 * LyOnLx) / 10)
    Coef43 = Val(Label43.Caption)
    Coef44 = Val(Label44.Caption)
    Coef45 = Val(Label45.Caption)
    Coef46 = Val(Label46.Caption)
    Coef47 = Val(Label47.Caption)
    Coef48 = Val(Label48.Caption)
    Lx8 = Val(In3)
        End If
End Sub

Private Sub PaneL9()
Dim LyOnLx As Double

LyOnLx = Val(In1) / Val(In8)
If Val(In1) < Val(In8) Then
   LyOnLx = 1 / LyOnLx
   End If
If LyOnLx > 2 Then
   LyOnLx = 2
   End If
   
 If Val(In1) >= Val(In8) Then
    Label49.Caption = 0
    Label50.Caption = 0.034
    Label51.Caption = 0.045
    Label52.Caption = "0" & Str(Int(1000 * TwoAdjEDiscoHog(LyOnLx)) / 1000)
    Label53.Caption = "0" & Str(Int(1000 * TwoAdjEDiscoSag(LyOnLx)) / 1000)
    Label54.Caption = 0
    Label71.Caption = "v" & Str(Int(10 * LyOnLx) / 10)
    Coef49 = Val(Label49.Caption)
    Coef50 = Val(Label50.Caption)
    Coef51 = Val(Label51.Caption)
    Coef52 = Val(Label52.Caption)
    Coef53 = Val(Label53.Caption)
    Coef54 = Val(Label54.Caption)
    Lx9 = Val(In8)
           Else
    Label49.Caption = 0
    Label50.Caption = "0" & Str(Int(1000 * TwoAdjEDiscoSag(LyOnLx)) / 1000)
    Label51.Caption = "0" & Str(Int(1000 * TwoAdjEDiscoHog(LyOnLx)) / 1000)
    Label52.Caption = 0.045
    Label53.Caption = 0.034
    Label54.Caption = 0
    Label71.Caption = ">" & Str(Int(10 * LyOnLx) / 10)
    Coef49 = Val(Label49.Caption)
    Coef50 = Val(Label50.Caption)
    Coef51 = Val(Label51.Caption)
    Coef52 = Val(Label52.Caption)
    Coef53 = Val(Label53.Caption)
    Coef54 = Val(Label54.Caption)
    Lx9 = Val(In1)
        End If
End Sub

Private Sub PaneL10()
Dim LyOnLx As Double

LyOnLx = Val(In2) / Val(In8)
If Val(In2) < Val(In8) Then
   LyOnLx = 1 / LyOnLx
   End If
If LyOnLx > 2 Then
   LyOnLx = 2
   End If
   
   
 If Val(In2) >= Val(In8) Then
    Label55.Caption = 0.045
    Label56.Caption = 0.034
    Label57.Caption = 0
    Label58.Caption = "0" & Str(Int(1000 * TwoAdjEDiscoHog(LyOnLx)) / 1000)
    Label59.Caption = "0" & Str(Int(1000 * TwoAdjEDiscoSag(LyOnLx)) / 1000)
    Label60.Caption = 0
    Label72.Caption = "v" & Str(Int(10 * LyOnLx) / 10)
    Coef55 = Val(Label55.Caption)
    Coef56 = Val(Label56.Caption)
    Coef57 = Val(Label57.Caption)
    Coef58 = Val(Label58.Caption)
    Coef59 = Val(Label59.Caption)
    Coef60 = Val(Label60.Caption)
    Lx10 = Val(In8)
           Else
    Label55.Caption = "0" & Str(Int(1000 * TwoAdjEDiscoHog(LyOnLx)) / 1000)
    Label56.Caption = "0" & Str(Int(1000 * TwoAdjEDiscoSag(LyOnLx)) / 1000)
    Label57.Caption = 0
    Label58.Caption = 0.045
    Label59.Caption = 0.034
    Label60.Caption = 0
    Label72.Caption = ">" & Str(Int(10 * LyOnLx) / 10)
    Coef55 = Val(Label55.Caption)
    Coef56 = Val(Label56.Caption)
    Coef57 = Val(Label57.Caption)
    Coef58 = Val(Label58.Caption)
    Coef59 = Val(Label59.Caption)
    Coef60 = Val(Label60.Caption)
    Lx10 = Val(In2)
        End If
End Sub

Private Sub PaneL11()
Dim LyOnLx As Double

LyOnLx = Val(In5) / Val(In8)
If Val(In5) < Val(In8) Then
   LyOnLx = 1 / LyOnLx
   End If
If LyOnLx > 2 Then
   LyOnLx = 2
   End If
   
   
 If Val(In5) >= Val(In8) Then
    Label61.Caption = 0.056
    Label62.Caption = "0" & Str(Int(1000 * FourEdgesDiscoSag(LyOnLx)) / 1000)
    Label73.Caption = "v" & Str(Int(10 * LyOnLx) / 10)
    Coef61 = Val(Label61.Caption)
    Coef62 = Val(Label62.Caption)
    Lx11 = Val(In8)
           Else
    Label61.Caption = "0" & Str(Int(1000 * FourEdgesDiscoSag(LyOnLx)) / 1000)
    Label62.Caption = 0.056
    Label73.Caption = ">" & Str(Int(10 * LyOnLx) / 10)
    Coef61 = Val(Label61.Caption)
    Coef62 = Val(Label62.Caption)
    Lx11 = Val(In5)
        End If
End Sub

Private Sub Command1_Click()
Label87.Caption = "Minimun Ast = 0.13% -> 0.24%bh."
HighlightPanel
Command11.Enabled = False
Command12.Enabled = False
Command13.Enabled = False

OpenDataFileLabel_Grid
HighlightGridInput
HighlightCircleGrid
HideRentang
HideCoefficient
HideBeban
HideLabel_Beban
HideLabel_Grid
HighlightRentang
DisableTxtOnetoEight

Text20.Text = LG1
Text21.Text = LG2
Text22.Text = LG3
Text23.Text = LG4
Text24.Text = LG5
Text25.Text = LG6
Text26.Text = LG7
Text27.Text = LG8
Text28.Text = LG9
Text29.Text = LG10
Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
Command7.Enabled = False
Command10.Enabled = True
Command14.Enabled = False

Shape30.Shape = 3
Shape31.Shape = 3
Shape32.Shape = 3
Shape33.Shape = 3
Shape34.Shape = 3
Shape35.Shape = 3
Shape36.Shape = 3
Shape37.Shape = 3
Shape38.Shape = 3
Shape39.Shape = 3

Shape30.FillColor = vbWhite
Shape31.FillColor = vbWhite
Shape32.FillColor = vbWhite
Shape33.FillColor = vbWhite
Shape34.FillColor = vbWhite
Shape35.FillColor = vbWhite
Shape36.FillColor = vbWhite
Shape37.FillColor = vbWhite
Shape38.FillColor = vbWhite
Shape39.FillColor = vbWhite


Dim fnum As Integer
Dim txtFile, Temp As String
fnum = FreeFile
txtFile = "C:\autodraf\rasuk\SpanOneGET.txt"

LabelNew.Value20 = Text20.Text
LabelNew.Value21 = Text21.Text
LabelNew.Value22 = Text22.Text
LabelNew.Value23 = Text23.Text
LabelNew.Value24 = Text24.Text
LabelNew.Value25 = Text25.Text
LabelNew.Value26 = Text26.Text
LabelNew.Value27 = Text27.Text
LabelNew.Value28 = Text28.Text
LabelNew.Value29 = Text29.Text

Text20.Text = LabelNew.G_ONE_AtoE
Text21.Text = LabelNew.G_TWO_AtoE
Text22.Text = LabelNew.G_THREE_BtoE
Text23.Text = LabelNew.G_FOUR_CtoE
Text24.Text = LabelNew.G_FIVE_DtoE
Text25.Text = LabelNew.G_A_ONEtoTWO
Text26.Text = LabelNew.G_B_ONEtoTHREE
Text27.Text = LabelNew.G_C_ONEtoFOUR
Text28.Text = LabelNew.G_D_ONEtoFIVE
Text29.Text = LabelNew.G_E_ONEtoFIVE

End Sub

Private Sub Command10_Click()
Label87.Caption = "Minimun Ast = 0.13% -> 0.24%bh."
HighlightPanel
Form1.Picture = LoadPicture(NamaFolder & "icon\datas.ico")
Dim fnum As Integer
Dim txtFile, Temp As String
fnum = FreeFile

Text1.BackColor = &H8000000A
Text2.BackColor = &H8000000A
Text3.BackColor = &H8000000A
Text4.BackColor = &H8000000A
Text5.BackColor = &H8000000A
Text6.BackColor = &H8000000A
Text7.BackColor = &H8000000A
Text8.BackColor = &H8000000A

Shape30.FillColor = &H8000000A
Shape31.FillColor = &H8000000A
Shape32.FillColor = &H8000000A
Shape33.FillColor = &H8000000A
Shape34.FillColor = &H8000000A
Shape35.FillColor = &H8000000A
Shape36.FillColor = &H8000000A
Shape37.FillColor = &H8000000A
Shape38.FillColor = &H8000000A
Shape39.FillColor = &H8000000A



If Command7.Enabled = True Then
Bh1 = Val(Text9.Text)
Bh2 = Val(Text10.Text)
Bh3 = Val(Text11.Text)
Bh4 = Val(Text12.Text)
Bh5 = Val(Text13.Text)
Bh6 = Val(Text14.Text)
Bh7 = Val(Text15.Text)
Bh8 = Val(Text16.Text)
Bh9 = Val(Text17.Text)
Bh10 = Val(Text18.Text)
Bh11 = Val(Text19.Text)
txtFile = NamaFolder & "papak\data_papak\Beban_Hidup.txt"
Open txtFile For Output As #fnum
Print #fnum, Bh1, Bh2, Bh3, Bh4, Bh5, Bh6, Bh7, Bh8, _
             Bh9, Bh10, Bh11
Close #fnum
Label63.Caption = "beban hidup/"
Label64.Caption = "beban hidup/"
Label65.Caption = "beban hidup/"
Label66.Caption = "beban hidup/"
Label67.Caption = "beban hidup/"
Label68.Caption = "beban hidup/"
Label69.Caption = "beban hidup/"
Label70.Caption = "beban hidup/"
Label71.Caption = "beban hidup/"
Label72.Caption = "beban hidup/"
Label73.Caption = "beban hidup/"
End If

If Command6.Enabled = True Then
Bk1 = Val(Text9.Text)
Bk2 = Val(Text10.Text)
Bk3 = Val(Text11.Text)
Bk4 = Val(Text12.Text)
Bk5 = Val(Text13.Text)
Bk6 = Val(Text14.Text)
Bk7 = Val(Text15.Text)
Bk8 = Val(Text16.Text)
Bk9 = Val(Text17.Text)
Bk10 = Val(Text18.Text)
Bk11 = Val(Text19.Text)
txtFile = NamaFolder & "papak\data_papak\Beban_Kemasan.txt"
Open txtFile For Output As #fnum
Print #fnum, Bk1, Bk2, Bk3, Bk4, Bk5, Bk6, Bk7, Bk8, _
             Bk9, Bk10, Bk11
Close #fnum
Label63.Caption = "kemasan/"
Label64.Caption = "kemasan/"
Label65.Caption = "kemasan/"
Label66.Caption = "kemasan/"
Label67.Caption = "kemasan/"
Label68.Caption = "kemasan/"
Label69.Caption = "kemasan/"
Label70.Caption = "kemasan/"
Label71.Caption = "kemasan/"
Label72.Caption = "kemasan/"
Label73.Caption = "kemasan/"
End If

If Command1.Enabled = True Then
LG1 = Text20.Text
LG2 = Text21.Text
LG3 = Text22.Text
LG4 = Text23.Text
LG5 = Text24.Text
LG6 = Text25.Text
LG7 = Text26.Text
LG8 = Text27.Text
LG9 = Text28.Text
LG10 = Text29.Text
txtFile = NamaFolder & "papak\data_papak\Label_Grid.txt"
Open txtFile For Output As #fnum
Print #fnum, LG1
Print #fnum, LG2
Print #fnum, LG3
Print #fnum, LG4
Print #fnum, LG5
Print #fnum, LG6

Print #fnum, LG7
Print #fnum, LG8
Print #fnum, LG9
Print #fnum, LG10
Close #fnum

End If

If Command3.Enabled = True Then
RB1 = Val(Text20.Text)
RB2 = Val(Text21.Text)
RB3 = Val(Text22.Text)
RB4 = Val(Text23.Text)
RB5 = Val(Text24.Text)
RB6 = Val(Text25.Text)
RB7 = Val(Text26.Text)
RB8 = Val(Text27.Text)
RB9 = Val(Text28.Text)
RB10 = Val(Text29.Text)
txtFile = NamaFolder & "papak\data_papak\Ukur_Rasuk_B.txt"
Open txtFile For Output As #fnum
Print #fnum, RB1, RB2, RB3, RB4, RB5, RB6, RB7, RB8, RB9, RB10
Close #fnum

End If

If Command4.Enabled = True Then
RH1 = Val(Text20.Text)
RH2 = Val(Text21.Text)
RH3 = Val(Text22.Text)
RH4 = Val(Text23.Text)
RH5 = Val(Text24.Text)
RH6 = Val(Text25.Text)
RH7 = Val(Text26.Text)
RH8 = Val(Text27.Text)
RH9 = Val(Text28.Text)
RH10 = Val(Text29.Text)
txtFile = NamaFolder & "papak\data_papak\Ukur_Rasuk_H.txt"
Open txtFile For Output As #fnum
Print #fnum, RH1, RH2, RH3, RH4, RH5, RH6, RH7, RH8, RH9, RH10
Close #fnum

End If

If Command2.Enabled = True Then

Command11.Enabled = True
Command12.Enabled = True
Command13.Enabled = True

In1 = Val(Text1.Text)
In2 = Val(Text2.Text)
In3 = Val(Text3.Text)
In4 = Val(Text4.Text)
In5 = Val(Text5.Text)
In6 = Val(Text6.Text)
In7 = Val(Text7.Text)
In8 = Val(Text8.Text)
txtFile = NamaFolder & "papak\data_papak\Ukur_Rentang.txt"
Open txtFile For Output As #fnum
Print #fnum, In1, In2, In3, In4, In5, In6, In7, In8
Close #fnum

End If

If Command5.Enabled = True Then
TP1 = Val(Text9.Text)
TP2 = Val(Text10.Text)
TP3 = Val(Text11.Text)
TP4 = Val(Text12.Text)
TP5 = Val(Text13.Text)
TP6 = Val(Text14.Text)
TP7 = Val(Text15.Text)
TP8 = Val(Text16.Text)
TP9 = Val(Text17.Text)
TP10 = Val(Text18.Text)
TP11 = Val(Text19.Text)
txtFile = NamaFolder & "papak\data_papak\Tebal_Papak.txt"
Open txtFile For Output As #fnum
Print #fnum, TP1, TP2, TP3, TP4, TP5, TP6, TP7, TP8, _
             TP9, TP10, TP11
Close #fnum
Label63.Caption = "tebal papak/"
Label64.Caption = "tebal papak/"
Label65.Caption = "tebal papak/"
Label66.Caption = "tebal papak/"
Label67.Caption = "tebal papak/"
Label68.Caption = "tebal papak/"
Label69.Caption = "tebal papak/"
Label70.Caption = "tebal papak/"
Label71.Caption = "tebal papak/"
Label72.Caption = "tebal papak/"
Label73.Caption = "tebal papak/"
End If


If Command14.Enabled = True Then
Command11.Enabled = True
Command12.Enabled = True
Command13.Enabled = True

DFCU = Val(Text9.Text)
DFYV = Val(Text10.Text)
DDIA = Val(Text11.Text)
DCVR = Val(Text12.Text)
DBMK = Val(Text13.Text)
txtFile = NamaFolder & "papak\data_papak\Dsg_Perimeters.txt"
Open txtFile For Output As #fnum
Print #fnum, DFCU, DFYV, DDIA, DCVR, DBMK
Close #fnum
Label63.Caption = " fcu _/"
Label64.Caption = " fy _/"
Label65.Caption = "bar dia _/"
Label66.Caption = "conc cvr _/"
Label67.Caption = "barmark _/"
Label68.Caption = "-"
Label69.Caption = "-"
Label70.Caption = "-"
Label71.Caption = "-"
Label72.Caption = "-"
Label73.Caption = "-"
End If


Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
Command7.Enabled = True
Command10.Enabled = False
Command14.Enabled = True

Label84.Caption = "PAPAK MS1195: fy = " & Str(DFYV)
End Sub

Private Sub Command11_Click()
''''''''''''''''''''''''''''''''
DefaultLocLabel63to73
Command10.Enabled = False
BoldLabel63to73
Form1.Picture = LoadPicture(NamaFolder & "icon\datam.ico")
Dim fnum As Integer
Dim txtFile, Temp As String
fnum = FreeFile

txtFile = NamaFolder & "papak\data_papak\Beban_Hidup.txt"
Open txtFile For Input As #fnum
Input #fnum, Bh1, Bh2, Bh3, Bh4, Bh5, Bh6, Bh7, Bh8, _
             Bh9, Bh10, Bh11
Close #fnum
'''''''''''''''''''''''''''''''''''''''''
txtFile = NamaFolder & "papak\data_papak\Beban_Kemasan.txt"
Open txtFile For Input As #fnum
Input #fnum, Bk1, Bk2, Bk3, Bk4, Bk5, Bk6, Bk7, Bk8, _
             Bk9, Bk10, Bk11
Close #fnum
''''''''''''''''''''''''''''''''''''''''''
txtFile = NamaFolder & "papak\data_papak\Label_Grid.txt"
Open txtFile For Input As #fnum
Input #fnum, LG1
Input #fnum, LG2
Input #fnum, LG3
Input #fnum, LG4
Input #fnum, LG5
Input #fnum, LG6

Input #fnum, LG7
Input #fnum, LG8
Input #fnum, LG9
Input #fnum, LG10
Close #fnum
''''''''''''''''''''''''''''''''''''''''''
txtFile = NamaFolder & "papak\data_papak\Ukur_Rasuk_B.txt"
Open txtFile For Input As #fnum
Input #fnum, RB1, RB2, RB3, RB4, RB5, RB6, RB7, RB8, RB9, RB10
Close #fnum
''''''''''''''''''''''''''''''''''''''''''
txtFile = NamaFolder & "papak\data_papak\Ukur_Rasuk_H.txt"
Open txtFile For Input As #fnum
Input #fnum, RH1, RH2, RH3, RH4, RH5, RH6, RH7, RH8, RH9, RH10
Close #fnum
''''''''''''''''''''''''''''''''''''''''''
txtFile = NamaFolder & "papak\data_papak\Ukur_Rentang.txt"
Open txtFile For Input As #fnum
Input #fnum, In1, In2, In3, In4, In5, In6, In7, In8
Close #fnum
''''''''''''''''''''''''''''''''''''''''''''
txtFile = NamaFolder & "papak\data_papak\Tebal_Papak.txt"
Open txtFile For Input As #fnum
Input #fnum, TP1, TP2, TP3, TP4, TP5, TP6, TP7, TP8, _
             TP9, TP10, TP11
Close #fnum


txtFile = NamaFolder & "papak\data_papak\Dsg_Perimeters.txt"
Open txtFile For Input As #fnum
Input #fnum, DFCU, DFYV, DDIA, DCVR, DBMK
Close #fnum
''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''
'In1 = Trim(Text1.Text)
'In2 = Trim(Text2.Text)
'In3 = Trim(Text3.Text)
'In4 = Trim(Text4.Text)
'In5 = Trim(Text5.Text)
'In6 = Trim(Text6.Text)
'In7 = Trim(Text7.Text)
'In8 = Trim(Text8.Text)

'''Command2_Click
HighlightRentang
DisableTxtOnetoEight
HighlightCoefficient
HighlightBeban
HighlightLabel_Beban
HighlightLabel_Grid
HighlightCircleGrid
HideGridInput
TajukCoefficient

PaneL1
PaneL2
PaneL3
PaneL4
PaneL5
PaneL6
PaneL7
PaneL8
PaneL9
PaneL10
PaneL11


'Label63.Caption = ""
'Label64.Caption = ""
'Label65.Caption = ""
'Label66.Caption = ""
'Label67.Caption = ""
'Label68.Caption = ""
'Label69.Caption = ""
'Label70.Caption = ""
'Label71.Caption = ""
'Label72.Caption = ""
'Label73.Caption = ""

EnableLabel_Beban
DisableTxtNineToNineteen
End Sub



Private Sub Command12_Click()
'''show load
'''Command2_Click
Command11_Click
UnBoldLabel63to73
HighlightBeban
HighlightLabel_Beban
HideCoefficient
TajukBeban
Label63_Click
Label64_Click
Label65_Click
Label66_Click
Label67_Click
Label68_Click
Label69_Click
Label70_Click
Label71_Click
Label72_Click
Label73_Click

DisableLabel_Beban
DisableTxtNineToNineteen
End Sub

Private Sub Command13_Click()
'''show moment
'''Command2_Click
DefaultLocLabel63to73
BoldLabel63to73
Command11_Click
HighlightBeban
'''HideLabel_Beban
HighlightCoefficient
TajukMoment
Text9_DblClick
Text10_DblClick
Text11_DblClick
Text12_DblClick
Text13_DblClick
Text14_DblClick
Text15_DblClick
Text16_DblClick
Text17_DblClick
Text18_DblClick
Text19_DblClick

'Label63.Caption = ""
'Label64.Caption = ""
'Label65.Caption = ""
'Label66.Caption = ""
'Label67.Caption = ""
'Label68.Caption = ""
'Label69.Caption = ""
'Label70.Caption = ""
'Label71.Caption = ""
'Label72.Caption = ""
'Label73.Caption = ""

DisableLabel_Beban
DisableTxtNineToNineteen
End Sub

Private Sub Command14_Click()
''dsg perimeters
Label87.Caption = "Minimun Ast = 0.13% -> 0.24%bh."
HighlightPanel
OpenDataFileUkur_Rentang
Text1.Text = In1
Text2.Text = In2
Text3.Text = In3
Text4.Text = In4
Text5.Text = In5
Text6.Text = In6
Text7.Text = In7
Text8.Text = In8
OpenDataFileDsg_Perimeters
Label63.Caption = "\ fcu"
Label64.Caption = "\ fy"
Label65.Caption = "\ bar dia"
Label66.Caption = "\ conc cvr"
Label67.Caption = "\ barmark"
Label68.Caption = "-"
Label69.Caption = "-"
Label70.Caption = "-"
Label71.Caption = "-"
Label72.Caption = "-"
Label73.Caption = "-"


Command11.Enabled = False
Command12.Enabled = False
Command13.Enabled = False

EnableTxtNineToNineteen
Text14.Text = "<->"
Text15.Text = "<->"
Text16.Text = "<->"
Text17.Text = "<->"
Text18.Text = "<->"
Text19.Text = "<->"



OpenDataFileTebal_Papak
HighlightBeban
HighlightLabel_Beban
HideCoefficient
''HideCircleGrid
''HideGridInput
''HideLabel_Grid
HighlightRentang
DisableTxtOnetoEight

HighlightLabel_Grid
HighlightCircleGrid
HideGridInput

Text9.Text = DFCU
Text10.Text = DFYV
Text11.Text = DDIA
Text12.Text = DCVR
Text13.Text = DBMK

Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
Command7.Enabled = False
Command10.Enabled = True
Command14.Enabled = True

DsgPeriNew.Value9 = Text9.Text
DsgPeriNew.Value10 = Text10.Text
DsgPeriNew.Value11 = Text11.Text
DsgPeriNew.Value12 = Text12.Text
DsgPeriNew.Value13 = Text13.Text

Text9.Text = DsgPeriNew.DSG_fcu
Text10.Text = DsgPeriNew.DSG_fy
Text11.Text = DsgPeriNew.DSG_bardia
Text12.Text = DsgPeriNew.DSG_cover
Text13.Text = DsgPeriNew.DSG_barmark


End Sub


Private Sub Command2_Click()
Label87.Caption = "Minimun Ast = 0.13% -> 0.24%bh."
HighlightPanel
OpenDataFileUkur_Rentang
Command11.Enabled = False
Command12.Enabled = False
Command13.Enabled = False

HighlightRentang
HideCoefficient
HideBeban
HideLabel_Beban
''HideCircleGrid
''HideGridInput
''HideLabel_Grid
EnableTxtOnetoEight

HighlightLabel_Grid
HighlightCircleGrid
HideGridInput

Text1.Text = In1
Text2.Text = In2
Text3.Text = In3
Text4.Text = In4
Text5.Text = In5
Text6.Text = In6
Text7.Text = In7
Text8.Text = In8

Text1.BackColor = vbWhite
Text2.BackColor = vbWhite
Text3.BackColor = vbWhite
Text4.BackColor = vbWhite
Text5.BackColor = vbWhite
Text6.BackColor = vbWhite
Text7.BackColor = vbWhite
Text8.BackColor = vbWhite

Command1.Enabled = False
Command2.Enabled = True
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
Command7.Enabled = False
Command10.Enabled = True
Command14.Enabled = False

Dim fnum As Integer
Dim txtFile, Temp As String
fnum = FreeFile
txtFile = NamaFolder & "rasuk\SpanOneGET.txt"

RentangNew.Value1 = Text1.Text
RentangNew.Value2 = Text2.Text
RentangNew.Value3 = Text3.Text
RentangNew.Value4 = Text4.Text
RentangNew.Value5 = Text5.Text
RentangNew.Value6 = Text6.Text
RentangNew.Value7 = Text7.Text
RentangNew.Value8 = Text8.Text

Text1.Text = RentangNew.Span_ONEtoTWO
Text2.Text = RentangNew.Span_TWOtoTHREE
Text3.Text = RentangNew.Span_THREEtoFOUR
Text4.Text = RentangNew.Span_FOURtoFIVE
Text5.Text = RentangNew.Span_EtoD
Text6.Text = RentangNew.Span_DtoC
Text7.Text = RentangNew.Span_CtoB
Text8.Text = RentangNew.Span_BtoA


End Sub

Private Sub Command3_Click()
Label87.Caption = "Minimun Ast = 0.13% -> 0.24%bh."
HighlightPanel
Command11.Enabled = False
Command12.Enabled = False
Command13.Enabled = False

OpenDataFileUkur_Rasuk_B
HighlightGridInput
HighlightCircleGrid
HideCoefficient
HideBeban
HideLabel_Beban
HideLabel_Grid
HighlightRentang
DisableTxtOnetoEight

Text20.Text = RB1
Text21.Text = RB2
Text22.Text = RB3
Text23.Text = RB4
Text24.Text = RB5
Text25.Text = RB6
Text26.Text = RB7
Text27.Text = RB8
Text28.Text = RB9
Text29.Text = RB10
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = True
Command4.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
Command7.Enabled = False
Command10.Enabled = True
Command14.Enabled = False

Shape30.Shape = 1
Shape31.Shape = 1
Shape32.Shape = 1
Shape33.Shape = 1
Shape34.Shape = 1
Shape35.Shape = 1
Shape36.Shape = 1
Shape37.Shape = 1
Shape38.Shape = 1
Shape39.Shape = 1

Shape30.FillColor = vbBlue
Shape31.FillColor = vbBlue
Shape32.FillColor = vbBlue
Shape33.FillColor = vbBlue
Shape34.FillColor = vbBlue
Shape35.FillColor = vbBlue
Shape36.FillColor = vbBlue
Shape37.FillColor = vbBlue
Shape38.FillColor = vbBlue
Shape39.FillColor = vbBlue

Dim fnum As Integer
Dim txtFile, Temp As String
fnum = FreeFile
txtFile = NamaFolder & "rasuk\SpanOneGET.txt"

Rasuk_BNew.Value20 = Text20.Text
Rasuk_BNew.Value21 = Text21.Text
Rasuk_BNew.Value22 = Text22.Text
Rasuk_BNew.Value23 = Text23.Text
Rasuk_BNew.Value24 = Text24.Text
Rasuk_BNew.Value25 = Text25.Text
Rasuk_BNew.Value26 = Text26.Text
Rasuk_BNew.Value27 = Text27.Text
Rasuk_BNew.Value28 = Text28.Text
Rasuk_BNew.Value29 = Text29.Text

Text20.Text = Rasuk_BNew.B_ONE_AtoE
Text21.Text = Rasuk_BNew.B_TWO_AtoE
Text22.Text = Rasuk_BNew.B_THREE_BtoE
Text23.Text = Rasuk_BNew.B_FOUR_CtoE
Text24.Text = Rasuk_BNew.B_FIVE_DtoE
Text25.Text = Rasuk_BNew.B_A_ONEtoTWO
Text26.Text = Rasuk_BNew.B_B_ONEtoTHREE
Text27.Text = Rasuk_BNew.B_C_ONEtoFOUR
Text28.Text = Rasuk_BNew.B_D_ONEtoFIVE
Text29.Text = Rasuk_BNew.B_E_ONEtoFIVE

End Sub

Private Sub Command4_Click()
Label87.Caption = "Minimun Ast = 0.13% -> 0.24%bh."
HighlightPanel
Command11.Enabled = False
Command12.Enabled = False
Command13.Enabled = False

OpenDataFileUkur_Rasuk_H
HighlightGridInput
HighlightCircleGrid
HideCoefficient
HideBeban
HideLabel_Beban
HideLabel_Grid
HighlightRentang
DisableTxtOnetoEight

Text20.Text = RH1
Text21.Text = RH2
Text22.Text = RH3
Text23.Text = RH4
Text24.Text = RH5
Text25.Text = RH6
Text26.Text = RH7
Text27.Text = RH8
Text28.Text = RH9
Text29.Text = RH10
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = True
Command5.Enabled = False
Command6.Enabled = False
Command7.Enabled = False
Command10.Enabled = True
Command14.Enabled = False

Shape30.Shape = 1
Shape31.Shape = 1
Shape32.Shape = 1
Shape33.Shape = 1
Shape34.Shape = 1
Shape35.Shape = 1
Shape36.Shape = 1
Shape37.Shape = 1
Shape38.Shape = 1
Shape39.Shape = 1

Shape30.FillColor = vbGreen
Shape31.FillColor = vbGreen
Shape32.FillColor = vbGreen
Shape33.FillColor = vbGreen
Shape34.FillColor = vbGreen
Shape35.FillColor = vbGreen
Shape36.FillColor = vbGreen
Shape37.FillColor = vbGreen
Shape38.FillColor = vbGreen
Shape39.FillColor = vbGreen



Rasuk_HNew.Value20 = Text20.Text
Rasuk_HNew.Value21 = Text21.Text
Rasuk_HNew.Value22 = Text22.Text
Rasuk_HNew.Value23 = Text23.Text
Rasuk_HNew.Value24 = Text24.Text
Rasuk_HNew.Value25 = Text25.Text
Rasuk_HNew.Value26 = Text26.Text
Rasuk_HNew.Value27 = Text27.Text
Rasuk_HNew.Value28 = Text28.Text
Rasuk_HNew.Value29 = Text29.Text

Text20.Text = Rasuk_HNew.H_ONE_AtoE
Text21.Text = Rasuk_HNew.H_TWO_AtoE
Text22.Text = Rasuk_HNew.H_THREE_BtoE
Text23.Text = Rasuk_HNew.H_FOUR_CtoE
Text24.Text = Rasuk_HNew.H_FIVE_DtoE
Text25.Text = Rasuk_HNew.H_A_ONEtoTWO
Text26.Text = Rasuk_HNew.H_B_ONEtoTHREE
Text27.Text = Rasuk_HNew.H_C_ONEtoFOUR
Text28.Text = Rasuk_HNew.H_D_ONEtoFIVE
Text29.Text = Rasuk_HNew.H_E_ONEtoFIVE

End Sub

Private Sub Command5_Click()
Label87.Caption = "Minimun Ast = 0.13% -> 0.24%bh."
HighlightPanel
Label63.Caption = "\tebal papak"
Label64.Caption = "\tebal papak"
Label65.Caption = "\tebal papak"
Label66.Caption = "\tebal papak"
Label67.Caption = "\tebal papak"
Label68.Caption = "\tebal papak"
Label69.Caption = "\tebal papak"
Label70.Caption = "\tebal papak"
Label71.Caption = "\tebal papak"
Label72.Caption = "\tebal papak"
Label73.Caption = "\tebal papak"

Command11.Enabled = False
Command12.Enabled = False
Command13.Enabled = False

EnableTxtNineToNineteen
OpenDataFileTebal_Papak
HighlightBeban
HighlightLabel_Beban
HideCoefficient
''HideCircleGrid
''HideGridInput
''HideLabel_Grid
HighlightRentang
DisableTxtOnetoEight

HighlightLabel_Grid
HighlightCircleGrid
HideGridInput

Text9.Text = TP1
Text10.Text = TP2
Text11.Text = TP3
Text12.Text = TP4
Text13.Text = TP5
Text14.Text = TP6
Text15.Text = TP7
Text16.Text = TP8
Text17.Text = TP9
Text18.Text = TP10
Text19.Text = TP11
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = True
Command6.Enabled = False
Command7.Enabled = False
Command10.Enabled = True
Command14.Enabled = False

TPapakNew.Value9 = Text9.Text
TPapakNew.Value10 = Text10.Text
TPapakNew.Value11 = Text11.Text
TPapakNew.Value12 = Text12.Text
TPapakNew.Value13 = Text13.Text
TPapakNew.Value14 = Text14.Text
TPapakNew.Value15 = Text15.Text
TPapakNew.Value16 = Text16.Text
TPapakNew.Value17 = Text17.Text
TPapakNew.Value18 = Text18.Text
TPapakNew.Value19 = Text19.Text

Text9.Text = TPapakNew.TPapak_ONE_D
Text10.Text = TPapakNew.TPapak_TWO_D
Text11.Text = TPapakNew.TPapak_THREE_D
Text12.Text = TPapakNew.TPapak_FOUR_D
Text13.Text = TPapakNew.TPapak_ONE_C
Text14.Text = TPapakNew.TPapak_TWO_C
Text15.Text = TPapakNew.TPapak_THREE_C
Text16.Text = TPapakNew.TPapak_ONE_B
Text17.Text = TPapakNew.TPapak_TWO_B
Text18.Text = TPapakNew.TPapak_ONE_A
Text19.Text = TPapakNew.TPapak_FOUR_A

End Sub

Private Sub Command6_Click()
Label87.Caption = "Minimun Ast = 0.13% -> 0.24%bh."
HighlightPanel
Label63.Caption = "\kemasan"
Label64.Caption = "\kemasan"
Label65.Caption = "\kemasan"
Label66.Caption = "\kemasan"
Label67.Caption = "\kemasan"
Label68.Caption = "\kemasan"
Label69.Caption = "\kemasan"
Label70.Caption = "\kemasan"
Label71.Caption = "\kemasan"
Label72.Caption = "\kemasan"
Label73.Caption = "\kemasan"

Command11.Enabled = False
Command12.Enabled = False
Command13.Enabled = False

EnableTxtNineToNineteen
OpenDataFileBeban_Kemasan
HighlightBeban
HighlightLabel_Beban
HideCoefficient
''HideCircleGrid
''HideGridInput
''HideLabel_Grid
HighlightRentang
DisableTxtOnetoEight

HighlightLabel_Grid
HighlightCircleGrid
HideGridInput

Text9.Text = Bk1
Text10.Text = Bk2
Text11.Text = Bk3
Text12.Text = Bk4
Text13.Text = Bk5
Text14.Text = Bk6
Text15.Text = Bk7
Text16.Text = Bk8
Text17.Text = Bk9
Text18.Text = Bk10
Text19.Text = Bk11
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Command6.Enabled = True
Command7.Enabled = False
Command10.Enabled = True
Command14.Enabled = False

BKemasNew.Value9 = Text9.Text
BKemasNew.Value10 = Text10.Text
BKemasNew.Value11 = Text11.Text
BKemasNew.Value12 = Text12.Text
BKemasNew.Value13 = Text13.Text
BKemasNew.Value14 = Text14.Text
BKemasNew.Value15 = Text15.Text
BKemasNew.Value16 = Text16.Text
BKemasNew.Value17 = Text17.Text
BKemasNew.Value18 = Text18.Text
BKemasNew.Value19 = Text19.Text

Text9.Text = BKemasNew.BKemas_ONE_D
Text10.Text = BKemasNew.BKemas_TWO_D
Text11.Text = BKemasNew.BKemas_THREE_D
Text12.Text = BKemasNew.BKemas_FOUR_D
Text13.Text = BKemasNew.BKemas_ONE_C
Text14.Text = BKemasNew.BKemas_TWO_C
Text15.Text = BKemasNew.BKemas_THREE_C
Text16.Text = BKemasNew.BKemas_ONE_B
Text17.Text = BKemasNew.BKemas_TWO_B
Text18.Text = BKemasNew.BKemas_ONE_A
Text19.Text = BKemasNew.BKemas_FOUR_A

End Sub

Private Sub Command7_Click()
Label87.Caption = "Minimun Ast = 0.13% -> 0.24%bh."
HighlightPanel
Label63.Caption = "\beban hidup"
Label64.Caption = "\beban hidup"
Label65.Caption = "\beban hidup"
Label66.Caption = "\beban hidup"
Label67.Caption = "\beban hidup"
Label68.Caption = "\beban hidup"
Label69.Caption = "\beban hidup"
Label70.Caption = "\beban hidup"
Label71.Caption = "\beban hidup"
Label72.Caption = "\beban hidup"
Label73.Caption = "\beban hidup"

Command11.Enabled = False
Command12.Enabled = False
Command13.Enabled = False

EnableTxtNineToNineteen
OpenDataFileBeban_Hidup
HighlightBeban
HighlightLabel_Beban
HideCoefficient
''HideCircleGrid
''HideGridInput
''HideLabel_Grid
HighlightRentang
DisableTxtOnetoEight

HighlightLabel_Grid
HighlightCircleGrid
HideGridInput

Text9.Text = Bh1
Text10.Text = Bh2
Text11.Text = Bh3
Text12.Text = Bh4
Text13.Text = Bh5
Text14.Text = Bh6
Text15.Text = Bh7
Text16.Text = Bh8
Text17.Text = Bh9
Text18.Text = Bh10
Text19.Text = Bh11
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
Command7.Enabled = True
Command10.Enabled = True
Command14.Enabled = False


BHidupNew.Value9 = Text9.Text
BHidupNew.Value10 = Text10.Text
BHidupNew.Value11 = Text11.Text
BHidupNew.Value12 = Text12.Text
BHidupNew.Value13 = Text13.Text
BHidupNew.Value14 = Text14.Text
BHidupNew.Value15 = Text15.Text
BHidupNew.Value16 = Text16.Text
BHidupNew.Value17 = Text17.Text
BHidupNew.Value18 = Text18.Text
BHidupNew.Value19 = Text19.Text

Text9.Text = BHidupNew.BHidup_ONE_D
Text10.Text = BHidupNew.BHidup_TWO_D
Text11.Text = BHidupNew.BHidup_THREE_D
Text12.Text = BHidupNew.BHidup_FOUR_D
Text13.Text = BHidupNew.BHidup_ONE_C
Text14.Text = BHidupNew.BHidup_TWO_C
Text15.Text = BHidupNew.BHidup_THREE_C
Text16.Text = BHidupNew.BHidup_ONE_B
Text17.Text = BHidupNew.BHidup_TWO_B
Text18.Text = BHidupNew.BHidup_ONE_A
Text19.Text = BHidupNew.BHidup_FOUR_A

End Sub



Private Sub OpenDataFileBeban_Hidup()
Form1.Picture = LoadPicture(NamaFolder & "icon\datam.ico")
Dim fnum As Integer
Dim txtFile, Temp As String
fnum = FreeFile
txtFile = NamaFolder & "papak\data_papak\Beban_Hidup.txt"
Open txtFile For Input As #fnum
Input #fnum, Bh1, Bh2, Bh3, Bh4, Bh5, Bh6, Bh7, Bh8, _
             Bh9, Bh10, Bh11
Close #fnum
End Sub
''''''''''''''''''''''''''''''''''''''''''
Private Sub OpenDataFileBeban_Kemasan()
Form1.Picture = LoadPicture(NamaFolder & "icon\datam.ico")
Dim fnum As Integer
Dim txtFile, Temp As String
fnum = FreeFile
txtFile = NamaFolder & "papak\data_papak\Beban_Kemasan.txt"
Open txtFile For Input As #fnum
Input #fnum, Bk1, Bk2, Bk3, Bk4, Bk5, Bk6, Bk7, Bk8, _
             Bk9, Bk10, Bk11
Close #fnum
End Sub
''''''''''''''''''''''''''''''''''''''''''

Private Sub OpenDataFileLabel_Grid()
Form1.Picture = LoadPicture(NamaFolder & "icon\datam.ico")
Dim fnum As Integer
Dim txtFile, Temp As String
fnum = FreeFile
txtFile = NamaFolder & "papak\data_papak\Label_Grid.txt"
Open txtFile For Input As #fnum
Input #fnum, LG1, LG2, LG3, LG4, LG5
Input #fnum, LG6
Input #fnum, LG7
Input #fnum, LG8
Input #fnum, LG9
Input #fnum, LG10
Close #fnum
End Sub
''''''''''''''''''''''''''''''''''''''''''
Private Sub OpenDataFileUkur_Rasuk_B()
Form1.Picture = LoadPicture(NamaFolder & "icon\datam.ico")
Dim fnum As Integer
Dim txtFile, Temp As String
fnum = FreeFile
txtFile = NamaFolder & "papak\data_papak\Ukur_Rasuk_B.txt"
Open txtFile For Input As #fnum
Input #fnum, RB1, RB2, RB3, RB4, RB5, RB6, RB7, RB8, RB9, RB10
Close #fnum
End Sub
''''''''''''''''''''''''''''''''''''''''''
Private Sub OpenDataFileUkur_Rasuk_H()
Form1.Picture = LoadPicture(NamaFolder & "icon\datam.ico")
Dim fnum As Integer
Dim txtFile, Temp As String
fnum = FreeFile
txtFile = NamaFolder & "papak\data_papak\Ukur_Rasuk_H.txt"
Open txtFile For Input As #fnum
Input #fnum, RH1, RH2, RH3, RH4, RH5, RH6, RH7, RH8, RH9, RH10
Close #fnum
End Sub
''''''''''''''''''''''''''''''''''''''''''
Private Sub OpenDataFileUkur_Rentang()
Form1.Picture = LoadPicture(NamaFolder & "icon\datam.ico")
Dim fnum As Integer
Dim txtFile, Temp As String
fnum = FreeFile
txtFile = NamaFolder & "papak\data_papak\Ukur_Rentang.txt"
Open txtFile For Input As #fnum
Input #fnum, In1, In2, In3, In4, In5, In6, In7, In8
Close #fnum
''''''''''''''''''''''''''''''''''''''''''''
End Sub

Private Sub OpenDataFileTebal_Papak()
Form1.Picture = LoadPicture(NamaFolder & "icon\datam.ico")
Dim fnum As Integer
Dim txtFile, Temp As String
fnum = FreeFile
txtFile = NamaFolder & "papak\data_papak\Tebal_Papak.txt"
Open txtFile For Input As #fnum
Input #fnum, TP1, TP2, TP3, TP4, TP5, TP6, TP7, TP8, _
             TP9, TP10, TP11
Close #fnum
''''''''''''''''''''''''''''''''''''''''''
End Sub
Private Sub OpenDataFileDsg_Perimeters()
Form1.Picture = LoadPicture(NamaFolder & "icon\datam.ico")
Dim fnum As Integer
Dim txtFile, Temp As String
fnum = FreeFile
txtFile = NamaFolder & "papak\data_papak\Dsg_Perimeters.txt"
Open txtFile For Input As #fnum
Input #fnum, DFCU, DFYV, DDIA, DCVR, DBMK

Close #fnum
''''''''''''''''''''''''''''''''''''''''''
End Sub


Private Sub Command8_Click()
Label87.Caption = "Minimun Ast = 0.13% -> 0.24%bh."
HighlightPanel
Load Form2
Form2.Visible = True

''''MsgBox "                                                               " & _
"uNder constrUcTion" & _
"                                                               " & _
Str(1) & Str(2) & Chr(3) & Chr(4) & Chr(5) & _
Chr(6) & Chr(7) & Chr(8) & Chr(9) & Chr(10) & _
Chr(11) & Chr(12) & Chr(13) & Chr(14) & Chr(15) & _
Chr(16) & Chr(17) & Chr(18) & Chr(19) & Chr(20) & _
Chr(21) & Chr(22) & Chr(23) & Chr(24) & Chr(25) & _
Chr(26) & Chr(27) & Chr(28) & Chr(29) & Chr(30) & _
Chr(31) & Chr(32) & Chr(33) & Chr(34) & Chr(35) & _
Chr(136) & Chr(137) & Chr(138) & Chr(139) & Chr(140) & _
Chr(141) & Chr(142) & Chr(143) & Chr(144) & Chr(145) & _
Chr(146) & Chr(147) & Chr(148) & Chr(149) & Chr(140) & _
Chr(151) & Chr(152) & Chr(153) & Chr(154) & Chr(155) & _
Chr(156) & Chr(157) & Chr(158) & Chr(159) & Chr(160) & _
Chr(161) & Chr(162) & Chr(163) & Chr(164) & Chr(165) & _
Chr(166) & Chr(167) & Chr(168) & " <> " & Chr(169) & Chr(169) & Chr(170) & _
Chr(171) & Chr(172) & Chr(173) & Chr(174) & Chr(174) & Chr(175) & _
Chr(176) & Chr(177) & Chr(178) & Chr(179) & Chr(180) & _
Chr(181) & Chr(182) & Chr(183) & Chr(184) & Chr(185) & _
Chr(186) & Chr(187) & Chr(188) & Chr(189) & Chr(190) & _
Chr(191) & Chr(192) & Chr(193) & Chr(194) & Chr(195), 4144, "link to autocad"
End Sub

Private Sub Command8_GotFocus()
Label85.Visible = True
End Sub

Private Sub Command8_LostFocus()
Label85.Visible = False
End Sub



Private Sub Form_Load()

NamaFolder = "C:\autodraf\"

End Sub

Private Sub Label11_Click()
Dim modFact As Double
If In6 / In2 < 2 Then
modFact = 0.55 + (477 - 5 * DFYV / 8) _
          / (120 * (0.9 + ((Val(Label11.Caption) * 1000000#) _
          / 1000 / ((TP2) - (DCVR) - (DDIA / 2)) ^ 2)))
Label64.Caption = "l/d=" & Int((In6) / ((TP2) - (DCVR) - (DDIA / 2))) _
& " m.f.=" & Int(modFact * 100) / 100
End If

If Val(Label11.Caption) >= Val(Label8.Caption) Then
Text10.Text = "Ast=" & Int(RequiredAst(DFCU, DFYV, DDIA, DCVR, _
TP2, DBMK, 0.9, Val(Label11.Caption))) & "mm2/m."
  Else
  Text10.Text = "Ast=" & Int(RequiredAst(DFCU, DFYV, 3 * DDIA, DCVR, _
   TP2, DBMK, 0.9, Val(Label11.Caption))) & "mm2/m."
    End If

Label87.Caption = "Panel_2 min. Ast = " & Int(TP2 * 1000 * _
                   (0.24 - (DFYV - 250) * 0.11 / 210) / 100) & " mm2/m."
HighlightPanel
Label7.FontBold = False
Label8.FontBold = False
Label9.FontBold = False
Label10.FontBold = False
Label11.FontBold = True
Label12.FontBold = False
End Sub

Private Sub Label11_DblClick()
Command13_Click
End Sub


Private Sub Label12_Click()
Text10.Text = "Ast=" & Int(RequiredAst(DFCU, DFYV, DDIA, DCVR, _
TP2, DBMK, 0.9, Val(Label12.Caption))) & "mm2/m."
Label87.Caption = "Panel_2 min. Ast = " & Int(TP2 * 1000 * _
                   (0.24 - (DFYV - 250) * 0.11 / 210) / 100) & " mm2/m."

HighlightPanel
Label7.FontBold = False
Label8.FontBold = False
Label9.FontBold = False
Label10.FontBold = False
Label11.FontBold = False
Label12.FontBold = True
End Sub

Private Sub Label12_DblClick()
Command13_Click
End Sub

Private Sub Label13_Click()
Text11.Text = "Ast=" & Int(RequiredAst(DFCU, DFYV, DDIA, DCVR, _
TP3, DBMK, 0.9, Val(Label13.Caption))) & "mm2/m."
Label87.Caption = "Panel_3 min. Ast = " & Int(TP3 * 1000 * _
                   (0.24 - (DFYV - 250) * 0.11 / 210) / 100) & " mm2/m."

HighlightPanel
Label13.FontBold = True
Label14.FontBold = False
Label15.FontBold = False
Label16.FontBold = False
Label17.FontBold = False
Label18.FontBold = False
End Sub

Private Sub Label13_DblClick()
Command13_Click
End Sub

Private Sub Label14_Click()
Dim modFact As Double
If In3 / In6 < 2 Then
modFact = 0.55 + (477 - 5 * DFYV / 8) _
          / (120 * (0.9 + ((Val(Label14.Caption) * 1000000#) _
          / 1000 / ((TP3) - (DCVR) - (DDIA / 2)) ^ 2)))
Label65.Caption = "l/d=" & Int((In3) / ((TP3) - (DCVR) - (DDIA / 2))) _
& " m.f.=" & Int(modFact * 100) / 100
End If

Text11.Text = "Ast=" & Int(RequiredAst(DFCU, DFYV, DDIA, DCVR, _
TP3, DBMK, 0.9, Val(Label14.Caption))) & "mm2/m."
Label87.Caption = "Panel_3 min. Ast = " & Int(TP3 * 1000 * _
                   (0.24 - (DFYV - 250) * 0.11 / 210) / 100) & " mm2/m."

HighlightPanel
Label13.FontBold = False
Label14.FontBold = True
Label15.FontBold = False
Label16.FontBold = False
Label17.FontBold = False
Label18.FontBold = False
End Sub

Private Sub Label14_DblClick()
Command13_Click
End Sub

Private Sub Label15_Click()
Text11.Text = "Ast=" & Int(RequiredAst(DFCU, DFYV, DDIA, DCVR, _
TP3, DBMK, 0.9, Val(Label15.Caption))) & "mm2/m."
Label87.Caption = "Panel_3 min. Ast = " & Int(TP3 * 1000 * _
                   (0.24 - (DFYV - 250) * 0.11 / 210) / 100) & " mm2/m."

HighlightPanel
Label13.FontBold = False
Label14.FontBold = False
Label15.FontBold = True
Label16.FontBold = False
Label17.FontBold = False
Label18.FontBold = False
End Sub

Private Sub Label15_DblClick()
Command13_Click
End Sub

Private Sub Label17_Click()
Dim modFact As Double
If In6 / In3 < 2 Then
modFact = 0.55 + (477 - 5 * DFYV / 8) _
          / (120 * (0.9 + ((Val(Label17.Caption) * 1000000#) _
          / 1000 / ((TP3) - (DCVR) - (DDIA / 2)) ^ 2)))
Label65.Caption = "l/d=" & Int((In6) / ((TP3) - (DCVR) - (DDIA / 2))) _
& " m.f.=" & Int(modFact * 100) / 100
End If

Text11.Text = "Ast=" & Int(RequiredAst(DFCU, DFYV, DDIA, DCVR, _
TP3, DBMK, 0.9, Val(Label17.Caption))) & "mm2/m."
Label87.Caption = "Panel_3 min. Ast = " & Int(TP3 * 1000 * _
                   (0.24 - (DFYV - 250) * 0.11 / 210) / 100) & " mm2/m."

HighlightPanel
Label13.FontBold = False
Label14.FontBold = False
Label15.FontBold = False
Label16.FontBold = False
Label17.FontBold = True
Label18.FontBold = False
End Sub

Private Sub Label17_DblClick()
Command13_Click
End Sub

Private Sub Label18_Click()
Text11.Text = "Ast=" & Int(RequiredAst(DFCU, DFYV, DDIA, DCVR, _
TP3, DBMK, 0.9, Val(Label18.Caption))) & "mm2/m."
Label87.Caption = "Panel_3 min. Ast = " & Int(TP3 * 1000 * _
                   (0.24 - (DFYV - 250) * 0.11 / 210) / 100) & " mm2/m."

HighlightPanel
Label13.FontBold = False
Label14.FontBold = False
Label15.FontBold = False
Label16.FontBold = False
Label17.FontBold = False
Label18.FontBold = True
End Sub

Private Sub Label18_DblClick()
Command13_Click
End Sub

Private Sub Label19_Click()
Text12.Text = "Ast=" & Int(RequiredAst(DFCU, DFYV, DDIA, DCVR, _
TP4, DBMK, 0.9, Val(Label19.Caption))) & "mm2/m."
Label87.Caption = "Panel_4 min. Ast = " & Int(TP4 * 1000 * _
                   (0.24 - (DFYV - 250) * 0.11 / 210) / 100) & " mm2/m."

HighlightPanel
Label19.FontBold = True
Label20.FontBold = False
Label21.FontBold = False
Label22.FontBold = False
Label23.FontBold = False
Label24.FontBold = False
End Sub

Private Sub Label19_DblClick()
Command13_Click
End Sub

Private Sub Label2_Click()
Dim modFact As Double
If In1 / In6 < 2 Then
modFact = 0.55 + (477 - 5 * DFYV / 8) _
          / (120 * (0.9 + ((Val(Label2.Caption) * 1000000#) _
          / 1000 / ((TP1) - (DCVR) - (DDIA / 2)) ^ 2)))
Label63.Caption = "l/d=" & Int((In1) / ((TP1) - (DCVR) - (DDIA / 2))) _
& " m.f.=" & Int(modFact * 100) / 100
End If

If Val(Label2.Caption) >= Val(Label5.Caption) Then
Text9.Text = "Ast=" & Int(RequiredAst(DFCU, DFYV, DDIA, DCVR, _
TP1, DBMK, 0.9, Val(Label2.Caption))) & "mm2/m."
 Else
 Text9.Text = "Ast=" & Int(RequiredAst(DFCU, DFYV, 3 * DDIA, DCVR, _
 TP1, DBMK, 0.9, Val(Label2.Caption))) & "mm2/m."
   End If
   
Label87.Caption = "Panel_1 min. Ast = " & Int(TP1 * 1000 * _
                   (0.24 - (DFYV - 250) * 0.11 / 210) / 100) & " mm2/m."

HighlightPanel
Label1.FontBold = False
Label2.FontBold = True
Label3.FontBold = False
Label4.FontBold = False
Label5.FontBold = False
Label6.FontBold = False

End Sub

Private Sub Label2_DblClick()
Command13_Click
End Sub


Private Sub Label20_Click()
Dim modFact As Double
If In4 / In6 < 2 Then
modFact = 0.55 + (477 - 5 * DFYV / 8) _
          / (120 * (0.9 + ((Val(Label20.Caption) * 1000000#) _
          / 1000 / ((TP4) - (DCVR) - (DDIA / 2)) ^ 2)))
Label66.Caption = "l/d=" & Int((In4) / ((TP4) - (DCVR) - (DDIA / 2))) _
& " m.f.=" & Int(modFact * 100) / 100
End If

Text12.Text = "Ast=" & Int(RequiredAst(DFCU, DFYV, DDIA, DCVR, _
TP4, DBMK, 0.9, Val(Label20.Caption))) & "mm2/m."
Label87.Caption = "Panel_4 min. Ast = " & Int(TP4 * 1000 * _
                   (0.24 - (DFYV - 250) * 0.11 / 210) / 100) & " mm2/m."

HighlightPanel
Label19.FontBold = False
Label20.FontBold = True
Label21.FontBold = False
Label22.FontBold = False
Label23.FontBold = False
Label24.FontBold = False
End Sub

Private Sub Label20_DblClick()
Command13_Click
End Sub

Private Sub Label21_Click()
Text12.Text = "Ast=" & Int(RequiredAst(DFCU, DFYV, DDIA, DCVR, _
TP4, DBMK, 0.9, Val(Label21.Caption))) & "mm2/m."
Label87.Caption = "Panel_4 min. Ast = " & Int(TP4 * 1000 * _
                   (0.24 - (DFYV - 250) * 0.11 / 210) / 100) & " mm2/m."

HighlightPanel
Label19.FontBold = False
Label20.FontBold = False
Label21.FontBold = True
Label22.FontBold = False
Label23.FontBold = False
Label24.FontBold = False
End Sub

Private Sub Label21_DblClick()
Command13_Click
End Sub

Private Sub Label23_Click()
Dim modFact As Double
If In6 / In4 < 2 Then
modFact = 0.55 + (477 - 5 * DFYV / 8) _
          / (120 * (0.9 + ((Val(Label23.Caption) * 1000000#) _
          / 1000 / ((TP4) - (DCVR) - (DDIA / 2)) ^ 2)))
Label66.Caption = "l/d=" & Int((In6) / ((TP4) - (DCVR) - (DDIA / 2))) _
& " m.f.=" & Int(modFact * 100) / 100
End If

Text12.Text = "Ast=" & Int(RequiredAst(DFCU, DFYV, DDIA, DCVR, _
TP4, DBMK, 0.9, Val(Label23.Caption))) & "mm2/m."
Label87.Caption = "Panel_4 min. Ast = " & Int(TP4 * 1000 * _
                   (0.24 - (DFYV - 250) * 0.11 / 210) / 100) & " mm2/m."

HighlightPanel
Label19.FontBold = False
Label20.FontBold = False
Label21.FontBold = False
Label22.FontBold = False
Label23.FontBold = True
Label24.FontBold = False
End Sub

Private Sub Label23_DblClick()
Command13_Click
End Sub

Private Sub Label25_Click()
Text13.Text = "Ast=" & Int(RequiredAst(DFCU, DFYV, DDIA, DCVR, _
TP5, DBMK, 0.9, Val(Label25.Caption))) & "mm2/m."
Label87.Caption = "Panel_5 min. Ast = " & Int(TP5 * 1000 * _
                   (0.24 - (DFYV - 250) * 0.11 / 210) / 100) & " mm2/m."

HighlightPanel
Label25.FontBold = True
Label26.FontBold = False
Label27.FontBold = False
Label28.FontBold = False
Label29.FontBold = False
Label30.FontBold = False
End Sub

Private Sub Label25_DblClick()
Command13_Click
End Sub

Private Sub Label26_Click()
Dim modFact As Double
If In5 / In6 < 2 Then
modFact = 0.55 + (477 - 5 * DFYV / 8) _
          / (120 * (0.9 + ((Val(Label26.Caption) * 1000000#) _
          / 1000 / ((TP5) - (DCVR) - (DDIA / 2)) ^ 2)))
Label67.Caption = "l/d=" & Int((In5) / ((TP5) - (DCVR) - (DDIA / 2))) _
& " m.f.=" & Int(modFact * 100) / 100
End If

Text13.Text = "Ast=" & Int(RequiredAst(DFCU, DFYV, DDIA, DCVR, _
TP5, DBMK, 0.9, Val(Label26.Caption))) & "mm2/m."
Label87.Caption = "Panel_5 min. Ast = " & Int(TP5 * 1000 * _
                   (0.24 - (DFYV - 250) * 0.11 / 210) / 100) & " mm2/m."

HighlightPanel
Label25.FontBold = False
Label26.FontBold = True
Label27.FontBold = False
Label28.FontBold = False
Label29.FontBold = False
Label30.FontBold = False
End Sub

Private Sub Label26_DblClick()
Command13_Click
End Sub

Private Sub Label29_Click()
Dim modFact As Double
If In6 / In5 < 2 Then
modFact = 0.55 + (477 - 5 * DFYV / 8) _
          / (120 * (0.9 + ((Val(Label29.Caption) * 1000000#) _
          / 1000 / ((TP5) - (DCVR) - (DDIA / 2)) ^ 2)))
Label67.Caption = "l/d=" & Int((In6) / ((TP5) - (DCVR) - (DDIA / 2))) _
& " m.f.=" & Int(modFact * 100) / 100
End If

Text13.Text = "Ast=" & Int(RequiredAst(DFCU, DFYV, DDIA, DCVR, _
TP5, DBMK, 0.9, Val(Label29.Caption))) & "mm2/m."
Label87.Caption = "Panel_5 min. Ast = " & Int(TP5 * 1000 * _
                   (0.24 - (DFYV - 250) * 0.11 / 210) / 100) & " mm2/m."

HighlightPanel
Label25.FontBold = False
Label26.FontBold = False
Label27.FontBold = False
Label28.FontBold = False
Label29.FontBold = True
Label30.FontBold = False
End Sub

Private Sub Label29_DblClick()
Command13_Click
End Sub

Private Sub Label3_Click()
Text9.Text = "Ast=" & Int(RequiredAst(DFCU, DFYV, DDIA, DCVR, _
TP1, DBMK, 0.9, Val(Label3.Caption))) & "mm2/m."
Label87.Caption = "Panel_1 min. Ast = " & Int(TP1 * 1000 * _
                   (0.24 - (DFYV - 250) * 0.11 / 210) / 100) & " mm2/m."

HighlightPanel
Label1.FontBold = False
Label2.FontBold = False
Label3.FontBold = True
Label4.FontBold = False
Label5.FontBold = False
Label6.FontBold = False
End Sub

Private Sub Label3_DblClick()
Command13_Click
End Sub

Private Sub Label32_Click()
Dim modFact As Double
If In1 / In7 < 2 Then
modFact = 0.55 + (477 - 5 * DFYV / 8) _
          / (120 * (0.9 + ((Val(Label32.Caption) * 1000000#) _
          / 1000 / ((TP6) - (DCVR) - (DDIA / 2)) ^ 2)))
Label68.Caption = "l/d=" & Int((In1) / ((TP6) - (DCVR) - (DDIA / 2))) _
& " m.f.=" & Int(modFact * 100) / 100
End If

Text14.Text = "Ast=" & Int(RequiredAst(DFCU, DFYV, DDIA, DCVR, _
TP6, DBMK, 0.9, Val(Label32.Caption))) & "mm2/m."
Label87.Caption = "Panel_6 min. Ast = " & Int(TP6 * 1000 * _
                   (0.24 - (DFYV - 250) * 0.11 / 210) / 100) & " mm2/m."

HighlightPanel
Label31.FontBold = False
Label32.FontBold = True
Label33.FontBold = False
Label34.FontBold = False
Label35.FontBold = False
Label36.FontBold = False
End Sub

Private Sub Label32_DblClick()
Command13_Click
End Sub

Private Sub Label33_Click()
Text14.Text = "Ast=" & Int(RequiredAst(DFCU, DFYV, DDIA, DCVR, _
TP6, DBMK, 0.9, Val(Label33.Caption))) & "mm2/m."
Label87.Caption = "Panel_6 min. Ast = " & Int(TP6 * 1000 * _
                   (0.24 - (DFYV - 250) * 0.11 / 210) / 100) & " mm2/m."

HighlightPanel
Label31.FontBold = False
Label32.FontBold = False
Label33.FontBold = True
Label34.FontBold = False
Label35.FontBold = False
Label36.FontBold = False
End Sub

Private Sub Label33_DblClick()
Command13_Click
End Sub

Private Sub Label34_Click()
Text14.Text = "Ast=" & Int(RequiredAst(DFCU, DFYV, DDIA, DCVR, _
TP6, DBMK, 0.9, Val(Label34.Caption))) & "mm2/m."
Label87.Caption = "Panel_6 min. Ast = " & Int(TP6 * 1000 * _
                   (0.24 - (DFYV - 250) * 0.11 / 210) / 100) & " mm2/m."

HighlightPanel
Label31.FontBold = False
Label32.FontBold = False
Label33.FontBold = False
Label34.FontBold = True
Label35.FontBold = False
Label36.FontBold = False
End Sub

Private Sub Label34_DblClick()
Command13_Click
End Sub

Private Sub Label35_Click()
Dim modFact As Double
If In7 / In1 < 2 Then
modFact = 0.55 + (477 - 5 * DFYV / 8) _
          / (120 * (0.9 + ((Val(Label35.Caption) * 1000000#) _
          / 1000 / ((TP6) - (DCVR) - (DDIA / 2)) ^ 2)))
Label68.Caption = "l/d=" & Int((In7) / ((TP6) - (DCVR) - (DDIA / 2))) _
& " m.f.=" & Int(modFact * 100) / 100
End If

Text14.Text = "Ast=" & Int(RequiredAst(DFCU, DFYV, DDIA, DCVR, _
TP6, DBMK, 0.9, Val(Label35.Caption))) & "mm2/m."
Label87.Caption = "Panel_6 min. Ast = " & Int(TP6 * 1000 * _
                   (0.24 - (DFYV - 250) * 0.11 / 210) / 100) & " mm2/m."

HighlightPanel
Label31.FontBold = False
Label32.FontBold = False
Label33.FontBold = False
Label34.FontBold = False
Label35.FontBold = True
Label36.FontBold = False
End Sub

Private Sub Label35_DblClick()
Command13_Click
End Sub

Private Sub Label36_Click()
Text14.Text = "Ast=" & Int(RequiredAst(DFCU, DFYV, DDIA, DCVR, _
TP6, DBMK, 0.9, Val(Label36.Caption))) & "mm2/m."
Label87.Caption = "Panel_6 min. Ast = " & Int(TP6 * 1000 * _
                   (0.24 - (DFYV - 250) * 0.11 / 210) / 100) & " mm2/m."

HighlightPanel
Label31.FontBold = False
Label32.FontBold = False
Label33.FontBold = False
Label34.FontBold = False
Label35.FontBold = False
Label36.FontBold = True
End Sub

Private Sub Label36_DblClick()
Command13_Click
End Sub

Private Sub Label37_Click()
Text15.Text = "Ast=" & Int(RequiredAst(DFCU, DFYV, DDIA, DCVR, _
TP7, DBMK, 0.9, Val(Label37.Caption))) & "mm2/m."
Label87.Caption = "Panel_7 min. Ast = " & Int(TP7 * 1000 * _
                   (0.24 - (DFYV - 250) * 0.11 / 210) / 100) & " mm2/m."

HighlightPanel
Label37.FontBold = True
Label38.FontBold = False
Label39.FontBold = False
Label40.FontBold = False
Label41.FontBold = False
Label42.FontBold = False
End Sub

Private Sub Label37_DblClick()
Command13_Click
End Sub

Private Sub Label38_Click()
Dim modFact As Double
If In2 / In7 < 2 Then
modFact = 0.55 + (477 - 5 * DFYV / 8) _
          / (120 * (0.9 + ((Val(Label38.Caption) * 1000000#) _
          / 1000 / ((TP7) - (DCVR) - (DDIA / 2)) ^ 2)))
Label69.Caption = "l/d=" & Int((In2) / ((TP7) - (DCVR) - (DDIA / 2))) _
& " m.f.=" & Int(modFact * 100) / 100
End If

Text15.Text = "Ast=" & Int(RequiredAst(DFCU, DFYV, DDIA, DCVR, _
TP7, DBMK, 0.9, Val(Label38.Caption))) & "mm2/m."
Label87.Caption = "Panel_7 min. Ast = " & Int(TP7 * 1000 * _
                   (0.24 - (DFYV - 250) * 0.11 / 210) / 100) & " mm2/m."

HighlightPanel
Label37.FontBold = False
Label38.FontBold = True
Label39.FontBold = False
Label40.FontBold = False
Label41.FontBold = False
Label42.FontBold = False
End Sub

Private Sub Label38_DblClick()
Command13_Click
End Sub



Private Sub Label39_Click()
Text15.Text = "Ast=" & Int(RequiredAst(DFCU, DFYV, DDIA, DCVR, _
TP7, DBMK, 0.9, Val(Label39.Caption))) & "mm2/m."
Label87.Caption = "Panel_7 min. Ast = " & Int(TP7 * 1000 * _
                   (0.24 - (DFYV - 250) * 0.11 / 210) / 100) & " mm2/m."

HighlightPanel
Label37.FontBold = False
Label38.FontBold = False
Label39.FontBold = True
Label40.FontBold = False
Label41.FontBold = False
Label42.FontBold = False
End Sub

Private Sub Label39_DblClick()
Command13_Click
End Sub

Private Sub Label40_Click()
Text15.Text = "Ast=" & Int(RequiredAst(DFCU, DFYV, DDIA, DCVR, _
TP7, DBMK, 0.9, Val(Label40.Caption))) & "mm2/m."
Label87.Caption = "Panel_7 min. Ast = " & Int(TP7 * 1000 * _
                   (0.24 - (DFYV - 250) * 0.11 / 210) / 100) & " mm2/m."

HighlightPanel
Label37.FontBold = False
Label38.FontBold = False
Label39.FontBold = False
Label40.FontBold = True
Label41.FontBold = False
Label42.FontBold = False
End Sub

Private Sub Label40_DblClick()
Command13_Click
End Sub

Private Sub Label41_Click()
Dim modFact As Double
If In7 / In2 < 2 Then
modFact = 0.55 + (477 - 5 * DFYV / 8) _
          / (120 * (0.9 + ((Val(Label41.Caption) * 1000000#) _
          / 1000 / ((TP7) - (DCVR) - (DDIA / 2)) ^ 2)))
Label69.Caption = "l/d=" & Int((In7) / ((TP7) - (DCVR) - (DDIA / 2))) _
& " m.f.=" & Int(modFact * 100) / 100
End If

Text15.Text = "Ast=" & Int(RequiredAst(DFCU, DFYV, DDIA, DCVR, _
TP7, DBMK, 0.9, Val(Label41.Caption))) & "mm2/m."
Label87.Caption = "Panel_7 min. Ast = " & Int(TP7 * 1000 * _
                   (0.24 - (DFYV - 250) * 0.11 / 210) / 100) & " mm2/m."

HighlightPanel
Label37.FontBold = False
Label38.FontBold = False
Label39.FontBold = False
Label40.FontBold = False
Label41.FontBold = True
Label42.FontBold = False
End Sub

Private Sub Label41_DblClick()
Command13_Click
End Sub

Private Sub Label42_Click()
Text15.Text = "Ast=" & Int(RequiredAst(DFCU, DFYV, DDIA, DCVR, _
TP7, DBMK, 0.9, Val(Label42.Caption))) & "mm2/m."
Label37.FontBold = False
Label38.FontBold = False
Label39.FontBold = False
Label40.FontBold = False
Label41.FontBold = False
Label42.FontBold = True
End Sub

Private Sub Label42_DblClick()
Command13_Click
End Sub

Private Sub Label43_Click()
Text16.Text = "Ast=" & Int(RequiredAst(DFCU, DFYV, DDIA, DCVR, _
TP8, DBMK, 0.9, Val(Label43.Caption))) & "mm2/m."
Label87.Caption = "Panel_8 min. Ast = " & Int(TP8 * 1000 * _
                   (0.24 - (DFYV - 250) * 0.11 / 210) / 100) & " mm2/m."

HighlightPanel
Label43.FontBold = True
Label44.FontBold = False
Label45.FontBold = False
Label46.FontBold = False
Label47.FontBold = False
Label48.FontBold = False
End Sub

Private Sub Label43_DblClick()
Command13_Click
End Sub

Private Sub Label44_Click()
Dim modFact As Double
If In3 / In7 < 2 Then
modFact = 0.55 + (477 - 5 * DFYV / 8) _
          / (120 * (0.9 + ((Val(Label44.Caption) * 1000000#) _
          / 1000 / ((TP8) - (DCVR) - (DDIA / 2)) ^ 2)))
Label70.Caption = "l/d=" & Int((In3) / ((TP8) - (DCVR) - (DDIA / 2))) _
& " m.f.=" & Int(modFact * 100) / 100
End If

Text16.Text = "Ast=" & Int(RequiredAst(DFCU, DFYV, DDIA, DCVR, _
TP8, DBMK, 0.9, Val(Label44.Caption))) & "mm2/m."
Label87.Caption = "Panel_8 min. Ast = " & Int(TP8 * 1000 * _
                   (0.24 - (DFYV - 250) * 0.11 / 210) / 100) & " mm2/m."

HighlightPanel
Label43.FontBold = False
Label44.FontBold = True
Label45.FontBold = False
Label46.FontBold = False
Label47.FontBold = False
Label48.FontBold = False
End Sub

Private Sub Label44_DblClick()
Command13_Click
End Sub

Private Sub Label46_Click()
Text16.Text = "Ast=" & Int(RequiredAst(DFCU, DFYV, DDIA, DCVR, _
TP8, DBMK, 0.9, Val(Label46.Caption))) & "mm2/m."
Label87.Caption = "Panel_8 min. Ast = " & Int(TP8 * 1000 * _
                   (0.24 - (DFYV - 250) * 0.11 / 210) / 100) & " mm2/m."

HighlightPanel
Label43.FontBold = False
Label44.FontBold = False
Label45.FontBold = False
Label46.FontBold = True
Label47.FontBold = False
Label48.FontBold = False
End Sub

Private Sub Label46_DblClick()
Command13_Click
End Sub

Private Sub Label47_Click()
Dim modFact As Double
If In7 / In3 < 2 Then
modFact = 0.55 + (477 - 5 * DFYV / 8) _
          / (120 * (0.9 + ((Val(Label47.Caption) * 1000000#) _
          / 1000 / ((TP8) - (DCVR) - (DDIA / 2)) ^ 2)))
Label70.Caption = "l/d=" & Int((In7) / ((TP8) - (DCVR) - (DDIA / 2))) _
& " m.f.=" & Int(modFact * 100) / 100
End If

Text16.Text = "Ast=" & Int(RequiredAst(DFCU, DFYV, DDIA, DCVR, _
TP8, DBMK, 0.9, Val(Label47.Caption))) & "mm2/m."
Label87.Caption = "Panel_8 min. Ast = " & Int(TP8 * 1000 * _
                   (0.24 - (DFYV - 250) * 0.11 / 210) / 100) & " mm2/m."

HighlightPanel
Label43.FontBold = False
Label44.FontBold = False
Label45.FontBold = False
Label46.FontBold = False
Label47.FontBold = True
Label48.FontBold = False
End Sub

Private Sub Label47_DblClick()
Command13_Click
End Sub

Private Sub Label5_Click()
Dim modFact As Double
If In6 / In1 < 2 Then
modFact = 0.55 + (477 - 5 * DFYV / 8) _
          / (120 * (0.9 + ((Val(Label5.Caption) * 1000000#) _
          / 1000 / ((TP1) - (DCVR) - (DDIA / 2)) ^ 2)))
Label63.Caption = "l/d=" & Int((In6) / ((TP1) - (DCVR) - (DDIA / 2))) _
& " m.f.=" & Int(modFact * 100) / 100
End If

If Val(Label5.Caption) >= Val(Label2.Caption) Then
Text9.Text = "Ast=" & Int(RequiredAst(DFCU, DFYV, DDIA, DCVR, _
TP1, DBMK, 0.9, Val(Label5.Caption))) & "mm2/m."
 Else
  Text9.Text = "Ast=" & Int(RequiredAst(DFCU, DFYV, 3 * DDIA, DCVR, _
   TP1, DBMK, 0.9, Val(Label5.Caption))) & "mm2/m."
    End If
    
Label87.Caption = "Panel_1 min. Ast = " & Int(TP1 * 1000 * _
           (0.24 - (DFYV - 250) * 0.11 / 210) / 100) & " mm2/m."
HighlightPanel
Label1.FontBold = False
Label2.FontBold = False
Label3.FontBold = False
Label4.FontBold = False
Label5.FontBold = True
Label6.FontBold = False
End Sub

Private Sub Label5_DblClick()
Command13_Click
End Sub

Private Sub Label50_Click()
Dim modFact As Double
If In1 / In8 < 2 Then
modFact = 0.55 + (477 - 5 * DFYV / 8) _
          / (120 * (0.9 + ((Val(Label50.Caption) * 1000000#) _
          / 1000 / ((TP9) - (DCVR) - (DDIA / 2)) ^ 2)))
Label71.Caption = "l/d=" & Int((In1) / ((TP9) - (DCVR) - (DDIA / 2))) _
& " m.f.=" & Int(modFact * 100) / 100
End If

Text17.Text = "Ast=" & Int(RequiredAst(DFCU, DFYV, DDIA, DCVR, _
TP9, DBMK, 0.9, Val(Label50.Caption))) & "mm2/m."
Label87.Caption = "Panel_9 min. Ast = " & Int(TP9 * 1000 * _
                   (0.24 - (DFYV - 250) * 0.11 / 210) / 100) & " mm2/m."

HighlightPanel
Label49.FontBold = False
Label50.FontBold = True
Label51.FontBold = False
Label52.FontBold = False
Label53.FontBold = False
Label54.FontBold = False
End Sub

Private Sub Label50_DblClick()
Command13_Click
End Sub

Private Sub Label51_Click()
Text17.Text = "Ast=" & Int(RequiredAst(DFCU, DFYV, DDIA, DCVR, _
TP9, DBMK, 0.9, Val(Label51.Caption))) & "mm2/m."
Label87.Caption = "Panel_9 min. Ast = " & Int(TP9 * 1000 * _
                   (0.24 - (DFYV - 250) * 0.11 / 210) / 100) & " mm2/m."

HighlightPanel
Label49.FontBold = False
Label50.FontBold = False
Label51.FontBold = True
Label52.FontBold = False
Label53.FontBold = False
Label54.FontBold = False
End Sub

Private Sub Label51_DblClick()
Command13_Click
End Sub

Private Sub Label52_Click()
Text17.Text = "Ast=" & Int(RequiredAst(DFCU, DFYV, DDIA, DCVR, _
TP9, DBMK, 0.9, Val(Label52.Caption))) & "mm2/m."
Label87.Caption = "Panel_9 min. Ast = " & Int(TP9 * 1000 * _
                   (0.24 - (DFYV - 250) * 0.11 / 210) / 100) & " mm2/m."

HighlightPanel
Label49.FontBold = False
Label50.FontBold = False
Label51.FontBold = False
Label52.FontBold = True
Label53.FontBold = False
Label54.FontBold = False
End Sub

Private Sub Label52_DblClick()
Command13_Click
End Sub

Private Sub Label53_Click()
Dim modFact As Double
If In8 / In1 < 2 Then
modFact = 0.55 + (477 - 5 * DFYV / 8) _
          / (120 * (0.9 + ((Val(Label53.Caption) * 1000000#) _
          / 1000 / ((TP9) - (DCVR) - (DDIA / 2)) ^ 2)))
Label71.Caption = "l/d=" & Int((In8) / ((TP9) - (DCVR) - (DDIA / 2))) _
& " m.f.=" & Int(modFact * 100) / 100
End If

Text17.Text = "Ast=" & Int(RequiredAst(DFCU, DFYV, DDIA, DCVR, _
TP9, DBMK, 0.9, Val(Label53.Caption))) & "mm2/m."
Label87.Caption = "Panel_9 min. Ast = " & Int(TP9 * 1000 * _
                   (0.24 - (DFYV - 250) * 0.11 / 210) / 100) & " mm2/m."

HighlightPanel
Label49.FontBold = False
Label50.FontBold = False
Label51.FontBold = False
Label52.FontBold = False
Label53.FontBold = True
Label54.FontBold = False
End Sub

Private Sub Label53_DblClick()
Command13_Click
End Sub

Private Sub Label55_Click()
Text18.Text = "Ast=" & Int(RequiredAst(DFCU, DFYV, DDIA, DCVR, _
TP10, DBMK, 0.9, Val(Label55.Caption))) & "mm2/m."
Label87.Caption = "Panel_10 min. Ast = " & Int(TP10 * 1000 * _
                   (0.24 - (DFYV - 250) * 0.11 / 210) / 100) & " mm2/m."

HighlightPanel
Label55.FontBold = True
Label56.FontBold = False
Label57.FontBold = False
Label58.FontBold = False
Label59.FontBold = False
Label60.FontBold = False
End Sub

Private Sub Label55_DblClick()
Command13_Click
End Sub

Private Sub Label56_Click()
Dim modFact As Double
If In2 / In8 < 2 Then
modFact = 0.55 + (477 - 5 * DFYV / 8) _
          / (120 * (0.9 + ((Val(Label56.Caption) * 1000000#) _
          / 1000 / ((TP10) - (DCVR) - (DDIA / 2)) ^ 2)))
Label72.Caption = "l/d=" & Int((In2) / ((TP10) - (DCVR) - (DDIA / 2))) _
& " m.f.=" & Int(modFact * 100) / 100
End If

Text18.Text = "Ast=" & Int(RequiredAst(DFCU, DFYV, DDIA, DCVR, _
TP10, DBMK, 0.9, Val(Label56.Caption))) & "mm2/m."
Label87.Caption = "Panel_10 min. Ast = " & Int(TP10 * 1000 * _
                   (0.24 - (DFYV - 250) * 0.11 / 210) / 100) & " mm2/m."

HighlightPanel
Label55.FontBold = False
Label56.FontBold = True
Label57.FontBold = False
Label58.FontBold = False
Label59.FontBold = False
Label60.FontBold = False
End Sub

Private Sub Label56_DblClick()
Command13_Click
End Sub

Private Sub Label58_Click()
Text18.Text = "Ast=" & Int(RequiredAst(DFCU, DFYV, DDIA, DCVR, _
TP10, DBMK, 0.9, Val(Label58.Caption))) & "mm2/m."
Label87.Caption = "Panel_10 min. Ast = " & Int(TP10 * 1000 * _
                   (0.24 - (DFYV - 250) * 0.11 / 210) / 100) & " mm2/m."

HighlightPanel
Label55.FontBold = False
Label56.FontBold = False
Label57.FontBold = False
Label58.FontBold = True
Label59.FontBold = False
Label60.FontBold = False
End Sub

Private Sub Label58_DblClick()
Command13_Click
End Sub

Private Sub Label59_Click()
Dim modFact As Double
If In8 / In2 < 2 Then
modFact = 0.55 + (477 - 5 * DFYV / 8) _
          / (120 * (0.9 + ((Val(Label59.Caption) * 1000000#) _
          / 1000 / ((TP10) - (DCVR) - (DDIA / 2)) ^ 2)))
Label72.Caption = "l/d=" & Int((In8) / ((TP10) - (DCVR) - (DDIA / 2))) _
& " m.f.=" & Int(modFact * 100) / 100
End If

Text18.Text = "Ast=" & Int(RequiredAst(DFCU, DFYV, DDIA, DCVR, _
TP10, DBMK, 0.9, Val(Label59.Caption))) & "mm2/m."
Label87.Caption = "Panel_10 min. Ast = " & Int(TP10 * 1000 * _
                   (0.24 - (DFYV - 250) * 0.11 / 210) / 100) & " mm2/m."

HighlightPanel
Label55.FontBold = False
Label56.FontBold = False
Label57.FontBold = False
Label58.FontBold = False
Label59.FontBold = True
Label60.FontBold = False
End Sub

Private Sub Label59_DblClick()
Command13_Click
End Sub

Private Sub Label6_Click()
Text9.Text = "Ast=" & Int(RequiredAst(DFCU, DFYV, DDIA, DCVR, _
TP1, DBMK, 0.9, Val(Label6.Caption))) & "mm2/m."
Label87.Caption = "Panel_1 min. Ast = " & Int(TP1 * 1000 * _
                   (0.24 - (DFYV - 250) * 0.11 / 210) / 100) & " mm2/m."

HighlightPanel
Label1.FontBold = False
Label2.FontBold = False
Label3.FontBold = False
Label4.FontBold = False
Label5.FontBold = False
Label6.FontBold = True
End Sub

Private Sub Label6_DblClick()
Command13_Click
End Sub

Private Sub Label61_Click()
Dim modFact As Double
If In5 / In8 < 2 Then
modFact = 0.55 + (477 - 5 * DFYV / 8) _
          / (120 * (0.9 + ((Val(Label61.Caption) * 1000000#) _
          / 1000 / ((TP11) - (DCVR) - (DDIA / 2)) ^ 2)))
Label73.Caption = "l/d=" & Int((In5) / ((TP11) - (DCVR) - (DDIA / 2))) _
& " m.f.=" & Int(modFact * 100) / 100
End If

Text19.Text = "Ast=" & Int(RequiredAst(DFCU, DFYV, DDIA, DCVR, _
TP11, DBMK, 0.9, Val(Label61.Caption))) & "mm2/m."
Label87.Caption = "Panel_11 min. Ast = " & Int(TP11 * 1000 * _
                   (0.24 - (DFYV - 250) * 0.11 / 210) / 100) & " mm2/m."

HighlightPanel
Label61.FontBold = True
Label62.FontBold = False
Label63.FontBold = False
Label64.FontBold = False
Label65.FontBold = False
Label66.FontBold = False
End Sub

Private Sub Label61_DblClick()
Command13_Click
End Sub

Private Sub Label62_Click()
Dim modFact As Double
If In8 / In5 < 2 Then
modFact = 0.55 + (477 - 5 * DFYV / 8) _
          / (120 * (0.9 + ((Val(Label62.Caption) * 1000000#) _
          / 1000 / ((TP11) - (DCVR) - (DDIA / 2)) ^ 2)))
Label73.Caption = "l/d=" & Int((In8) / ((TP11) - (DCVR) - (DDIA / 2))) _
& " m.f.=" & Int(modFact * 100) / 100
End If

Text19.Text = "Ast=" & Int(RequiredAst(DFCU, DFYV, DDIA, DCVR, _
TP11, DBMK, 0.9, Val(Label62.Caption))) & "mm2/m."
Label87.Caption = "Panel_11 min. Ast = " & Int(TP11 * 1000 * _
                   (0.24 - (DFYV - 250) * 0.11 / 210) / 100) & " mm2/m."

HighlightPanel
Label61.FontBold = False
Label62.FontBold = True
Label63.FontBold = False
Label64.FontBold = False
Label65.FontBold = False
Label66.FontBold = False
End Sub

Private Sub Label62_DblClick()
Command13_Click
End Sub

Private Sub Label63_Click()
Label63.Height = 1100
Label63.Left = 1300
Label63.Top = 1350
Label63.Width = 1700
Label63.Caption = "Tbl papak = " & Trim(TP1) _
& "  Beban kemasan = " & Trim(Bk1) & "  Beban hidup = " & Trim(Bh1)
Label1.Visible = False
Label2.Visible = False
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
'''Text9.Text = "Coeff."
End Sub

Private Sub Label64_Click()
Label64.Height = 1100
Label64.Left = 3550
Label64.Top = 1350
Label64.Width = 1700
Label64.Caption = "Tbl papak = " & Trim(TP2) _
& "  Beban kemasan = " & Trim(Bk2) & "  Beban hidup = " & Trim(Bh2)
Label7.Visible = False
Label8.Visible = False
Label9.Visible = False
Label10.Visible = False
Label11.Visible = False
Label12.Visible = False
'''Text10.Text = "Coeff."
End Sub

Private Sub Label65_Click()
Label65.Height = 1100
Label65.Left = 5500
Label65.Top = 1350
Label65.Width = 1700
Label65.Caption = "Tbl papak = " & Trim(TP3) _
& "  Beban kemasan = " & Trim(Bk3) & "  Beban hidup = " & Trim(Bh3)
Label13.Visible = False
Label14.Visible = False
Label15.Visible = False
Label16.Visible = False
Label17.Visible = False
Label18.Visible = False
'''Text11.Text = "Coeff."
End Sub

Private Sub Label66_Click()
Label66.Height = 1100
Label66.Left = 7400
Label66.Top = 1350
Label66.Width = 1700
Label66.Caption = "Tbl papak = " & Trim(TP4) _
& "  Beban kemasan = " & Trim(Bk4) & "  Beban hidup = " & Trim(Bh4)
Label19.Visible = False
Label20.Visible = False
Label21.Visible = False
Label22.Visible = False
Label23.Visible = False
Label24.Visible = False
'''Text12.Text = "Coeff."
End Sub

Private Sub Label67_Click()
Label67.Height = 1100
Label67.Left = 9350
Label67.Top = 1350
Label67.Width = 1700
Label67.Caption = "Tbl papak = " & Trim(TP5) _
& "  Beban kemasan = " & Trim(Bk5) & "  Beban hidup = " & Trim(Bh5)
Label25.Visible = False
Label26.Visible = False
Label27.Visible = False
Label28.Visible = False
Label29.Visible = False
Label30.Visible = False
'''Text13.Text = "Coeff."
End Sub

Private Sub Label68_Click()
Label68.Height = 1100
Label68.Left = 1300
Label68.Top = 3050
Label68.Width = 1700
Label68.Caption = "Tbl papak = " & Trim(TP6) _
& "  Beban kemasan = " & Trim(Bk6) & "  Beban hidup = " & Trim(Bh6)
Label31.Visible = False
Label32.Visible = False
Label33.Visible = False
Label34.Visible = False
Label35.Visible = False
Label36.Visible = False
'''Text14.Text = "Coeff."
End Sub

Private Sub Label69_Click()
Label69.Height = 1100
Label69.Left = 3550
Label69.Top = 3050
Label69.Width = 1700
Label69.Caption = "Tbl papak = " & Trim(TP7) _
& "  Beban kemasan = " & Trim(Bk7) & "  Beban hidup = " & Trim(Bh7)
Label37.Visible = False
Label38.Visible = False
Label39.Visible = False
Label40.Visible = False
Label41.Visible = False
Label42.Visible = False
'''Text15.Text = "Coeff."
End Sub

Private Sub Label7_Click()
Text10.Text = "Ast=" & Int(RequiredAst(DFCU, DFYV, DDIA, DCVR, _
TP2, DBMK, 0.9, Val(Label7.Caption))) & "mm2/m."
Label87.Caption = "Panel_2 min. Ast = " & Int(TP2 * 1000 * _
                   (0.24 - (DFYV - 250) * 0.11 / 210) / 100) & " mm2/m."

HighlightPanel
Label7.FontBold = True
Label8.FontBold = False
Label9.FontBold = False
Label10.FontBold = False
Label11.FontBold = False
Label12.FontBold = False
End Sub

Private Sub Label7_DblClick()
Command13_Click
End Sub

Private Sub Label70_Click()
Label70.Height = 1100
Label70.Left = 5500
Label70.Top = 3050
Label70.Width = 1700
Label70.Caption = "Tbl papak = " & Trim(TP8) _
& "  Beban kemasan = " & Trim(Bk8) & "  Beban hidup = " & Trim(Bh8)
Label43.Visible = False
Label44.Visible = False
Label45.Visible = False
Label46.Visible = False
Label47.Visible = False
Label48.Visible = False
'''Text16.Text = "Coeff."
End Sub

Private Sub Label71_Click()
Label71.Height = 1100
Label71.Left = 1300
Label71.Top = 4600
Label71.Width = 1700
Label71.Caption = "Tbl papak = " & Trim(TP9) _
& "  Beban kemasan = " & Trim(Bk9) & "  Beban hidup = " & Trim(Bh9)
Label49.Visible = False
Label50.Visible = False
Label51.Visible = False
Label52.Visible = False
Label53.Visible = False
Label54.Visible = False
'''Text17.Text = "Coeff."
End Sub

Private Sub Label72_Click()
Label72.Height = 1100
Label72.Left = 3550
Label72.Top = 4600
Label72.Width = 1700
Label72.Caption = "Tbl papak = " & Trim(TP10) _
& "  Beban kemasan = " & Trim(Bk10) & " Beban hidup = " & Trim(Bh10)
Label55.Visible = False
Label56.Visible = False
Label57.Visible = False
Label58.Visible = False
Label59.Visible = False
Label60.Visible = False
'''Text18.Text = "Coeff."
End Sub

Private Sub Label73_Click()
Label73.Height = 1100
Label73.Left = 9350
Label73.Top = 4600
Label73.Width = 1700
Label73.Caption = "Tbl papak = " & Trim(TP11) _
& "  Beban kemasan = " & Trim(Bk11) & " Beban hidup = " & Trim(Bh11)
Label61.Visible = False
Label62.Visible = False
'''Text19.Text = "Coeff."
End Sub



Private Sub Label8_Click()
Dim modFact As Double
If In2 / In6 < 2 Then
modFact = 0.55 + (477 - 5 * DFYV / 8) _
          / (120 * (0.9 + ((Val(Label8.Caption) * 1000000#) _
          / 1000 / ((TP2) - (DCVR) - (DDIA / 2)) ^ 2)))
Label64.Caption = "l/d=" & Int((In2) / ((TP2) - (DCVR) - (DDIA / 2))) _
& " m.f.=" & Int(modFact * 100) / 100
End If

If Val(Label8.Caption) >= Val(Label11.Caption) Then
Text10.Text = "Ast=" & Int(RequiredAst(DFCU, DFYV, DDIA, DCVR, _
TP2, DBMK, 0.9, Val(Label8.Caption))) & "mm2/m."
  Else
   Text10.Text = "Ast=" & Int(RequiredAst(DFCU, DFYV, 3 * DDIA, DCVR, _
    TP2, DBMK, 0.9, Val(Label8.Caption))) & "mm2/m."
   End If
   
Label87.Caption = "Panel_2 min. Ast = " & Int(TP2 * 1000 * _
                   (0.24 - (DFYV - 250) * 0.11 / 210) / 100) & " mm2/m."
HighlightPanel
Label7.FontBold = False
Label8.FontBold = True
Label9.FontBold = False
Label10.FontBold = False
Label11.FontBold = False
Label12.FontBold = False
End Sub

Private Sub Label8_DblClick()
Command13_Click
End Sub



Private Sub Label9_Click()
Text10.Text = "Ast=" & Int(RequiredAst(DFCU, DFYV, DDIA, DCVR, _
TP2, DBMK, 0.9, Val(Label9.Caption))) & "mm2/m."
Label87.Caption = "Panel_2 min. Ast = " & Int(TP2 * 1000 * _
                   (0.24 - (DFYV - 250) * 0.11 / 210) / 100) & " mm2/m."

HighlightPanel
Label7.FontBold = False
Label8.FontBold = False
Label9.FontBold = True
Label10.FontBold = False
Label11.FontBold = False
Label12.FontBold = False
End Sub

Private Sub Label9_DblClick()
Command13_Click
End Sub

Private Sub mnuFileExit_Click()
End
End Sub

Private Sub mnuFileOpenDwg_Click()
Dim chkFile As String
Dim caseNo As Integer
mnuFileExit.Enabled = True
mnuFileOpenDwg = True

    CommonDialog1.Filter = "Dwg files (*.DWG)|*.DWG"
    CommonDialog1.ShowOpen       'display Open dialog box
    dwgName = CommonDialog1.FileName
    Label5.Caption = "..." & Right(dwgName, 50)
    mnuFileExit.Enabled = True
    mnuFileOpenDwg = True
        
If dwgName = "" Then
 mnuFileExit.Enabled = True
 mnuFileOpenDwg.Enabled = True
 
 Form1.Picture = LoadPicture("C:\autodraf\icon\ukad.ico")
Else
 mnuFileExit.Enabled = True
 mnuFileOpenDwg.Enabled = False
 
 Form1.Picture = LoadPicture("C:\autodraf\icon\ukad1.ico")
 ''mnuFile.Enabled = False
End If
    Command1.Enabled = True
End Sub







Private Sub Text1_Change()
If Val(Text1.Text) <= 0 Then
   Text1.Text = 1000
     End If
End Sub

Private Sub Text10_Click()
If Command1.Enabled = True And _
   Command2.Enabled = True And _
   Command3.Enabled = True And _
   Command4.Enabled = True And _
   Command5.Enabled = True Then
Label7.Visible = True
Label8.Visible = True
Label9.Visible = True
Label10.Visible = True
Label11.Visible = True
Label12.Visible = True
Text10.Text = "Coef."
PaneL2
    Else
Label7.Visible = False
Label8.Visible = False
Label9.Visible = False
Label10.Visible = False
Label11.Visible = False
Label12.Visible = False
'''Text10.Text = ""
End If
End Sub

Private Sub Text10_DblClick()
If Command1.Enabled = True And _
   Command2.Enabled = True And _
   Command3.Enabled = True And _
   Command4.Enabled = True And _
   Command5.Enabled = True Then
Label7.Visible = True
Label8.Visible = True
Label9.Visible = True
Label10.Visible = True
Label11.Visible = True
Label12.Visible = True
Text10.Text = "Moment"
    Label7.Caption = Int(10 * Coef7 * Lx2 ^ 2 * (1.4 * (TP2 / 1000 * 24 + Bk2) _
                         + 1.6 * Bh2) / 1000000#) / 10
    Label8.Caption = Int(10 * Coef8 * Lx2 ^ 2 * (1.4 * (TP2 / 1000 * 24 + Bk2) _
                         + 1.6 * Bh2) / 1000000#) / 10
    Label9.Caption = Int(10 * Coef9 * Lx2 ^ 2 * (1.4 * (TP2 / 1000 * 24 + Bk2) _
                         + 1.6 * Bh2) / 1000000#) / 10
    Label10.Caption = Int(10 * Coef10 * Lx2 ^ 2 * (1.4 * (TP2 / 1000 * 24 + Bk2) _
                         + 1.6 * Bh2) / 1000000#) / 10
    Label11.Caption = Int(10 * Coef11 * Lx2 ^ 2 * (1.4 * (TP2 / 1000 * 24 + Bk2) _
                         + 1.6 * Bh2) / 1000000#) / 10
    Label12.Caption = Int(10 * Coef12 * Lx2 ^ 2 * (1.4 * (TP2 / 1000 * 24 + Bk2) _
                         + 1.6 * Bh2) / 1000000#) / 10
Else
Label7.Visible = False
Label8.Visible = False
Label9.Visible = False
Label10.Visible = False
Label11.Visible = False
Label12.Visible = False
'''Text10.Text = ""

End If
End Sub

Private Sub Text10_GotFocus()
Label64.Height = 500
Label64.Left = 4320
Label64.Top = 1920
Label64.Width = 860
If Command5.Enabled = True Then
   Label64.Caption = "tebal papak"
     End If
If Command6.Enabled = True Then
   Label64.Caption = "beban kemasan"
     End If
If Command7.Enabled = True Then
   Label64.Caption = "beban hidup"
     End If

End Sub

Private Sub Text10_LostFocus()
Label64.Height = 500
Label64.Left = 4320
Label64.Top = 1920
Label64.Width = 860
Label64.Caption = Chr(202)

End Sub

Private Sub Text11_Click()
If Command1.Enabled = True And _
   Command2.Enabled = True And _
   Command3.Enabled = True And _
   Command4.Enabled = True And _
   Command5.Enabled = True Then
Label13.Visible = True
Label14.Visible = True
Label15.Visible = True
Label16.Visible = True
Label17.Visible = True
Label18.Visible = True
Text11.Text = "Coef."
PaneL3
    Else
Label13.Visible = False
Label14.Visible = False
Label15.Visible = False
Label16.Visible = False
Label17.Visible = False
Label18.Visible = False
'''Text11.Text = ""
End If
End Sub

Private Sub Text11_DblClick()
If Command1.Enabled = True And _
   Command2.Enabled = True And _
   Command3.Enabled = True And _
   Command4.Enabled = True And _
   Command5.Enabled = True Then
Label13.Visible = True
Label14.Visible = True
Label15.Visible = True
Label16.Visible = True
Label16.Visible = True
Label17.Visible = True
Text11.Text = "Moment"
    Label13.Caption = Int(10 * Coef13 * Lx3 ^ 2 * (1.4 * (TP3 / 1000 * 24 + Bk3) _
                         + 1.6 * Bh3) / 1000000#) / 10
    Label14.Caption = Int(10 * Coef14 * Lx3 ^ 2 * (1.4 * (TP3 / 1000 * 24 + Bk3) _
                         + 1.6 * Bh3) / 1000000#) / 10
    Label15.Caption = Int(10 * Coef15 * Lx3 ^ 2 * (1.4 * (TP3 / 1000 * 24 + Bk3) _
                         + 1.6 * Bh3) / 1000000#) / 10
    Label16.Caption = Int(10 * Coef16 * Lx3 ^ 2 * (1.4 * (TP3 / 1000 * 24 + Bk3) _
                         + 1.6 * Bh3) / 1000000#) / 10
    Label17.Caption = Int(10 * Coef17 * Lx3 ^ 2 * (1.4 * (TP3 / 1000 * 24 + Bk3) _
                         + 1.6 * Bh3) / 1000000#) / 10
    Label18.Caption = Int(10 * Coef18 * Lx3 ^ 2 * (1.4 * (TP3 / 1000 * 24 + Bk3) _
                         + 1.6 * Bh3) / 1000000#) / 10

Else
Label13.Visible = False
Label14.Visible = False
Label15.Visible = False
Label16.Visible = False
Label17.Visible = False
Label18.Visible = False
'''Text11.Text = ""

End If
End Sub

Private Sub Text11_GotFocus()
Label65.Height = 500
Label65.Left = 6369
Label65.Top = 1920
Label65.Width = 860
If Command5.Enabled = True Then
   Label65.Caption = "tebal papak"
     End If
If Command6.Enabled = True Then
   Label65.Caption = "beban kemasan"
     End If
If Command7.Enabled = True Then
   Label65.Caption = "beban hidup"
     End If

End Sub

Private Sub Text11_LostFocus()
Label65.Height = 500
Label65.Left = 6369
Label65.Top = 1920
Label65.Width = 860
Label65.Caption = Chr(212)

End Sub

Private Sub Text12_Click()
If Command1.Enabled = True And _
   Command2.Enabled = True And _
   Command3.Enabled = True And _
   Command4.Enabled = True And _
   Command5.Enabled = True Then
Label19.Visible = True
Label20.Visible = True
Label21.Visible = True
Label22.Visible = True
Label23.Visible = True
Label24.Visible = True
Text12.Text = "Coef."
PaneL4
    Else
Label19.Visible = False
Label20.Visible = False
Label21.Visible = False
Label22.Visible = False
Label23.Visible = False
Label24.Visible = False
'''Text12.Text = ""
End If
End Sub

Private Sub Text12_DblClick()
If Command1.Enabled = True And _
   Command2.Enabled = True And _
   Command3.Enabled = True And _
   Command4.Enabled = True And _
   Command5.Enabled = True Then
Label19.Visible = True
Label20.Visible = True
Label21.Visible = True
Label22.Visible = True
Label23.Visible = True
Label24.Visible = True
Text12.Text = "Moment"
    Label19.Caption = Int(10 * Coef19 * Lx4 ^ 2 * (1.4 * (TP4 / 1000 * 24 + Bk4) _
                         + 1.6 * Bh4) / 1000000#) / 10
    Label20.Caption = Int(10 * Coef20 * Lx4 ^ 2 * (1.4 * (TP4 / 1000 * 24 + Bk4) _
                         + 1.6 * Bh4) / 1000000#) / 10
    Label21.Caption = Int(10 * Coef21 * Lx4 ^ 2 * (1.4 * (TP4 / 1000 * 24 + Bk4) _
                         + 1.6 * Bh4) / 1000000#) / 10
    Label22.Caption = Int(10 * Coef22 * Lx4 ^ 2 * (1.4 * (TP4 / 1000 * 24 + Bk4) _
                         + 1.6 * Bh4) / 1000000#) / 10
    Label23.Caption = Int(10 * Coef23 * Lx4 ^ 2 * (1.4 * (TP4 / 1000 * 24 + Bk4) _
                         + 1.6 * Bh4) / 1000000#) / 10
    Label24.Caption = Int(10 * Coef24 * Lx4 ^ 2 * (1.4 * (TP4 / 1000 * 24 + Bk4) _
                         + 1.6 * Bh4) / 1000000#) / 10

Else
Label19.Visible = False
Label20.Visible = False
Label21.Visible = False
Label22.Visible = False
Label23.Visible = False
Label24.Visible = False
'''Text12.Text = ""

End If
End Sub

Private Sub Text12_GotFocus()
Label66.Height = 500
Label66.Left = 8280
Label66.Top = 1920
Label66.Width = 860
If Command5.Enabled = True Then
   Label66.Caption = "tebal papak"
     End If
If Command6.Enabled = True Then
   Label66.Caption = "beban kemasan"
     End If
If Command7.Enabled = True Then
   Label66.Caption = "beban hidup"
     End If

End Sub

Private Sub Text12_LostFocus()
Label66.Height = 500
Label66.Left = 8280
Label66.Top = 1920
Label66.Width = 860
Label66.Caption = Chr(90)

End Sub

Private Sub Text13_Click()
If Command1.Enabled = True And _
   Command2.Enabled = True And _
   Command3.Enabled = True And _
   Command4.Enabled = True And _
   Command5.Enabled = True Then
Label25.Visible = True
Label26.Visible = True
Label27.Visible = True
Label28.Visible = True
Label29.Visible = True
Label30.Visible = True
Text13.Text = "Coef."
PaneL5
    Else
Label25.Visible = False
Label26.Visible = False
Label27.Visible = False
Label28.Visible = False
Label29.Visible = False
Label30.Visible = False
'''Text13.Text = ""
End If
End Sub

Private Sub Text13_DblClick()
If Command1.Enabled = True And _
   Command2.Enabled = True And _
   Command3.Enabled = True And _
   Command4.Enabled = True And _
   Command5.Enabled = True Then
Label25.Visible = True
Label26.Visible = True
Label27.Visible = True
Label28.Visible = True
Label29.Visible = True
Label30.Visible = True
Text13.Text = "Moment"
    Label25.Caption = Int(10 * Coef25 * Lx5 ^ 2 * (1.4 * (TP5 / 1000 * 24 + Bk5) _
                         + 1.6 * Bh5) / 1000000#) / 10
    Label26.Caption = Int(10 * Coef26 * Lx5 ^ 2 * (1.4 * (TP5 / 1000 * 24 + Bk5) _
                         + 1.6 * Bh5) / 1000000#) / 10
    Label27.Caption = Int(10 * Coef27 * Lx5 ^ 2 * (1.4 * (TP5 / 1000 * 24 + Bk5) _
                         + 1.6 * Bh5) / 1000000#) / 10
    Label28.Caption = Int(10 * Coef28 * Lx5 ^ 2 * (1.4 * (TP5 / 1000 * 24 + Bk5) _
                         + 1.6 * Bh5) / 1000000#) / 10
    Label29.Caption = Int(10 * Coef29 * Lx5 ^ 2 * (1.4 * (TP5 / 1000 * 24 + Bk5) _
                         + 1.6 * Bh5) / 1000000#) / 10
    Label30.Caption = Int(10 * Coef30 * Lx5 ^ 2 * (1.4 * (TP5 / 1000 * 24 + Bk5) _
                         + 1.6 * Bh5) / 1000000#) / 10

Else
Label25.Visible = False
Label26.Visible = False
Label27.Visible = False
Label28.Visible = False
Label29.Visible = False
Label30.Visible = False
'''Text13.Text = ""

End If
End Sub

Private Sub Text13_GotFocus()
Label67.Height = 500
Label67.Left = 10200
Label67.Top = 1920
Label67.Width = 860
If Command5.Enabled = True Then
   Label67.Caption = "tebal papak"
     End If
If Command6.Enabled = True Then
   Label67.Caption = "beban kemasan"
     End If
If Command7.Enabled = True Then
   Label67.Caption = "beban hidup"
     End If

End Sub

Private Sub Text13_LostFocus()
Label67.Height = 500
Label67.Left = 10200
Label67.Top = 1920
Label67.Width = 860
Label67.Caption = Chr(77)

End Sub

Private Sub Text14_Click()
If Command1.Enabled = True And _
   Command2.Enabled = True And _
   Command3.Enabled = True And _
   Command4.Enabled = True And _
   Command5.Enabled = True Then
Label31.Visible = True
Label32.Visible = True
Label33.Visible = True
Label34.Visible = True
Label35.Visible = True
Label36.Visible = True
Text14.Text = "Coef."
PaneL6
    Else
Label31.Visible = False
Label32.Visible = False
Label33.Visible = False
Label34.Visible = False
Label35.Visible = False
Label36.Visible = False
'''Text14.Text = ""
End If
End Sub

Private Sub Text14_DblClick()
If Command1.Enabled = True And _
   Command2.Enabled = True And _
   Command3.Enabled = True And _
   Command4.Enabled = True And _
   Command5.Enabled = True Then
Label31.Visible = True
Label32.Visible = True
Label33.Visible = True
Label34.Visible = True
Label35.Visible = True
Label36.Visible = True
Text14.Text = "Moment"
    Label31.Caption = Int(10 * Coef31 * Lx6 ^ 2 * (1.4 * (TP6 / 1000 * 24 + Bk6) _
                         + 1.6 * Bh6) / 1000000#) / 10
    Label32.Caption = Int(10 * Coef32 * Lx6 ^ 2 * (1.4 * (TP6 / 1000 * 24 + Bk6) _
                         + 1.6 * Bh6) / 1000000#) / 10
    Label33.Caption = Int(10 * Coef33 * Lx6 ^ 2 * (1.4 * (TP6 / 1000 * 24 + Bk6) _
                         + 1.6 * Bh6) / 1000000#) / 10
    Label34.Caption = Int(10 * Coef34 * Lx6 ^ 2 * (1.4 * (TP6 / 1000 * 24 + Bk6) _
                         + 1.6 * Bh6) / 1000000#) / 10
    Label35.Caption = Int(10 * Coef35 * Lx6 ^ 2 * (1.4 * (TP6 / 1000 * 24 + Bk6) _
                         + 1.6 * Bh6) / 1000000#) / 10
    Label36.Caption = Int(10 * Coef36 * Lx6 ^ 2 * (1.4 * (TP6 / 1000 * 24 + Bk6) _
                         + 1.6 * Bh6) / 1000000#) / 10

Else
Label31.Visible = False
Label32.Visible = False
Label33.Visible = False
Label34.Visible = False
Label35.Visible = False
Label36.Visible = False
'''Text14.Text = ""

End If
End Sub

Private Sub Text14_GotFocus()
Label68.Height = 500
Label68.Left = 2400
Label68.Top = 3480
Label68.Width = 860
If Command5.Enabled = True Then
   Label68.Caption = "tebal papak"
     End If
If Command6.Enabled = True Then
   Label68.Caption = "beban kemasan"
     End If
If Command7.Enabled = True Then
   Label68.Caption = "beban hidup"
     End If

End Sub

Private Sub Text14_LostFocus()
Label68.Height = 500
Label68.Left = 2400
Label68.Top = 3480
Label68.Width = 860
Label68.Caption = Chr(250)

End Sub

Private Sub Text15_Click()
If Command1.Enabled = True And _
   Command2.Enabled = True And _
   Command3.Enabled = True And _
   Command4.Enabled = True And _
   Command5.Enabled = True Then
Label37.Visible = True
Label38.Visible = True
Label39.Visible = True
Label40.Visible = True
Label41.Visible = True
Label42.Visible = True
Text15.Text = "Coef."
PaneL7
    Else
Label37.Visible = False
Label38.Visible = False
Label39.Visible = False
Label40.Visible = False
Label41.Visible = False
Label42.Visible = False
'''Text15.Text = ""
End If
End Sub

Private Sub Text15_DblClick()
If Command1.Enabled = True And _
   Command2.Enabled = True And _
   Command3.Enabled = True And _
   Command4.Enabled = True And _
   Command5.Enabled = True Then
Label37.Visible = True
Label38.Visible = True
Label39.Visible = True
Label40.Visible = True
Label41.Visible = True
Label42.Visible = True
Text15.Text = "Moment"
    Label37.Caption = Int(10 * Coef37 * Lx7 ^ 2 * (1.4 * (TP7 / 1000 * 24 + Bk7) _
                         + 1.6 * Bh7) / 1000000#) / 10
    Label38.Caption = Int(10 * Coef38 * Lx7 ^ 2 * (1.4 * (TP7 / 1000 * 24 + Bk7) _
                         + 1.6 * Bh7) / 1000000#) / 10
    Label39.Caption = Int(10 * Coef39 * Lx7 ^ 2 * (1.4 * (TP7 / 1000 * 24 + Bk7) _
                         + 1.6 * Bh7) / 1000000#) / 10
    Label40.Caption = Int(10 * Coef40 * Lx7 ^ 2 * (1.4 * (TP7 / 1000 * 24 + Bk7) _
                         + 1.6 * Bh7) / 1000000#) / 10
    Label41.Caption = Int(10 * Coef41 * Lx7 ^ 2 * (1.4 * (TP7 / 1000 * 24 + Bk7) _
                         + 1.6 * Bh7) / 1000000#) / 10
    Label42.Caption = Int(10 * Coef42 * Lx7 ^ 2 * (1.4 * (TP7 / 1000 * 24 + Bk7) _
                         + 1.6 * Bh7) / 1000000#) / 10

Else
Label37.Visible = False
Label38.Visible = False
Label39.Visible = False
Label40.Visible = False
Label41.Visible = False
Label42.Visible = False
'''Text15.Text = ""

End If
End Sub

Private Sub Text15_GotFocus()
Label69.Height = 500
Label69.Left = 4320
Label69.Top = 3480
Label69.Width = 860
If Command5.Enabled = True Then
   Label69.Caption = "tebal papak"
     End If
If Command6.Enabled = True Then
   Label69.Caption = "beban kemasan"
     End If
If Command7.Enabled = True Then
   Label69.Caption = "beban hidup"
     End If

End Sub

Private Sub Text15_LostFocus()
Label69.Height = 500
Label69.Left = 4320
Label69.Top = 3480
Label69.Width = 860
Label69.Caption = Chr(225)

End Sub

Private Sub Text16_Click()
If Command1.Enabled = True And _
   Command2.Enabled = True And _
   Command3.Enabled = True And _
   Command4.Enabled = True And _
   Command5.Enabled = True Then
Label43.Visible = True
Label44.Visible = True
Label45.Visible = True
Label46.Visible = True
Label47.Visible = True
Label48.Visible = True
Text16.Text = "Coef."
PaneL8
    Else
Label43.Visible = False
Label44.Visible = False
Label45.Visible = False
Label46.Visible = False
Label47.Visible = False
Label48.Visible = False
'''Text16.Text = ""
End If
End Sub

Private Sub Text16_DblClick()
If Command1.Enabled = True And _
   Command2.Enabled = True And _
   Command3.Enabled = True And _
   Command4.Enabled = True And _
   Command5.Enabled = True Then
Label43.Visible = True
Label44.Visible = True
Label45.Visible = True
Label46.Visible = True
Label47.Visible = True
Label48.Visible = True
Text16.Text = "Moment"
    Label43.Caption = Int(10 * Coef43 * Lx8 ^ 2 * (1.4 * (TP8 / 1000 * 24 + Bk8) _
                         + 1.6 * Bh8) / 1000000#) / 10
    Label44.Caption = Int(10 * Coef44 * Lx8 ^ 2 * (1.4 * (TP8 / 1000 * 24 + Bk8) _
                         + 1.6 * Bh8) / 1000000#) / 10
    Label45.Caption = Int(10 * Coef45 * Lx8 ^ 2 * (1.4 * (TP8 / 1000 * 24 + Bk8) _
                         + 1.6 * Bh8) / 1000000#) / 10
    Label46.Caption = Int(10 * Coef46 * Lx8 ^ 2 * (1.4 * (TP8 / 1000 * 24 + Bk8) _
                         + 1.6 * Bh8) / 1000000#) / 10
    Label47.Caption = Int(10 * Coef47 * Lx8 ^ 2 * (1.4 * (TP8 / 1000 * 24 + Bk8) _
                         + 1.6 * Bh8) / 1000000#) / 10
    Label48.Caption = Int(10 * Coef48 * Lx8 ^ 2 * (1.4 * (TP8 / 1000 * 24 + Bk8) _
                         + 1.6 * Bh8) / 1000000#) / 10

Else
Label43.Visible = False
Label44.Visible = False
Label45.Visible = False
Label46.Visible = False
Label47.Visible = False
Label48.Visible = False
'''Text16.Text = ""

End If
End Sub

Private Sub Text16_GotFocus()
Label70.Height = 500
Label70.Left = 6369
Label70.Top = 3480
Label70.Width = 860
If Command5.Enabled = True Then
   Label70.Caption = "tebal papak"
     End If
If Command6.Enabled = True Then
   Label70.Caption = "beban kemasan"
     End If
If Command7.Enabled = True Then
   Label70.Caption = "beban hidup"
     End If

End Sub

Private Sub Text16_LostFocus()
Label70.Height = 500
Label70.Left = 6369
Label70.Top = 3480
Label70.Width = 860
Label70.Caption = Chr(150)

End Sub

Private Sub Text17_Click()
If Command1.Enabled = True And _
   Command2.Enabled = True And _
   Command3.Enabled = True And _
   Command4.Enabled = True And _
   Command5.Enabled = True Then
Label49.Visible = True
Label50.Visible = True
Label51.Visible = True
Label52.Visible = True
Label53.Visible = True
Label54.Visible = True
Text17.Text = "Coef."
PaneL9
    Else
Label49.Visible = False
Label50.Visible = False
Label51.Visible = False
Label52.Visible = False
Label53.Visible = False
Label54.Visible = False
'''Text17.Text = ""
End If
End Sub

Private Sub Text17_DblClick()
If Command1.Enabled = True And _
   Command2.Enabled = True And _
   Command3.Enabled = True And _
   Command4.Enabled = True And _
   Command5.Enabled = True Then
Label49.Visible = True
Label50.Visible = True
Label51.Visible = True
Label52.Visible = True
Label53.Visible = True
Label54.Visible = True
Text17.Text = "Moment"
    Label49.Caption = Int(10 * Coef49 * Lx9 ^ 2 * (1.4 * (TP9 / 1000 * 24 + Bk9) _
                         + 1.6 * Bh9) / 1000000#) / 10
    Label50.Caption = Int(10 * Coef50 * Lx9 ^ 2 * (1.4 * (TP9 / 1000 * 24 + Bk9) _
                         + 1.6 * Bh9) / 1000000#) / 10
    Label51.Caption = Int(10 * Coef51 * Lx9 ^ 2 * (1.4 * (TP9 / 1000 * 24 + Bk9) _
                         + 1.6 * Bh9) / 1000000#) / 10
    Label52.Caption = Int(10 * Coef52 * Lx9 ^ 2 * (1.4 * (TP9 / 1000 * 24 + Bk9) _
                         + 1.6 * Bh9) / 1000000#) / 10
    Label53.Caption = Int(10 * Coef53 * Lx9 ^ 2 * (1.4 * (TP9 / 1000 * 24 + Bk9) _
                         + 1.6 * Bh9) / 1000000#) / 10
    Label54.Caption = Int(10 * Coef54 * Lx9 ^ 2 * (1.4 * (TP9 / 1000 * 24 + Bk9) _
                         + 1.6 * Bh9) / 1000000#) / 10

Else
Label49.Visible = False
Label50.Visible = False
Label51.Visible = False
Label52.Visible = False
Label53.Visible = False
Label54.Visible = False
'''Text17.Text = ""

End If
End Sub

Private Sub Text17_GotFocus()
Label71.Height = 500
Label71.Left = 2400
Label71.Top = 5040
Label71.Width = 860
If Command5.Enabled = True Then
   Label71.Caption = "tebal papak"
     End If
If Command6.Enabled = True Then
   Label71.Caption = "beban kemasan"
     End If
If Command7.Enabled = True Then
   Label71.Caption = "beban hidup"
     End If

End Sub

Private Sub Text17_LostFocus()
Label71.Height = 500
Label71.Left = 2400
Label71.Top = 5040
Label71.Width = 860
Label71.Caption = Chr(250)

End Sub

Private Sub Text18_Click()
If Command1.Enabled = True And _
   Command2.Enabled = True And _
   Command3.Enabled = True And _
   Command4.Enabled = True And _
   Command5.Enabled = True Then
Label55.Visible = True
Label56.Visible = True
Label57.Visible = True
Label58.Visible = True
Label59.Visible = True
Label60.Visible = True
Text18.Text = "Coef."
PaneL10
    Else
Label55.Visible = False
Label56.Visible = False
Label57.Visible = False
Label58.Visible = False
Label59.Visible = False
Label60.Visible = False
'''Text18.Text = ""
End If
End Sub

Private Sub Text18_DblClick()
If Command1.Enabled = True And _
   Command2.Enabled = True And _
   Command3.Enabled = True And _
   Command4.Enabled = True And _
   Command5.Enabled = True Then
Label55.Visible = True
Label56.Visible = True
Label57.Visible = True
Label58.Visible = True
Label59.Visible = True
Label60.Visible = True
Text18.Text = "Moment"
    Label55.Caption = Int(10 * Coef55 * Lx10 ^ 2 * (1.4 * (TP10 / 1000 * 24 + Bk10) _
                         + 1.6 * Bh10) / 1000000#) / 10
    Label56.Caption = Int(10 * Coef56 * Lx10 ^ 2 * (1.4 * (TP10 / 1000 * 24 + Bk10) _
                         + 1.6 * Bh10) / 1000000#) / 10
    Label57.Caption = Int(10 * Coef57 * Lx10 ^ 2 * (1.4 * (TP10 / 1000 * 24 + Bk10) _
                         + 1.6 * Bh10) / 1000000#) / 10
    Label58.Caption = Int(10 * Coef58 * Lx10 ^ 2 * (1.4 * (TP10 / 1000 * 24 + Bk10) _
                         + 1.6 * Bh10) / 1000000#) / 10
    Label59.Caption = Int(10 * Coef59 * Lx10 ^ 2 * (1.4 * (TP10 / 1000 * 24 + Bk10) _
                         + 1.6 * Bh10) / 1000000#) / 10
    Label60.Caption = Int(10 * Coef60 * Lx10 ^ 2 * (1.4 * (TP10 / 1000 * 24 + Bk10) _
                         + 1.6 * Bh10) / 1000000#) / 10

Else
Label55.Visible = False
Label56.Visible = False
Label57.Visible = False
Label58.Visible = False
Label59.Visible = False
Label60.Visible = False
'''Text18.Text = ""

End If
End Sub

Private Sub Text18_GotFocus()
Label72.Height = 500
Label72.Left = 4320
Label72.Top = 5040
Label72.Width = 860
If Command5.Enabled = True Then
   Label72.Caption = "tebal papak"
     End If
If Command6.Enabled = True Then
   Label72.Caption = "beban kemasan"
     End If
If Command7.Enabled = True Then
   Label72.Caption = "beban hidup"
     End If

End Sub

Private Sub Text18_LostFocus()
Label72.Height = 500
Label72.Left = 4320
Label72.Top = 5040
Label72.Width = 860
Label72.Caption = Chr(110)

End Sub

Private Sub Text19_Click()
If Command1.Enabled = True And _
   Command2.Enabled = True And _
   Command3.Enabled = True And _
   Command4.Enabled = True And _
   Command5.Enabled = True Then
Label61.Visible = True
Label62.Visible = True
Text19.Text = "Coef."
PaneL11
    Else
Label61.Visible = False
Label62.Visible = False
'''Text19.Text = ""
End If
End Sub

Private Sub Text19_DblClick()
If Command1.Enabled = True And _
   Command2.Enabled = True And _
   Command3.Enabled = True And _
   Command4.Enabled = True And _
   Command5.Enabled = True Then
Label61.Visible = True
Label62.Visible = True
Text19.Text = "Moment"
    Label61.Caption = Int(10 * Coef61 * Lx11 ^ 2 * (1.4 * (TP11 / 1000 * 24 + Bk11) _
                         + 1.6 * Bh11) / 1000000#) / 10
    Label62.Caption = Int(10 * Coef62 * Lx11 ^ 2 * (1.4 * (TP11 / 1000 * 24 + Bk11) _
                         + 1.6 * Bh11) / 1000000#) / 10
    
Else
Label61.Visible = False
Label62.Visible = False
'''Text19.Text = ""

End If
End Sub

Private Sub Text19_GotFocus()
Label73.Height = 500
Label73.Left = 10200
Label73.Top = 5040
Label73.Width = 860
If Command5.Enabled = True Then
   Label73.Caption = "tebal papak"
     End If
If Command6.Enabled = True Then
   Label73.Caption = "beban kemasan"
     End If
If Command7.Enabled = True Then
   Label73.Caption = "beban hidup"
     End If

End Sub

Private Sub Text19_LostFocus()
Label73.Height = 500
Label73.Left = 10200
Label73.Top = 5040
Label73.Width = 860
Label73.Caption = Chr(75)

End Sub

Private Sub Text2_Change()
If Val(Text2.Text) <= 0 Then
   Text2.Text = 1000
     End If
End Sub

Private Sub Text3_Change()
If Val(Text3.Text) <= 0 Then
   Text3.Text = 1000
     End If
End Sub

Private Sub Text4_Change()
If Val(Text4.Text) <= 0 Then
   Text4.Text = 1000
     End If
End Sub

Private Sub Text5_Change()
If Val(Text5.Text) <= 0 Then
   Text5.Text = 1000
     End If
End Sub

Private Sub Text6_Change()
If Val(Text6.Text) <= 0 Then
   Text6.Text = 1000
     End If
End Sub

Private Sub Text7_Change()
If Val(Text7.Text) <= 0 Then
   Text7.Text = 1000
     End If
End Sub

Private Sub Text8_Change()
If Val(Text8.Text) <= 0 Then
   Text8.Text = 1000
     End If
End Sub

Private Sub Text9_Click()
If Command1.Enabled = True And _
   Command2.Enabled = True And _
   Command3.Enabled = True And _
   Command4.Enabled = True And _
   Command5.Enabled = True Then
Label1.Visible = True
Label2.Visible = True
Label3.Visible = True
Label4.Visible = True
Label5.Visible = True
Label6.Visible = True
Text9.Text = "Coef."
PaneL1
    Else
Label1.Visible = False
Label2.Visible = False
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
'''Text9.Text = ""
End If

End Sub

Private Sub Text9_DblClick()
If Command1.Enabled = True And _
   Command2.Enabled = True And _
   Command3.Enabled = True And _
   Command4.Enabled = True And _
   Command5.Enabled = True Then
Label1.Visible = True
Label2.Visible = True
Label3.Visible = True
Label4.Visible = True
Label5.Visible = True
Label6.Visible = True
Text9.Text = "Moment"
    Label1.Caption = Int(10 * Coef1 * Lx1 ^ 2 * (1.4 * (TP1 / 1000 * 24 + Bk1) _
                         + 1.6 * Bh1) / 1000000#) / 10
    Label2.Caption = Int(10 * Coef2 * Lx1 ^ 2 * (1.4 * (TP1 / 1000 * 24 + Bk1) _
                         + 1.6 * Bh1) / 1000000#) / 10
    Label3.Caption = Int(10 * Coef3 * Lx1 ^ 2 * (1.4 * (TP1 / 1000 * 24 + Bk1) _
                         + 1.6 * Bh1) / 1000000#) / 10
    Label4.Caption = Int(10 * Coef4 * Lx1 ^ 2 * (1.4 * (TP1 / 1000 * 24 + Bk1) _
                         + 1.6 * Bh1) / 1000000#) / 10
    Label5.Caption = Int(10 * Coef5 * Lx1 ^ 2 * (1.4 * (TP1 / 1000 * 24 + Bk1) _
                         + 1.6 * Bh1) / 1000000#) / 10
    Label6.Caption = Int(10 * Coef6 * Lx1 ^ 2 * (1.4 * (TP1 / 1000 * 24 + Bk1) _
                         + 1.6 * Bh1) / 1000000#) / 10

Else
Label1.Visible = False
Label2.Visible = False
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
'''Text9.Text = ""

End If
End Sub

Private Sub Text9_GotFocus()
Label63.Height = 500
Label63.Left = 2400
Label63.Top = 1920
Label63.Width = 860
If Command5.Enabled = True Then
   Label63.Caption = "tebal papak"
     End If
If Command6.Enabled = True Then
   Label63.Caption = "beban kemasan"
     End If
If Command7.Enabled = True Then
   Label63.Caption = "beban hidup"
     End If
End Sub
Private Sub Text9_LostFocus()
Label63.Height = 500
Label63.Left = 2400
Label63.Top = 1920
Label63.Width = 860
Label63.Caption = Chr(250)

End Sub

Private Sub HighlightRentang()

Text1.Visible = True
Text2.Visible = True
Text3.Visible = True
Text4.Visible = True
Text5.Visible = True
Text6.Visible = True
Text7.Visible = True
Text8.Visible = True
 

End Sub

Private Sub HideRentang()

Text1.Visible = False
Text2.Visible = False
Text3.Visible = False
Text4.Visible = False
Text5.Visible = False
Text6.Visible = False
Text7.Visible = False
Text8.Visible = False
 
End Sub


Private Sub HighlightGridInput()

Text20.Visible = True
Text21.Visible = True
Text22.Visible = True
Text23.Visible = True
Text24.Visible = True
Text25.Visible = True
Text26.Visible = True
Text27.Visible = True
Text28.Visible = True
Text29.Visible = True


End Sub


Private Sub HideGridInput()

Text20.Visible = False
Text21.Visible = False
Text22.Visible = False
Text23.Visible = False
Text24.Visible = False
Text25.Visible = False
Text26.Visible = False
Text27.Visible = False
Text28.Visible = False
Text29.Visible = False


End Sub

Private Sub HighlightCircleGrid()

Shape30.Visible = True
Shape31.Visible = True
Shape32.Visible = True
Shape33.Visible = True
Shape34.Visible = True
Shape35.Visible = True
Shape36.Visible = True
Shape37.Visible = True
Shape38.Visible = True
Shape39.Visible = True

End Sub
Private Sub HideCircleGrid()

Shape30.Visible = False
Shape31.Visible = False
Shape32.Visible = False
Shape33.Visible = False
Shape34.Visible = False
Shape35.Visible = False
Shape36.Visible = False
Shape37.Visible = False
Shape38.Visible = False
Shape39.Visible = False

End Sub


Private Sub HighlightBeban()

Text9.Visible = True
Text10.Visible = True
Text11.Visible = True
Text12.Visible = True
Text13.Visible = True
Text14.Visible = True
Text15.Visible = True
Text16.Visible = True
Text17.Visible = True
Text18.Visible = True
Text19.Visible = True

End Sub
Private Sub HideBeban()

Text9.Visible = False
Text10.Visible = False
Text11.Visible = False
Text12.Visible = False
Text13.Visible = False
Text14.Visible = False
Text15.Visible = False
Text16.Visible = False
Text17.Visible = False
Text18.Visible = False
Text19.Visible = False

End Sub
Private Sub EnableTxtNineToNineteen()

Text9.Enabled = True
Text10.Enabled = True
Text11.Enabled = True
Text12.Enabled = True
Text13.Enabled = True
Text14.Enabled = True
Text15.Enabled = True
Text16.Enabled = True
Text17.Enabled = True
Text18.Enabled = True
Text19.Enabled = True

End Sub
Private Sub DisableTxtNineToNineteen()

Text9.Enabled = False
Text10.Enabled = False
Text11.Enabled = False
Text12.Enabled = False
Text13.Enabled = False
Text14.Enabled = False
Text15.Enabled = False
Text16.Enabled = False
Text17.Enabled = False
Text18.Enabled = False
Text19.Enabled = False

End Sub



Private Sub HighlightCoefficient()

Label1.Visible = True
Label2.Visible = True
Label3.Visible = True
Label4.Visible = True
Label5.Visible = True
Label6.Visible = True
Label7.Visible = True
Label8.Visible = True
Label9.Visible = True
Label10.Visible = True
Label11.Visible = True
Label12.Visible = True
Label13.Visible = True
Label14.Visible = True
Label15.Visible = True
Label16.Visible = True
Label17.Visible = True
Label18.Visible = True
Label19.Visible = True
Label20.Visible = True
Label21.Visible = True
Label22.Visible = True
Label23.Visible = True
Label24.Visible = True
Label25.Visible = True
Label26.Visible = True
Label27.Visible = True
Label28.Visible = True
Label29.Visible = True
Label30.Visible = True
Label31.Visible = True
Label32.Visible = True
Label33.Visible = True
Label34.Visible = True
Label35.Visible = True
Label36.Visible = True
Label37.Visible = True
Label38.Visible = True
Label39.Visible = True
Label40.Visible = True
Label41.Visible = True
Label42.Visible = True
Label43.Visible = True
Label44.Visible = True
Label45.Visible = True
Label46.Visible = True
Label47.Visible = True
Label48.Visible = True
Label49.Visible = True
Label50.Visible = True
Label51.Visible = True
Label52.Visible = True
Label53.Visible = True
Label54.Visible = True
Label55.Visible = True
Label56.Visible = True
Label57.Visible = True
Label58.Visible = True
Label59.Visible = True
Label60.Visible = True
Label61.Visible = True
Label62.Visible = True

End Sub
'''''''''''''''''''''''''''

Private Sub HideCoefficient()

Label1.Visible = False
Label2.Visible = False
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
Label7.Visible = False
Label8.Visible = False
Label9.Visible = False
Label10.Visible = False
Label11.Visible = False
Label12.Visible = False
Label13.Visible = False
Label14.Visible = False
Label15.Visible = False
Label16.Visible = False
Label17.Visible = False
Label18.Visible = False
Label19.Visible = False
Label20.Visible = False
Label21.Visible = False
Label22.Visible = False
Label23.Visible = False
Label24.Visible = False
Label25.Visible = False
Label26.Visible = False
Label27.Visible = False
Label28.Visible = False
Label29.Visible = False
Label30.Visible = False
Label31.Visible = False
Label32.Visible = False
Label33.Visible = False
Label34.Visible = False
Label35.Visible = False
Label36.Visible = False
Label37.Visible = False
Label38.Visible = False
Label39.Visible = False
Label40.Visible = False
Label41.Visible = False
Label42.Visible = False
Label43.Visible = False
Label44.Visible = False
Label45.Visible = False
Label46.Visible = False
Label47.Visible = False
Label48.Visible = False
Label49.Visible = False
Label50.Visible = False
Label51.Visible = False
Label52.Visible = False
Label53.Visible = False
Label54.Visible = False
Label55.Visible = False
Label56.Visible = False
Label57.Visible = False
Label58.Visible = False
Label59.Visible = False
Label60.Visible = False
Label61.Visible = False
Label62.Visible = False


End Sub

'''''''''''''''''''''''''''
Private Sub HighlightLabel_Beban()
Label63.Visible = True
Label64.Visible = True
Label65.Visible = True
Label66.Visible = True
Label67.Visible = True
Label68.Visible = True
Label69.Visible = True
Label70.Visible = True
Label71.Visible = True
Label72.Visible = True
Label73.Visible = True
 

End Sub

Private Sub HideLabel_Beban()
Label63.Visible = False
Label64.Visible = False
Label65.Visible = False
Label66.Visible = False
Label67.Visible = False
Label68.Visible = False
Label69.Visible = False
Label70.Visible = False
Label71.Visible = False
Label72.Visible = False
Label73.Visible = False

End Sub
Private Sub EnableLabel_Beban()
Label63.Enabled = True
Label64.Enabled = True
Label65.Enabled = True
Label66.Enabled = True
Label67.Enabled = True
Label68.Enabled = True
Label69.Enabled = True
Label70.Enabled = True
Label71.Enabled = True
Label72.Enabled = True
Label73.Enabled = True

End Sub
Private Sub DisableLabel_Beban()
Label63.Enabled = False
Label64.Enabled = False
Label65.Enabled = False
Label66.Enabled = False
Label67.Enabled = False
Label68.Enabled = False
Label69.Enabled = False
Label70.Enabled = False
Label71.Enabled = False
Label72.Enabled = False
Label73.Enabled = False

End Sub
Private Sub HighlightLabel_Grid()

Label74.Visible = True
Label75.Visible = True
Label76.Visible = True
Label77.Visible = True
Label78.Visible = True
Label79.Visible = True
Label80.Visible = True
Label81.Visible = True
Label82.Visible = True
Label83.Visible = True


Label74.Caption = LG1
Label75.Caption = LG2
Label76.Caption = LG3
Label77.Caption = LG4
Label78.Caption = LG5
Label79.Caption = LG6
Label80.Caption = LG7
Label81.Caption = LG8
Label82.Caption = LG9
Label83.Caption = LG10

Shape30.Shape = 3
Shape31.Shape = 3
Shape32.Shape = 3
Shape33.Shape = 3
Shape34.Shape = 3
Shape35.Shape = 3
Shape36.Shape = 3
Shape37.Shape = 3
Shape38.Shape = 3
Shape39.Shape = 3

End Sub

Private Sub HideLabel_Grid()

Label74.Visible = False
Label75.Visible = False
Label76.Visible = False
Label77.Visible = False
Label78.Visible = False
Label79.Visible = False
Label80.Visible = False
Label81.Visible = False
Label82.Visible = False
Label83.Visible = False


End Sub

Private Sub TajukBeban()

Text9.Text = "Beban1"
Text10.Text = "Beban2"
Text11.Text = "Beban3"
Text12.Text = "Beban4"
Text13.Text = "Beban5"
Text14.Text = "Beban6"
Text15.Text = "Beban7"
Text16.Text = "Beban8"
Text17.Text = "Beban9"
Text18.Text = "Beban10"
Text19.Text = "Beban11"

End Sub
Private Sub TajukCoefficient()

Text9.Text = "Coeff-1"
Text10.Text = "Coeff-2"
Text11.Text = "Coeff-3"
Text12.Text = "Coeff-4"
Text13.Text = "Coeff-5"
Text14.Text = "Coeff-6"
Text15.Text = "Coeff-7"
Text16.Text = "Coeff-8"
Text17.Text = "Coeff-9"
Text18.Text = "Coeff-10"
Text19.Text = "Coeff-11"

End Sub
Private Sub TajukMoment()

Text9.Text = "Moment-1"
Text10.Text = "Moment-2"
Text11.Text = "Moment-3"
Text12.Text = "Moment-4"
Text13.Text = "Moment-5"
Text14.Text = "Moment-6"
Text15.Text = "Moment-7"
Text16.Text = "Moment-8"
Text17.Text = "Moment-9"
Text18.Text = "Moment-10"
Text19.Text = "Moment-11"

End Sub


Private Sub EnableTxtOnetoEight()

Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text8.Enabled = True

End Sub

Private Sub DisableTxtOnetoEight()

Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
Text7.Enabled = False
Text8.Enabled = False

End Sub
Private Sub DefaultLocLabel63to73()
Label63.Height = 500
Label63.Left = 2200
Label63.Top = 1920
Label63.Width = 900

Label64.Height = 500
Label64.Left = 4300
Label64.Top = 1920
Label64.Width = 860

Label65.Height = 500
Label65.Left = 6200
Label65.Top = 1920
Label65.Width = 860

Label66.Height = 500
Label66.Left = 8100
Label66.Top = 1920
Label66.Width = 860

Label67.Height = 500
Label67.Left = 10000
Label67.Top = 1920
Label67.Width = 860

Label68.Height = 500
Label68.Left = 2200
Label68.Top = 3480
Label68.Width = 860

Label69.Height = 500
Label69.Left = 4300
Label69.Top = 3480
Label69.Width = 860

Label70.Height = 500
Label70.Left = 6200
Label70.Top = 3480
Label70.Width = 860

Label71.Height = 500
Label71.Left = 2200
Label71.Top = 5040
Label71.Width = 860

Label72.Height = 500
Label72.Left = 4300
Label72.Top = 5040
Label72.Width = 860

Label73.Height = 500
Label73.Left = 10000
Label73.Top = 5040
Label73.Width = 860

End Sub

Private Sub BoldLabel63to73()
Label63.FontBold = True
Label64.FontBold = True
Label65.FontBold = True
Label66.FontBold = True
Label67.FontBold = True
Label68.FontBold = True
Label69.FontBold = True
Label70.FontBold = True
Label71.FontBold = True
Label72.FontBold = True
Label73.FontBold = True

End Sub


Private Sub UnBoldLabel63to73()
Label63.FontBold = False
Label64.FontBold = False
Label65.FontBold = False
Label66.FontBold = False
Label67.FontBold = False
Label68.FontBold = False
Label69.FontBold = False
Label70.FontBold = False
Label71.FontBold = False
Label72.FontBold = False
Label73.FontBold = False

End Sub

Private Sub HighlightPanel()
If Left(Label87.Caption, 8) = "Panel_1 " Then
        Shape1.BorderWidth = 2
          Else
            Shape1.BorderWidth = 1
              End If

 If Left(Label87.Caption, 7) = "Panel_2" Then
        Shape2.BorderWidth = 2
          Else
            Shape2.BorderWidth = 1
              End If
              
If Left(Label87.Caption, 7) = "Panel_3" Then
        Shape7.BorderWidth = 2
          Else
            Shape7.BorderWidth = 1
              End If
              
If Left(Label87.Caption, 7) = "Panel_4" Then
        Shape9.BorderWidth = 2
          Else
            Shape9.BorderWidth = 1
              End If
              
If Left(Label87.Caption, 7) = "Panel_5" Then
        Shape40.BorderWidth = 2
          Else
            Shape40.BorderWidth = 1
              End If

If Left(Label87.Caption, 7) = "Panel_6" Then
        Shape3.BorderWidth = 2
          Else
            Shape3.BorderWidth = 1
              End If


If Left(Label87.Caption, 7) = "Panel_7" Then
        Shape4.BorderWidth = 2
          Else
            Shape4.BorderWidth = 1
              End If


If Left(Label87.Caption, 7) = "Panel_8" Then
        Shape8.BorderWidth = 2
          Else
            Shape8.BorderWidth = 1
              End If

If Left(Label87.Caption, 7) = "Panel_9" Then
        Shape5.BorderWidth = 2
          Else
            Shape5.BorderWidth = 1
              End If


If Left(Label87.Caption, 8) = "Panel_10" Then
        Shape6.BorderWidth = 2
          Else
            Shape6.BorderWidth = 1
              End If


If Left(Label87.Caption, 8) = "Panel_11" Then
        Shape12.BorderWidth = 2
          Else
            Shape12.BorderWidth = 1
              End If

If Left(Label87.Caption, 6) = "Minimun" Then
        Shape1.BorderWidth = 1
        Shape2.BorderWidth = 1
        Shape7.BorderWidth = 1
        Shape9.BorderWidth = 1
        Shape3.BorderWidth = 1
        Shape4.BorderWidth = 1
        Shape8.BorderWidth = 1
        Shape5.BorderWidth = 1
        Shape6.BorderWidth = 1
        Shape12.BorderWidth = 1
              End If


End Sub


























