VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00C0C000&
   Caption         =   "Reka & Lukis Tiang (versi:02/2011)"
   ClientHeight    =   7725
   ClientLeft      =   255
   ClientTop       =   735
   ClientWidth     =   10890
   FillColor       =   &H0000C000&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Lucida Handwriting"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000000FF&
   Icon            =   "STiang.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "STiang.frx":030A
   ScaleHeight     =   7198.931
   ScaleMode       =   0  'User
   ScaleWidth      =   10890
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000C000&
      Cancel          =   -1  'True
      Caption         =   "Kriteria Rekabentuk dsb"
      Height          =   225
      Left            =   4560
      MaskColor       =   &H0000C000&
      TabIndex        =   137
      Top             =   360
      UseMaskColor    =   -1  'True
      Width           =   4935
   End
   Begin VB.OptionButton Option9 
      BackColor       =   &H00C0C000&
      Caption         =   "Option9"
      Enabled         =   0   'False
      Height          =   225
      Left            =   10440
      TabIndex        =   127
      Top             =   1200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.OptionButton Option8 
      BackColor       =   &H00C0C000&
      Caption         =   "Option8"
      Enabled         =   0   'False
      Height          =   225
      Left            =   10440
      TabIndex        =   126
      Top             =   1560
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.OptionButton Option7 
      BackColor       =   &H00C0C000&
      Caption         =   "Option7"
      Enabled         =   0   'False
      Height          =   225
      Left            =   10440
      TabIndex        =   125
      Top             =   1920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.OptionButton Option6 
      BackColor       =   &H00C0C000&
      Caption         =   "Option6"
      Enabled         =   0   'False
      Height          =   225
      Left            =   10440
      TabIndex        =   124
      Top             =   2280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C000&
      Caption         =   "Interaction Diagram"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   5160
      MaskColor       =   &H00C0C000&
      TabIndex        =   114
      Top             =   2618
      Width           =   2300
   End
   Begin VB.ListBox List1 
      BackColor       =   &H80000007&
      ForeColor       =   &H80000005&
      Height          =   735
      ItemData        =   "STiang.frx":0614
      Left            =   1080
      List            =   "STiang.frx":061B
      TabIndex        =   120
      Top             =   1560
      Width           =   8355
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   495
      Left            =   0
      Picture         =   "STiang.frx":0625
      ScaleHeight     =   8.202
      ScaleMode       =   6  'Millimeter
      ScaleWidth      =   9.79
      TabIndex        =   119
      Top             =   2880
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.TextBox Text62 
      BackColor       =   &H00C0C000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   330
      Left            =   3840
      TabIndex        =   115
      Tag             =   " "
      Text            =   "UNBRACED"
      Top             =   3600
      Width           =   700
   End
   Begin VB.TextBox Text61 
      BackColor       =   &H00C0C000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   315
      Left            =   3840
      TabIndex        =   112
      Tag             =   " "
      Text            =   "UNBRACED"
      Top             =   3960
      Width           =   700
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C000&
      Caption         =   "Lukis"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   7900
      MaskColor       =   &H00C0C000&
      TabIndex        =   116
      Top             =   2618
      Width           =   1400
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C000&
      Caption         =   "ok"
      Enabled         =   0   'False
      Height          =   644
      Left            =   9600
      TabIndex        =   111
      Top             =   600
      Width           =   350
   End
   Begin VB.TextBox Text60 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0;(0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      Height          =   330
      Left            =   8640
      MaxLength       =   4
      TabIndex        =   107
      Text            =   "101"
      Top             =   840
      Width           =   780
   End
   Begin VB.OptionButton Option5 
      BackColor       =   &H00C0C000&
      Caption         =   "Option5"
      Enabled         =   0   'False
      Height          =   225
      Left            =   4480
      TabIndex        =   106
      Top             =   2640
      Width           =   220
   End
   Begin VB.OptionButton Option4 
      BackColor       =   &H00C0C000&
      Caption         =   "Option4"
      Enabled         =   0   'False
      Height          =   225
      Left            =   3480
      TabIndex        =   105
      Top             =   2640
      Width           =   220
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00C0C000&
      Caption         =   "Option3"
      Enabled         =   0   'False
      Height          =   225
      Left            =   2480
      TabIndex        =   104
      Top             =   2640
      Width           =   220
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00C0C000&
      Caption         =   "Option2"
      Enabled         =   0   'False
      Height          =   225
      Left            =   1480
      TabIndex        =   103
      Top             =   2640
      Width           =   220
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0C000&
      Caption         =   "Option1"
      Enabled         =   0   'False
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   480
      MaskColor       =   &H000000FF&
      TabIndex        =   102
      Top             =   2640
      Width           =   220
   End
   Begin VB.TextBox Text59 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0;(0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      Height          =   330
      Left            =   7800
      MaxLength       =   4
      TabIndex        =   92
      Text            =   "200"
      Top             =   840
      Width           =   780
   End
   Begin VB.TextBox Text58 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0;(0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      Height          =   330
      Left            =   6960
      MaxLength       =   3
      TabIndex        =   91
      Text            =   "10"
      Top             =   840
      Width           =   780
   End
   Begin VB.TextBox Text57 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00;(0.00)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      Height          =   330
      Left            =   6120
      MaxLength       =   4
      TabIndex        =   90
      Text            =   "2.5"
      Top             =   840
      Width           =   780
   End
   Begin VB.TextBox Text56 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00000;(0.00000)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      Height          =   330
      Left            =   5280
      MaxLength       =   7
      TabIndex        =   89
      Text            =   "0.0003"
      Top             =   840
      Width           =   780
   End
   Begin VB.TextBox Text55 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0;(0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      Height          =   330
      Left            =   4440
      MaxLength       =   4
      TabIndex        =   88
      Text            =   "250"
      Top             =   840
      Width           =   780
   End
   Begin VB.TextBox Text54 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0;(0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      Height          =   330
      Left            =   3600
      MaxLength       =   4
      TabIndex        =   87
      Text            =   "460"
      Top             =   840
      Width           =   780
   End
   Begin VB.TextBox Text53 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0;(0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      Height          =   330
      Left            =   2760
      MaxLength       =   3
      TabIndex        =   86
      Text            =   "30"
      Top             =   840
      Width           =   765
   End
   Begin VB.TextBox Text52 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      Height          =   330
      Left            =   1920
      TabIndex        =   85
      Text            =   "0"
      Top             =   840
      Width           =   780
   End
   Begin VB.TextBox Text51 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      Height          =   330
      Left            =   1080
      TabIndex        =   84
      Text            =   "0"
      Top             =   840
      Width           =   780
   End
   Begin VB.TextBox Text50 
      BackColor       =   &H00C0C000&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0;(0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      ForeColor       =   &H80000004&
      Height          =   375
      Left            =   7800
      MaxLength       =   3
      TabIndex        =   81
      Text            =   "40"
      Top             =   7080
      Width           =   500
   End
   Begin VB.TextBox Text49 
      BackColor       =   &H00C0C000&
      Enabled         =   0   'False
      ForeColor       =   &H80000004&
      Height          =   375
      Left            =   8400
      TabIndex        =   80
      Text            =   "A/1"
      Top             =   7080
      Width           =   1665
   End
   Begin VB.TextBox Text48 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0;(0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   376
      Left            =   8880
      MaxLength       =   3
      TabIndex        =   64
      Text            =   "20"
      Top             =   6120
      Width           =   500
   End
   Begin VB.TextBox Text47 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0;(0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   376
      Left            =   8160
      MaxLength       =   4
      TabIndex        =   63
      Text            =   "3"
      Top             =   6120
      Width           =   500
   End
   Begin VB.TextBox Text46 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0;(0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   376
      Left            =   8880
      MaxLength       =   3
      TabIndex        =   62
      Text            =   "20"
      Top             =   3840
      Width           =   500
   End
   Begin VB.TextBox Text45 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0;(0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   376
      Left            =   8160
      MaxLength       =   3
      TabIndex        =   61
      Text            =   "3"
      Top             =   3840
      Width           =   500
   End
   Begin VB.TextBox Text44 
      BackColor       =   &H00C0C000&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0;(0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   376
      Left            =   9720
      MaxLength       =   3
      TabIndex        =   60
      Text            =   "20"
      Top             =   5280
      Width           =   500
   End
   Begin VB.TextBox Text43 
      BackColor       =   &H00C0C000&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0;(0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   376
      Left            =   9720
      MaxLength       =   3
      TabIndex        =   59
      Text            =   "1"
      Top             =   4680
      Width           =   500
   End
   Begin VB.TextBox Text42 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C000&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0;(0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   375
      Left            =   7320
      MaxLength       =   3
      TabIndex        =   58
      Text            =   "20"
      Top             =   5280
      Width           =   500
   End
   Begin VB.TextBox Text41 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C000&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0;(0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   376
      Left            =   7320
      MaxLength       =   3
      TabIndex        =   57
      Text            =   "1"
      Top             =   4680
      Width           =   500
   End
   Begin VB.TextBox Text40 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000014&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0;(0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   451
      Left            =   3840
      MaxLength       =   6
      TabIndex        =   56
      Text            =   "1000"
      Top             =   5280
      Width           =   720
   End
   Begin VB.TextBox Text39 
      BackColor       =   &H00C0C000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   330
      Left            =   3840
      ScrollBars      =   2  'Vertical
      TabIndex        =   44
      Text            =   "RESIST"
      Top             =   7080
      Width           =   700
   End
   Begin VB.TextBox Text38 
      BackColor       =   &H8000000A&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0;(0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      Height          =   330
      Left            =   3000
      MaxLength       =   6
      TabIndex        =   43
      Text            =   "1500"
      Top             =   7200
      Width           =   800
   End
   Begin VB.TextBox Text37 
      BackColor       =   &H8000000A&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0;(0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      Height          =   330
      Left            =   2160
      MaxLength       =   5
      TabIndex        =   42
      Text            =   "400"
      Top             =   7200
      Width           =   800
   End
   Begin VB.TextBox Text36 
      BackColor       =   &H8000000A&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0;(0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      Height          =   330
      Left            =   1320
      MaxLength       =   5
      TabIndex        =   41
      Text            =   "350"
      Top             =   7200
      Width           =   800
   End
   Begin VB.TextBox Text35 
      BackColor       =   &H8000000B&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0;(0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      Height          =   330
      Left            =   3000
      MaxLength       =   6
      TabIndex        =   40
      Text            =   "3000"
      Top             =   6840
      Width           =   800
   End
   Begin VB.TextBox Text34 
      BackColor       =   &H8000000B&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0;(0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      Height          =   330
      Left            =   2160
      MaxLength       =   5
      TabIndex        =   39
      Text            =   "600"
      Top             =   6840
      Width           =   800
   End
   Begin VB.TextBox Text33 
      BackColor       =   &H8000000B&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0;(0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      Height          =   330
      Left            =   1320
      MaxLength       =   5
      TabIndex        =   38
      Text            =   "200"
      Top             =   6840
      Width           =   800
   End
   Begin VB.TextBox Text32 
      BackColor       =   &H8000000B&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0;(0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      Height          =   330
      Left            =   3000
      MaxLength       =   6
      TabIndex        =   37
      Text            =   "5000"
      Top             =   6480
      Width           =   800
   End
   Begin VB.TextBox Text31 
      BackColor       =   &H8000000B&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0;(0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      Height          =   330
      Left            =   2160
      MaxLength       =   5
      TabIndex        =   36
      Text            =   "600"
      Top             =   6480
      Width           =   800
   End
   Begin VB.TextBox Text30 
      BackColor       =   &H8000000B&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0;(0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      Height          =   330
      Left            =   1320
      MaxLength       =   5
      TabIndex        =   35
      Text            =   "200"
      Top             =   6480
      Width           =   800
   End
   Begin VB.TextBox Text29 
      BackColor       =   &H8000000A&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0;(0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      Height          =   330
      Left            =   3000
      MaxLength       =   6
      TabIndex        =   34
      Text            =   "6000"
      Top             =   6120
      Width           =   800
   End
   Begin VB.TextBox Text28 
      BackColor       =   &H8000000A&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0;(0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      Height          =   330
      Left            =   2160
      MaxLength       =   5
      TabIndex        =   33
      Text            =   "450"
      Top             =   6120
      Width           =   800
   End
   Begin VB.TextBox Text27 
      BackColor       =   &H8000000A&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0;(0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      Height          =   330
      Left            =   1320
      MaxLength       =   5
      TabIndex        =   32
      Text            =   "200"
      Top             =   6120
      Width           =   800
   End
   Begin VB.TextBox Text26 
      BackColor       =   &H8000000A&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0;(0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      Height          =   330
      Left            =   3000
      MaxLength       =   6
      TabIndex        =   31
      Text            =   "4000"
      Top             =   5760
      Width           =   800
   End
   Begin VB.TextBox Text25 
      BackColor       =   &H8000000A&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0;(0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      Height          =   330
      Left            =   2160
      MaxLength       =   5
      TabIndex        =   30
      Text            =   "500"
      Top             =   5760
      Width           =   800
   End
   Begin VB.TextBox Text24 
      BackColor       =   &H8000000A&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0;(0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      Height          =   330
      Left            =   1320
      MaxLength       =   5
      TabIndex        =   29
      Text            =   "200"
      Top             =   5760
      Width           =   800
   End
   Begin VB.TextBox Text23 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000014&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   456
      Left            =   4560
      MaxLength       =   5
      TabIndex        =   28
      Text            =   "10"
      Top             =   5760
      Width           =   600
   End
   Begin VB.TextBox Text22 
      BackColor       =   &H80000014&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   456
      Left            =   3840
      MaxLength       =   5
      TabIndex        =   27
      Text            =   "55"
      Top             =   5760
      Width           =   600
   End
   Begin VB.TextBox Text21 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000014&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   456
      Left            =   4560
      MaxLength       =   5
      TabIndex        =   26
      Text            =   "15"
      Top             =   4800
      Width           =   600
   End
   Begin VB.TextBox Text20 
      BackColor       =   &H80000014&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   456
      Left            =   3840
      MaxLength       =   5
      TabIndex        =   25
      Text            =   "75"
      Top             =   4800
      Width           =   600
   End
   Begin VB.TextBox Text19 
      Alignment       =   2  'Center
      BackColor       =   &H80000014&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0;(0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   456
      Left            =   3000
      MaxLength       =   6
      TabIndex        =   24
      Text            =   "3600"
      Top             =   5280
      Width           =   800
   End
   Begin VB.TextBox Text18 
      Alignment       =   2  'Center
      BackColor       =   &H80000014&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0;(0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   456
      Left            =   2160
      MaxLength       =   5
      TabIndex        =   23
      Text            =   "350"
      Top             =   5280
      Width           =   800
   End
   Begin VB.TextBox Text17 
      Alignment       =   2  'Center
      BackColor       =   &H80000014&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0;(0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   456
      Left            =   1320
      MaxLength       =   5
      TabIndex        =   22
      Text            =   "300"
      Top             =   5280
      Width           =   800
   End
   Begin VB.TextBox Text16 
      BackColor       =   &H8000000A&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0;(0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      Height          =   330
      Left            =   3000
      MaxLength       =   6
      TabIndex        =   21
      Text            =   "4000"
      Top             =   4920
      Width           =   800
   End
   Begin VB.TextBox Text15 
      BackColor       =   &H8000000A&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0;(0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      Height          =   330
      Left            =   2160
      MaxLength       =   5
      TabIndex        =   20
      Text            =   "500"
      Top             =   4920
      Width           =   800
   End
   Begin VB.TextBox Text14 
      BackColor       =   &H8000000A&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0;(0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      Height          =   330
      Left            =   1320
      MaxLength       =   5
      TabIndex        =   19
      Text            =   "200"
      Top             =   4920
      Width           =   800
   End
   Begin VB.TextBox Text13 
      BackColor       =   &H8000000A&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0;(0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      Height          =   330
      Left            =   3000
      MaxLength       =   6
      TabIndex        =   18
      Text            =   "6000"
      Top             =   4560
      Width           =   800
   End
   Begin VB.TextBox Text12 
      BackColor       =   &H8000000A&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0;(0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      Height          =   330
      Left            =   2160
      MaxLength       =   5
      TabIndex        =   17
      Text            =   "450"
      Top             =   4560
      Width           =   800
   End
   Begin VB.TextBox Text11 
      BackColor       =   &H8000000A&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0;(0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      Height          =   330
      Left            =   1320
      MaxLength       =   5
      TabIndex        =   16
      Text            =   "200"
      Top             =   4560
      Width           =   800
   End
   Begin VB.TextBox Text10 
      BackColor       =   &H8000000B&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0;(0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      Height          =   330
      Left            =   3000
      MaxLength       =   6
      TabIndex        =   15
      Text            =   "3000"
      Top             =   4200
      Width           =   800
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H8000000B&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0;(0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      Height          =   330
      Left            =   2160
      MaxLength       =   5
      TabIndex        =   14
      Text            =   "600"
      Top             =   4200
      Width           =   800
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H8000000B&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0;(0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      Height          =   330
      Left            =   1320
      MaxLength       =   5
      TabIndex        =   13
      Text            =   "200"
      Top             =   4200
      Width           =   800
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H8000000B&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0;(0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      Height          =   330
      Left            =   3000
      MaxLength       =   6
      TabIndex        =   12
      Text            =   "3000"
      Top             =   3840
      Width           =   800
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H8000000B&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0;(0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      Height          =   330
      Left            =   2160
      MaxLength       =   5
      TabIndex        =   11
      Text            =   "600"
      Top             =   3840
      Width           =   800
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H8000000B&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0;(0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      Height          =   330
      Left            =   1320
      MaxLength       =   5
      TabIndex        =   10
      Text            =   "200"
      Top             =   3840
      Width           =   800
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0C000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   330
      Left            =   3840
      TabIndex        =   9
      Text            =   "RESIST"
      Top             =   6720
      Width           =   700
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H8000000A&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0;(0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      Height          =   330
      Left            =   3000
      MaxLength       =   6
      TabIndex        =   8
      Text            =   "3900"
      Top             =   3480
      Width           =   800
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H8000000A&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0;(0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      Height          =   330
      Left            =   2160
      MaxLength       =   5
      TabIndex        =   7
      Text            =   "350"
      Top             =   3480
      Width           =   800
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000A&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0;(0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      Height          =   330
      Left            =   1320
      MaxLength       =   5
      MousePointer    =   1  'Arrow
      TabIndex        =   6
      Text            =   "300"
      Top             =   3480
      Width           =   800
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10200
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "TIANG - 1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   268
      Left            =   110
      MaskColor       =   &H00C0C0FF&
      TabIndex        =   113
      Top             =   2280
      UseMaskColor    =   -1  'True
      Width           =   1000
   End
   Begin VB.Image Image1 
      Height          =   675
      Left            =   5640
      Picture         =   "STiang.frx":092F
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   945
   End
   Begin VB.Label Label61 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C000&
      Caption         =   "me@wanluqman.com"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   210
      Left            =   9360
      TabIndex        =   136
      Top             =   2610
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label60 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(My 1,2)"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8280
      TabIndex        =   135
      Top             =   6760
      Width           =   735
   End
   Begin VB.Label Label59 
      BackStyle       =   0  'Transparent
      Caption         =   "(Mx 1,2)"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6660
      TabIndex        =   134
      Top             =   5312
      Width           =   735
   End
   Begin VB.Label Label58 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C000&
      Caption         =   "My2"
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   4560
      TabIndex        =   133
      Top             =   4560
      Width           =   615
   End
   Begin VB.Label Label57 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C000&
      Caption         =   "My1"
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   4560
      TabIndex        =   132
      Top             =   6240
      Width           =   615
   End
   Begin VB.Label Label56 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      Caption         =   "axis"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   8340
      TabIndex        =   131
      Top             =   6480
      Width           =   855
   End
   Begin VB.Label Label55 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      Caption         =   "minor"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   8340
      TabIndex        =   130
      Top             =   3540
      Width           =   855
   End
   Begin VB.Line Line26 
      Visible         =   0   'False
      X1              =   6120
      X2              =   5640
      Y1              =   3578.498
      Y2              =   3466.67
   End
   Begin VB.Line Line25 
      Visible         =   0   'False
      X1              =   6120
      X2              =   6600
      Y1              =   3578.498
      Y2              =   3466.67
   End
   Begin VB.Line Line24 
      Visible         =   0   'False
      X1              =   5640
      X2              =   5640
      Y1              =   3354.842
      Y2              =   3466.67
   End
   Begin VB.Line Line23 
      Visible         =   0   'False
      X1              =   6600
      X2              =   6600
      Y1              =   3354.842
      Y2              =   3466.67
   End
   Begin VB.Line Line22 
      Visible         =   0   'False
      X1              =   5640
      X2              =   6120
      Y1              =   3354.842
      Y2              =   3578.498
   End
   Begin VB.Line Line21 
      Visible         =   0   'False
      X1              =   6600
      X2              =   6120
      Y1              =   3354.842
      Y2              =   3578.498
   End
   Begin VB.Shape Shape26 
      Height          =   495
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   4507
      Width           =   495
   End
   Begin VB.Shape Shape25 
      Height          =   495
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5956
      Width           =   495
   End
   Begin VB.Shape Shape24 
      Height          =   375
      Left            =   5880
      Top             =   6960
      Width           =   495
   End
   Begin VB.Label Label47 
      BackColor       =   &H00C0C000&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   6840
      TabIndex        =   129
      Top             =   6480
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C000&
      Caption         =   "Y"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   6720
      TabIndex        =   128
      Top             =   5760
      Width           =   255
   End
   Begin VB.Line Line19 
      X1              =   5640
      X2              =   6120
      Y1              =   4249.466
      Y2              =   4473.122
   End
   Begin VB.Line Line18 
      X1              =   6600
      X2              =   6120
      Y1              =   4584.95
      Y2              =   4361.294
   End
   Begin VB.Line Line17 
      X1              =   6600
      X2              =   6120
      Y1              =   4249.466
      Y2              =   4473.122
   End
   Begin VB.Line Line16 
      X1              =   5640
      X2              =   6120
      Y1              =   4584.95
      Y2              =   4361.294
   End
   Begin VB.Line Line15 
      X1              =   5640
      X2              =   6120
      Y1              =   5591.403
      Y2              =   5815.059
   End
   Begin VB.Line Line14 
      X1              =   6120
      X2              =   6600
      Y1              =   5703.23
      Y2              =   5926.887
   End
   Begin VB.Line Line13 
      X1              =   6600
      X2              =   6120
      Y1              =   5591.403
      Y2              =   5815.059
   End
   Begin VB.Line Line12 
      X1              =   5640
      X2              =   6120
      Y1              =   5926.887
      Y2              =   5703.23
   End
   Begin VB.Shape Shape23 
      Height          =   855
      Left            =   6000
      Top             =   3840
      Width           =   255
   End
   Begin VB.Shape Shape22 
      Height          =   1335
      Left            =   6000
      Top             =   4800
      Width           =   255
   End
   Begin VB.Shape Shape21 
      Height          =   735
      Left            =   6000
      Top             =   6240
      Width           =   255
   End
   Begin VB.Shape Shape20 
      FillColor       =   &H00FFFF00&
      Height          =   255
      Left            =   5640
      Top             =   4320
      Width           =   135
   End
   Begin VB.Shape Shape19 
      FillColor       =   &H00FFFF00&
      Height          =   255
      Left            =   6480
      Top             =   4920
      Width           =   135
   End
   Begin VB.Shape Shape18 
      FillColor       =   &H0000FFFF&
      Height          =   255
      Left            =   6480
      Top             =   4320
      Width           =   135
   End
   Begin VB.Shape Shape17 
      FillColor       =   &H0000FFFF&
      Height          =   255
      Left            =   5640
      Top             =   4920
      Width           =   135
   End
   Begin VB.Shape Shape16 
      FillColor       =   &H00FFFF00&
      Height          =   255
      Left            =   5640
      Top             =   5760
      Width           =   135
   End
   Begin VB.Shape Shape15 
      FillColor       =   &H00FFFF00&
      Height          =   255
      Left            =   6480
      Top             =   6360
      Width           =   135
   End
   Begin VB.Shape Shape14 
      FillColor       =   &H0000FFFF&
      Height          =   255
      Left            =   6480
      Top             =   5760
      Width           =   135
   End
   Begin VB.Shape Shape13 
      FillColor       =   &H0000FFFF&
      Height          =   255
      Left            =   5640
      Top             =   6360
      Width           =   135
   End
   Begin VB.Line Line11 
      BorderColor     =   &H80000005&
      BorderStyle     =   3  'Dot
      X1              =   6720
      X2              =   5520
      Y1              =   5479.575
      Y2              =   6038.715
   End
   Begin VB.Line Line10 
      BorderColor     =   &H80000005&
      BorderStyle     =   3  'Dot
      X1              =   5520
      X2              =   6960
      Y1              =   5479.575
      Y2              =   6150.543
   End
   Begin VB.Line Line9 
      BorderColor     =   &H80000005&
      BorderStyle     =   3  'Dot
      X1              =   6720
      X2              =   5520
      Y1              =   4137.638
      Y2              =   4696.778
   End
   Begin VB.Line Line7 
      BorderColor     =   &H80000005&
      BorderStyle     =   3  'Dot
      X1              =   5520
      X2              =   6720
      Y1              =   4137.638
      Y2              =   4696.778
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      BorderStyle     =   3  'Dot
      X1              =   6120
      X2              =   6120
      Y1              =   3354.842
      Y2              =   6709.683
   End
   Begin VB.Shape Shape12 
      BorderColor     =   &H80000014&
      Height          =   751
      Left            =   960
      Top             =   483
      Width           =   8655
   End
   Begin VB.Line Line5 
      BorderColor     =   &H8000000E&
      X1              =   100
      X2              =   9500
      Y1              =   2399.644
      Y2              =   2399.644
   End
   Begin VB.Line Line4 
      BorderColor     =   &H8000000E&
      X1              =   100
      X2              =   9500
      Y1              =   2683.873
      Y2              =   2682.941
   End
   Begin VB.Line Line3 
      X1              =   1320
      X2              =   1320
      Y1              =   4473.122
      Y2              =   6709.683
   End
   Begin VB.Shape Shape11 
      BorderColor     =   &H80000014&
      BorderWidth     =   3
      FillColor       =   &H00008000&
      Height          =   375
      Left            =   960
      Top             =   45
      Width           =   3480
   End
   Begin VB.Label Label54 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      Caption         =   " ........"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   0
      TabIndex        =   123
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label53 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      Caption         =   "Desg Axial Load"
      ForeColor       =   &H80000014&
      Height          =   495
      Left            =   0
      TabIndex        =   122
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      BorderStyle     =   4  'Dash-Dot
      X1              =   8760
      X2              =   8760
      Y1              =   3578.498
      Y2              =   6038.715
   End
   Begin VB.Shape Shape10 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   2
      FillColor       =   &H80000005&
      Height          =   1502
      Left            =   8040
      Shape           =   4  'Rounded Rectangle
      Top             =   4440
      Width           =   1450
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H80000001&
      BorderColor     =   &H00008000&
      FillColor       =   &H8000000E&
      Height          =   1717
      Left            =   7920
      Top             =   4320
      Width           =   1700
   End
   Begin VB.Label Label52 
      BackColor       =   &H00C0C000&
      Caption         =   "      Axes:        Ncap.           Mcap.            M/bh2.          N/bh.          K             Mdsgn"
      ForeColor       =   &H80000014&
      Height          =   240
      Left            =   1080
      TabIndex        =   121
      Top             =   1320
      Width           =   8355
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000C0&
      FillColor       =   &H000000C0&
      FillStyle       =   0  'Solid
      Height          =   268
      Left            =   9260
      Shape           =   3  'Circle
      Top             =   5040
      Width           =   200
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000C0&
      FillColor       =   &H000000C0&
      FillStyle       =   0  'Solid
      Height          =   270
      Left            =   8060
      Shape           =   3  'Circle
      Top             =   5040
      Width           =   195
   End
   Begin VB.Shape Shape6 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   210
      Left            =   9230
      Shape           =   3  'Circle
      Top             =   5677
      Width           =   200
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   210
      Left            =   8640
      Shape           =   3  'Circle
      Top             =   5698
      Width           =   250
   End
   Begin VB.Shape Shape4 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   210
      Left            =   8100
      Shape           =   3  'Circle
      Top             =   5677
      Width           =   200
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   210
      Left            =   9230
      Shape           =   3  'Circle
      Top             =   4485
      Width           =   200
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   210
      Left            =   8640
      Shape           =   3  'Circle
      Top             =   4464
      Width           =   250
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   210
      Left            =   8100
      Shape           =   3  'Circle
      Top             =   4485
      Width           =   200
   End
   Begin VB.Label Label51 
      BackColor       =   &H00C0C000&
      Caption         =   "Y braced  FRAME ^"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   435
      Left            =   4560
      TabIndex        =   118
      Top             =   4080
      Width           =   855
   End
   Begin VB.Label Label50 
      BackColor       =   &H00C0C000&
      Caption         =   "X braced"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   255
      Left            =   4560
      TabIndex        =   117
      Top             =   3720
      Width           =   855
   End
   Begin VB.Label Label49 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      Caption         =   "axis"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   9960
      TabIndex        =   110
      Top             =   5040
      Width           =   615
   End
   Begin VB.Label Label48 
      BackColor       =   &H00C0C000&
      Caption         =   "major"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   6840
      TabIndex        =   109
      Top             =   5040
      Width           =   735
   End
   Begin VB.Line Line8 
      BorderColor     =   &H80000005&
      BorderStyle     =   4  'Dash-Dot
      X1              =   7320
      X2              =   10200
      Y1              =   4808.606
      Y2              =   4808.606
   End
   Begin VB.Label Label46 
      Alignment       =   2  'Center
      Caption         =   "Bmark."
      Height          =   255
      Left            =   8640
      TabIndex        =   108
      Top             =   600
      Width           =   900
   End
   Begin VB.Label Label45 
      Caption         =   "-spacing"
      Height          =   255
      Left            =   7800
      TabIndex        =   101
      Top             =   600
      Width           =   900
   End
   Begin VB.Label Label44 
      Alignment       =   2  'Center
      Caption         =   "Linkdia."
      Height          =   255
      Left            =   6960
      TabIndex        =   100
      Top             =   600
      Width           =   900
   End
   Begin VB.Label Label43 
      Alignment       =   2  'Center
      Caption         =   "creep."
      Height          =   255
      Left            =   6120
      TabIndex        =   99
      Top             =   600
      Width           =   900
   End
   Begin VB.Label Label42 
      Alignment       =   2  'Center
      Caption         =   "shrink."
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   98
      Top             =   600
      Width           =   900
   End
   Begin VB.Label Label41 
      Alignment       =   2  'Center
      Caption         =   "fyv"
      Height          =   255
      Left            =   4440
      TabIndex        =   97
      Top             =   600
      Width           =   900
   End
   Begin VB.Label Label40 
      Alignment       =   2  'Center
      Caption         =   "fy"
      Height          =   255
      Left            =   3600
      TabIndex        =   96
      Top             =   600
      Width           =   900
   End
   Begin VB.Label Label39 
      Alignment       =   2  'Center
      Caption         =   "fcu"
      Height          =   255
      Left            =   2760
      TabIndex        =   95
      Top             =   600
      Width           =   885
   End
   Begin VB.Label Label38 
      Alignment       =   2  'Center
      Caption         =   "Y  "
      Height          =   255
      Left            =   1920
      TabIndex        =   94
      Top             =   600
      Width           =   900
   End
   Begin VB.Label Label37 
      Alignment       =   2  'Center
      Caption         =   "X  "
      Height          =   255
      Left            =   1080
      TabIndex        =   93
      Top             =   600
      Width           =   900
   End
   Begin VB.Label Label36 
      BackColor       =   &H00C0C000&
      Caption         =   "Concrete cover."
      Height          =   495
      Left            =   6840
      TabIndex        =   83
      Top             =   7080
      Width           =   855
   End
   Begin VB.Label Label35 
      BackColor       =   &H00C0C000&
      Caption         =   "Label tiang."
      Height          =   495
      Left            =   10080
      TabIndex        =   82
      Top             =   7080
      Width           =   735
   End
   Begin VB.Label Label34 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      Caption         =   "   Reinforcements  layout  relative to axes.      ( Design plain :-  global  X- axis )"
      Height          =   465
      Left            =   6600
      TabIndex        =   79
      Top             =   3000
      Width           =   4095
   End
   Begin VB.Label Label33 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      Caption         =   "Reactions  and Ends  conditions."
      Height          =   495
      Left            =   3840
      TabIndex        =   78
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label Label32 
      BackColor       =   &H00C0C000&
      Caption         =   "Bar dia."
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   6960
      TabIndex        =   77
      Top             =   5640
      Width           =   855
   End
   Begin VB.Label Label31 
      BackColor       =   &H00C0C000&
      Caption         =   "No of bars @ yy"
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   6840
      TabIndex        =   76
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label Label30 
      BackColor       =   &H00C0C000&
      Caption         =   "Bar dia."
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   9360
      TabIndex        =   75
      Top             =   6120
      Width           =   855
   End
   Begin VB.Label Label29 
      BackColor       =   &H00C0C000&
      Caption         =   "No of bars           @ xx"
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   7200
      TabIndex        =   74
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Label Label28 
      BackColor       =   &H00C0C000&
      Caption         =   "Bar dia."
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   9360
      TabIndex        =   73
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label27 
      BackColor       =   &H00C0C000&
      Caption         =   "No of bars "
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   7200
      TabIndex        =   72
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0C000&
      Caption         =   "Bar dia."
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   9720
      TabIndex        =   71
      Top             =   5640
      Width           =   855
   End
   Begin VB.Label Label25 
      BackColor       =   &H00C0C000&
      Caption         =   "No of bars"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   9720
      TabIndex        =   70
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label Label24 
      BackColor       =   &H00C0C000&
      Caption         =   "Y res-mnt"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   375
      Left            =   4560
      TabIndex        =   69
      Top             =   7080
      Width           =   915
   End
   Begin VB.Label Label23 
      BackColor       =   &H00C0C000&
      Caption         =   "Mx1           My1"
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   3840
      TabIndex        =   68
      Top             =   6240
      Width           =   585
   End
   Begin VB.Label Label22 
      BackColor       =   &H00C0C000&
      Caption         =   "Axial load"
      ForeColor       =   &H0000FFFF&
      Height          =   420
      Left            =   4680
      TabIndex        =   67
      Top             =   5280
      Width           =   675
   End
   Begin VB.Label Label21 
      BackColor       =   &H00C0C000&
      Caption         =   "Mx2  "
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   3840
      TabIndex        =   66
      Top             =   4560
      Width           =   555
   End
   Begin VB.Label Label20 
      BackColor       =   &H00C0C000&
      Caption         =   "BASE v        X res-mnt"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   435
      Left            =   4560
      TabIndex        =   65
      Top             =   6480
      Width           =   885
   End
   Begin VB.Label Label19 
      BackColor       =   &H00C0C000&
      Caption         =   "low col"
      Height          =   311
      Left            =   0
      TabIndex        =   55
      Top             =   7200
      Width           =   1305
   End
   Begin VB.Label Label18 
      BackColor       =   &H00C0C000&
      Caption         =   "west beam x"
      Height          =   311
      Left            =   0
      TabIndex        =   54
      Top             =   6840
      Width           =   1305
   End
   Begin VB.Label Label17 
      BackColor       =   &H00C0C000&
      Caption         =   "east beam  x"
      Height          =   311
      Left            =   0
      TabIndex        =   53
      Top             =   6480
      Width           =   1305
   End
   Begin VB.Label Label16 
      BackColor       =   &H00C0C000&
      Caption         =   "south beam y"
      Height          =   311
      Left            =   0
      TabIndex        =   52
      Top             =   6120
      Width           =   1305
   End
   Begin VB.Label Label15 
      BackColor       =   &H00C0C000&
      Caption         =   "north beam y"
      Height          =   311
      Left            =   0
      TabIndex        =   51
      Top             =   5760
      Width           =   1305
   End
   Begin VB.Label Label14 
      BackColor       =   &H00C0C000&
      Caption         =   " Column :"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   381
      Left            =   0
      TabIndex        =   50
      Top             =   5312
      Width           =   1725
   End
   Begin VB.Label Label13 
      BackColor       =   &H00C0C000&
      Caption         =   "north beam y"
      Height          =   311
      Left            =   0
      TabIndex        =   49
      Top             =   4920
      Width           =   1305
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0C000&
      Caption         =   "south beam y"
      Height          =   311
      Left            =   0
      TabIndex        =   48
      Top             =   4560
      Width           =   1305
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0C000&
      Caption         =   "east beam x"
      Height          =   311
      Left            =   0
      TabIndex        =   47
      Top             =   4200
      Width           =   1305
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0C000&
      Caption         =   "west beam  x"
      Height          =   311
      Left            =   0
      TabIndex        =   46
      Top             =   3840
      Width           =   1305
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0C000&
      Caption         =   "upper col"
      Height          =   311
      Left            =   0
      TabIndex        =   45
      Top             =   3480
      Width           =   1305
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   456
      Left            =   3000
      TabIndex        =   5
      Top             =   3000
      Width           =   795
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   " H"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   456
      Left            =   2160
      TabIndex        =   4
      Top             =   3000
      Width           =   795
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "B "
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   456
      Left            =   1320
      TabIndex        =   3
      Top             =   3000
      Width           =   795
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "GothicE"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   15
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      Caption         =   "FAIL DWG"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   161
      Left            =   5160
      TabIndex        =   1
      Top             =   105
      Width           =   5760
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "      TIANG MS1195."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   960
      TabIndex        =   0
      Top             =   90
      Width           =   2715
   End
   Begin VB.Menu mnuItemFile 
      Caption         =   "&File"
      Begin VB.Menu mnuItemOpenDwg 
         Caption         =   "&OpenDWG"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuItemExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
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
'''THE PROGRAM IS AN INPUT TEMPLATE FOR STRUCTURAL REINFORCED CONCRETE COLUMN,  '''
'''TO INTERFACE INTO AUTOCAD ENVIRONMENT WITH PRE-DESIGNED ALGORITHMS THAT    '''
'''ENABLE TO CONVERT THE INPUT DATA AUTOMATICALLY INTO STRUCT. R.C. DRAWINGS. '''
'''CREATED IN 2001 BY : WAN SOHAIMI BIN WAN MOHAMED.                          '''
'''(LATEST REVISION FEB 2002)- [butiran tiang sahaja]                         '''
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

Option Base 1

Dim slabThick, cVr, gAp, stirupD, StirupSPACE As Double
Dim fcu, fy, fyv, shrink, creep, Ast, dPrime As Double
Dim BarMark, i As Integer
Dim Xinsertion, Yinsertion As Double
Dim NbhXX(1 To 3000), Mbh2XX(1 To 3000) As Double
Dim NbhYY(1 To 3000), Mbh2YY(1 To 3000) As Double
Dim DesgMntXX(1 To 3000), DesgMntYY(1 To 3000) As Double
Dim PointX, PointY As Integer
Dim EffHeightXX(1 To 9), EffHeightYY(1 To 9) As Double

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''FOR TIANG'''''''''''''''''''''''''''''''''
Dim CoverTng(1 To 9) As Double
Dim NamaJadualTiang, GridTng(1 To 9) As String
Dim bilJenisTng As Integer
Dim bTng(1 To 9), hTng(1 To 9), FloorHgt(1 To 9) As Double
Dim BarXTopN(1 To 9), BarXTopD(1 To 9) As Integer
Dim BarXBotN(1 To 9), BarXBotD(1 To 9) As Integer
Dim BarYLefN(1 To 9), BarYLefD(1 To 9) As Integer
Dim BarYRigN(1 To 9), BarYRigD(1 To 9) As Integer
Dim DesgAxial(1 To 9), Mx1(1 To 9), Mx2(1 To 9) As Double
Dim My1(1 To 9), My2(1 To 9) As Double

Dim LowCoHgt(1 To 9), LowCoB(1 To 9), LowCoH(1 To 9) As Double
Dim UppCoHgt(1 To 9), UppCoB(1 To 9), UppCoH(1 To 9) As Double

Dim xBracedCol(1 To 9), xBASEresistmnt(1 To 9) As String
Dim xLowLeBl(1 To 9), xLowLeBb(1 To 9), xLowLeBh(1 To 9) As Double
Dim xLowRiBl(1 To 9), xLowRiBb(1 To 9), xLowRiBh(1 To 9) As Double
Dim xUppLeBl(1 To 9), xUppLeBb(1 To 9), xUppLeBh(1 To 9) As Double
Dim xUppRiBl(1 To 9), xUppRiBb(1 To 9), xUppRiBh(1 To 9) As Double

Dim yBracedCol(1 To 9), yBASEresistmnt(1 To 10) As String
Dim yLowLeBl(1 To 9), yLowLeBb(1 To 9), yLowLeBh(1 To 9) As Double
Dim yLowRiBl(1 To 9), yLowRiBb(1 To 9), yLowRiBh(1 To 9) As Double
Dim yUppLeBl(1 To 9), yUppLeBb(1 To 9), yUppLeBh(1 To 9) As Double
Dim yUppRiBl(1 To 9), yUppRiBb(1 To 9), yUppRiBh(1 To 9) As Double
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim Components As Integer
Public DataFile, NamaFolder As String
Public dwgName As String
Public acadApp As Object
Public acadDoc As Object
Public moSpace As Object
Public paSpace As Object
''''''''''''''''''''''''''''''''''''''''''
Dim StressCapacity1 As New StressCapacity
''''''''''''''''''''''''''''''''''''''''''
Dim Dimensi1 As New Dimensi
Dim Tetulang1 As New Tetulang
Dim AppliedStress1 As New AppliedStress
''''''''''''''''''''''''''''''''''''''''''
Dim Dimensi2 As New Dimensi
Dim Tetulang2 As New Tetulang
Dim AppliedStress2 As New AppliedStress
''''''''''''''''''''''''''''''''''''''''''
Dim Dimensi3 As New Dimensi
Dim Tetulang3 As New Tetulang
Dim AppliedStress3 As New AppliedStress
''''''''''''''''''''''''''''''''''''''''''
Dim Dimensi4 As New Dimensi
Dim Tetulang4 As New Tetulang
Dim AppliedStress4 As New AppliedStress
''''''''''''''''''''''''''''''''''''''''''
Dim Dimensi5 As New Dimensi
Dim Tetulang5 As New Tetulang
Dim AppliedStress5 As New AppliedStress
''''''''''''''''''''''''''''''''''''''''''
Dim Dimensi6 As New Dimensi
Dim Tetulang6 As New Tetulang
Dim AppliedStress6 As New AppliedStress
''''''''''''''''''''''''''''''''''''''''''
Dim Dimensi7 As New Dimensi
Dim Tetulang7 As New Tetulang
Dim AppliedStress7 As New AppliedStress
''''''''''''''''''''''''''''''''''''''''''
Dim Dimensi8 As New Dimensi
Dim Tetulang8 As New Tetulang
Dim AppliedStress8 As New AppliedStress
''''''''''''''''''''''''''''''''''''''''''
Dim Dimensi9 As New Dimensi
Dim Tetulang9 As New Tetulang
Dim AppliedStress9 As New AppliedStress
''''''''''''''''''''''''''''''''''''''''''







Private Sub Command1_Click()
ReaDFileTiang
bilJenisTng = Val(Right(Command2.Caption, 1))
Form1.Picture = LoadPicture(NamaFolder & "icon\datam.ico")
'''CloseAllVisibility
Command2.Enabled = False
Command3.Enabled = True
Command4.Enabled = False
Command5.Enabled = False
Dim fnum As Integer
Dim txtFile As String
fnum = FreeFile
               
        Picture1.Enabled = False
        Picture1.Top = 900
        Picture1.Left = 0
        Picture1.Height = 50
        Picture1.Width = 50
        Picture1.Visible = False
        Picture1.Cls
        
Command1.Enabled = False
Option1.Enabled = False
Option2.Enabled = False
Option3.Enabled = False
Option4.Enabled = False
Option5.Enabled = False
Option6.Enabled = False
Option7.Enabled = False
Option8.Enabled = False
Option9.Enabled = False

'If mnuItemOpenDwg.Enabled = True Or Command1.Enabled = True Then
'Command4.Enabled = False
'Else
'Command4.Enabled = True
'End If

Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
Text7.Enabled = False
Text8.Enabled = False
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
Text20.Enabled = False
Text21.Enabled = False
Text22.Enabled = False
Text23.Enabled = False   ''''
Text24.Enabled = False
Text25.Enabled = False
Text26.Enabled = False   ''''
Text27.Enabled = False
Text28.Enabled = False
Text29.Enabled = False
Text30.Enabled = False
Text31.Enabled = False
Text32.Enabled = False
Text33.Enabled = False
Text34.Enabled = False
Text35.Enabled = False
Text36.Enabled = False
Text37.Enabled = False
Text38.Enabled = False
Text39.Enabled = False
Text40.Enabled = False
Text41.Enabled = False
Text42.Enabled = False
Text43.Enabled = False
Text44.Enabled = False
Text45.Enabled = False
Text46.Enabled = False
Text47.Enabled = False
Text48.Enabled = False
Text49.Enabled = False
Text50.Enabled = False

Command1.Enabled = False
Text51.Enabled = True
Text52.Enabled = True
Text53.Enabled = True
Text54.Enabled = True
Text55.Enabled = True
Text56.Enabled = True
Text57.Enabled = True
Text58.Enabled = True
Text59.Enabled = True
Text60.Enabled = True

Text61.Enabled = False
Text62.Enabled = False

''''***********************************************************''''
Text51.text = Xinsertion
Text52.text = Yinsertion
Text53.text = fcu
Text54.text = fy
Text55.text = fyv
Text56.text = shrink
Text57.text = creep
Text58.text = stirupD
Text59.text = StirupSPACE
Text60.text = BarMark

''*************************************************************''

StressCapacity1.Value51 = Text51.text
StressCapacity1.Value52 = Text52.text
StressCapacity1.Value53 = Text53.text
StressCapacity1.Value54 = Text54.text
StressCapacity1.Value55 = Text55.text
StressCapacity1.Value56 = Text56.text
StressCapacity1.Value57 = Text57.text
StressCapacity1.Value58 = Text58.text
StressCapacity1.Value59 = Text59.text
StressCapacity1.Value60 = Text60.text

''''**********************************************************''''

Text51.text = StressCapacity1.XinsertPoint
Text52.text = StressCapacity1.YinsertPoint
Text53.text = StressCapacity1.fcuNEW
Text54.text = StressCapacity1.fyNEW
Text55.text = StressCapacity1.fyvNEW
Text56.text = StressCapacity1.ShrinkNEW
Text57.text = StressCapacity1.CreepNEW
Text58.text = StressCapacity1.LinkDiameter
Text59.text = StressCapacity1.LinkSpacing
Text60.text = StressCapacity1.BarMarkNEW

''************************************************************''


End Sub

'''''''''''''''''''''''''''TIANG'''''''''''''''''''''''''''''''''
Private Sub Command2_Click()
Dim Locker As Integer
Locker = 0
If Locker = 0 Then
ShowAllVisibility
Locker = 1
End If

'''Command1.Enabled = True
      
Command2.Enabled = False
Command5.Enabled = True
Dim fnum As Integer
Dim txtFile As String
fnum = FreeFile
Label54.Caption = DesgAxial(bilJenisTng) & " kN."
Form1.Picture = LoadPicture(NamaFolder & "icon\ukad4.ico")
'''''''''''''''''''''''''
        Picture1.Enabled = False
        Picture1.Top = 900
        Picture1.Left = 0
        Picture1.Height = 50
        Picture1.Width = 50
        Picture1.Visible = False
        Picture1.Cls

Option1.Enabled = True
Option2.Enabled = True
Option3.Enabled = True
Option4.Enabled = True
Option5.Enabled = True
Option6.Enabled = True
Option7.Enabled = True
Option8.Enabled = True
Option9.Enabled = True

If mnuItemOpenDwg.Enabled = True Then
Command4.Enabled = False
Else
Command4.Enabled = True
End If

DisableText1To62
''''*****************************************************''''

If Command2.Left = 110 Then
Option1.Enabled = False
txtFile = NamaFolder & "tiang\datainput\ColOneGET.txt"
Call TransferIData(1)
SetBraceAndBase (1)
'''''''''save data''''''''''''''
Open txtFile For Output As #fnum


Print #fnum, UppCoB(1), UppCoH(1), UppCoHgt(1)

Print #fnum, yUppLeBb(1), yUppLeBh(1), yUppLeBl(1)
Print #fnum, yUppRiBb(1), yUppRiBh(1), yUppRiBl(1)

Print #fnum, xUppRiBb(1), xUppRiBh(1), xUppRiBl(1)
Print #fnum, xUppLeBb(1), xUppLeBh(1), xUppLeBl(1)

Print #fnum, bTng(1), hTng(1), FloorHgt(1)

Print #fnum, Mx2(1), My2(1)
Print #fnum, Mx1(1), My1(1)
Print #fnum, DesgAxial(1), CoverTng(1)

Print #fnum, xLowLeBb(1), xLowLeBh(1), xLowLeBl(1)
Print #fnum, xLowRiBb(1), xLowRiBh(1), xLowRiBl(1)

Print #fnum, yLowRiBb(1), yLowRiBh(1), yLowRiBl(1)
Print #fnum, yLowLeBb(1), yLowLeBh(1), yLowLeBl(1)

Print #fnum, LowCoB(1), LowCoH(1), LowCoHgt(1)

Print #fnum, BarYLefN(1), BarYLefD(1)
Print #fnum, BarYRigN(1), BarYRigD(1)

Print #fnum, BarXTopN(1), BarXTopD(1)
Print #fnum, BarXBotN(1), BarXBotD(1)

Print #fnum, xBracedCol(1)
Print #fnum, yBracedCol(1)
Print #fnum, xBASEresistmnt(1)
Print #fnum, yBASEresistmnt(1)
Print #fnum, GridTng(1)
Close #fnum
End If

''''''*********************************************************''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

If Command2.Left = 1110 Then
Option2.Enabled = False
txtFile = NamaFolder & "tiang\datainput\ColTwoGET.txt"
Call TransferIData(2)
SetBraceAndBase (2)

'''''''''save data''''''''''''''
Open txtFile For Output As #fnum

Print #fnum, UppCoB(2), UppCoH(2), UppCoHgt(2)

Print #fnum, yUppLeBb(2), yUppLeBh(2), yUppLeBl(2)
Print #fnum, yUppRiBb(2), yUppRiBh(2), yUppRiBl(2)

Print #fnum, xUppRiBb(2), xUppRiBh(2), xUppRiBl(2)
Print #fnum, xUppLeBb(2), xUppLeBh(2), xUppLeBl(2)

Print #fnum, bTng(2), hTng(2), FloorHgt(2)

Print #fnum, Mx2(2), My2(2)
Print #fnum, Mx1(2), My1(2)
Print #fnum, DesgAxial(2), CoverTng(2)

Print #fnum, xLowLeBb(2), xLowLeBh(2), xLowLeBl(2)
Print #fnum, xLowRiBb(2), xLowRiBh(2), xLowRiBl(2)

Print #fnum, yLowRiBb(2), yLowRiBh(2), yLowRiBl(2)
Print #fnum, yLowLeBb(2), yLowLeBh(2), yLowLeBl(2)

Print #fnum, LowCoB(2), LowCoH(2), LowCoHgt(2)

Print #fnum, BarYLefN(2), BarYLefD(2)
Print #fnum, BarYRigN(2), BarYRigD(2)

Print #fnum, BarXTopN(2), BarXTopD(2)
Print #fnum, BarXBotN(2), BarXBotD(2)

Print #fnum, xBracedCol(2)
Print #fnum, yBracedCol(2)
Print #fnum, xBASEresistmnt(2)
Print #fnum, yBASEresistmnt(2)
Print #fnum, GridTng(2)
Close #fnum
End If

''''''*********************************************************''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

If Command2.Left = 2110 Then
Option3.Enabled = False
txtFile = NamaFolder & "tiang\datainput\ColThreeGET.txt"
Call TransferIData(3)
SetBraceAndBase (3)
'''''''''save data''''''''''''''
Open txtFile For Output As #fnum

Print #fnum, UppCoB(3), UppCoH(3), UppCoHgt(3)

Print #fnum, yUppLeBb(3), yUppLeBh(3), yUppLeBl(3)
Print #fnum, yUppRiBb(3), yUppRiBh(3), yUppRiBl(3)

Print #fnum, xUppRiBb(3), xUppRiBh(3), xUppRiBl(3)
Print #fnum, xUppLeBb(3), xUppLeBh(3), xUppLeBl(3)

Print #fnum, bTng(3), hTng(3), FloorHgt(3)

Print #fnum, Mx2(3), My2(3)
Print #fnum, Mx1(3), My1(3)
Print #fnum, DesgAxial(3), CoverTng(3)

Print #fnum, xLowLeBb(3), xLowLeBh(3), xLowLeBl(3)
Print #fnum, xLowRiBb(3), xLowRiBh(3), xLowRiBl(3)

Print #fnum, yLowRiBb(3), yLowRiBh(3), yLowRiBl(3)
Print #fnum, yLowLeBb(3), yLowLeBh(3), yLowLeBl(3)

Print #fnum, LowCoB(3), LowCoH(3), LowCoHgt(3)

Print #fnum, BarYLefN(3), BarYLefD(3)
Print #fnum, BarYRigN(3), BarYRigD(3)

Print #fnum, BarXTopN(3), BarXTopD(3)
Print #fnum, BarXBotN(3), BarXBotD(3)

Print #fnum, xBracedCol(3)
Print #fnum, yBracedCol(3)
Print #fnum, xBASEresistmnt(3)
Print #fnum, yBASEresistmnt(3)
Print #fnum, GridTng(3)

Close #fnum
End If

''''''*********************************************************''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

If Command2.Left = 3110 Then
Option4.Enabled = False
txtFile = NamaFolder & "tiang\datainput\ColFourGET.txt"
Call TransferIData(4)
SetBraceAndBase (4)
'''''''''save data''''''''''''''
Open txtFile For Output As #fnum

Print #fnum, UppCoB(4), UppCoH(4), UppCoHgt(4)

Print #fnum, yUppLeBb(4), yUppLeBh(4), yUppLeBl(4)
Print #fnum, yUppRiBb(4), yUppRiBh(4), yUppRiBl(4)

Print #fnum, xUppRiBb(4), xUppRiBh(4), xUppRiBl(4)
Print #fnum, xUppLeBb(4), xUppLeBh(4), xUppLeBl(4)

Print #fnum, bTng(4), hTng(4), FloorHgt(4)

Print #fnum, Mx2(4), My2(4)
Print #fnum, Mx1(4), My1(4)
Print #fnum, DesgAxial(4), CoverTng(4)

Print #fnum, xLowLeBb(4), xLowLeBh(4), xLowLeBl(4)
Print #fnum, xLowRiBb(4), xLowRiBh(4), xLowRiBl(4)

Print #fnum, yLowRiBb(4), yLowRiBh(4), yLowRiBl(4)
Print #fnum, yLowLeBb(4), yLowLeBh(4), yLowLeBl(4)

Print #fnum, LowCoB(4), LowCoH(4), LowCoHgt(4)

Print #fnum, BarYLefN(4), BarYLefD(4)
Print #fnum, BarYRigN(4), BarYRigD(4)

Print #fnum, BarXTopN(4), BarXTopD(4)
Print #fnum, BarXBotN(4), BarXBotD(4)

Print #fnum, xBracedCol(4)
Print #fnum, yBracedCol(4)
Print #fnum, xBASEresistmnt(4)
Print #fnum, yBASEresistmnt(4)
Print #fnum, GridTng(4)
Close #fnum
End If


''''''*********************************************************''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

If Command2.Left = 4110 Then
Option5.Enabled = False
txtFile = NamaFolder & "tiang\datainput\ColFiveGET.txt"
Call TransferIData(5)
SetBraceAndBase (5)
'''''''''save data''''''''''''''
Open txtFile For Output As #fnum

Print #fnum, UppCoB(5), UppCoH(5), UppCoHgt(5)

Print #fnum, yUppLeBb(5), yUppLeBh(5), yUppLeBl(5)
Print #fnum, yUppRiBb(5), yUppRiBh(5), yUppRiBl(5)

Print #fnum, xUppRiBb(5), xUppRiBh(5), xUppRiBl(5)
Print #fnum, xUppLeBb(5), xUppLeBh(5), xUppLeBl(5)

Print #fnum, bTng(5), hTng(5), FloorHgt(5)


Print #fnum, Mx2(5), My2(5)
Print #fnum, Mx1(5), My1(5)
Print #fnum, DesgAxial(5), CoverTng(5)

Print #fnum, xLowLeBb(5), xLowLeBh(5), xLowLeBl(5)
Print #fnum, xLowRiBb(5), xLowRiBh(5), xLowRiBl(5)

Print #fnum, yLowRiBb(5), yLowRiBh(5), yLowRiBl(5)
Print #fnum, yLowLeBb(5), yLowLeBh(5), yLowLeBl(5)

Print #fnum, LowCoB(5), LowCoH(5), LowCoHgt(5)

Print #fnum, BarYLefN(5), BarYLefD(5)
Print #fnum, BarYRigN(5), BarYRigD(5)

Print #fnum, BarXTopN(5), BarXTopD(5)
Print #fnum, BarXBotN(5), BarXBotD(5)

Print #fnum, xBracedCol(5)
Print #fnum, yBracedCol(5)
Print #fnum, xBASEresistmnt(5)
Print #fnum, yBASEresistmnt(5)
Print #fnum, GridTng(5)

Close #fnum
End If

''''''*********************************************************''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

If Command2.Left = 5110 Then
Option6.Enabled = False
txtFile = NamaFolder & "tiang\datainput\ColSixGET.txt"
Call TransferIData(6)
SetBraceAndBase (6)
'''''''''save data''''''''''''''
Open txtFile For Output As #fnum


Print #fnum, UppCoB(6), UppCoH(6), UppCoHgt(6)

Print #fnum, yUppLeBb(6), yUppLeBh(6), yUppLeBl(6)
Print #fnum, yUppRiBb(6), yUppRiBh(6), yUppRiBl(6)

Print #fnum, xUppRiBb(6), xUppRiBh(6), xUppRiBl(6)
Print #fnum, xUppLeBb(6), xUppLeBh(6), xUppLeBl(6)

Print #fnum, bTng(6), hTng(6), FloorHgt(6)


Print #fnum, Mx2(6), My2(6)
Print #fnum, Mx1(6), My1(6)
Print #fnum, DesgAxial(6), CoverTng(6)

Print #fnum, xLowLeBb(6), xLowLeBh(6), xLowLeBl(6)
Print #fnum, xLowRiBb(6), xLowRiBh(6), xLowRiBl(6)

Print #fnum, yLowRiBb(6), yLowRiBh(6), yLowRiBl(6)
Print #fnum, yLowLeBb(6), yLowLeBh(6), yLowLeBl(6)

Print #fnum, LowCoB(6), LowCoH(6), LowCoHgt(6)

Print #fnum, BarYLefN(6), BarYLefD(6)
Print #fnum, BarYRigN(6), BarYRigD(6)

Print #fnum, BarXTopN(6), BarXTopD(6)
Print #fnum, BarXBotN(6), BarXBotD(6)

Print #fnum, xBracedCol(6)
Print #fnum, yBracedCol(6)
Print #fnum, xBASEresistmnt(6)
Print #fnum, yBASEresistmnt(6)
Print #fnum, GridTng(6)

Close #fnum
End If


''''''''''''''''''''***********************''''''''''''''''''

If Command2.Left = 6110 Then
Option7.Enabled = False
txtFile = NamaFolder & "tiang\datainput\ColSevenGET.txt"
Call TransferIData(7)
SetBraceAndBase (7)
'''''''''save data''''''''''''''
Open txtFile For Output As #fnum

Print #fnum, UppCoB(7), UppCoH(7), UppCoHgt(7)

Print #fnum, yUppLeBb(7), yUppLeBh(7), yUppLeBl(7)
Print #fnum, yUppRiBb(7), yUppRiBh(7), yUppRiBl(7)

Print #fnum, xUppRiBb(7), xUppRiBh(7), xUppRiBl(7)
Print #fnum, xUppLeBb(7), xUppLeBh(7), xUppLeBl(7)

Print #fnum, bTng(7), hTng(7), FloorHgt(7)

Print #fnum, Mx2(7), My2(7)
Print #fnum, Mx1(7), My1(7)
Print #fnum, DesgAxial(7), CoverTng(7)

Print #fnum, xLowLeBb(7), xLowLeBh(7), xLowLeBl(7)
Print #fnum, xLowRiBb(7), xLowRiBh(7), xLowRiBl(7)

Print #fnum, yLowRiBb(7), yLowRiBh(7), yLowRiBl(7)
Print #fnum, yLowLeBb(7), yLowLeBh(7), yLowLeBl(7)

Print #fnum, LowCoB(7), LowCoH(7), LowCoHgt(7)

Print #fnum, BarYLefN(7), BarYLefD(7)
Print #fnum, BarYRigN(7), BarYRigD(7)

Print #fnum, BarXTopN(7), BarXTopD(7)
Print #fnum, BarXBotN(7), BarXBotD(7)

Print #fnum, xBracedCol(7)
Print #fnum, yBracedCol(7)
Print #fnum, xBASEresistmnt(7)
Print #fnum, yBASEresistmnt(7)
Print #fnum, GridTng(7)

Close #fnum
End If

''''''''''''''''''''***********************''''''''''''''''''

If Command2.Left = 7110 Then
Option8.Enabled = False
txtFile = NamaFolder & "tiang\datainput\ColEightGET.txt"
Call TransferIData(8)
SetBraceAndBase (8)
'''''''''save data''''''''''''''
Open txtFile For Output As #fnum

Print #fnum, UppCoB(8), UppCoH(8), UppCoHgt(8)

Print #fnum, yUppLeBb(8), yUppLeBh(8), yUppLeBl(8)
Print #fnum, yUppRiBb(8), yUppRiBh(8), yUppRiBl(8)

Print #fnum, xUppRiBb(8), xUppRiBh(8), xUppRiBl(8)
Print #fnum, xUppLeBb(8), xUppLeBh(8), xUppLeBl(8)

Print #fnum, bTng(8), hTng(8), FloorHgt(8)

Print #fnum, Mx2(8), My2(8)
Print #fnum, Mx1(8), My1(8)
Print #fnum, DesgAxial(8), CoverTng(8)

Print #fnum, xLowLeBb(8), xLowLeBh(8), xLowLeBl(8)
Print #fnum, xLowRiBb(8), xLowRiBh(8), xLowRiBl(8)

Print #fnum, yLowRiBb(8), yLowRiBh(8), yLowRiBl(8)
Print #fnum, yLowLeBb(8), yLowLeBh(8), yLowLeBl(8)

Print #fnum, LowCoB(8), LowCoH(8), LowCoHgt(8)

Print #fnum, BarYLefN(8), BarYLefD(8)
Print #fnum, BarYRigN(8), BarYRigD(8)

Print #fnum, BarXTopN(8), BarXTopD(8)
Print #fnum, BarXBotN(8), BarXBotD(8)

Print #fnum, xBracedCol(8)
Print #fnum, yBracedCol(8)
Print #fnum, xBASEresistmnt(8)
Print #fnum, yBASEresistmnt(8)
Print #fnum, GridTng(8)

Close #fnum
End If

''''''''''''''''''''***********************''''''''''''''''''

If Command2.Left = 8110 Then
Option9.Enabled = False
txtFile = NamaFolder & "tiang\datainput\ColNineGET.txt"
Call TransferIData(9)
SetBraceAndBase (9)
'''''''''save data''''''''''''''
Open txtFile For Output As #fnum

Print #fnum, UppCoB(9), UppCoH(9), UppCoHgt(9)

Print #fnum, yUppLeBb(9), yUppLeBh(9), yUppLeBl(9)
Print #fnum, yUppRiBb(9), yUppRiBh(9), yUppRiBl(9)

Print #fnum, xUppRiBb(9), xUppRiBh(9), xUppRiBl(9)
Print #fnum, xUppLeBb(9), xUppLeBh(9), xUppLeBl(9)

Print #fnum, bTng(9), hTng(9), FloorHgt(9)

Print #fnum, Mx2(9), My2(9)
Print #fnum, Mx1(9), My1(9)
Print #fnum, DesgAxial(9), CoverTng(9)

Print #fnum, xLowLeBb(9), xLowLeBh(9), xLowLeBl(9)
Print #fnum, xLowRiBb(9), xLowRiBh(9), xLowRiBl(9)

Print #fnum, yLowRiBb(9), yLowRiBh(9), yLowRiBl(9)
Print #fnum, yLowLeBb(9), yLowLeBh(9), yLowLeBl(9)

Print #fnum, LowCoB(9), LowCoH(9), LowCoHgt(9)

Print #fnum, BarYLefN(9), BarYLefD(9)
Print #fnum, BarYRigN(9), BarYRigD(9)

Print #fnum, BarXTopN(9), BarXTopD(9)
Print #fnum, BarXBotN(9), BarXBotD(9)

Print #fnum, xBracedCol(9)
Print #fnum, yBracedCol(9)
Print #fnum, xBASEresistmnt(9)
Print #fnum, yBASEresistmnt(9)
Print #fnum, GridTng(9)

Close #fnum
End If

''MsgBox "no tng :", , bilJenisTng
ReaDFileTiang

ShowExistingMember (Int(Right(Command2.Caption, 1)))
bilJenisTng = Val(Right(Command2.Caption, 1))
Call GraphicOne(Val(bTng(bilJenisTng)), Val(hTng(bilJenisTng)), _
       Val(BarXBotN(bilJenisTng)) * Val(BarXBotD(bilJenisTng)) ^ 2 * 3.14 / 4, _
       Val(BarXBotD(bilJenisTng)) ^ 2 * 3.14 / 2 + _
       Val(BarYLefN(bilJenisTng)) * Val(BarYLefD(bilJenisTng)) ^ 2 * 3.14 / 4)
       
Form1.Picture = LoadPicture(NamaFolder & "icon\datas.ico")
Label54.Caption = DesgAxial(bilJenisTng) & " kN."
End Sub

Private Sub CalColumnStrength()

Dim delX, delY, delXt, delYr As Double
Dim barLayerYY As Integer


If (BarXBotN(i) - 1) < 1 Then
   delX = (bTng(i) - 2 * CoverTng(i) - 2 * stirupD) / 2
   Else
   delX = (bTng(i) - 2 * CoverTng(i) - 2 * stirupD) / (BarXBotN(i) - 1)
End If
delY = (hTng(i) - 2 * CoverTng(i) - 2 * stirupD) / (BarYLefN(i) + 1)
If (BarXTopN(i) - 1) < 1 Then
   delXt = (bTng(i) - 2 * CoverTng(i) - 2 * stirupD) / 2
   Else
   delXt = (bTng(i) - 2 * CoverTng(i) - 2 * stirupD) / (BarXTopN(i) - 1)
End If
delYr = (hTng(i) - 2 * CoverTng(i) - 2 * stirupD) / (BarYRigN(i) + 1)
'''''''''''''''''''''''''''''''''''''''''''''''''


Call Column(Xinsertion, Yinsertion, _
    delY, bTng(i), hTng(i), CoverTng(i) + stirupD + BarXBotD(i) / 2, BarYLefN(i) + 2, _
    BarXBotD(i), BarYLefD(i), BarXBotN(i), BarYLefN(i), DesgAxial(i), _
    Mx1(i), Mx2(i), "XX", "BeamName")
    
      If BarXBotN(i) = 1 Then
           barLayerYY = BarXBotN(i) + 2
        Else
           barLayerYY = BarXBotN(i)
      End If
            
Call Column(Xinsertion, Yinsertion, _
    delX, hTng(i), bTng(i), CoverTng(i) + stirupD + BarYLefD(i) / 2, barLayerYY, _
    BarXBotD(i), BarYLefD(i), BarXBotN(i), BarYLefN(i), DesgAxial(i), _
    My1(i), My2(i), "YY", "BeamName")
    
    
End Sub
Private Sub StartAutoCAD()

On Error Resume Next
  
Set acadApp = GetObject(, "AutoCAD.Application")
If Err Then
   Err.Clear
   Set acadApp = CreateObject("AutoCAD.Application")
      If Err Then
       MsgBox Err.Description
       Exit Sub
      End If
End If

Set acadDoc = acadApp.ActiveDocument
acadDoc.Open dwgName
acadApp.Visible = True
acadApp.Top = 0
acadApp.Left = 0
acadApp.Width = 900
acadApp.Height = 800
Form1.Picture = LoadPicture(NamaFolder & "icon\statacad.ico")

End Sub


''''FOR TIANG''''
Private Sub ReaDFileTiang()

''''*********************************************'''

Dim fnum As Integer
Dim txtFile, Temp As String
fnum = FreeFile

txtFile = NamaFolder & "tiang\datainput\DefaultStressTiang.txt"
Open txtFile For Input As #fnum
Input #fnum, Xinsertion
Input #fnum, Yinsertion
''Input #fnum, NamaTiang
Input #fnum, fcu
Input #fnum, fy
Input #fnum, fyv
Input #fnum, shrink
Input #fnum, creep
''Input #fnum, cVr
''Input #fnum, slabThick
Input #fnum, stirupD
Input #fnum, StirupSPACE
Input #fnum, BarMark
Close #fnum

txtFile = NamaFolder & "tiang\datainput\ColOneGET.txt"
Open txtFile For Input As #fnum

Input #fnum, UppCoB(1), UppCoH(1), UppCoHgt(1)

Input #fnum, yUppLeBb(1), yUppLeBh(1), yUppLeBl(1)
Input #fnum, yUppRiBb(1), yUppRiBh(1), yUppRiBl(1)

Input #fnum, xUppRiBb(1), xUppRiBh(1), xUppRiBl(1)
Input #fnum, xUppLeBb(1), xUppLeBh(1), xUppLeBl(1)

Input #fnum, bTng(1), hTng(1), FloorHgt(1)

Input #fnum, Mx2(1), My2(1)
Input #fnum, Mx1(1), My1(1)
Input #fnum, DesgAxial(1), CoverTng(1)

Input #fnum, xLowLeBb(1), xLowLeBh(1), xLowLeBl(1)
Input #fnum, xLowRiBb(1), xLowRiBh(1), xLowRiBl(1)

Input #fnum, yLowRiBb(1), yLowRiBh(1), yLowRiBl(1)
Input #fnum, yLowLeBb(1), yLowLeBh(1), yLowLeBl(1)

Input #fnum, LowCoB(1), LowCoH(1), LowCoHgt(1)

Input #fnum, BarYLefN(1), BarYLefD(1)
Input #fnum, BarYRigN(1), BarYRigD(1)

Input #fnum, BarXTopN(1), BarXTopD(1)
Input #fnum, BarXBotN(1), BarXBotD(1)

Input #fnum, xBracedCol(1)
Input #fnum, yBracedCol(1)
Input #fnum, xBASEresistmnt(1)
Input #fnum, yBASEresistmnt(1)
Input #fnum, GridTng(1)
Close #fnum
''''''*****************************************************''''''

txtFile = NamaFolder & "tiang\datainput\ColTwoGET.txt"
Open txtFile For Input As #fnum

Input #fnum, UppCoB(2), UppCoH(2), UppCoHgt(2)

Input #fnum, yUppLeBb(2), yUppLeBh(2), yUppLeBl(2)
Input #fnum, yUppRiBb(2), yUppRiBh(2), yUppRiBl(2)

Input #fnum, xUppRiBb(2), xUppRiBh(2), xUppRiBl(2)
Input #fnum, xUppLeBb(2), xUppLeBh(2), xUppLeBl(2)

Input #fnum, bTng(2), hTng(2), FloorHgt(2)

Input #fnum, Mx2(2), My2(2)
Input #fnum, Mx1(2), My1(2)
Input #fnum, DesgAxial(2), CoverTng(2)

Input #fnum, xLowLeBb(2), xLowLeBh(2), xLowLeBl(2)
Input #fnum, xLowRiBb(2), xLowRiBh(2), xLowRiBl(2)

Input #fnum, yLowRiBb(2), yLowRiBh(2), yLowRiBl(2)
Input #fnum, yLowLeBb(2), yLowLeBh(2), yLowLeBl(2)

Input #fnum, LowCoB(2), LowCoH(2), LowCoHgt(2)

Input #fnum, BarYLefN(2), BarYLefD(2)
Input #fnum, BarYRigN(2), BarYRigD(2)

Input #fnum, BarXTopN(2), BarXTopD(2)
Input #fnum, BarXBotN(2), BarXBotD(2)

Input #fnum, xBracedCol(2)
Input #fnum, yBracedCol(2)
Input #fnum, xBASEresistmnt(2)
Input #fnum, yBASEresistmnt(2)
Input #fnum, GridTng(2)
Close #fnum
''''''''********************************************************''''''

txtFile = NamaFolder & "tiang\datainput\ColThreeGET.txt"
Open txtFile For Input As #fnum

Input #fnum, UppCoB(3), UppCoH(3), UppCoHgt(3)

Input #fnum, yUppLeBb(3), yUppLeBh(3), yUppLeBl(3)
Input #fnum, yUppRiBb(3), yUppRiBh(3), yUppRiBl(3)

Input #fnum, xUppRiBb(3), xUppRiBh(3), xUppRiBl(3)
Input #fnum, xUppLeBb(3), xUppLeBh(3), xUppLeBl(3)

Input #fnum, bTng(3), hTng(3), FloorHgt(3)


Input #fnum, Mx2(3), My2(3)
Input #fnum, Mx1(3), My1(3)
Input #fnum, DesgAxial(3), CoverTng(3)

Input #fnum, xLowLeBb(3), xLowLeBh(3), xLowLeBl(3)
Input #fnum, xLowRiBb(3), xLowRiBh(3), xLowRiBl(3)

Input #fnum, yLowRiBb(3), yLowRiBh(3), yLowRiBl(3)
Input #fnum, yLowLeBb(3), yLowLeBh(3), yLowLeBl(3)

Input #fnum, LowCoB(3), LowCoH(3), LowCoHgt(3)

Input #fnum, BarYLefN(3), BarYLefD(3)
Input #fnum, BarYRigN(3), BarYRigD(3)

Input #fnum, BarXTopN(3), BarXTopD(3)
Input #fnum, BarXBotN(3), BarXBotD(3)

Input #fnum, xBracedCol(3)
Input #fnum, yBracedCol(3)
Input #fnum, xBASEresistmnt(3)
Input #fnum, yBASEresistmnt(3)
Input #fnum, GridTng(3)

Close #fnum


''''''''********************************************************''''''

txtFile = NamaFolder & "tiang\datainput\ColFourGET.txt"
Open txtFile For Input As #fnum

Input #fnum, UppCoB(4), UppCoH(4), UppCoHgt(4)

Input #fnum, yUppLeBb(4), yUppLeBh(4), yUppLeBl(4)
Input #fnum, yUppRiBb(4), yUppRiBh(4), yUppRiBl(4)

Input #fnum, xUppRiBb(4), xUppRiBh(4), xUppRiBl(4)
Input #fnum, xUppLeBb(4), xUppLeBh(4), xUppLeBl(4)

Input #fnum, bTng(4), hTng(4), FloorHgt(4)


Input #fnum, Mx2(4), My2(4)
Input #fnum, Mx1(4), My1(4)
Input #fnum, DesgAxial(4), CoverTng(4)

Input #fnum, xLowLeBb(4), xLowLeBh(4), xLowLeBl(4)
Input #fnum, xLowRiBb(4), xLowRiBh(4), xLowRiBl(4)

Input #fnum, yLowRiBb(4), yLowRiBh(4), yLowRiBl(4)
Input #fnum, yLowLeBb(4), yLowLeBh(4), yLowLeBl(4)

Input #fnum, LowCoB(4), LowCoH(4), LowCoHgt(4)

Input #fnum, BarYLefN(4), BarYLefD(4)
Input #fnum, BarYRigN(4), BarYRigD(4)

Input #fnum, BarXTopN(4), BarXTopD(4)
Input #fnum, BarXBotN(4), BarXBotD(4)

Input #fnum, xBracedCol(4)
Input #fnum, yBracedCol(4)
Input #fnum, xBASEresistmnt(4)
Input #fnum, yBASEresistmnt(4)
Input #fnum, GridTng(4)

Close #fnum


''''''''********************************************************''''''

txtFile = NamaFolder & "tiang\datainput\ColFiveGET.txt"
Open txtFile For Input As #fnum

Input #fnum, UppCoB(5), UppCoH(5), UppCoHgt(5)

Input #fnum, yUppLeBb(5), yUppLeBh(5), yUppLeBl(5)
Input #fnum, yUppRiBb(5), yUppRiBh(5), yUppRiBl(5)

Input #fnum, xUppRiBb(5), xUppRiBh(5), xUppRiBl(5)
Input #fnum, xUppLeBb(5), xUppLeBh(5), xUppLeBl(5)

Input #fnum, bTng(5), hTng(5), FloorHgt(5)


Input #fnum, Mx2(5), My2(5)
Input #fnum, Mx1(5), My1(5)
Input #fnum, DesgAxial(5), CoverTng(5)

Input #fnum, xLowLeBb(5), xLowLeBh(5), xLowLeBl(5)
Input #fnum, xLowRiBb(5), xLowRiBh(5), xLowRiBl(5)

Input #fnum, yLowRiBb(5), yLowRiBh(5), yLowRiBl(5)
Input #fnum, yLowLeBb(5), yLowLeBh(5), yLowLeBl(5)

Input #fnum, LowCoB(5), LowCoH(5), LowCoHgt(5)

Input #fnum, BarYLefN(5), BarYLefD(5)
Input #fnum, BarYRigN(5), BarYRigD(5)

Input #fnum, BarXTopN(5), BarXTopD(5)
Input #fnum, BarXBotN(5), BarXBotD(5)

Input #fnum, xBracedCol(5)
Input #fnum, yBracedCol(5)
Input #fnum, xBASEresistmnt(5)
Input #fnum, yBASEresistmnt(5)
Input #fnum, GridTng(5)

Close #fnum



''''''''********************************************************''''''

txtFile = NamaFolder & "tiang\datainput\ColSixGET.txt"
Open txtFile For Input As #fnum

Input #fnum, UppCoB(6), UppCoH(6), UppCoHgt(6)

Input #fnum, yUppLeBb(6), yUppLeBh(6), yUppLeBl(6)
Input #fnum, yUppRiBb(6), yUppRiBh(6), yUppRiBl(6)

Input #fnum, xUppRiBb(6), xUppRiBh(6), xUppRiBl(6)
Input #fnum, xUppLeBb(6), xUppLeBh(6), xUppLeBl(6)

Input #fnum, bTng(6), hTng(6), FloorHgt(6)


Input #fnum, Mx2(6), My2(6)
Input #fnum, Mx1(6), My1(6)
Input #fnum, DesgAxial(6), CoverTng(6)

Input #fnum, xLowLeBb(6), xLowLeBh(6), xLowLeBl(6)
Input #fnum, xLowRiBb(6), xLowRiBh(6), xLowRiBl(6)

Input #fnum, yLowRiBb(6), yLowRiBh(6), yLowRiBl(6)
Input #fnum, yLowLeBb(6), yLowLeBh(6), yLowLeBl(6)

Input #fnum, LowCoB(6), LowCoH(6), LowCoHgt(6)

Input #fnum, BarYLefN(6), BarYLefD(6)
Input #fnum, BarYRigN(6), BarYRigD(6)

Input #fnum, BarXTopN(6), BarXTopD(6)
Input #fnum, BarXBotN(6), BarXBotD(6)

Input #fnum, xBracedCol(6)
Input #fnum, yBracedCol(6)
Input #fnum, xBASEresistmnt(6)
Input #fnum, yBASEresistmnt(6)
Input #fnum, GridTng(6)

Close #fnum


''''''''********************************************************''''''

txtFile = NamaFolder & "tiang\datainput\ColSevenGET.txt"
Open txtFile For Input As #fnum

Input #fnum, UppCoB(7), UppCoH(7), UppCoHgt(7)

Input #fnum, yUppLeBb(7), yUppLeBh(7), yUppLeBl(7)
Input #fnum, yUppRiBb(7), yUppRiBh(7), yUppRiBl(7)

Input #fnum, xUppRiBb(7), xUppRiBh(7), xUppRiBl(7)
Input #fnum, xUppLeBb(7), xUppLeBh(7), xUppLeBl(7)

Input #fnum, bTng(7), hTng(7), FloorHgt(7)


Input #fnum, Mx2(7), My2(7)
Input #fnum, Mx1(7), My1(7)
Input #fnum, DesgAxial(7), CoverTng(7)

Input #fnum, xLowLeBb(7), xLowLeBh(7), xLowLeBl(7)
Input #fnum, xLowRiBb(7), xLowRiBh(7), xLowRiBl(7)

Input #fnum, yLowRiBb(7), yLowRiBh(7), yLowRiBl(7)
Input #fnum, yLowLeBb(7), yLowLeBh(7), yLowLeBl(7)

Input #fnum, LowCoB(7), LowCoH(7), LowCoHgt(7)

Input #fnum, BarYLefN(7), BarYLefD(7)
Input #fnum, BarYRigN(7), BarYRigD(7)

Input #fnum, BarXTopN(7), BarXTopD(7)
Input #fnum, BarXBotN(7), BarXBotD(7)

Input #fnum, xBracedCol(7)
Input #fnum, yBracedCol(7)
Input #fnum, xBASEresistmnt(7)
Input #fnum, yBASEresistmnt(7)
Input #fnum, GridTng(7)

Close #fnum


''''''''********************************************************''''''

txtFile = NamaFolder & "tiang\datainput\ColEightGET.txt"
Open txtFile For Input As #fnum

Input #fnum, UppCoB(8), UppCoH(8), UppCoHgt(8)

Input #fnum, yUppLeBb(8), yUppLeBh(8), yUppLeBl(8)
Input #fnum, yUppRiBb(8), yUppRiBh(8), yUppRiBl(8)

Input #fnum, xUppRiBb(8), xUppRiBh(8), xUppRiBl(8)
Input #fnum, xUppLeBb(8), xUppLeBh(8), xUppLeBl(8)

Input #fnum, bTng(8), hTng(8), FloorHgt(8)


Input #fnum, Mx2(8), My2(8)
Input #fnum, Mx1(8), My1(8)
Input #fnum, DesgAxial(8), CoverTng(8)


Input #fnum, xLowLeBb(8), xLowLeBh(8), xLowLeBl(8)
Input #fnum, xLowRiBb(8), xLowRiBh(8), xLowRiBl(8)

Input #fnum, yLowRiBb(8), yLowRiBh(8), yLowRiBl(8)
Input #fnum, yLowLeBb(8), yLowLeBh(8), yLowLeBl(8)

Input #fnum, LowCoB(8), LowCoH(8), LowCoHgt(8)

Input #fnum, BarYLefN(8), BarYLefD(8)
Input #fnum, BarYRigN(8), BarYRigD(8)

Input #fnum, BarXTopN(8), BarXTopD(8)
Input #fnum, BarXBotN(8), BarXBotD(8)

Input #fnum, xBracedCol(8)
Input #fnum, yBracedCol(8)
Input #fnum, xBASEresistmnt(8)
Input #fnum, yBASEresistmnt(8)
Input #fnum, GridTng(8)

Close #fnum

''''''''********************************************************''''''

txtFile = NamaFolder & "tiang\datainput\ColNineGET.txt"
Open txtFile For Input As #fnum

Input #fnum, UppCoB(9), UppCoH(9), UppCoHgt(9)

Input #fnum, yUppLeBb(9), yUppLeBh(9), yUppLeBl(9)
Input #fnum, yUppRiBb(9), yUppRiBh(9), yUppRiBl(9)

Input #fnum, xUppRiBb(9), xUppRiBh(9), xUppRiBl(9)
Input #fnum, xUppLeBb(9), xUppLeBh(9), xUppLeBl(9)

Input #fnum, bTng(9), hTng(9), FloorHgt(9)


Input #fnum, Mx2(9), My2(9)
Input #fnum, Mx1(9), My1(9)
Input #fnum, DesgAxial(9), CoverTng(9)

Input #fnum, xLowLeBb(9), xLowLeBh(9), xLowLeBl(9)
Input #fnum, xLowRiBb(9), xLowRiBh(9), xLowRiBl(9)

Input #fnum, yLowRiBb(9), yLowRiBh(9), yLowRiBl(9)
Input #fnum, yLowLeBb(9), yLowLeBh(9), yLowLeBl(9)

Input #fnum, LowCoB(9), LowCoH(9), LowCoHgt(9)

Input #fnum, BarYLefN(9), BarYLefD(9)
Input #fnum, BarYRigN(9), BarYRigD(9)

Input #fnum, BarXTopN(9), BarXTopD(9)
Input #fnum, BarXBotN(9), BarXBotD(9)

Input #fnum, xBracedCol(9)
Input #fnum, yBracedCol(9)
Input #fnum, xBASEresistmnt(9)
Input #fnum, yBASEresistmnt(9)
Input #fnum, GridTng(9)

Close #fnum


End Sub



''''FOR TIANG''''
Private Sub DrawFWTiang()
   
Set moSpace = acadDoc.ModelSpace
Dim locx, locy, locXx, locYy As Double
Dim lbrBox, tgiBox As Double


locx = Xinsertion
locy = Yinsertion
locXx = locx
locYy = locy

Dim pt(0 To 3) As Double
Dim PolyPt As Object
Dim ArcPt As Object
Dim GridObj As Object
Dim center(0 To 2) As Double
Dim angStat, angEnd, Radius As Double
Dim locXsec, locYsec As Double
Dim delX, delY, delXt, delYr As Double
Dim barLayerYY As Integer

For i = 1 To bilJenisTng
lbrBox = 6 * bTng(i)
tgiBox = 700
locXx = locXx + lbrBox

Call DrawPRect(locXx - lbrBox / 2, locy + tgiBox / 2, lbrBox, tgiBox, "Slab")
Call VerLShape(locXx - 5.25 * bTng(i), locYy + 90, bTng(i), 700 - 90 + 1000 + _
               40 * BarXBotD(i), BarXBotD(i), "RebarSupt") '''5.25>>4.25
Call DropShape(locXx - 6 * bTng(i) + 50, locYy + 90 + 75, 50, 0, 0, 50, _
               BarXBotD(i), "RebarLink")
Call DropShape(locXx - 6 * bTng(i) + 50, locYy + 700 - 75, 50, 0, 0, 50, _
               BarXBotD(i), "RebarLink")
Call ArrowVertical(locXx - 6 * bTng(i) + 100, locYy + 90 + 75, 460, 100, _
                   20, "BeamDimension")
Call ArrowVertical(locXx - 6 * bTng(i) + 100, locYy + 700 - 75, 150, 1, _
                   1, "BeamDimension") '''''connection
Call DropShape(locXx - 6 * bTng(i) + 50, locYy + 700 + 75, 50, 0, 0, 50, _
               BarXBotD(i), "RebarLink")
Call DropShape(locXx - 6 * bTng(i) + 50, locYy + 1700 - yLowRiBh(i) - 75, _
               50, 0, 0, 50, BarXBotD(i), "RebarLink")
Call ArrowVertical(locXx - 6 * bTng(i) + 100, locYy + 700 + 75, _
                   1000 - yLowRiBh(i) - 2 * 75, 100, 20, "BeamDimension")
BarMark = BarMark + 1
Call LabelRbar(locXx - 6 * bTng(i) + 225, locYy + 500, 6, 10, BarMark, 0, _
               1.57, 50, "R", "RebarLink")
'''''''''''''''''
                   
Call VerColShape(locXx - 4.25 * bTng(i) + 2 * BarXBotD(i), locYy + 1700 + 75, _
                 40 * BarXBotD(i) - 75, 400, -BarXBotD(i), _
                 FloorHgt(i) - 400, _
                 BarXBotD(i), "RebarSupt")   ''+ 300>>4.25
Call DropShape(locXx - 6 * bTng(i) + 50, locYy + 1700 - yLowRiBh(i) + 125, _
               50, 0, 0, 50, BarXBotD(i), "RebarLink")
Call DropShape(locXx - 6 * bTng(i) + 50, locYy + 1700 - 125, 50, 0, 0, 50, _
               BarXBotD(i), "RebarLink")
Call ArrowVertical(locXx - 6 * bTng(i) + 100, locYy + 1700 - yLowRiBh(i) + 125, _
                   yLowRiBh(i) - 250, 75, 20, "BeamDimension")
Call ArrowVertical(locXx - 6 * bTng(i) + 100, locYy + 1700 - 125, 200, 1, _
                   1, "BeamDimension") '''''connection
Call DropShape(locXx - 6 * bTng(i) + 50, locYy + 1700 + 75, 50, 0, 0, 50, _
               BarXBotD(i), "RebarLink")
Call DropShape(locXx - 6 * bTng(i) + 50, locYy + 1700 + FloorHgt(i) - _
               yUppRiBh(i) - 75, 50, 0, 0, 50, BarXBotD(i), "RebarLink")
Call ArrowVertical(locXx - 6 * bTng(i) + 100, locYy + 1700 + 75, _
                   FloorHgt(i) - yUppRiBh(i) - 2 * 75, 100, 20, "BeamDimension")
BarMark = BarMark + 1
Call LabelRbar(locXx - 6 * bTng(i) + 175, locYy + 1700 + FloorHgt(i) / 3, _
               Int((FloorHgt(i) - yUppRiBh(i) - 2 * 75) / 200 + 1 + 2), 10, _
               BarMark, 200, 1.57, 50, "R", "RebarLink")
               
BarMark = BarMark + 1
Call LabelRbar(locXx - 5 * bTng(i) + bTng(i) / 2, locYy + 1700 + FloorHgt(i) _
               / 4, 2 * BarXBotN(i), BarXBotD(i), _
               BarMark, 0, 1.57, 50, "T", "RebarSupt")
               
BarMark = BarMark + 1
Call LabelRbar(locXx - 5 * bTng(i) + bTng(i) / 2, locYy + 1700 + FloorHgt(i) _
               / 4 + 13 * 50, 2 * BarYLefN(i), BarYLefD(i), _
               BarMark, 0, 1.57, 50, "T", "RebarSupt")



Call DrawPRect(locXx - 5 * bTng(i), locy + 1200, 2 * bTng(i), 1000, "Column")
Call DrawPRect(locXx - 2 * bTng(i), locy + 1200, 4 * bTng(i), 1000, "Column")
pt(0) = locXx - 6 * bTng(i)
pt(1) = locy + 1700 - yLowRiBh(i)
pt(2) = pt(0) + 3 * bTng(i)
pt(3) = pt(1)
Set PolyPt = moSpace.AddLightWeightPolyline(pt)
    PolyPt.Layer = "BeamSection"
    PolyPt.Update

Call DrawPRect(locXx - 5.5 * bTng(i), locy + 1700 + FloorHgt(i) / 2, _
               bTng(i), FloorHgt(i), "Column")
Call DrawPRect(locXx - 4.5 * bTng(i), locy + 1700 + FloorHgt(i) / 2, _
               bTng(i), FloorHgt(i), "Column")
Call DrawPRect(locXx - 2 * bTng(i), locy + 1700 + FloorHgt(i) / 2, _
               4 * bTng(i), FloorHgt(i), "Column")
pt(0) = locXx - 6 * bTng(i)
pt(1) = locy + 1700 + FloorHgt(i) - yUppRiBh(i)
pt(2) = pt(0) + 3 * bTng(i)
pt(3) = pt(1)
Set PolyPt = moSpace.AddLightWeightPolyline(pt)
    PolyPt.Layer = "BeamSection"
    PolyPt.Update
    
'''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''xsection + dimension''''''''''''''''''''''''''''''
locXsec = locXx - 2 * bTng(i)
locYsec = locYy + 1700 + FloorHgt(i) / 2

Call DrawPRect(locXsec, locYsec, bTng(i), hTng(i), "FormWork")
Call DrawPCircle(locXsec, locYsec + hTng(i) / 2 + bTng(i) / 2, _
                 160, "Grid")
center(0) = locXsec - 30
center(1) = locYsec + hTng(i) / 2 + bTng(i) / 2 - 40
center(2) = 0
Set GridObj = moSpace.AddText(Mid$(GridTng(i), 1, 1), center, 80)
    GridObj.Layer = "BeamDimension"
    GridObj.Update


Call DrawPCircle(locXsec - bTng(i) / 2 - bTng(i) / 2, locYsec, 160, "Grid")
center(0) = locXsec - bTng(i) / 2 - bTng(i) / 2 - 25
center(1) = locYsec - 30
center(2) = 0
Set GridObj = moSpace.AddText(Mid$(GridTng(i), 3, 1), center, 80)
    GridObj.Layer = "BeamDimension"
    GridObj.Update


 
'''''''''''''''''
pt(0) = locXsec + bTng(i)
pt(1) = locYsec - hTng(i) / 2 - CoverTng(i)
pt(2) = pt(0)
pt(3) = pt(1) + hTng(i) + 2 * CoverTng(i)
Set PolyPt = moSpace.AddLightWeightPolyline(pt)
    PolyPt.Layer = "Grid"
    PolyPt.Update
Call SlantLine(pt(0), pt(1) + CoverTng(i), 25, _
               2.3, "Grid")
Call SlantLine(pt(2), pt(3) - CoverTng(i), 25, _
               2.3, "Grid")
    
    
center(0) = locXsec + bTng(i) + 50 + 60
center(1) = locYsec - 2 * 60
center(2) = 0
Set GridObj = moSpace.AddText(Str(hTng(i)), center, _
                              60)
Call GridObj.Rotate(center, 1.57)

''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''
                 

''''''''''''''''''''''''''''''''''''''''
pt(0) = locXsec - bTng(i) / 2 - CoverTng(i)
pt(1) = locYsec - hTng(i) / 2 - bTng(i) / 2
pt(2) = pt(0) + bTng(i) + 2 * CoverTng(i)
pt(3) = pt(1)
Set PolyPt = moSpace.AddLightWeightPolyline(pt)
    PolyPt.Layer = "Grid"
    PolyPt.Update
Call SlantLine(pt(0) + CoverTng(i), pt(1), 25, _
               2.3, "Grid")
Call SlantLine(pt(2) - CoverTng(i), pt(3), 25, _
               2.3, "Grid")
    
    
center(0) = locXsec - 2 * 60
center(1) = pt(1) - 50 - 60
center(2) = 0
Set GridObj = moSpace.AddText(Str(bTng(i)), center, 60)

''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''loose link''''''''''''''''''''''''''''''
''locYy +1700 + FloorHgt(i) - yUppRiBh(i)/2
center(0) = locXsec - ((bTng(i) - 2 * CoverTng(i) - 4 * stirupD) / 2)
center(1) = locYy + 1700 + FloorHgt(i) - yUppRiBh(i) / 2 + ((hTng(i) - 2 * CoverTng(i) - 5 * stirupD) / 2)
center(2) = 0
Radius = 1.5 * stirupD
angStat = 1.57
angEnd = 3.14
Set ArcPt = moSpace.Addarc(center, Radius, angStat, angEnd)
    ArcPt.Layer = "RebarLink"
    ArcPt.Update
pt(0) = locXsec - ((bTng(i) - 2 * CoverTng(i) - 4 * stirupD) / 2)
pt(1) = locYy + 1700 + FloorHgt(i) - yUppRiBh(i) / 2 + ((hTng(i) - 2 * CoverTng(i) - 2 * stirupD) / 2)
pt(2) = locXsec + ((bTng(i) - 2 * CoverTng(i) - 4 * stirupD) / 2)
pt(3) = pt(1)
Set PolyPt = moSpace.AddLightWeightPolyline(pt)
    PolyPt.Layer = "RebarLink"
    PolyPt.Update
    
center(0) = locXsec + ((bTng(i) - 2 * CoverTng(i) - 4 * stirupD) / 2)
center(1) = locYy + 1700 + FloorHgt(i) - yUppRiBh(i) / 2 + ((hTng(i) - 2 * CoverTng(i) - 5 * stirupD) / 2)
angStat = 0
angEnd = 1.57
Set ArcPt = moSpace.Addarc(center, Radius, angStat, angEnd)
    ArcPt.Layer = "RebarLink"
    ArcPt.Update
pt(0) = 1.5 * stirupD + locXsec + ((bTng(i) - 2 * CoverTng(i) - 1 * stirupD) / 2)
pt(1) = locYy + 1700 + FloorHgt(i) - yUppRiBh(i) / 2 + ((hTng(i) - 2 * CoverTng(i) - 5 * stirupD) / 2)
pt(2) = pt(0)
pt(3) = locYy + 1700 + FloorHgt(i) - yUppRiBh(i) / 2 - ((hTng(i) - 2 * CoverTng(i) - 5 * stirupD) / 2)
Set PolyPt = moSpace.AddLightWeightPolyline(pt)
    PolyPt.Layer = "RebarLink"
    PolyPt.Update
  
pt(0) = locXsec + ((bTng(i) - 2 * CoverTng(i) - 1 * stirupD) / 2)
pt(1) = locYy + 1700 + FloorHgt(i) - yUppRiBh(i) / 2 + ((hTng(i) - 2 * CoverTng(i) - 5 * stirupD) / 2)
pt(2) = pt(0)
pt(3) = locYy + 1700 + FloorHgt(i) - yUppRiBh(i) / 2 - ((hTng(i) - 2 * CoverTng(i) - 5 * stirupD) / 2)
Set PolyPt = moSpace.AddLightWeightPolyline(pt)
    PolyPt.Layer = "RebarLink"
    PolyPt.Update
  
  
  
center(0) = 1.5 * stirupD + locXsec + ((bTng(i) - 2 * CoverTng(i) - 4 * stirupD) / 2)
center(1) = locYy + 1700 + FloorHgt(i) - yUppRiBh(i) / 2 - ((hTng(i) - 2 * CoverTng(i) - 5 * stirupD) / 2)
angStat = 4.71
angEnd = 6.28
Set ArcPt = moSpace.Addarc(center, Radius, angStat, angEnd)
    ArcPt.Layer = "RebarLink"
    ArcPt.Update
pt(0) = 1.5 * stirupD + locXsec + ((bTng(i) - 2 * CoverTng(i) - 4 * stirupD) / 2)
pt(1) = locYy + 1700 + FloorHgt(i) - yUppRiBh(i) / 2 - ((hTng(i) - 2 * CoverTng(i) - 2 * stirupD) / 2)
pt(2) = 1.5 * stirupD + locXsec - ((bTng(i) - 2 * CoverTng(i) - 4 * stirupD) / 2)
pt(3) = pt(1)
Set PolyPt = moSpace.AddLightWeightPolyline(pt)
    PolyPt.Layer = "RebarLink"
    PolyPt.Update
    
center(0) = 1.5 * stirupD + locXsec - ((bTng(i) - 2 * CoverTng(i) - 4 * stirupD) / 2)
center(1) = locYy + 1700 + FloorHgt(i) - yUppRiBh(i) / 2 - ((hTng(i) - 2 * CoverTng(i) - 5 * stirupD) / 2)
angStat = 3.14
angEnd = 4.71
Set ArcPt = moSpace.Addarc(center, Radius, angStat, angEnd)
    ArcPt.Layer = "RebarLink"
    ArcPt.Update
pt(0) = 1.5 * stirupD + locXsec - ((bTng(i) - 2 * CoverTng(i) - 1 * stirupD) / 2)
pt(1) = locYy + 1700 + FloorHgt(i) - yUppRiBh(i) / 2 - ((hTng(i) - 2 * CoverTng(i) - 5 * stirupD) / 2)
pt(2) = pt(0)
pt(3) = locYy + 1700 + FloorHgt(i) - yUppRiBh(i) / 2 + ((hTng(i) - 2 * CoverTng(i) - 5 * stirupD) / 2)
Set PolyPt = moSpace.AddLightWeightPolyline(pt)
    PolyPt.Layer = "RebarLink"
    PolyPt.Update
    
pt(0) = locXsec - ((bTng(i) - 2 * CoverTng(i) - 1 * stirupD) / 2)
pt(1) = locYy + 1700 + FloorHgt(i) - yUppRiBh(i) / 2 - ((hTng(i) - 2 * CoverTng(i) - 5 * stirupD) / 2)
pt(2) = pt(0)
pt(3) = locYy + 1700 + FloorHgt(i) - yUppRiBh(i) / 2 + ((hTng(i) - 2 * CoverTng(i) - 5 * stirupD) / 2)
Set PolyPt = moSpace.AddLightWeightPolyline(pt)
    PolyPt.Layer = "RebarLink"
    PolyPt.Update
    
'''''''''''''''''end of loose link'''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''
    
 
 
 
If yLowRiBh(i) <> 0 Or yLowRiBb(i) <> 0 Or yLowRiBl(i) <> 0 Then
center(0) = locXx - 3.25 * bTng(i) + 75
center(1) = locYy + 1700 - yLowRiBh(i) / 2 - 2.5 * 60
center(2) = 0
Set GridObj = moSpace.AddText("Rasuk", center, 60)
Call GridObj.Rotate(center, 1.57)
     GridObj.Layer = "BeamDimension"
     GridObj.Update
Call VerColShape(locXx - 3.25 * bTng(i), locYy + 1700 - yLowRiBh(i), _
                 yLowRiBh(i), 0, 0, 0, BarXBotD(i), "BeamSection")
Call SlantLine(locXx - 3.25 * bTng(i), locYy + 1700 - yLowRiBh(i), 25, _
               2.3, "BeamSection")
Call SlantLine(locXx - 3.25 * bTng(i), locYy + 1700, 25, _
               2.3, "BeamSection")
End If


If yUppRiBh(i) <> 0 Or yUppRiBb(i) <> 0 Or yUppRiBl(i) <> 0 Then
center(0) = locXx - 3.25 * bTng(i) + 75
center(1) = locYy + 1700 + FloorHgt(i) - yUppRiBh(i) / 2 - 2.5 * 60
center(2) = 0
Set GridObj = moSpace.AddText("Rasuk", center, 60)
Call GridObj.Rotate(center, 1.57)
     GridObj.Layer = "BeamDimension"
     GridObj.Update

Call VerColShape(locXx - 3.25 * bTng(i), locYy + 1700 + FloorHgt(i) - _
                 yUppRiBh(i), yUppRiBh(i), 0, 0, 0, BarXBotD(i), "BeamSection")
Call SlantLine(locXx - 3.25 * bTng(i), locYy + 1700 + FloorHgt(i) - _
               yUppRiBh(i), 25, 2.3, "BeamSection")
Call SlantLine(locXx - 3.25 * bTng(i), locYy + 1700 + FloorHgt(i), _
               25, 2.3, "BeamSection")
End If


center(0) = locXx - 4 * bTng(i) + 200
center(1) = locYy + 700 + 50
center(2) = 0
Set GridObj = moSpace.AddText(" Aras Asas ", center, 60)
    GridObj.Layer = "BeamDimension"
    GridObj.Update
Call ArrowTail(locXx - 4 * bTng(i) + 100, locYy + 700, _
               0, 100, 100, 0, "BeamSection")

center(0) = locXx - 4 * bTng(i) + 150
center(1) = locYy + 1700 + 50
center(2) = 0
Set GridObj = moSpace.AddText(" Aras Bawah", center, 60)
    GridObj.Layer = "BeamDimension"
    GridObj.Update
Call ArrowTail(locXx - 4 * bTng(i) + 100, locYy + 1700, _
               0, 100, 100, 0, "BeamSection")
center(1) = locYy + 1700 + 150
Set GridObj = moSpace.AddText("0.00 mm", center, 60)
    GridObj.Layer = "BeamDimension"
    GridObj.Update
'''
 
center(0) = locXx - 4 * bTng(i) + 150
center(1) = locYy + 1700 + FloorHgt(i) + 50
center(2) = 0
Set GridObj = moSpace.AddText(" Aras Satu", center, 60)
    GridObj.Layer = "BeamDimension"
    GridObj.Update
Call ArrowTail(locXx - 4 * bTng(i) + 100, locYy + 1700 + FloorHgt(i), _
               0, 100, 100, 0, "BeamSection")
center(1) = locYy + 1700 + FloorHgt(i) + 150
Set GridObj = moSpace.AddText(Str(FloorHgt(i)) & " mm", center, 60)
    GridObj.Layer = "BeamDimension"
    GridObj.Update
               
'''''''''''''''''''''''''''''''''''''''''''




If (BarXBotN(i) - 1) < 1 Then
   delX = (bTng(i) - 2 * CoverTng(i) - 4 * stirupD - BarXBotD(i)) / 2
   Else
   delX = (bTng(i) - 2 * CoverTng(i) - 4 * stirupD - BarXBotD(i)) / (BarXBotN(i) - 1)
End If
delY = (hTng(i) - 2 * CoverTng(i) - 4 * stirupD - BarXBotD(i) / 2 - BarXTopD(i) / 2) / (BarYLefN(i) + 1)
If (BarXTopN(i) - 1) < 1 Then
   delXt = (bTng(i) - 2 * CoverTng(i) - 4 * stirupD - BarXTopD(i)) / 2
   Else
   delXt = (bTng(i) - 2 * CoverTng(i) - 4 * stirupD - BarXTopD(i)) / (BarXTopN(i) - 1)
End If
delYr = (hTng(i) - 2 * CoverTng(i) - 4 * stirupD - BarXBotD(i) / 2 - BarXTopD(i) / 2) / (BarYRigN(i) + 1)
'''''''''''''''''''''''''''''''''''''''''''''''''




''''''''''''''''''link proper'''''''''''''''''''''''
center(0) = locXsec - ((bTng(i) - 2 * CoverTng(i) - 4 * stirupD) / 2)
center(1) = locYsec + ((hTng(i) - 2 * CoverTng(i) - 4 * stirupD) / 2)
center(2) = 0
Radius = stirupD
angStat = 1.57
angEnd = 3.14
Set ArcPt = moSpace.Addarc(center, Radius, angStat, angEnd)
    ArcPt.Layer = "RebarLink"
    ArcPt.Update
pt(0) = locXsec - ((bTng(i) - 2 * CoverTng(i) - 4 * stirupD) / 2)
pt(1) = locYsec + ((hTng(i) - 2 * CoverTng(i) - 2 * stirupD) / 2)
pt(2) = locXsec + ((bTng(i) - 2 * CoverTng(i) - 4 * stirupD) / 2)
pt(3) = pt(1)
Set PolyPt = moSpace.AddLightWeightPolyline(pt)
    PolyPt.Layer = "RebarLink"
    PolyPt.Update
    
center(0) = locXsec + ((bTng(i) - 2 * CoverTng(i) - 4 * stirupD) / 2)
center(1) = locYsec + ((hTng(i) - 2 * CoverTng(i) - 4 * stirupD) / 2)
angStat = 0
angEnd = 1.57
Set ArcPt = moSpace.Addarc(center, Radius, angStat, angEnd)
    ArcPt.Layer = "RebarLink"
    ArcPt.Update
pt(0) = locXsec + ((bTng(i) - 2 * CoverTng(i) - 2 * stirupD) / 2)
pt(1) = locYsec + ((hTng(i) - 2 * CoverTng(i) - 4 * stirupD) / 2)
pt(2) = pt(0)
pt(3) = locYsec - ((hTng(i) - 2 * CoverTng(i) - 4 * stirupD) / 2)
Set PolyPt = moSpace.AddLightWeightPolyline(pt)
    PolyPt.Layer = "RebarLink"
    PolyPt.Update
  
center(0) = locXsec + ((bTng(i) - 2 * CoverTng(i) - 4 * stirupD) / 2)
center(1) = locYsec - ((hTng(i) - 2 * CoverTng(i) - 4 * stirupD) / 2)
angStat = 4.71
angEnd = 6.28
Set ArcPt = moSpace.Addarc(center, Radius, angStat, angEnd)
    ArcPt.Layer = "RebarLink"
    ArcPt.Update
pt(0) = locXsec + ((bTng(i) - 2 * CoverTng(i) - 4 * stirupD) / 2)
pt(1) = locYsec - ((hTng(i) - 2 * CoverTng(i) - 2 * stirupD) / 2)
pt(2) = locXsec - ((bTng(i) - 2 * CoverTng(i) - 4 * stirupD) / 2)
pt(3) = pt(1)
Set PolyPt = moSpace.AddLightWeightPolyline(pt)
    PolyPt.Layer = "RebarLink"
    PolyPt.Update
    
center(0) = locXsec - ((bTng(i) - 2 * CoverTng(i) - 4 * stirupD) / 2)
center(1) = locYsec - ((hTng(i) - 2 * CoverTng(i) - 4 * stirupD) / 2)
angStat = 3.14
angEnd = 4.71
Set ArcPt = moSpace.Addarc(center, Radius, angStat, angEnd)
    ArcPt.Layer = "RebarLink"
    ArcPt.Update
pt(0) = locXsec - ((bTng(i) - 2 * CoverTng(i) - 2 * stirupD) / 2)
pt(1) = locYsec - ((hTng(i) - 2 * CoverTng(i) - 4 * stirupD) / 2)
pt(2) = pt(0)
pt(3) = locYsec + ((hTng(i) - 2 * CoverTng(i) - 4 * stirupD) / 2)
Set PolyPt = moSpace.AddLightWeightPolyline(pt)
    PolyPt.Layer = "RebarLink"
    PolyPt.Update
    
Dim locXxsec, locYysec As Double
Dim N As Integer
locXxsec = locXsec - ((bTng(i) - 2 * CoverTng(i) - 4 * stirupD) / 2) + BarXBotD(i) / 2
locYysec = locYsec - ((hTng(i) - 2 * CoverTng(i) - 4 * stirupD) / 2) + BarXBotD(i) / 2
     For N = 1 To BarXBotN(i)
      If BarXBotN(i) = 1 Then locXxsec = locXxsec + ((bTng(i) - 2 * CoverTng(i) - 4 * stirupD) / 2)
      Call DrawPCircle(locXxsec, locYysec, BarXBotD(i), "RebarSupt")
      locXxsec = locXxsec + delX
     Next N
locXxsec = locXsec - ((bTng(i) - 2 * CoverTng(i) - 4 * stirupD) / 2) + BarXTopD(i) / 2
locYysec = locYsec + ((hTng(i) - 2 * CoverTng(i) - 4 * stirupD) / 2) - BarXTopD(i) / 2
     For N = 1 To BarXTopN(i)
      If BarXTopN(i) = 1 Then locXxsec = locXxsec + ((bTng(i) - 2 * CoverTng(i) - 4 * stirupD) / 2)
      Call DrawPCircle(locXxsec, locYysec, BarXTopD(i), "RebarSupt")
      locXxsec = locXxsec + delXt
     Next N
locXxsec = locXsec - ((bTng(i) - 2 * CoverTng(i) - 4 * stirupD) / 2) + BarYLefD(i) / 2
locYysec = locYsec - ((hTng(i) - 2 * CoverTng(i) - 4 * stirupD) / 2) + BarXBotD(i) / 2 + delY
     For N = 1 To BarYLefN(i)
      Call DrawPCircle(locXxsec, locYysec, BarYLefD(i), "RebarSupt")
      locYysec = locYysec + delY
     Next N
locXxsec = locXsec + ((bTng(i) - 2 * CoverTng(i) - 4 * stirupD) / 2) - BarYRigD(i) / 2
locYysec = locYsec - ((hTng(i) - 2 * CoverTng(i) - 4 * stirupD) / 2) + BarXBotD(i) / 2 + delYr
     For N = 1 To BarYRigN(i)
      Call DrawPCircle(locXxsec, locYysec, BarYRigD(i), "RebarSupt")
      locYysec = locYysec + delYr
     Next N
''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''
Call DrawPRect(locXx - 5.5 * bTng(i), locy + 3500 + FloorHgt(i), _
               bTng(i), 700, "FormWork")
center(0) = locXx - 5.5 * bTng(i) + 50
center(1) = locy + 3500 + FloorHgt(i) - 300
center(2) = 0
Set GridObj = moSpace.AddText("PENGIKAT", center, 80)
Call GridObj.Rotate(center, 1.57)
     GridObj.Layer = "BeamDimension"
     GridObj.Update
Call DrawPRect(locXx - 4.5 * bTng(i), locy + 3500 + FloorHgt(i), _
               bTng(i), 700, "FormWork")
center(0) = locXx - 4.5 * bTng(i) + 50
center(1) = locy + 3500 + FloorHgt(i) - 300
center(2) = 0
Set GridObj = moSpace.AddText("TETULANG", center, 80)
Call GridObj.Rotate(center, 1.57)
     GridObj.Layer = "BeamDimension"
     GridObj.Update

Call DrawPRect(locXx - 2 * bTng(i), locy + 3500 + FloorHgt(i), _
               4 * bTng(i), 700, "FormWork")
center(0) = locXx - 2.9 * bTng(i)
center(1) = locy + 3500 + FloorHgt(i) ''- 300
center(2) = 0
Set GridObj = moSpace.AddText("KERATAN", center, 80)
    GridObj.Layer = "BeamDimension"
    GridObj.Update
'Call GridObj.Rotate(center, 1.57)


Call DrawPRect(locXx - 3 * bTng(i), locy + 4500 + FloorHgt(i), _
               6 * bTng(i), 800, "Grid")
center(0) = locXx - 5.5 * bTng(i)
center(1) = locy + 4250 + FloorHgt(i)
center(2) = 0
Set GridObj = moSpace.AddText(GridTng(i), center, 110)
GridObj.Layer = "BeamDimension"
GridObj.Color = 50
GridObj.Update
            
Next i
Form1.Picture = LoadPicture(NamaFolder & "icon\tiang1.ico")

End Sub







''''GENERAL''''
Public Sub SetLayer()
'''''Form1.Picture = LoadPicture(NamaFolder & "icon\layer.ico")
  
Dim layerObj As Object


Set layerObj = acadDoc.Layers.Add("FormWork")
layerObj.Color = 120

Set layerObj = acadDoc.Layers.Add("Column")
layerObj.Color = 254

Set layerObj = acadDoc.Layers.Add("Slab")
layerObj.Color = 120

Set layerObj = acadDoc.Layers.Add("Grid")
layerObj.Color = 254   '' text label color 0

Set layerObj = acadDoc.Layers.Add("RebarSupt")
layerObj.Color = 20

Set layerObj = acadDoc.Layers.Add("LabelRebarSupt")
layerObj.Color = 255

Set layerObj = acadDoc.Layers.Add("RebarMTop")
layerObj.Color = 20

Set layerObj = acadDoc.Layers.Add("LabelRebarMTop")
layerObj.Color = 255

Set layerObj = acadDoc.Layers.Add("RebarMBot")
layerObj.Color = 20

Set layerObj = acadDoc.Layers.Add("LabelRebarMBot")
layerObj.Color = 255

Set layerObj = acadDoc.Layers.Add("RebarLink")
layerObj.Color = 90  '' asal 50

Set layerObj = acadDoc.Layers.Add("LabelRebarLink")
layerObj.Color = 255

Set layerObj = acadDoc.Layers.Add("Curtailment")
layerObj.Color = 255

Set layerObj = acadDoc.Layers.Add("BeamDimension")
layerObj.Color = 255

Set layerObj = acadDoc.Layers.Add("BeamSection")
layerObj.Color = 120   '' for formwork; 51 link; 30 rebar

Set layerObj = acadDoc.Layers.Add("BeamName")
layerObj.Color = 90

Set layerObj = acadDoc.Layers.Add("Structural_Strength")
layerObj.Color = 254

End Sub





''''GENERAL''''
Private Function ArrowHorizontal(ByVal locx As Double, ByVal _
locy As Double, ByVal Mlength As Double, ByVal ArrowWidth _
As Double, ByVal ArrowHeight As Double, _
LayerName As String) As Object

Dim pt(0 To 19) As Double
pt(0) = locx + ArrowWidth
pt(1) = locy
pt(2) = pt(0)
pt(3) = pt(1) - ArrowHeight / 2
pt(4) = pt(2) - ArrowWidth
pt(5) = pt(3) + ArrowHeight / 2
pt(6) = pt(4) + ArrowWidth
pt(7) = pt(5) + ArrowHeight / 2
pt(8) = pt(6)
pt(9) = pt(7) - ArrowHeight / 2
pt(10) = pt(8) + Mlength - 2 * ArrowWidth
pt(11) = pt(9)
pt(12) = pt(10)
pt(13) = pt(11) + ArrowHeight / 2
pt(14) = pt(12) + ArrowWidth
pt(15) = pt(13) - ArrowHeight / 2
pt(16) = pt(14) - ArrowWidth
pt(17) = pt(15) - ArrowHeight / 2
pt(18) = pt(16)
pt(19) = pt(17) + ArrowHeight / 2

Dim PolyPt As Object
Set PolyPt = moSpace.AddLightWeightPolyline(pt)
Set PolyPt = moSpace.Item(moSpace.Count - 1)
PolyPt.Color = 254
PolyPt.Layer = LayerName
PolyPt.Update

End Function

''''GENERAL''''
Private Function ArrowVertical(ByVal locx As Double, ByVal _
locy As Double, ByVal Mlength As Double, ByVal ArrowWidth _
As Double, ByVal ArrowHeight As Double, _
LayerName As String) As Object

Dim pt(0 To 19) As Double
pt(0) = locx
pt(1) = locy + ArrowWidth
pt(2) = pt(0) + ArrowHeight / 2
pt(3) = pt(1)
pt(4) = pt(2) - ArrowHeight / 2
pt(5) = pt(3) - ArrowWidth
pt(6) = pt(4) - ArrowHeight / 2
pt(7) = pt(5) + ArrowWidth
pt(8) = pt(6) + ArrowHeight / 2
pt(9) = pt(7)
pt(10) = pt(8)
pt(11) = pt(9) + Mlength - 2 * ArrowWidth
pt(12) = pt(10) - ArrowHeight / 2
pt(13) = pt(11)
pt(14) = pt(12) + ArrowHeight / 2
pt(15) = pt(13) + ArrowWidth
pt(16) = pt(14) + ArrowHeight / 2
pt(17) = pt(15) - ArrowWidth
pt(18) = pt(16) - ArrowHeight / 2
pt(19) = pt(17)

Dim PolyPt As Object
Set PolyPt = moSpace.AddLightWeightPolyline(pt)
Set PolyPt = moSpace.Item(moSpace.Count - 1)
PolyPt.Color = 254
PolyPt.Layer = LayerName
PolyPt.Update

End Function

''''GENERAL''''
Private Function ArrowTail(ByVal locx As Double, ByVal _
locy As Double, ByVal Mlength As Double, ByVal ArrowWidth _
As Double, ByVal ArrowHeight As Double, ByVal Mwidth As Double, _
LayerName As String) As Object

Dim pt(0 To 13) As Double
pt(0) = locx
pt(1) = locy + ArrowWidth
pt(2) = pt(0) + ArrowHeight / 2
pt(3) = pt(1)
pt(4) = pt(2) - ArrowHeight / 2
pt(5) = pt(3) - ArrowWidth
pt(6) = pt(4) - ArrowHeight / 2
pt(7) = pt(5) + ArrowWidth
pt(8) = pt(6) + ArrowHeight / 2
pt(9) = pt(7)
pt(10) = pt(8)
pt(11) = pt(9) + Mlength
pt(12) = pt(10) + Mwidth
pt(13) = pt(11)

Dim PolyPt As Object
Set PolyPt = moSpace.AddLightWeightPolyline(pt)
Set PolyPt = moSpace.Item(moSpace.Count - 1)
PolyPt.Color = 254
PolyPt.Layer = LayerName
PolyPt.Update

locx = pt(12) + ArrowWidth
locy = pt(13) - ArrowWidth / 2

End Function

''''GENERAL''''
Private Function LabelRbar(ByVal locx As Double, ByVal _
locy As Double, ByVal rbarNo As Integer, ByVal rbarDia _
As Integer, ByVal rbarMark As Integer, ByVal RbarSpacing _
As Double, ByVal Rotate As Double, ByVal txtHeight As Integer, _
barType As String, laPisan As String) As Object
 
Dim acadObj As Object
Dim corner1(0 To 2) As Double
Dim text As String

corner1(0) = locx
corner1(1) = locy
corner1(2) = 0
If txtHeight < 20 Then
txtHeight = 20
End If
If RbarSpacing = 0 Then
  text = Trim(Str(rbarNo)) & barType & Trim(Str(rbarDia)) & " - " & _
       Trim(Str(rbarMark))
Else
  text = Trim(Str(rbarNo)) & barType & Trim(Str(rbarDia)) & " - " & _
       Trim(Str(rbarMark)) & "-" & Trim(Str(RbarSpacing))
End If

Set acadObj = moSpace.AddText(text, corner1, txtHeight)
Call acadObj.Rotate(corner1, Rotate)
acadObj.Layer = laPisan
acadObj.Color = 0
acadObj.Update

End Function

''''GENERAL''''
Private Function LabelDimension(ByVal locx As Double, ByVal _
locy As Double, ByVal curtailLength As Double) As Object
Dim acadObj As Object
Dim corner1(0 To 2) As Double
Dim theight As Double
Dim text As String

corner1(0) = locx
corner1(1) = locy
corner1(2) = 0
theight = 40
text = Str(curtailLength)
Set acadObj = moSpace.AddText(text, corner1, theight)

End Function


''''FOR RASUK & COLUMN''''
Private Function DrawLinkSection(ByVal locx As Double, ByVal _
locy As Double, ByVal LinkDia As Double, _
ByVal b As Double, ByVal h As Double, ByVal SlabT As Double, _
ByVal Bar1No As Integer, ByVal Bar1Dia As Double, ByVal Bar1BM _
As Integer, ByVal Bar2No As Integer, ByVal Bar2Dia As Double, _
ByVal Bar2BM As Integer, ByVal Bar3No As Integer, ByVal Bar3Dia _
As Double, ByVal Bar3BM As Integer, ByVal Bar4No As Integer, _
ByVal Bar4Dia As Double, ByVal Bar4BM As Integer, ByVal Bar5No _
As Integer, ByVal Bar5Dia As Double, ByVal Bar5BM As Integer, _
ByVal Bar6No As Integer, ByVal Bar6Dia As Double, ByVal Bar6BM _
As Integer) As Object


Dim pt(0 To 53) As Double
pt(0) = locx + SlabT + cVr + 3 * LinkDia ''2.5
pt(1) = locy - cVr - 0.5 * LinkDia
pt(2) = pt(0) + b - 2 * cVr - 6 * LinkDia ''5.5
pt(3) = pt(1)
pt(4) = pt(2) + 3 / 6 * 2.5 * LinkDia
pt(5) = pt(3) - 1 / 6 * 2.5 * LinkDia
pt(6) = pt(4) + 2 / 6 * 2.5 * LinkDia
pt(7) = pt(5) - 2 / 6 * 2.5 * LinkDia
pt(8) = pt(6) + 1 / 6 * 2.5 * LinkDia
pt(9) = pt(7) - 3 / 6 * 2.5 * LinkDia
pt(10) = pt(8)
pt(11) = pt(9) - h + 2 * cVr + 6 * LinkDia
pt(12) = pt(10) - 1 / 6 * 2.5 * LinkDia
pt(13) = pt(11) - 3 / 6 * 2.5 * LinkDia
pt(14) = pt(12) - 2 / 6 * 2.5 * LinkDia
pt(15) = pt(13) - 2 / 6 * 2.5 * LinkDia
pt(16) = pt(14) - 3 / 6 * 2.5 * LinkDia
pt(17) = pt(15) - 1 / 6 * 2.5 * LinkDia
pt(18) = pt(16) - b + 2 * cVr + 6 * LinkDia
pt(19) = pt(17)
pt(20) = pt(18) - 3 / 6 * 2.5 * LinkDia
pt(21) = pt(19) + 1 / 6 * 2.5 * LinkDia
pt(22) = pt(20) - 2 / 6 * 2.5 * LinkDia
pt(23) = pt(21) + 2 / 6 * 2.5 * LinkDia
pt(24) = pt(22) - 1 / 6 * 2.5 * LinkDia
pt(25) = pt(23) + 3 / 6 * 2.5 * LinkDia
pt(26) = pt(24)
pt(27) = pt(25) + h - 2 * cVr - 6 * LinkDia
pt(28) = pt(26) + 1 / 6 * 2.5 * LinkDia
pt(29) = pt(27) + 3 / 6 * 2.5 * LinkDia
pt(30) = pt(28) + 2 / 6 * 2.5 * LinkDia
pt(31) = pt(29) + 2 / 6 * 2.5 * LinkDia
pt(32) = pt(30) + 3 / 6 * 2.5 * LinkDia
pt(33) = pt(31) + 1 / 6 * 2.5 * LinkDia

pt(34) = pt(0) + b - 2 * cVr - 6 * LinkDia
pt(35) = pt(1)
pt(36) = pt(2) + 3 / 6 * 2.5 * LinkDia
pt(37) = pt(3) - 1 / 6 * 2.5 * LinkDia
pt(38) = pt(4) + 2 / 6 * 2.5 * LinkDia
pt(39) = pt(5) - 2 / 6 * 2.5 * LinkDia
pt(40) = pt(6) + 1 / 6 * 2.5 * LinkDia
pt(41) = pt(7) - 3 / 6 * 2.5 * LinkDia
pt(42) = pt(8) - 0.5 * LinkDia
pt(43) = pt(9) - 6 * LinkDia

pt(44) = pt(6) + 1 / 6 * 2.5 * LinkDia
pt(45) = pt(7) - 3 / 6 * 2.5 * LinkDia
pt(46) = pt(4) + 2 / 6 * 2.5 * LinkDia
pt(47) = pt(5) - 2 / 6 * 2.5 * LinkDia
pt(48) = pt(2) + 3 / 6 * 2.5 * LinkDia
pt(49) = pt(3) - 1 / 6 * 2.5 * LinkDia
pt(50) = pt(0) + b - 2 * cVr - 6 * LinkDia
pt(51) = pt(1)
pt(52) = pt(2) - 6 * LinkDia
pt(53) = pt(1) - 0.5 * LinkDia



Dim PolyPt As Object
Set PolyPt = moSpace.AddLightWeightPolyline(pt)
PolyPt.Color = 90   '' asal 51
PolyPt.Layer = "BeamSect"
PolyPt.Update

Call CircleBarMark(locx, locy, LinkDia, b, h, SlabT, _
Bar1No, Bar1Dia, Bar1BM, _
Bar2No, Bar2Dia, Bar2BM, _
Bar3No, Bar3Dia, Bar3BM, _
Bar4No, Bar4Dia, Bar4BM, _
Bar5No, Bar5Dia, Bar5BM, _
Bar6No, Bar6Dia, Bar6BM)

End Function
''''FOR RASUK''''
Private Function CircleBarMark(ByVal locx As Double, _
ByVal locy As Double, ByVal LinkDia As Double, _
ByVal b As Double, ByVal h As Double, ByVal SlabT As Double, _
ByVal Bar1No As Integer, ByVal Bar1Dia As Double, ByVal Bar1BM _
As Integer, ByVal Bar2No As Integer, ByVal Bar2Dia As Double, _
ByVal Bar2BM As Integer, ByVal Bar3No As Integer, ByVal Bar3Dia _
As Double, ByVal Bar3BM As Integer, ByVal Bar4No As Integer, _
ByVal Bar4Dia As Double, ByVal Bar4BM As Integer, ByVal Bar5No _
As Integer, ByVal Bar5Dia As Double, ByVal Bar5BM As Integer, _
ByVal Bar6No As Integer, ByVal Bar6Dia As Double, ByVal Bar6BM _
As Integer) As Object
  
Dim circleObj As Object
Dim center(0 To 2) As Double
Dim Radius As Double
Dim barmarkObj As Object
Dim insPnt(0 To 2) As Double
Dim textHgt As Double
Dim textStr As String
Dim Rotate As Double
Dim j As Integer
Dim deltaX, deltaY As Double

If Bar1No <> 0 Then
  center(0) = locx + SlabT + cVr + 3 * LinkDia
  center(1) = locy - cVr - LinkDia - Bar3Dia - Bar1Dia / 2
  deltaX = (b - 2 * cVr - 6 * LinkDia) / (Bar1No - 1)
    
  For j = 1 To Bar1No
  center(0) = center(0)
  center(1) = center(1)
  center(2) = 0
  Radius = Bar1Dia / 2
  Set circleObj = moSpace.AddCircle(center, Radius)
  circleObj.Layer = "BeamSect"
  circleObj.Color = 30
  circleObj.Update
  
  insPnt(0) = center(0)
  insPnt(1) = locy + 4 * cVr
  insPnt(2) = 0
  textHgt = 30
  textStr = Trim(Str(Bar1BM))
  Set barmarkObj = moSpace.AddText(textStr, insPnt, textHgt)
  Rotate = 1.57  '' 90 degrees
  Call barmarkObj.Rotate(insPnt, Rotate)
  center(0) = center(0) + deltaX
  barmarkObj.Layer = "BeamSect"
  barmarkObj.Color = 255
  barmarkObj.Update
  Next
End If

If Bar2No <> 0 Then
  center(0) = locx + SlabT + cVr + 3 * LinkDia + Bar1Dia / 2 + _
              Bar2Dia / 2
  center(1) = locy - cVr - LinkDia - Bar3Dia - 2 * Bar1Dia - _
              Bar2Dia / 2
  deltaX = (b - 2 * cVr - 6 * LinkDia - Bar1Dia - Bar2Dia) / _
           (Bar2No - 1)
    
  For j = 1 To Bar2No
  center(0) = center(0)
  center(1) = center(1)
  center(2) = 0
  Radius = Bar2Dia / 2
  Set circleObj = moSpace.AddCircle(center, Radius)
  circleObj.Layer = "BeamSect"
  circleObj.Color = 30
  circleObj.Update
  
  insPnt(0) = center(0)
  insPnt(1) = locy + 7 * cVr
  insPnt(2) = 0
  textHgt = 30
  textStr = Trim(Str(Bar2BM))
  Set barmarkObj = moSpace.AddText(textStr, insPnt, textHgt)
  Rotate = 1.57  '' 90 degrees
  Call barmarkObj.Rotate(insPnt, Rotate)
  center(0) = center(0) + deltaX
  barmarkObj.Layer = "BeamSect"
  barmarkObj.Color = 255
  barmarkObj.Update
  Next
End If

If Bar3No <> 0 Then
  center(0) = locx + SlabT + cVr + 3 * LinkDia
  center(1) = locy - cVr - LinkDia - Bar3Dia / 2
  deltaX = (b - 2 * cVr - 6 * LinkDia) / (Bar3No - 1)
 
  For j = 1 To Bar3No
  center(0) = center(0)
  center(1) = center(1)
  center(2) = 0
  Radius = Bar3Dia / 2
  Set circleObj = moSpace.AddCircle(center, Radius)
  circleObj.Layer = "BeamSect"
  circleObj.Color = 30
  circleObj.Update
    
  insPnt(0) = center(0)
  insPnt(1) = locy + cVr
  insPnt(2) = 0
  textHgt = 30
  textStr = Trim(Str(Bar3BM))
  Set barmarkObj = moSpace.AddText(textStr, insPnt, textHgt)
  Rotate = 1.57  '' 90 degrees
  Call barmarkObj.Rotate(insPnt, Rotate)
  center(0) = center(0) + deltaX
  barmarkObj.Layer = "BeamSect"
  barmarkObj.Color = 255
  barmarkObj.Update
  Next
End If

If Bar4No <> 0 Then
  center(0) = locx + SlabT + cVr + 3 * LinkDia
  center(1) = locy - h + cVr + LinkDia + Bar4Dia / 2
  deltaX = (b - 2 * cVr - 6 * LinkDia) / (Bar4No - 1)
 
  For j = 1 To Bar4No
  center(0) = center(0)
  center(1) = center(1)
  center(2) = 0
  Radius = Bar4Dia / 2
  Set circleObj = moSpace.AddCircle(center, Radius)
  circleObj.Layer = "BeamSect"
  circleObj.Color = 30
  circleObj.Update
    
  insPnt(0) = center(0)
  insPnt(1) = locy - h - 3 * cVr
  insPnt(2) = 0
  textHgt = 30
  textStr = Trim(Str(Bar4BM))
  Set barmarkObj = moSpace.AddText(textStr, insPnt, textHgt)
  Rotate = 1.57  '' 90 degrees
  Call barmarkObj.Rotate(insPnt, Rotate)
  center(0) = center(0) + deltaX
  barmarkObj.Layer = "BeamSect"
  barmarkObj.Color = 255
  barmarkObj.Update
  Next
End If

If Bar5No <> 0 Then
  center(0) = locx + SlabT + cVr + 3 * LinkDia
  center(1) = locy - h + cVr + LinkDia + Bar4Dia + Bar6Dia _
              + Bar5Dia / 2
  deltaX = (b - 2 * cVr - 6 * LinkDia) / (Bar5No - 1)
 
  For j = 1 To Bar5No
  center(0) = center(0)
  center(1) = center(1)
  center(2) = 0
  Radius = Bar5Dia / 2
  Set circleObj = moSpace.AddCircle(center, Radius)
  circleObj.Layer = "BeamSect"
  circleObj.Color = 245
  circleObj.Update
    
  insPnt(0) = center(0)
  insPnt(1) = locy - h - 9 * cVr
  insPnt(2) = 0
  textHgt = 30
  textStr = Trim(Str(Bar5BM))
  Set barmarkObj = moSpace.AddText(textStr, insPnt, textHgt)
  Rotate = 1.57  '' 90 degrees
  Call barmarkObj.Rotate(insPnt, Rotate)
  center(0) = center(0) + deltaX
  barmarkObj.Layer = "BeamSect"
  barmarkObj.Color = 255
  barmarkObj.Update
  Next
End If

If Bar6No <> 0 Then
  center(0) = locx + SlabT + cVr + 3 * LinkDia
  center(1) = locy - h + cVr + LinkDia + Bar4Dia + Bar6Dia / 2
  deltaX = (b - 2 * cVr - 6 * LinkDia) / (Bar6No - 1)
 
  For j = 1 To Bar6No
  center(0) = center(0)
  center(1) = center(1)
  center(2) = 0
  Radius = Bar6Dia / 2
  Set circleObj = moSpace.AddCircle(center, Radius)
  circleObj.Layer = "BeamSect"
  circleObj.Color = 30
  circleObj.Update
    
  insPnt(0) = center(0)
  insPnt(1) = locy - h - 6 * cVr
  insPnt(2) = 0
  textHgt = 30
  textStr = Trim(Str(Bar6BM))
  Set barmarkObj = moSpace.AddText(textStr, insPnt, textHgt)
  Rotate = 1.57  '' 90 degrees
  Call barmarkObj.Rotate(insPnt, Rotate)
  center(0) = center(0) + deltaX
  barmarkObj.Layer = "BeamSect"
  barmarkObj.Color = 255
  barmarkObj.Update
  Next
End If

End Function

'''''''DRAW FUNCTIONS'''''''

Private Function DrawPRect(ByVal locx As Double, ByVal locy As Double, _
ByVal WthX As Double, ByVal WthY As Double, laPisan As String) As Object

Dim pt(0 To 9) As Double
Dim PolyPt As Object

pt(0) = locx - WthX / 2
pt(1) = locy - WthY / 2
pt(2) = pt(0)
pt(3) = pt(1) + WthY
pt(4) = pt(2) + WthX
pt(5) = pt(3)
pt(6) = pt(4)
pt(7) = pt(5) - WthY
pt(8) = pt(0)
pt(9) = pt(1)

Set PolyPt = moSpace.AddLightWeightPolyline(pt)
   PolyPt.Layer = laPisan
   PolyPt.Update
End Function
Private Function DrawPCircle(ByVal locx As Double, ByVal locy As Double, _
ByVal Diameter As Double, laPisan As String) As Object

Dim center(0 To 2) As Double
Dim Radius As Double
Dim circleObj As Object

  center(0) = locx
  center(1) = locy
  center(2) = 0
  Radius = Diameter / 2
  Set circleObj = moSpace.AddCircle(center, Radius)
  circleObj.Layer = laPisan
  circleObj.Update
End Function

'''''''REBAR SHAPE FUNCTIONS'''''''
Private Function StraitXShape(ByVal LocBarX As Double, ByVal LocBarY As Double, _
ByVal Length1 As Double, ByVal BarDia As Double, laPisan As String) _
As Object

Dim pt(0 To 3) As Double
Dim PolyPt As Object

pt(0) = LocBarX
pt(1) = LocBarY
pt(2) = pt(0) + Length1
pt(3) = pt(1)
Set PolyPt = moSpace.AddLightWeightPolyline(pt)
    PolyPt.Layer = laPisan
    PolyPt.Update
End Function
Private Function StraitYShape(ByVal LocBarX As Double, ByVal LocBarY As Double, _
ByVal Length1 As Double, ByVal BarDia As Double, laPisan As String) _
As Object

Dim pt(0 To 3) As Double
Dim PolyPt As Object

pt(0) = LocBarX
pt(1) = LocBarY
pt(2) = pt(0)
pt(3) = pt(1) + Length1
Set PolyPt = moSpace.AddLightWeightPolyline(pt)
    PolyPt.Layer = laPisan
    PolyPt.Update
End Function


Private Function LeftUShape(ByVal LocBarX As Double, ByVal LocBarY As Double, _
ByVal Length1 As Double, ByVal Length2 As Double, ByVal Length3 As Double, _
ByVal BarDia As Double, laPisan As String) As Object

Dim pt(0 To 19) As Double
Dim PolyPt As Object

pt(0) = LocBarX
pt(1) = LocBarY
pt(2) = pt(0) - Length1 + 3 * BarDia
pt(3) = pt(1)
pt(4) = pt(2) - 1.5 * BarDia
pt(5) = pt(3) - 0.5 * BarDia
pt(6) = pt(4) - 1 * BarDia
pt(7) = pt(5) - 1 * BarDia
pt(8) = pt(6) - 0.5 * BarDia
pt(9) = pt(7) - 1.5 * BarDia
pt(10) = pt(8)
pt(11) = pt(9) - Length2 + 6 * BarDia
pt(12) = pt(10) + 0.5 * BarDia
pt(13) = pt(11) - 1.5 * BarDia
pt(14) = pt(12) + 1 * BarDia
pt(15) = pt(13) - 1 * BarDia
pt(16) = pt(14) + 1.5 * BarDia
pt(17) = pt(15) - 0.5 * BarDia
pt(18) = pt(16) + Length3 - 3 * BarDia
pt(19) = pt(17)
Set PolyPt = moSpace.AddLightWeightPolyline(pt)
    PolyPt.Layer = laPisan
    PolyPt.Update
End Function
Private Function RightUShape(ByVal LocBarX As Double, ByVal LocBarY As Double, _
ByVal Length1 As Double, ByVal Length2 As Double, ByVal Length3 As Double, _
ByVal BarDia As Double, laPisan As String) As Object

Dim pt(0 To 19) As Double
Dim PolyPt As Object

pt(0) = LocBarX
pt(1) = LocBarY
pt(2) = pt(0) + Length1 + 3 * BarDia
pt(3) = pt(1)
pt(4) = pt(2) + 1.5 * BarDia
pt(5) = pt(3) - 0.5 * BarDia
pt(6) = pt(4) + 1 * BarDia
pt(7) = pt(5) - 1 * BarDia
pt(8) = pt(6) + 0.5 * BarDia
pt(9) = pt(7) - 1.5 * BarDia
pt(10) = pt(8)
pt(11) = pt(9) - Length2 + 6 * BarDia
pt(12) = pt(10) - 0.5 * BarDia
pt(13) = pt(11) - 1.5 * BarDia
pt(14) = pt(12) - 1 * BarDia
pt(15) = pt(13) - 1 * BarDia
pt(16) = pt(14) - 1.5 * BarDia
pt(17) = pt(15) - 0.5 * BarDia
pt(18) = pt(16) - Length3 - 3 * BarDia
pt(19) = pt(17)
Set PolyPt = moSpace.AddLightWeightPolyline(pt)
    PolyPt.Layer = laPisan
    PolyPt.Update
End Function

Private Function LeftHorLShape(ByVal LocBarX As Double, ByVal LocBarY As Double, _
ByVal Length1 As Double, ByVal Length2 As Double, ByVal BarDia As Double, _
laPisan As String) As Object

Dim pt(0 To 11) As Double
Dim PolyPt As Object

pt(0) = LocBarX
pt(1) = LocBarY
pt(2) = pt(0) - Length1 + 3 * BarDia
pt(3) = pt(1)
pt(4) = pt(2) - 1.5 * BarDia
pt(5) = pt(3) - 0.5 * BarDia
pt(6) = pt(4) - 1 * BarDia
pt(7) = pt(5) - 1 * BarDia
pt(8) = pt(6) - 0.5 * BarDia
pt(9) = pt(7) - 1.5 * BarDia
pt(10) = pt(8)
pt(11) = pt(9) - Length2
Set PolyPt = moSpace.AddLightWeightPolyline(pt)
    PolyPt.Layer = laPisan
    PolyPt.Update
End Function

Private Function RightHorLShape(ByVal LocBarX As Double, ByVal LocBarY As Double, _
ByVal Length1 As Double, ByVal Length2 As Double, ByVal BarDia As Double, _
laPisan As String) As Object

Dim pt(0 To 11) As Double
Dim PolyPt As Object

pt(0) = LocBarX
pt(1) = LocBarY
pt(2) = pt(0) - Length1 + 3 * BarDia
pt(3) = pt(1)
pt(4) = pt(2) + 1.5 * BarDia
pt(5) = pt(3) - 0.5 * BarDia
pt(6) = pt(4) + 1 * BarDia
pt(7) = pt(5) - 1 * BarDia
pt(8) = pt(6) + 0.5 * BarDia
pt(9) = pt(7) - 1.5 * BarDia
pt(10) = pt(8)
pt(11) = pt(9) - Length2
Set PolyPt = moSpace.AddLightWeightPolyline(pt)
    PolyPt.Layer = laPisan
    PolyPt.Update
End Function
Private Function DropShape(ByVal LocBarX As Double, ByVal LocBarY As Double, _
ByVal Length1 As Double, ByVal Length2 As Double, ByVal DropUp As Double, _
ByVal Length3 As Double, ByVal BarDia As Double, laPisan As String) _
As Object

Dim pt(0 To 7) As Double
Dim PolyPt As Object

pt(0) = LocBarX
pt(1) = LocBarY
pt(2) = pt(0) + Length1
pt(3) = pt(1)
pt(4) = pt(2) + Length2
pt(5) = pt(3) + DropUp
pt(6) = pt(4) + Length3
pt(7) = pt(5)
Set PolyPt = moSpace.AddLightWeightPolyline(pt)
    PolyPt.Layer = laPisan
    PolyPt.Update
End Function



Private Function CageUpShape(ByVal LocBarX As Double, ByVal LocBarY As Double, _
ByVal Length1 As Double, ByVal Length2 As Double, ByVal BarDia As Double, _
laPisan As String) As Object

Dim pt(0 To 19) As Double
Dim PolyPt As Object

pt(0) = LocBarX
pt(1) = LocBarY
pt(2) = pt(0)
pt(3) = pt(1) - Length1 + 3 * BarDia
pt(4) = pt(2) + 0.5 * BarDia
pt(5) = pt(3) - 1.5 * BarDia
pt(6) = pt(4) + 1 * BarDia
pt(7) = pt(5) - 1 * BarDia
pt(8) = pt(6) + 1.5 * BarDia
pt(9) = pt(7) - 0.5 * BarDia
pt(10) = pt(8) + Length2 - 6 * BarDia
pt(11) = pt(9)
pt(12) = pt(10) + 1.5 * BarDia
pt(13) = pt(11) + 0.5 * BarDia
pt(14) = pt(12) + 1 * BarDia
pt(15) = pt(13) + 1 * BarDia
pt(16) = pt(14) + 0.5 * BarDia
pt(17) = pt(15) + 1.5 * BarDia
pt(18) = pt(16)
pt(19) = pt(17) + Length1 - 3 * BarDia
Set PolyPt = moSpace.AddLightWeightPolyline(pt)
    PolyPt.Layer = laPisan
    PolyPt.Update
End Function
Private Function CageDownShape(ByVal LocBarX As Double, ByVal LocBarY As Double, _
ByVal Length1 As Double, ByVal Length2 As Double, ByVal BarDia As Double, _
laPisan As String) As Object

Dim pt(0 To 17) As Double
Dim PolyPt As Object

pt(0) = LocBarX
pt(1) = LocBarY
pt(2) = pt(0) + Length1 - 3 * BarDia
pt(3) = pt(1)
pt(4) = pt(2) + 1.5 * BarDia
pt(5) = pt(3) + 0.5 * BarDia
pt(6) = pt(4) + 1 * BarDia
pt(7) = pt(5) + 1 * BarDia
pt(8) = pt(6) + 0.5 * BarDia
pt(9) = pt(7) + 1.5 * BarDia
pt(10) = pt(8) + Length2 - 6 * BarDia
pt(11) = pt(9) + 1.5 * BarDia
pt(12) = pt(10) - 0.5 * BarDia
pt(13) = pt(11) + 1 * BarDia
pt(14) = pt(12) - 1 * BarDia
pt(15) = pt(13) + 0.5 * BarDia
pt(16) = pt(14) - 1.5 * BarDia
pt(17) = pt(15) - Length1 + 3 * BarDia
Set PolyPt = moSpace.AddLightWeightPolyline(pt)
    PolyPt.Layer = laPisan
    PolyPt.Update
End Function
Private Function VerLShape(ByVal LocBarX As Double, ByVal LocBarY As Double, _
ByVal Length1 As Double, ByVal Length2 As Double, ByVal BarDia As Double, _
laPisan As String) As Object

Dim pt(0 To 11) As Double
Dim PolyPt As Object

pt(0) = LocBarX
pt(1) = LocBarY
pt(2) = pt(0) + Length1 - 3 * BarDia
pt(3) = pt(1)
pt(4) = pt(2) + 1.5 * BarDia
pt(5) = pt(3) + 0.5 * BarDia
pt(6) = pt(4) + 1 * BarDia
pt(7) = pt(5) + 1 * BarDia
pt(8) = pt(6) + 0.5 * BarDia
pt(9) = pt(7) + 1.5 * BarDia
pt(10) = pt(8)
pt(11) = pt(9) + Length2 - 3 * BarDia
Set PolyPt = moSpace.AddLightWeightPolyline(pt)
    PolyPt.Layer = laPisan
    PolyPt.Update
End Function
Private Function VerColShape(ByVal LocBarX As Double, ByVal LocBarY As Double, _
ByVal Length1 As Double, ByVal Length2 As Double, ByVal SwayLR As Double, _
ByVal Length3 As Double, ByVal BarDia As Double, laPisan As String) As Object

Dim pt(0 To 7) As Double
Dim PolyPt As Object

pt(0) = LocBarX
pt(1) = LocBarY
pt(2) = pt(0)
pt(3) = pt(1) + Length1
pt(4) = pt(2) + SwayLR
pt(5) = pt(3) + Length2
pt(6) = pt(4)
pt(7) = pt(5) + Length3
Set PolyPt = moSpace.AddLightWeightPolyline(pt)
    PolyPt.Layer = laPisan
    PolyPt.Update
End Function
Private Function SlantLine(ByVal LocBarX As Double, ByVal LocBarY As Double, _
ByVal Length As Double, ByVal SlantAng As Double, laPisan As String) _
As Object

Dim pt(0 To 3) As Double
Dim PolyPt As Object

pt(0) = LocBarX - Length * Cos(SlantAng)
pt(1) = LocBarY + Length * Sin(SlantAng)
pt(2) = LocBarX + Length * Cos(SlantAng)
pt(3) = LocBarY - Length * Sin(SlantAng)

Set PolyPt = moSpace.AddLightWeightPolyline(pt)
    PolyPt.Layer = laPisan
    PolyPt.Update
End Function


''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''
''''''FUNCTIONS FOR STRUCTURAL CAPACITY'''''
''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''
Private Function CalcMoment( _
ByVal fc As Double, _
ByVal fy As Double, _
ByVal bf As Double, _
ByVal bw As Double, _
ByVal hf As Double, _
ByVal d As Double, _
ByVal dp As Double, _
ByVal t As Double, _
ByVal c As Double, _
ByVal pu As Double, _
ByVal ap As Double, _
ByVal dt As Double, _
ByVal k As Double) As Integer


Dim g, u As Integer
Dim k8, k9, er, e1, e2, e3, e4, e5, e6 As Double
Dim w, w9, r, r9, f4, f5, f6, x, A, mu As Double
103: If hf > d Or hf = 0 Then hf = d
104: If hf = 0 Then bw = bf
105: k8 = 0.45
110: k9 = 0.9
115: er = (1 - k / 1000) * 0.75 * pu / 200000#
120: If t = 0 Then d = 0
130: e1 = fy / 1.15 / 200000#
135: e2 = 4 * pu / 1000000# / 1.15
140: e3 = 0.005 + 5 * pu / 1000000# / 1.15
145: u = 100: w = 1: r = 0: g = 1
150: If pu = 0 Or ap = 0 Then dt = d
155: x = dp + 1
170: x = x - dt / u
175: x = x + dt / u
180: w9 = w: r9 = r
185: f4 = 0: f6 = 0
190: If t <> 0 Then GoSub 340
195: If ap <> 0 Then GoSub 370
200: If c <> 0 Then GoSub 400
205: A = bf * k9 * x
210: If k9 * x > hf Then A = bf * hf + bw * (k9 * x - hf)
215: r = k8 * fc * A + c * f6
220: w = t * f4 + ap * f5
225: If w - r > 0 And w9 - r9 > 0 Then GoTo 175
230: If g = 2 Then GoTo 245
235: x = x - dt / u: u = 2000: g = 2
240: w = 1: r = 0: GoTo 170
245: If k9 * x > hf Then GoTo 260
250: mu = k8 * fc * bf * k9 * x * (dt - k9 * 0.5 * x) - _
          t * f4 * (dt - d) + c * f6 * (dt - dp)
255: GoTo 265
260: mu = k8 * fc * bf * hf * (dt - hf / 2) + k8 * fc * _
          bw * (k9 * x - hf) * (dt - k9 * x / 2 - hf / 2) - _
          t * f4 * (dt - d) + c * f6 * (dt - dp)
265: 'MsgBox "Mult=", , mu
266: If d = 0 Or t = 0 Then d = dt
270: 'MsgBox "x/d=", , x / d
275: CalcMoment = Int(mu / 1000000)
280: GoTo 430
'''''''''''''''''
340: Rem stress in tension steel
345: e4 = 0.0035 * (d - x) / x
350: If Abs(e4) >= e1 Then f4 = fy / 1.15 * Sgn(e4)
355: If Abs(e4) < e1 Then f4 = e4 * 200000# * Sgn(e4)
360: Return
''''''''''''''''
370: Rem stress in prestressed steel
375: e5 = 0.0035 * (dt - x) / x + er
380: If e5 >= e3 Then f5 = pu / 1.15
385: If e5 < e3 And e5 > e2 Then f5 = 0.2 * pu * (0.02 + e5) / _
     (0.00575 + pu / 1000000#)
390: If e5 <= e2 Then f5 = 200000# * e5
395: Return
''''''''''''''''''
400: Rem stress in compression steel
403: If x < dp Then x = dp
405: e6 = 0.0035 * (x - dp) / x
410: If Abs(e6) >= e1 Then f6 = fy / 1.15 * Sgn(e6)
415: If Abs(e6) < e1 Then f6 = e6 * 200000 * Sgn(e6)
420: Return
'''''''''''''''''''
'''''''''''''''''''
430 Rem
End Function
Private Function CalcShear( _
ByVal Link As Double, _
ByVal Cover As Double, _
ByVal sv As Double, _
ByVal asv As Double, _
ByVal asvt As Double, _
ByVal atr As Double, _
ByVal fcu As Double, _
ByVal fy As Double, _
ByVal fyv As Double, _
ByVal bw As Double, _
ByVal ht As Double, _
ByVal d As Double, _
ByVal dp As Double, _
ByVal Ast As Double, _
ByVal asc As Double) As Integer

750 Rem shear capacity and torsion moment
753 Dim x11, y11, j1, j2, j3, trc, vs, v1, v2, vc, vct As Double
754 If sv = 0 Or asv = 0 Then GoTo 805  ''' for nil link rebar ie pad
755 x11 = bw - 2 * Cover - Link
760 y11 = ht - 2 * Cover - Link
762 atr = atr + asc
765 j1 = atr * fy / (fyv * (x11 + y11))
770 j2 = asvt / sv
775 If Abs(j1 - j2) < 0.5 Then MsgBox " balanced torsion linkage"
780 j3 = j1
785 If j1 >= j2 Then j3 = j2
790 trc = j3 * 0.8 * x11 * y11 * 0.87 * fyv
795 'MsgBox "Mtor=", , trc
800 vs = asv * fyv * d / sv
805 v1 = 100 * Ast / bw / d
810 If v1 > 3 Then v1 = 3
815 v2 = 400 / d
820 If v2 < 1 Then v2 = 1
825 vc = 0.79 * v1 ^ (1 / 3) * v2 ^ (1 / 4) / 1.25
826 If fcu > 40 Then fcu = 40
828 vc = vc * (fcu / 25) ^ (1 / 3)
830 vct = vs + vc * bw * d
835 'MsgBox vc & " -- " & vs, , "vc" & "vs"
840 CalcShear = Int(vct / 1000)
End Function

Private Function ACurvatureOne( _
ByVal fcu As Double, _
ByVal shrink As Double, _
ByVal creep As Double, _
ByVal Mtt As Double, _
ByVal Mpm As Double, _
ByVal b As Double, _
ByVal h As Double, _
ByVal d As Double, _
ByVal dp As Double, _
ByVal Ast As Double, _
ByVal asc As Double, _
ByVal SLTerm As String) As Double

Dim ec, es, el, rd, rt, rc, rs, rl As Double
Dim i1, i2, i3, i4, i5, i6, i7 As Double
Dim x, x111, x222, x333, s As Double
Dim c1, c2, c3, c4, c5 As Double
Dim fct, xsl, isl, ae, mfct, mtn, mpn, fcc, fst, fsc As Double

1330 Rem deflection
1331 ec = 24 + (fcu - 20) / 5
1332 es = 200: rd = dp / d: rt = Ast / b / d
1334 rc = asc / b / d: el = ec / (1 + creep)
1336 rs = es / ec: rl = es / el
1338 i1 = -1 * rs * rt - (rs - 1) * rc
1340 i2 = rs * rt * (rs * rt + 2) + 2 * (rs - 1) * rc * (rs * rt + rd)
1342 i3 = i1 + Sqr(i2)
1344 i4 = i3 ^ 3 / 3 + rs * rt * (1 - i3) ^ 2 + (rs - 1) * rc * (i3 - rd) ^ 2
1346 i5 = i4 * b * d ^ 3
1348 x111 = -1 * rl * rt - (rl - 1) * rc
1350 x222 = rl * rt * (rl * rt + 2) + 2 * (rl - 1) * rc * (rl * rt + rd)
1352 x333 = x111 + Sqr(x222)
1354 x = x333 * d
1356 i6 = x333 ^ 3 / 3 + rl * rt * (1 - x333) ^ 2 + (rl - 1) * rc * (x333 - rd) ^ 2
1358 i7 = i6 * b * d ^ 3
1360 s = Ast * (d - x) - asc * (x - dp)
1361 GoTo 1470
1362 c1 = (mtn - mpn) * 1000# / i5 / ec
1364 c2 = mpn * 1000# / i7 / el
1366 c3 = shrink * rl * s / i7
1372 c4 = c1 + c2     ''c6  ''c8
1373 '''If sltd$ = "s" Then c4 = mtn * 1000# / i5 / ec
1374 c5 = c3          ''c7  ''c9
1375 ACurvatureOne = c4 '''If sltd$ = "s" Then c5 = 0
1376 GoTo 1510
''''''''''''''''''

1470 Rem calc mnet
1472 fct = 0.55: xsl = x: isl = i7: ae = rl
1474 '''If sltd$ = "s" Then fct = 1
1475 '''If sltd$ = "s" Then isl = i5
1476 '''If sltd$ = "s" Then xsl = i3 * d
1478 mfct = fct * b * (h - xsl) ^ 3 / (d - xsl) / 3 / 1000000#
1480 mtn = Mtt - mfct
1482 mpn = Mpm - mfct
1484 '''If sltd$ = "s" Then ae = rs
1486 fcc = mtn * 1000000# * xsl / isl
1488 fst = ae * mtn * 1000000# * (d - xsl) / isl
1490 fsc = ae * mtn * 1000000# * (xsl - dp) / isl
1500 GoTo 1362
1510 Rem
End Function

Private Function ACurvatureTwo( _
ByVal fcu As Double, _
ByVal shrink As Double, _
ByVal creep As Double, _
ByVal Mtt As Double, _
ByVal Mpm As Double, _
ByVal b As Double, _
ByVal h As Double, _
ByVal d As Double, _
ByVal dp As Double, _
ByVal Ast As Double, _
ByVal asc As Double, _
ByVal SLTerm As String) As Double

Dim ec, es, el, rd, rt, rc, rs, rl As Double
Dim i1, i2, i3, i4, i5, i6, i7 As Double
Dim x, x111, x222, x333, s As Double
Dim c1, c2, c3, c4, c5 As Double
Dim fct, xsl, isl, ae, mfct, mtn, mpn, fcc, fst, fsc As Double

1330 Rem deflection
1331 ec = 24 + (fcu - 20) / 5
1332 es = 200: rd = dp / d: rt = Ast / b / d
1334 rc = asc / b / d: el = ec / (1 + creep)
1336 rs = es / ec: rl = es / el
1338 i1 = -1 * rs * rt - (rs - 1) * rc
1340 i2 = rs * rt * (rs * rt + 2) + 2 * (rs - 1) * rc * (rs * rt + rd)
1342 i3 = i1 + Sqr(i2)
1344 i4 = i3 ^ 3 / 3 + rs * rt * (1 - i3) ^ 2 + (rs - 1) * rc * (i3 - rd) ^ 2
1346 i5 = i4 * b * d ^ 3
1348 x111 = -1 * rl * rt - (rl - 1) * rc
1350 x222 = rl * rt * (rl * rt + 2) + 2 * (rl - 1) * rc * (rl * rt + rd)
1352 x333 = x111 + Sqr(x222)
1354 x = x333 * d
1356 i6 = x333 ^ 3 / 3 + rl * rt * (1 - x333) ^ 2 + (rl - 1) * rc * (x333 - rd) ^ 2
1358 i7 = i6 * b * d ^ 3
1360 s = Ast * (d - x) - asc * (x - dp)
1361 GoTo 1470
1362 c1 = (mtn - mpn) * 1000# / i5 / ec
1364 c2 = mpn * 1000# / i7 / el
1366 c3 = shrink * rl * s / i7
1372 c4 = c1 + c2     ''c6  ''c8
1373 '''If sltd$ = "s" Then c4 = mtn * 1000# / i5 / ec
1374 c5 = c3          ''c7  ''c9
1375 ACurvatureTwo = c5 '''If sltd$ = "s" Then c5 = 0
1376 GoTo 1510
''''''''''''''''''

1470 Rem calc mnet
1472 fct = 0.55: xsl = x: isl = i7: ae = rl
1474 '''If sltd$ = "s" Then fct = 1
1475 '''If sltd$ = "s" Then isl = i5
1476 '''If sltd$ = "s" Then xsl = i3 * d
1478 mfct = fct * b * (h - xsl) ^ 3 / (d - xsl) / 3 / 1000000#
1480 mtn = Mtt - mfct
1482 mpn = Mpm - mfct
1484 '''If sltd$ = "s" Then ae = rs
1486 fcc = mtn * 1000000# * xsl / isl
1488 fst = ae * mtn * 1000000# * (d - xsl) / isl
1490 fsc = ae * mtn * 1000000# * (xsl - dp) / isl
1500 GoTo 1362
1510 Rem
End Function

Private Function ADeflect( _
ByVal l As Double, _
ByVal c4, c5, c6, c7, c8, c9 As Double) As Double

1390 Dim c As Integer
1392 Dim bt, dM, dn, dft As Double
1400 Rem total deflection
1402 '''MsgBox "shape parabolic="
1403 c = 1
1404 '''input #1, " 1=para  2=tria";c
1406 bt = (c6 + c8) / c4
1408 dM = l ^ 2 * c4
1409 dn = l ^ 2 * (c5 + c7 + c9) / 3 / 8
1410 '''On c GoTo 1412, 1414
1412 dft = (1 + bt / 10) * dM / 9.6 + dn  '''': GoTo 1418
1414 '''dft = (1 + bt / 4) * dm / 12 + dn: GoTo 1418
1418 '''Return
1419 ADeflect = Int(dft * 100) / 100
''''''''''
End Function
Private Function Column(ByVal locx As Double, ByVal locy As Double, _
ByVal s As Double, ByVal b As Double, ByVal h As Double, _
ByVal dp As Double, ByVal bl As Double, ByVal xBarD As Double, _
ByVal yBarD As Double, ByVal xBarLN As Double, ByVal yBarLN As Double, _
ByVal nc As Double, ByVal m1 As Double, ByVal m2 As Double, _
BendAxis As String, laPisan As String) As Object

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim ff, fm, la, x, xp, xh, N, m, ns, ms, e, f, ef, el, at As Double
Dim j, k, JK, q As Integer
Dim A(1 To 20) As Double
Dim Kfactor, kf, ma, mt, bp, mi, em As Double
Dim mType, bType, pAxis As String
'''**********************'''
Dim CountX, CountY As Integer
CountX = 1
CountY = 1
''>>>>>>>>>>>>>>>>
340 Rem eff height
Dim leh(0 To 6), beh(0 To 6), heh(0 To 6), keh(0 To 6) As Double
Dim kc1, kc2, kb1, kb2, a1, a2, am, L1, L2, Le, Lo As Double
Dim kc1T, kc2T, kb1T, kb2T As Double
If BendAxis = "XX" Then
  leh(0) = FloorHgt(i): beh(0) = bTng(i): heh(0) = hTng(i)
  leh(1) = LowCoHgt(i): beh(1) = LowCoB(i): heh(1) = LowCoH(i)
  leh(2) = UppCoHgt(i): beh(2) = UppCoB(i): heh(2) = UppCoH(i)
  leh(3) = xLowLeBl(i): beh(3) = xLowLeBb(i): heh(3) = xLowLeBh(i)
  leh(4) = xLowRiBl(i): beh(4) = xLowRiBb(i): heh(4) = xLowRiBh(i)
  leh(5) = xUppLeBl(i): beh(5) = xUppLeBb(i): heh(5) = xUppLeBh(i)
  leh(6) = xUppRiBl(i): beh(6) = xUppRiBb(i): heh(6) = xUppRiBh(i)
Else
  leh(0) = FloorHgt(i): beh(0) = hTng(i): heh(0) = bTng(i)
  leh(1) = LowCoHgt(i): beh(1) = LowCoH(i): heh(1) = LowCoB(i)
  leh(2) = UppCoHgt(i): beh(2) = UppCoH(i): heh(2) = UppCoB(i)
  leh(3) = yLowLeBl(i): beh(3) = yLowLeBb(i): heh(3) = yLowLeBh(i)
  leh(4) = yLowRiBl(i): beh(4) = yLowRiBb(i): heh(4) = yLowRiBh(i)
  leh(5) = yUppLeBl(i): beh(5) = yUppLeBb(i): heh(5) = yUppLeBh(i)
  leh(6) = yUppRiBl(i): beh(6) = yUppRiBb(i): heh(6) = yUppRiBh(i)
End If

492 ''Rem subr supt conditions
494 For q = 0 To 6
496 keh(q) = 0
498 If leh(q) = 0 Then GoTo 502
500 keh(q) = beh(q) * heh(q) ^ 3 / (12 * leh(q))
502 Next q
504 kc1 = keh(0) + keh(1): kb1 = keh(3) + keh(4)
506 kc2 = keh(0) + keh(2): kb2 = keh(5) + keh(6)
      kc1T = kc1: kb1T = kb1
      kc2T = kc2: kb2T = kb2
508 If kb1 = 0 Then kb1T = 1
510 If kb1 = 0 Then kc1T = 10
512 If kb2 = 0 Then kb2T = 1
514 If kb2 = 0 Then kc2T = 10
516 a1 = kc1T / kb1T
    If kc1T / kb1T > 10 Then a1 = 10
518 a2 = kc2T / kb2T
    If kc2T / kb2T > 10 Then a2 = 10
520 am = a1
522 If a1 > a2 Then am = a2

''If BendAxis = "XX" Then leh(0) = leh(0) - (xUppLeBh(i) + xUppRiBh(i)) / 2 Else _
    leh(0) = leh(0) - (yUppLeBh(i) + yUppRiBh(i)) / 2
    
If BendAxis = "XX" Then
    Lo = leh(0) - xUppLeBh(i)
      If Lo > leh(0) - xUppRiBh(i) Then
          Lo = leh(0) - xUppRiBh(i)
      End If
End If
    
If BendAxis = "YY" Then
    Lo = leh(0) - yUppLeBh(i)
      If Lo > leh(0) - yUppRiBh(i) Then
          Lo = leh(0) - yUppRiBh(i)
      End If
End If
    
    
''526 If leh(1) = 0 Then GoTo 530 Else GoTo 540
    '''check base resist mnt or not.
530 If Left(xBASEresistmnt(i), 5) = "FIXED" Or _
       Left(yBASEresistmnt(i), 5) = "FIXED" Or _
       Left(xBASEresistmnt(i), 5) = "fixed" Or _
       Left(yBASEresistmnt(i), 5) = "fixed" Or _
       Left(xBASEresistmnt(i), 1) = "Y" Or _
       Left(yBASEresistmnt(i), 1) = "Y" Or _
       Left(xBASEresistmnt(i), 1) = "y" Or _
       Left(yBASEresistmnt(i), 1) = "y" And a1 > 1 Then
       a1 = 1
       Else
       a1 = a1
    End If
    
    If a1 = 1 And a1 < a2 Then
       am = a1
       Else
       am = a2
    End If
    
540 If keh(5) = 0 And keh(6) = 0 Then GoTo 542 Else GoTo 380
542 a2 = 7
    '''check braced or unbraced column.
380 If Left(xBracedCol(i), 1) = "B" Or _
       Left(yBracedCol(i), 1) = "B" Or _
       Left(xBracedCol(i), 1) = "b" Or _
       Left(yBracedCol(i), 1) = "b" Or _
       Left(xBracedCol(i), 1) = "Y" Or _
       Left(yBracedCol(i), 1) = "Y" Or _
       Left(xBracedCol(i), 1) = "y" Or _
       Left(yBracedCol(i), 1) = "y" _
       Then GoTo 382 Else GoTo 388
       
382 L1 = Lo * (0.7 + 0.05 * (a1 + a2))
384 L2 = Lo * (0.85 + 0.05 * am)
386 GoTo 392
388 L1 = Lo * (1 + 0.15 * (a1 + a2))
390 L2 = Lo * (2 + 0.3 * am)
392 Le = Int(L1)
394 If L1 > L2 Then Le = Int(L2)

If BendAxis = "XX" Then
   EffHeightXX(i) = Le
     End If
If BendAxis = "YY" Then
   EffHeightYY(i) = Le
     End If
     
''<<<<<<<<<<<<<<<<
If bl < 2 Then bl = 2
If BendAxis = "XX" Then
   For k = 1 To bl
   If BarYLefN(i) <> 0 Then
   A(k) = 3.14 * yBarD ^ 2 / 4 + 3.14 * BarYRigD(i) ^ 2 / 4 * BarYRigN(i) / _
          BarYLefN(i)
   End If
   If k = 1 Then A(k) = 3.14 * xBarD ^ 2 / 4 * xBarLN
   If k = bl Then A(k) = 3.14 * BarXTopD(i) ^ 2 / 4 * BarXTopN(i)
   Next k
Else
   For k = 1 To bl
   If BarXBotN(i) <> 0 Then
      A(k) = 3.14 * xBarD ^ 2 / 4 + 3.14 * BarXTopD(i) ^ 2 / 4 * BarXTopN(i) / _
             BarXBotN(i)
   End If
   If k = 1 Then A(k) = 3.14 * xBarD ^ 2 / 4 + _
          3.14 * yBarD ^ 2 / 4 * yBarLN + 3.14 * BarXTopD(i) ^ 2 / 4
      If k = 1 And BarXBotN(i) = 1 Then A(k) = 3.14 * yBarD ^ 2 / 4 * yBarLN + _
             3.14 * BarXTopD(i) ^ 2 / 4
          If k = 1 And BarXBotN(i) = 1 And BarXTopN(i) = 1 Then _
                             A(k) = 3.14 * yBarD ^ 2 / 4 * yBarLN
             
   If k = bl Then A(k) = 3.14 * xBarD ^ 2 / 4 + _
          3.14 * BarYRigD(i) ^ 2 / 4 * BarYRigN(i) + 3.14 * BarXTopD(i) ^ 2 / 4
      If k = bl And BarXBotN(i) = 1 Then A(k) = 3.14 * BarYRigD(i) ^ 2 / 4 * _
             BarYRigN(i) + 3.14 * BarXTopD(i) ^ 2 / 4
          If k = bl And BarXBotN(i) = 1 And BarXTopN(i) = 1 Then _
                    A(k) = 3.14 * BarYRigD(i) ^ 2 / 4 * BarYRigN(i)
             
   Next k
End If

588 Rem subr centroidal axis
590 ''s = (h - 2 * dp) / (bl - 1)
592 ff = 0.45 * fcu * b * h
594 fm = ff * h / 2
    at = 0
596 For k = 1 To bl
598 ff = ff + 0.87 * fy * A(k)
600 la = dp + (k - 1) * s
602 fm = fm + 0.87 * fy * A(k) * la
    at = at + A(k)
604 Next k
606 xp = fm / ff
607 Ast = at
608 dPrime = dp
'''''''''''''
610 Rem subr n&m of concrete

Dim CntAxial As Integer
CntAxial = 20


   ''List1.ForeColor = vbBlue
   List1.FontSize = 8
      
'''*****************************************************************'''
611 x = 1#: JK = 1: locy = locy + 200     ''>>>>>>>>>>>>>v
612 For j = 1 To Int(hTng(i) * 3)
613 N = 0.45 * fcu * b * h: xh = 0.5 * h
614 If h < 0.9 * x Then GoTo 618      '''''''''''CHECK!!!!!!!!!!
616 N = 0.405 * fcu * b * x: xh = 0.45 * x
618 m = N * (xp - xh)
620 ms = m: ns = N
622 ''Return
''''''''''

624 Rem subr n&m of steel
625 at = 0
626 For k = 1 To bl
627 at = at + A(k)
628 la = dp + s * (k - 1)
630 e = 0.0035 * (x - la) / x
632 f = e * 200000#
634 If Abs(f) > 0.87 * fy Then f = 0.87 * fy * Sgn(e)
636 'if e<=-0.002 then f=-.87*fy
638 'if e<0.002 then f=2e5*e
640 'if e>0.002 then f=-2e5*e
642 N = N + f * A(k)
644 m = m + f * (xp - la) * A(k)
646 If k = 1 Then ef = e
648 If k = bl Then el = e
650 Next k

Kfactor = 0.45 * fcu * (b * h - at) + 0.87 * fy * at
Kfactor = (Kfactor - N) / (Kfactor - 0.25 * fcu * b * (h - dp))
Kfactor = Int(Kfactor * 100) / 100
756 If h < b Then bp = h Else bp = b
758 mType = "":  kf = Kfactor
759 If Kfactor > 1 Then kf = 1
760 ma = nc * 1000 * h * (Le / bp) ^ 2 * kf / 2000
762 mt = m2 * 1000000
764 mi = (0.4 * m1 + 0.6 * m2) * 1000000
766 If mt < ma + mi Then mt = ma + mi
768 mi = 0.4 * m2 * 1000000
770 If mt < ma + mi Then mt = ma + mi
772 If mt < m1 * 1000000 + ma / 2 Then mt = m1 * 1000000 + ma / 2
774 em = 0.05 * bp
776 If em > 0.02 Then em = 0.02
778 If mt < nc * 1000 * em Then mt = nc * 1000 * em
779 If mt = nc * 1000 * em Then mType = "e"
784 ''Return

                   
If Int(100 * m / bTng(i) / hTng(i) ^ 2) / 100 > 0 And _
   Int(100 * N / bTng(i) / hTng(i)) / 100 > 0 And _
       Kfactor > 0 And Kfactor < 1.5 Then
                         
    If BendAxis = "XX" Then
      Mbh2XX(CountX) = m / bTng(i) / hTng(i) ^ 2
        NbhXX(CountX) = N / bTng(i) / hTng(i)
        DesgMntXX(CountX) = mt
         PointX = CountX
           CountX = CountX + 1
                         
             If Abs(N / 1000 - nc) < CntAxial / 2 Then
                 List1.AddItem " < " & BendAxis & " > " & _
                  "           " & Str(Int(N / 1000)) & _
                   "           " & Str(Int(10 * m / 1000000) / 10) & _
                    "           " & Str(Int(100 * m / bTng(i) / hTng(i) ^ 2) / 100) & _
                     "            " & Str(Int(100 * N / bTng(i) / hTng(i)) / 100) & _
                      "           " & Str(Kfactor) & _
                       "           " & Str(Int(10 * mt / 1000000) / 10)
                         End If
                         
                            Else
                            
                         Mbh2YY(CountY) = m / hTng(i) / bTng(i) ^ 2
                        NbhYY(CountY) = N / bTng(i) / hTng(i)
                        DesgMntYY(CountY) = mt
                       PointY = CountY
                      CountY = CountY + 1
                      
                      If Abs(N / 1000 - nc) < CntAxial / 2 Then
                     List1.AddItem " < " & BendAxis & " > " & _
                    "           " & Str(Int(N / 1000)) & _
                   "           " & Str(Int(10 * m / 1000000) / 10) & _
                  "           " & Str(Int(100 * m / hTng(i) / bTng(i) ^ 2) / 100) & _
                 "            " & Str(Int(100 * N / bTng(i) / hTng(i)) / 100) & _
                "           " & Str(Kfactor) & _
               "           " & Str(Int(10 * mt / 1000000) / 10)
             End If
    End If
End If
           
       If Kfactor < 0 Then
          Exit Function
       End If

x = x + 0.5
Next j                 '''>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>^

End Function
Private Sub SetBraceAndBase(ByVal NO As Integer)
 If Left(xBASEresistmnt(NO), 3) = "FIX" Or _
    Left(xBASEresistmnt(NO), 3) = "fix" Or _
    Left(xBASEresistmnt(NO), 1) = "Y" Or _
    Left(xBASEresistmnt(NO), 1) = "y" Then
       xBASEresistmnt(NO) = "FIXED"
       Else
       xBASEresistmnt(NO) = "FREE"
       End If
 If Left(yBASEresistmnt(NO), 3) = "FIX" Or _
    Left(yBASEresistmnt(NO), 3) = "fix" Or _
    Left(yBASEresistmnt(NO), 1) = "Y" Or _
    Left(yBASEresistmnt(NO), 1) = "y" Then
       yBASEresistmnt(NO) = "FIXED"
       Else
       yBASEresistmnt(NO) = "FREE"
       End If
  
 If Left(xBracedCol(NO), 1) = "B" Or _
    Left(xBracedCol(NO), 1) = "b" Or _
    Left(xBracedCol(NO), 1) = "Y" Or _
    Left(xBracedCol(NO), 1) = "y" Then
       xBracedCol(NO) = "BRACED"
       Else
       xBracedCol(NO) = "UNBRACED"
       End If
  If Left(yBracedCol(NO), 1) = "B" Or _
     Left(yBracedCol(NO), 1) = "b" Or _
     Left(yBracedCol(NO), 1) = "Y" Or _
     Left(yBracedCol(NO), 1) = "y" Then
       yBracedCol(NO) = "BRACED"
       Else
       yBracedCol(NO) = "UNBRACED"
       End If
       
End Sub

Private Sub Command3_Click()

ShowAllVisibility

'If Command1.Enabled = True Then
      Command1.Enabled = False
      Command1.BackColor = &H8000000A
'         End If

Command2.Enabled = False
Command3.Enabled = False
Command5.Enabled = False
Dim fnum As Integer
Dim txtFile As String
fnum = FreeFile
Form1.Picture = LoadPicture(NamaFolder & "icon\pilihT.ico")
Option1.Enabled = True
Option2.Enabled = True
Option3.Enabled = True
Option4.Enabled = True
Option5.Enabled = True
Option6.Enabled = True
Option7.Enabled = True
Option8.Enabled = True
Option9.Enabled = True
If mnuItemOpenDwg.Enabled = True Then
Command4.Enabled = False
''Else
''Command4.Enabled = True
End If
DisableText1To62

txtFile = NamaFolder & "tiang\datainput\DefaultStressTiang.txt"

Xinsertion = Text51.text
Yinsertion = Text52.text
fcu = Text53.text
fy = Text54.text
fyv = Text55.text
shrink = Text56.text
creep = Text57.text
stirupD = Text58.text
StirupSPACE = Text59.text
BarMark = Text60.text

Open txtFile For Output As #fnum

Print #fnum, Xinsertion
Print #fnum, Yinsertion
Print #fnum, fcu
Print #fnum, fy
Print #fnum, fyv
Print #fnum, shrink
Print #fnum, creep
Print #fnum, stirupD
Print #fnum, StirupSPACE
Print #fnum, BarMark
Close #fnum

End Sub

Private Sub Command4_Click()

Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Option1.Enabled = True
Option2.Enabled = True
Option3.Enabled = True
Option4.Enabled = True
Option5.Enabled = True
Option6.Enabled = True
Option7.Enabled = True
Option8.Enabled = True
Option9.Enabled = True
Command1.Enabled = True

Picture1.Enabled = False
  Shape2.FillStyle = 0
  Shape5.FillStyle = 0
  Shape7.FillStyle = 0
  Shape8.FillStyle = 0
        Picture1.Top = 900
        Picture1.Left = 0
        Picture1.Height = 50
        Picture1.Width = 50
        Picture1.Visible = False
        Picture1.Cls
Form1.Picture = LoadPicture(NamaFolder & "icon\ukad4.ico")
DrawFWTiang

End Sub

Private Sub Command5_Click()
Dim SetLuput As Date

Form1.Picture = LoadPicture(NamaFolder & "icon\strength.ico")
SetLuput = DateValue("03/07/3013")
If Date >= SetLuput Then
 MsgBox ":::Sila hubungi Wan Sohaimi Wan Mohamed @ 603-61574717::: " _
 , , "   To reinstall   "
 Exit Sub
End If
''''**********************************'''''
i = Int(Val(Right(Command2.Caption, 1)))
Label54.Caption = DesgAxial(i) & " kN."

Line4.BorderColor = vbWhite
Line4.BorderWidth = 1
Command1.Enabled = True
Command1.Visible = True

List1.Clear
ReaDFileTiang
CalColumnStrength

Option1.Enabled = True
Option2.Enabled = True
Option3.Enabled = True
Option4.Enabled = True
Option5.Enabled = True
Option6.Enabled = True
Option7.Enabled = True
Option8.Enabled = True
Option9.Enabled = True

'''CalColumnStrength


Dim AxisX, AxisY, ChkPoint, ChkPointYY As Integer
Dim ChkAxial, ChkMoment, ChkAxialYY, ChkMomentYY As Double
Dim DMntX, DMntY As Double
Dim JK As Integer
AxisX = Mbh2XX(1)
AxisY = NbhXX(1)

For JK = 1 To PointX
  If AxisX < Mbh2XX(JK) Then
     AxisX = Mbh2XX(JK)
  End If
Next

For JK = 1 To PointX
  If AxisY < NbhXX(JK) Then
     AxisY = NbhXX(JK)
  End If
Next

For JK = 1 To PointX
   ChkPoint = JK
   ChkAxial = NbhXX(JK) * bTng(i) * hTng(i)
   If ChkAxial > DesgAxial(i) * 1000 Then
   GoTo 123  'JumpOut
   End If
Next
123 'JumpOut:

For JK = 1 To PointY
   ChkPointYY = JK
   ChkAxialYY = NbhYY(JK) * bTng(i) * hTng(i)
   If ChkAxialYY > DesgAxial(i) * 1000 Then
   GoTo 456  'JumpOut
   End If
Next
456 'JumpOut:


''MsgBox " ", , Str(ChkAxial) & "   " & Str(DesgAxial(i))
DMntX = DesgMntXX(ChkPoint) / bTng(i) / hTng(i) ^ 2
DMntY = DesgMntYY(ChkPointYY) / hTng(i) / bTng(i) ^ 2
ChkMoment = Mbh2XX(ChkPoint)
ChkAxial = NbhXX(ChkPoint)
ChkMomentYY = Mbh2YY(ChkPointYY)
ChkAxialYY = NbhYY(ChkPointYY)
''MsgBox Str(DMntX), , Str(DMntY)
''MsgBox Str(ChkAxial) & "<xx     yy>" & Str(ChkAxialYY), , "Axiaload"
   AxisX = 1.5 * AxisX
   AxisY = 1.5 * AxisY
If AxisX < DMntX * 1.25 Then
   AxisX = DMntX * 1.25
   End If
If AxisX < DMntY * 1.25 Then
   AxisX = DMntY * 1.25
   End If
        Picture1.Enabled = True
        Picture1.Top = 2680
        Picture1.Left = 0
        Picture1.Height = 4488
        Picture1.Width = 10880
        Picture1.Visible = True
        ''Picture1.BackColor = RGB(0, 0, 0)
        Picture1.BackColor = vbWhite
        Picture1.ForeColor = RGB(0, 0, 0)
        Picture1.ScaleHeight = AxisY
        Picture1.ScaleWidth = AxisX
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim Count As Integer
Dim SkelX, SkelY As Double
SkelX = 1
SkelY = 1

    Picture1.Line (AxisX / 50, 4 * AxisY / 50)-(AxisX / 50, 49 * AxisY / 50), vbBlack
    Picture1.Line (AxisX / 50, 49 * AxisY / 50)-(49 * AxisX / 50, 49 * AxisY / 50), vbBlack
    '''Picture1.FontSize = 10
78   If SkelY > AxisY / 1.5 Then GoTo 87
        Picture1.Line (AxisX / 50, 49 * AxisY / 50 - SkelY)-(0.5 * AxisX / 50, 49 * AxisY / 50 - SkelY), vbBlack
         SkelY = SkelY + 1
          GoTo 78
87

89   If SkelX > AxisX / 1.25 Then GoTo 98
        Picture1.Line (SkelX + AxisX / 50, 49 * AxisY / 50)-(SkelX + AxisX / 50, 49.85 * AxisY / 50), vbBlack
         SkelX = SkelX + 1
          GoTo 89
98
     SkelX = 0.5
895   If SkelX > AxisX / 1.25 Then GoTo 985
        Picture1.Line (SkelX + AxisX / 50, 49 * AxisY / 50)-(SkelX + AxisX / 50, 49.5 * AxisY / 50), vbWhite
         SkelX = SkelX + 1
          GoTo 895
985
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Picture1.Line (AxisX / 50, 49 * AxisY / 50 - ChkAxial)-(48 * AxisX / 50, 49 * AxisY / 50 - ChkAxial), vbGreen
    Picture1.Line (AxisX / 50, 49 * AxisY / 50 - ChkAxialYY)-(48 * AxisX / 50, 49 * AxisY / 50 - ChkAxialYY), vbGreen
    Picture1.Line (ChkMoment + AxisX / 50, 49 * AxisY / 50)-(ChkMoment + AxisX / 50, 49 * AxisY / 50 - ChkAxial), vbCyan
    Picture1.Line (ChkMomentYY + AxisX / 50, 49 * AxisY / 50)-(ChkMomentYY + AxisX / 50, 49 * AxisY / 50 - ChkAxialYY), vbMagenta
    Picture1.Line (DMntX + AxisX / 50, 49 * AxisY / 50)-(DMntX + AxisX / 50, 9 * AxisY / 50), vbCyan
    Picture1.Line (DMntY + AxisX / 50, 49 * AxisY / 50)-(DMntY + AxisX / 50, 11 * AxisY / 50), vbMagenta
    
    Picture1.CurrentX = DMntX - AxisX / 50
    Picture1.CurrentY = 7 * AxisY / 50
    Picture1.Print Str(Int(100 * DMntX) / 100)
    
    Picture1.CurrentX = DMntY - AxisX / 50
    Picture1.CurrentY = 9 * AxisY / 50
    Picture1.Print Str(Int(100 * DMntY) / 100)
    
    Picture1.CurrentX = 2 * AxisX / 50
    Picture1.CurrentY = 49 * AxisY / 50 - ChkAxial
    Picture1.Print Str(Int(100 * ChkAxial) / 100)
    
    Picture1.CurrentX = ChkMoment + AxisX / 50
    Picture1.CurrentY = 44 * AxisY / 50
    Picture1.Print Str(Int(100 * ChkMoment) / 100)
    
    Picture1.CurrentX = 5 * AxisX / 50
    Picture1.CurrentY = 49 * AxisY / 50 - ChkAxialYY
    Picture1.Print Str(Int(100 * ChkAxialYY) / 100)
    
    Picture1.CurrentX = ChkMomentYY + AxisX / 50
    Picture1.CurrentY = 47 * AxisY / 50
    Picture1.Print Str(Int(100 * ChkMomentYY) / 100)
    
    Picture1.CurrentX = 45 * AxisX / 50
    Picture1.CurrentY = 44 * AxisY / 50
    Picture1.Print "M/bh2"
    
    Picture1.CurrentX = 2 * AxisX / 50
    Picture1.CurrentY = 5 * AxisY / 50
    Picture1.Print "N/bh"

    
    Picture1.CurrentX = 8 * AxisX / 50
    Picture1.CurrentY = 4 * AxisY / 50
    Picture1.Print " (merah+cyan @x-x & biru+magenta @y-y)"
    
    Picture1.CurrentX = 7 * AxisX / 50
    Picture1.CurrentY = AxisY / 50
    Picture1.Print "COLUMN INTERACTION DIAGRAM " & Command2.Caption
            
    Picture1.CurrentX = 33 * AxisX / 50
    Picture1.CurrentY = 11 * AxisY / 50
    Picture1.Print "b = " & Str(Int(bTng(i))) & "    h = " & Str(Int(hTng(i)))
    
    Picture1.CurrentX = 33 * AxisX / 50
    Picture1.CurrentY = 14 * AxisY / 50
    Picture1.Print "100Asc/bh = " & Str(Int(10000 * Ast / bTng(i) / hTng(i)) / 100)
    
    Picture1.CurrentX = 33 * AxisX / 50
    Picture1.CurrentY = 17 * AxisY / 50
    Picture1.Print "d/h = " & Str(Int(100 * (hTng(i) - dPrime) / hTng(i)) / 100) & _
                   "   d/b = "; Str(Int(100 * (bTng(i) - dPrime) / bTng(i)) / 100)
    
    Picture1.CurrentX = 33 * AxisX / 50
    Picture1.CurrentY = 20 * AxisY / 50
    Picture1.Print "Le/h = " & Str(Int(10 * EffHeightXX(i) / (hTng(i)) / 10)) & _
                   "   Le/b = " & Str(Int(10 * EffHeightYY(i) / (bTng(i)) / 10))
 
    ''''''''''''''''''''''
    Picture1.CurrentX = 37 * AxisX / 50
    Picture1.CurrentY = 23 * AxisY / 50
    Picture1.Print Str(BarXTopN(i)) & "T" & Str(BarXTopD(i))
    
    Picture1.CurrentX = 33 * AxisX / 50
    Picture1.CurrentY = 25 * AxisY / 50
    Picture1.Print Str(BarYLefN(i)) & "T" & Str(BarYLefD(i))
    
    Picture1.CurrentX = 41 * AxisX / 50
    Picture1.CurrentY = 25 * AxisY / 50
    Picture1.Print Str(BarYRigN(i)) & "T" & Str(BarYRigD(i))
    
    Picture1.CurrentX = 37 * AxisX / 50
    Picture1.CurrentY = 27 * AxisY / 50
    Picture1.Print Str(BarXBotN(i)) & "T" & Str(BarXBotD(i))
    ''''''''''''''''''''''


For Count = 1 To PointX
    Picture1.PSet (Mbh2XX(Count) + AxisX / 50, AxisY - NbhXX(Count) - _
    AxisY / 50), vbRed
    DoEvents
Next


For Count = 1 To PointY
   Picture1.PSet (Mbh2YY(Count) + AxisX / 50, AxisY - NbhYY(Count) - _
   AxisY / 50), vbBlue
   DoEvents
Next


If Command2.Left = 110 Then
    SavePicture Picture1.Image, NamaFolder & "tiang\resultoutput\inter_chart1.bmp"
        End If
If Command2.Left = 1110 Then
    SavePicture Picture1.Image, NamaFolder & "tiang\resultoutput\inter_chart2.bmp"
        End If
If Command2.Left = 2110 Then
    SavePicture Picture1.Image, NamaFolder & "tiang\resultoutput\inter_chart3.bmp"
        End If
If Command2.Left = 3110 Then
    SavePicture Picture1.Image, NamaFolder & "tiang\resultoutput\inter_chart4.bmp"
        End If
If Command2.Left = 4110 Then
    SavePicture Picture1.Image, NamaFolder & "tiang\resultoutput\inter_chart5.bmp"
        End If
If Command2.Left = 5110 Then
    SavePicture Picture1.Image, NamaFolder & "tiang\resultoutput\inter_chart6.bmp"
        End If
If Command2.Left = 6110 Then
    SavePicture Picture1.Image, NamaFolder & "tiang\resultoutput\inter_chart7.bmp"
        End If
If Command2.Left = 7110 Then
    SavePicture Picture1.Image, NamaFolder & "tiang\resultoutput\inter_chart8.bmp"
        End If
If Command2.Left = 8110 Then
    SavePicture Picture1.Image, NamaFolder & "tiang\resultoutput\inter_chart9.bmp"
        End If
       
Form1.Picture = LoadPicture(NamaFolder & "icon\pilihT.ico")
End Sub


Private Sub Form_Load()
NamaFolder = "C:\autodraf\"
End Sub

'''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''
'''''''FILES MANAGEMENT''''''''''''''
'''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''

Private Sub mnuItemOpenDwg_Click() 'when user clicks Close command
Dim chkFile As String
Dim caseNo As Integer
mnuItemExit.Enabled = True
mnuItemOpenDwg = True

    CommonDialog1.Filter = "Dwg files (*.DWG)|*.DWG"
    CommonDialog1.ShowOpen       'display Open dialog box
    dwgName = CommonDialog1.FileName
    Label5.Caption = "..." & Right(dwgName, 50)
    mnuItemExit.Enabled = True
    mnuItemOpenDwg = True
If dwgName = "" Then
 mnuItemExit.Enabled = True
 mnuItemOpenDwg.Enabled = True
 
 Form1.Picture = LoadPicture(NamaFolder & "icon\ukad.ico")
  
Else
 mnuItemExit.Enabled = True
 mnuItemOpenDwg.Enabled = False
 
 Form1.Picture = LoadPicture(NamaFolder & "icon\ukad1.ico")
 ''mnuItemFile.Enabled = False
 ''Command1.Enabled = True
 ''''''''''''''''''''''''''''
 ''''''''''''''''''''''''''''
        Picture1.Enabled = False
        Picture1.Top = 900
        Picture1.Left = 0
        Picture1.Height = 50
        Picture1.Width = 50
        Picture1.Visible = False
        Picture1.Cls
'''''''AUTOCAD''''''''''''''
''''''''''''''''''''''''''''
Dim SetLuput As Date
SetLuput = DateValue("03/07/3013")
If Date >= SetLuput Then
 MsgBox ":::Sila hubungi Wan Sohaimi Wan Mohamed @ 603-61574717::: " _
 , , "   To reinstall   "
 Exit Sub
End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
If dwgName = "" Then
MsgBox "Sila pilih fail dwg untuk kerja.", , "NOTA AM:"
Exit Sub
End If

ReaDFileTiang

StartAutoCAD
SetLayer
Form1.Picture = LoadPicture(NamaFolder & "icon\statacad.ico")

Command4.Enabled = False
Option1.Enabled = True
Option2.Enabled = True
Option3.Enabled = True
Option4.Enabled = True
Option5.Enabled = True
Option6.Enabled = True
Option7.Enabled = True
Option8.Enabled = True
Option9.Enabled = True

 mnuItemExit.Enabled = True
 mnuItemOpenDwg.Enabled = False
 mnuItemFile.Enabled = True
 mnuItemFile.Visible = True
 mnuItemFile.Caption = "Klik di sini!"
 '''''''''''''''''''''''''''''''''''''''
 
End If
       
        Picture1.Enabled = False
        Picture1.Top = 900
        Picture1.Left = 0
        Picture1.Height = 50
        Picture1.Width = 50
        Picture1.Visible = False
        Picture1.Cls
End Sub


Private Sub mnuItemExit_Click()  'when user clicks Exit command
    End                          'quit program
End Sub



Private Sub Option1_Click()
List1.Clear
Command1.Visible = False
Form1.Picture = LoadPicture(NamaFolder & "icon\datam.ico")
Image1.Picture = LoadPicture(NamaFolder & "icon\ukad3.ico")
Picture1.Enabled = False
  Shape2.FillStyle = 0
  Shape5.FillStyle = 0
  Shape7.FillStyle = 0
  Shape8.FillStyle = 0
        Picture1.Top = 900
        Picture1.Left = 0
        Picture1.Height = 50
        Picture1.Width = 50
        Picture1.Visible = False
        Picture1.Cls
           
bilJenisTng = 1
ReaDFileTiang
ShowAllVisibility
Shape25.Shape = 3
'''ShowExistingMember (Int(Right(Command2.Caption, 1)))
Call GraphicOne(Val(bTng(bilJenisTng)), Val(hTng(bilJenisTng)), _
       Val(BarXBotN(bilJenisTng)) * Val(BarXBotD(bilJenisTng)) ^ 2 * 3.14 / 4, _
       Val(BarXBotD(bilJenisTng)) ^ 2 * 3.14 / 2 + Val(BarYLefN(bilJenisTng)) * Val(BarYLefD(bilJenisTng)) ^ 2 * 3.14 / 4)
         
Dim fnum As Integer
Dim txtFile, Temp As String
fnum = FreeFile
txtFile = NamaFolder & "tiang\datainput\ColOneGET.txt"
Command2.Enabled = True
Command2.Left = 110
Command2.Caption = "Tiang - 1"
Command4.Enabled = False
Command5.Enabled = False
If Command1.Enabled = True Then
   Command1.Enabled = False
      End If
''''NoOfSpan = 1
bilJenisTng = 1
Option1.Enabled = True
Option2.Enabled = False
Option3.Enabled = False
Option4.Enabled = False
Option5.Enabled = False
Option6.Enabled = False
Option7.Enabled = False
Option8.Enabled = False
Option9.Enabled = False

EnableText1To62



''************************************''

Text1.text = UppCoB(1)
Text2.text = UppCoH(1)
Text3.text = UppCoHgt(1)

Text4.text = xBASEresistmnt(1)

Text5.text = yUppLeBb(1)
Text6.text = yUppLeBh(1)
Text7.text = yUppLeBl(1)
Text8.text = yUppRiBb(1)
Text9.text = yUppRiBh(1)
Text10.text = yUppRiBl(1)
Text11.text = xUppRiBb(1)
Text12.text = xUppRiBh(1)
Text13.text = xUppRiBl(1)
Text14.text = xUppLeBb(1)
Text15.text = xUppLeBh(1)
Text16.text = xUppLeBl(1)

Text17.text = bTng(1)
Text18.text = hTng(1)
Text19.text = FloorHgt(1)

Text20.text = Mx2(1)
Text21.text = My2(1)
Text22.text = Mx1(1)
Text23.text = My1(1)

Text24.text = xLowLeBb(1)
Text25.text = xLowLeBh(1)
Text26.text = xLowLeBl(1)
Text27.text = xLowRiBb(1)
Text28.text = xLowRiBh(1)
Text29.text = xLowRiBl(1)

Text30.text = yLowRiBb(1)
Text31.text = yLowRiBh(1)
Text32.text = yLowRiBl(1)
Text33.text = yLowLeBb(1)
Text34.text = yLowLeBh(1)
Text35.text = yLowLeBl(1)

Text36.text = LowCoB(1)
Text37.text = LowCoH(1)
Text38.text = LowCoHgt(1)

Text39.text = yBASEresistmnt(1)
Text40.text = DesgAxial(1)

Text41.text = BarYLefN(1)
Text42.text = BarYLefD(1)
Text43.text = BarYRigN(1)
Text44.text = BarYRigD(1)

Text45.text = BarXTopN(1)
Text46.text = BarXTopD(1)
Text47.text = BarXBotN(1)
Text48.text = BarXBotD(1)

Text49.text = GridTng(1)
Text50.text = CoverTng(1)

Text61.text = xBracedCol(1)
Text62.text = yBracedCol(1)



''*************************************************************''

Dimensi1.Value1 = Text1.text
Dimensi1.Value2 = Text2.text
Dimensi1.Value3 = Text3.text

AppliedStress1.Value4 = Text4.text

Dimensi1.Value5 = Text5.text
Dimensi1.Value6 = Text6.text
Dimensi1.Value7 = Text7.text
Dimensi1.Value8 = Text8.text
Dimensi1.Value9 = Text9.text
Dimensi1.Value10 = Text10.text
Dimensi1.Value11 = Text11.text
Dimensi1.Value12 = Text12.text
Dimensi1.Value13 = Text13.text
Dimensi1.Value14 = Text14.text
Dimensi1.Value15 = Text15.text
Dimensi1.Value16 = Text16.text

Dimensi1.Value17 = Text17.text
Dimensi1.Value18 = Text18.text
Dimensi1.Value19 = Text19.text

AppliedStress1.Value20 = Text20.text
AppliedStress1.Value21 = Text21.text
AppliedStress1.Value22 = Text22.text
AppliedStress1.Value23 = Text23.text

Dimensi1.Value24 = Text24.text
Dimensi1.Value25 = Text25.text
Dimensi1.Value26 = Text26.text
Dimensi1.Value27 = Text27.text
Dimensi1.Value28 = Text28.text
Dimensi1.Value29 = Text29.text

Dimensi1.Value30 = Text30.text
Dimensi1.Value31 = Text31.text
Dimensi1.Value32 = Text32.text
Dimensi1.Value33 = Text33.text
Dimensi1.Value34 = Text34.text
Dimensi1.Value35 = Text35.text

Dimensi1.Value36 = Text36.text
Dimensi1.Value37 = Text37.text
Dimensi1.Value38 = Text38.text

AppliedStress1.Value39 = Text39.text
AppliedStress1.Value40 = Text40.text

Tetulang1.Value41 = Text41.text
Tetulang1.Value42 = Text42.text
Tetulang1.Value43 = Text43.text
Tetulang1.Value44 = Text44.text

Tetulang1.Value45 = Text45.text
Tetulang1.Value46 = Text46.text
Tetulang1.Value47 = Text47.text
Tetulang1.Value48 = Text48.text

Dimensi1.Value49 = Text49.text
Dimensi1.Value50 = Text50.text

AppliedStress1.Value61 = Text61.text
AppliedStress1.Value62 = Text62.text

''''**********************************************************''''



Text1.text = Dimensi1.BtngAtas
Text2.text = Dimensi1.HtngAtas
Text3.text = Dimensi1.TtngAtas

Text4.text = AppliedStress1.XbaseCondition

Text5.text = Dimensi1.BrskWAtas
Text6.text = Dimensi1.HrskWAtas
Text7.text = Dimensi1.PrskWAtas
Text8.text = Dimensi1.BrskEAtas
Text9.text = Dimensi1.HrskEAtas
Text10.text = Dimensi1.PrskEAtas
Text11.text = Dimensi1.BrskSAtas
Text12.text = Dimensi1.HrskSAtas
Text13.text = Dimensi1.PrskSAtas
Text14.text = Dimensi1.BrskNAtas
Text15.text = Dimensi1.HrskNAtas
Text16.text = Dimensi1.PrskNAtas

Text17.text = Dimensi1.BtianG
Text18.text = Dimensi1.HtianG
Text19.text = Dimensi1.TtianG

Text20.text = AppliedStress1.MomentX22
Text21.text = AppliedStress1.MomentY22
Text22.text = AppliedStress1.MomentX11
Text23.text = AppliedStress1.MomentY11

Text24.text = Dimensi1.BrskNBawah
Text25.text = Dimensi1.HrskNBawah
Text26.text = Dimensi1.PrskNBawah
Text27.text = Dimensi1.BrskSBawah
Text28.text = Dimensi1.HrskSBawah
Text29.text = Dimensi1.PrskSBawah

Text30.text = Dimensi1.BrskEBawah
Text31.text = Dimensi1.HrskEBawah
Text32.text = Dimensi1.PrskEBawah
Text33.text = Dimensi1.BrskWBawah
Text34.text = Dimensi1.HrskWBawah
Text35.text = Dimensi1.PrskWBawah

Text36.text = Dimensi1.BtngBawah
Text37.text = Dimensi1.HtngBawah
Text38.text = Dimensi1.TtngBawah

Text39.text = AppliedStress1.YbaseCondition
Text40.text = AppliedStress1.AxiaLoaD

Text41.text = Tetulang1.BarOnYWno
Text42.text = Tetulang1.BarOnYWdia
Text43.text = Tetulang1.BarOnYEno
Text44.text = Tetulang1.BarOnYEdia

Text45.text = Tetulang1.BarOnXNno
Text46.text = Tetulang1.BarOnXNdia
Text47.text = Tetulang1.BarOnXSno
Text48.text = Tetulang1.BarOnXSdia

Text49.text = Dimensi1.NamaTiang
Text50.text = Dimensi1.Cover

Text61.text = AppliedStress1.XBracedFrame
Text62.text = AppliedStress1.YBracedFrame

''************************************************************

End Sub

Private Sub Option2_Click()
List1.Clear
Command1.Visible = False

Form1.Picture = LoadPicture(NamaFolder & "icon\datam.ico")
Image1.Picture = LoadPicture(NamaFolder & "icon\cskp.ico")
Picture1.Enabled = False
  Shape2.FillStyle = 0
  Shape5.FillStyle = 0
  Shape7.FillStyle = 0
  Shape8.FillStyle = 0
        Picture1.Top = 900
        Picture1.Left = 0
        Picture1.Height = 50
        Picture1.Width = 50
        Picture1.Visible = False
        Picture1.Cls

bilJenisTng = 2
ReaDFileTiang
ShowAllVisibility
Shape25.Shape = 3
Call GraphicOne(Val(bTng(bilJenisTng)), Val(hTng(bilJenisTng)), _
       Val(BarXBotN(bilJenisTng)) * Val(BarXBotD(bilJenisTng)) ^ 2 * 3.14 / 4, _
       Val(BarXBotD(bilJenisTng)) ^ 2 * 3.14 / 2 + Val(BarYLefN(bilJenisTng)) * Val(BarYLefD(bilJenisTng)) ^ 2 * 3.14 / 4)
         
Dim fnum As Integer
Dim txtFile, Temp As String
fnum = FreeFile
txtFile = NamaFolder & "tiang\datainput\ColTwoGET.txt"
Command2.Enabled = True
Command2.Left = 1110
Command2.Caption = "Tiang - 2"
Command4.Enabled = False
Command5.Enabled = False
If Command1.Enabled = True Then
   Command1.Enabled = False
      End If

Option1.Enabled = False
Option2.Enabled = True
Option3.Enabled = False
Option4.Enabled = False
Option5.Enabled = False
Option6.Enabled = False
Option7.Enabled = False
Option8.Enabled = False
Option9.Enabled = False

EnableText1To62
''************************************''

Text1.text = UppCoB(2)
Text2.text = UppCoH(2)
Text3.text = UppCoHgt(2)

Text4.text = xBASEresistmnt(2)

Text5.text = yUppLeBb(2)
Text6.text = yUppLeBh(2)
Text7.text = yUppLeBl(2)
Text8.text = yUppRiBb(2)
Text9.text = yUppRiBh(2)
Text10.text = yUppRiBl(2)
Text11.text = xUppRiBb(2)
Text12.text = xUppRiBh(2)
Text13.text = xUppRiBl(2)
Text14.text = xUppLeBb(2)
Text15.text = xUppLeBh(2)
Text16.text = xUppLeBl(2)

Text17.text = bTng(2)
Text18.text = hTng(2)
Text19.text = FloorHgt(2)

Text20.text = Mx2(2)
Text21.text = My2(2)
Text22.text = Mx1(2)
Text23.text = My1(2)

Text24.text = xLowLeBb(2)
Text25.text = xLowLeBh(2)
Text26.text = xLowLeBl(2)
Text27.text = xLowRiBb(2)
Text28.text = xLowRiBh(2)
Text29.text = xLowRiBl(2)

Text30.text = yLowRiBb(2)
Text31.text = yLowRiBh(2)
Text32.text = yLowRiBl(2)
Text33.text = yLowLeBb(2)
Text34.text = yLowLeBh(2)
Text35.text = yLowLeBl(2)

Text36.text = LowCoB(2)
Text37.text = LowCoH(2)
Text38.text = LowCoHgt(2)

Text39.text = yBASEresistmnt(2)
Text40.text = DesgAxial(2)

Text41.text = BarYLefN(2)
Text42.text = BarYLefD(2)
Text43.text = BarYRigN(2)
Text44.text = BarYRigD(2)

Text45.text = BarXTopN(2)
Text46.text = BarXTopD(2)
Text47.text = BarXBotN(2)
Text48.text = BarXBotD(2)

Text49.text = GridTng(2)
Text50.text = CoverTng(2)

Text61.text = xBracedCol(2)
Text62.text = yBracedCol(2)

''*************************************************************''

Dimensi2.Value1 = Text1.text
Dimensi2.Value2 = Text2.text
Dimensi2.Value3 = Text3.text

AppliedStress1.Value4 = Text4.text

Dimensi2.Value5 = Text5.text
Dimensi2.Value6 = Text6.text
Dimensi2.Value7 = Text7.text
Dimensi2.Value8 = Text8.text
Dimensi2.Value9 = Text9.text
Dimensi2.Value10 = Text10.text
Dimensi2.Value11 = Text11.text
Dimensi2.Value12 = Text12.text
Dimensi2.Value13 = Text13.text
Dimensi2.Value14 = Text14.text
Dimensi2.Value15 = Text15.text
Dimensi2.Value16 = Text16.text

Dimensi2.Value17 = Text17.text
Dimensi2.Value18 = Text18.text
Dimensi2.Value19 = Text19.text

AppliedStress2.Value20 = Text20.text
AppliedStress2.Value21 = Text21.text
AppliedStress2.Value22 = Text22.text
AppliedStress2.Value23 = Text23.text

Dimensi2.Value24 = Text24.text
Dimensi2.Value25 = Text25.text
Dimensi2.Value26 = Text26.text
Dimensi2.Value27 = Text27.text
Dimensi2.Value28 = Text28.text
Dimensi2.Value29 = Text29.text

Dimensi2.Value30 = Text30.text
Dimensi2.Value31 = Text31.text
Dimensi2.Value32 = Text32.text
Dimensi2.Value33 = Text33.text
Dimensi2.Value34 = Text34.text
Dimensi2.Value35 = Text35.text

Dimensi2.Value36 = Text36.text
Dimensi2.Value37 = Text37.text
Dimensi2.Value38 = Text38.text

AppliedStress2.Value39 = Text39.text
AppliedStress2.Value40 = Text40.text

Tetulang2.Value41 = Text41.text
Tetulang2.Value42 = Text42.text
Tetulang2.Value43 = Text43.text
Tetulang2.Value44 = Text44.text

Tetulang2.Value45 = Text45.text
Tetulang2.Value46 = Text46.text
Tetulang2.Value47 = Text47.text
Tetulang2.Value48 = Text48.text

Dimensi2.Value49 = Text49.text
Dimensi2.Value50 = Text50.text

AppliedStress2.Value61 = Text61.text
AppliedStress2.Value62 = Text62.text

''''**********************************************************''''

Text1.text = Dimensi2.BtngAtas
Text2.text = Dimensi2.HtngAtas
Text3.text = Dimensi2.TtngAtas

Text4.text = AppliedStress1.XbaseCondition

Text5.text = Dimensi2.BrskWAtas
Text6.text = Dimensi2.HrskWAtas
Text7.text = Dimensi2.PrskWAtas
Text8.text = Dimensi2.BrskEAtas
Text9.text = Dimensi2.HrskEAtas
Text10.text = Dimensi2.PrskEAtas
Text11.text = Dimensi2.BrskSAtas
Text12.text = Dimensi2.HrskSAtas
Text13.text = Dimensi2.PrskSAtas
Text14.text = Dimensi2.BrskNAtas
Text15.text = Dimensi2.HrskNAtas
Text16.text = Dimensi2.PrskNAtas

Text17.text = Dimensi2.BtianG
Text18.text = Dimensi2.HtianG
Text19.text = Dimensi2.TtianG

Text20.text = AppliedStress2.MomentX22
Text21.text = AppliedStress2.MomentY22
Text22.text = AppliedStress2.MomentX11
Text23.text = AppliedStress2.MomentY11

Text24.text = Dimensi2.BrskNBawah
Text25.text = Dimensi2.HrskNBawah
Text26.text = Dimensi2.PrskNBawah
Text27.text = Dimensi2.BrskSBawah
Text28.text = Dimensi2.HrskSBawah
Text29.text = Dimensi2.PrskSBawah

Text30.text = Dimensi2.BrskEBawah
Text31.text = Dimensi2.HrskEBawah
Text32.text = Dimensi2.PrskEBawah
Text33.text = Dimensi2.BrskWBawah
Text34.text = Dimensi2.HrskWBawah
Text35.text = Dimensi2.PrskWBawah

Text36.text = Dimensi2.BtngBawah
Text37.text = Dimensi2.HtngBawah
Text38.text = Dimensi2.TtngBawah

Text39.text = AppliedStress2.YbaseCondition
Text40.text = AppliedStress2.AxiaLoaD

Text41.text = Tetulang2.BarOnYWno
Text42.text = Tetulang2.BarOnYWdia
Text43.text = Tetulang2.BarOnYEno
Text44.text = Tetulang2.BarOnYEdia

Text45.text = Tetulang2.BarOnXNno
Text46.text = Tetulang2.BarOnXNdia
Text47.text = Tetulang2.BarOnXSno
Text48.text = Tetulang2.BarOnXSdia

Text49.text = Dimensi2.NamaTiang
Text50.text = Dimensi2.Cover

Text61.text = AppliedStress2.XBracedFrame
Text62.text = AppliedStress2.YBracedFrame

''************************************************************


End Sub

Private Sub Option3_Click()
List1.Clear
Command1.Visible = False
Form1.Picture = LoadPicture(NamaFolder & "icon\datam.ico")
Image1.Picture = LoadPicture(NamaFolder & "icon\ukad4.ico")
Picture1.Enabled = False
  Shape2.FillStyle = 0
  Shape5.FillStyle = 0
  Shape7.FillStyle = 0
  Shape8.FillStyle = 0
        Picture1.Top = 900
        Picture1.Left = 0
        Picture1.Height = 50
        Picture1.Width = 50
        Picture1.Visible = False
        Picture1.Cls
        
bilJenisTng = 3
ReaDFileTiang
ShowAllVisibility
Shape25.Shape = 3
Call GraphicOne(Val(bTng(bilJenisTng)), Val(hTng(bilJenisTng)), _
       Val(BarXBotN(bilJenisTng)) * Val(BarXBotD(bilJenisTng)) ^ 2 * 3.14 / 4, _
       Val(BarXBotD(bilJenisTng)) ^ 2 * 3.14 / 2 + Val(BarYLefN(bilJenisTng)) * Val(BarYLefD(bilJenisTng)) ^ 2 * 3.14 / 4)
         
Dim fnum As Integer
Dim txtFile, Temp As String
fnum = FreeFile
txtFile = NamaFolder & "tiang\datainput\ColThreeGET.txt"
Command2.Enabled = True
Command2.Left = 2110
Command2.Caption = "Tiang - 3"
Command4.Enabled = False
Command5.Enabled = False
If Command1.Enabled = True Then
   Command1.Enabled = False
      End If

Option1.Enabled = False
Option2.Enabled = False
Option3.Enabled = True
Option4.Enabled = False
Option5.Enabled = False
Option6.Enabled = False
Option7.Enabled = False
Option8.Enabled = False
Option9.Enabled = False

EnableText1To62


''************************************''

Text1.text = UppCoB(3)
Text2.text = UppCoH(3)
Text3.text = UppCoHgt(3)

Text4.text = xBASEresistmnt(3)

Text5.text = yUppLeBb(3)
Text6.text = yUppLeBh(3)
Text7.text = yUppLeBl(3)
Text8.text = yUppRiBb(3)
Text9.text = yUppRiBh(3)
Text10.text = yUppRiBl(3)
Text11.text = xUppRiBb(3)
Text12.text = xUppRiBh(3)
Text13.text = xUppRiBl(3)
Text14.text = xUppLeBb(3)
Text15.text = xUppLeBh(3)
Text16.text = xUppLeBl(3)

Text17.text = bTng(3)
Text18.text = hTng(3)
Text19.text = FloorHgt(3)

Text20.text = Mx2(3)
Text21.text = My2(3)
Text22.text = Mx1(3)
Text23.text = My1(3)

Text24.text = xLowLeBb(3)
Text25.text = xLowLeBh(3)
Text26.text = xLowLeBl(3)
Text27.text = xLowRiBb(3)
Text28.text = xLowRiBh(3)
Text29.text = xLowRiBl(3)

Text30.text = yLowRiBb(3)
Text31.text = yLowRiBh(3)
Text32.text = yLowRiBl(3)
Text33.text = yLowLeBb(3)
Text34.text = yLowLeBh(3)
Text35.text = yLowLeBl(3)

Text36.text = LowCoB(3)
Text37.text = LowCoH(3)
Text38.text = LowCoHgt(3)

Text39.text = yBASEresistmnt(3)
Text40.text = DesgAxial(3)

Text41.text = BarYLefN(3)
Text42.text = BarYLefD(3)
Text43.text = BarYRigN(3)
Text44.text = BarYRigD(3)

Text45.text = BarXTopN(3)
Text46.text = BarXTopD(3)
Text47.text = BarXBotN(3)
Text48.text = BarXBotD(3)

Text49.text = GridTng(3)
Text50.text = CoverTng(3)

Text61.text = xBracedCol(3)
Text62.text = yBracedCol(3)



''*************************************************************''

Dimensi3.Value1 = Text1.text
Dimensi3.Value2 = Text2.text
Dimensi3.Value3 = Text3.text

AppliedStress3.Value4 = Text4.text

Dimensi3.Value5 = Text5.text
Dimensi3.Value6 = Text6.text
Dimensi3.Value7 = Text7.text
Dimensi3.Value8 = Text8.text
Dimensi3.Value9 = Text9.text
Dimensi3.Value10 = Text10.text
Dimensi3.Value11 = Text11.text
Dimensi3.Value12 = Text12.text
Dimensi3.Value13 = Text13.text
Dimensi3.Value14 = Text14.text
Dimensi3.Value15 = Text15.text
Dimensi3.Value16 = Text16.text

Dimensi3.Value17 = Text17.text
Dimensi3.Value18 = Text18.text
Dimensi3.Value19 = Text19.text

AppliedStress3.Value20 = Text20.text
AppliedStress3.Value21 = Text21.text
AppliedStress3.Value22 = Text22.text
AppliedStress3.Value23 = Text23.text

Dimensi3.Value24 = Text24.text
Dimensi3.Value25 = Text25.text
Dimensi3.Value26 = Text26.text
Dimensi3.Value27 = Text27.text
Dimensi3.Value28 = Text28.text
Dimensi3.Value29 = Text29.text

Dimensi3.Value30 = Text30.text
Dimensi3.Value31 = Text31.text
Dimensi3.Value32 = Text32.text
Dimensi3.Value33 = Text33.text
Dimensi3.Value34 = Text34.text
Dimensi3.Value35 = Text35.text

Dimensi3.Value36 = Text36.text
Dimensi3.Value37 = Text37.text
Dimensi3.Value38 = Text38.text

AppliedStress3.Value39 = Text39.text
AppliedStress3.Value40 = Text40.text

Tetulang3.Value41 = Text41.text
Tetulang3.Value42 = Text42.text
Tetulang3.Value43 = Text43.text
Tetulang3.Value44 = Text44.text

Tetulang3.Value45 = Text45.text
Tetulang3.Value46 = Text46.text
Tetulang3.Value47 = Text47.text
Tetulang3.Value48 = Text48.text

Dimensi3.Value49 = Text49.text
Dimensi3.Value50 = Text50.text

AppliedStress3.Value61 = Text61.text
AppliedStress3.Value62 = Text62.text

''''**********************************************************''''



Text1.text = Dimensi3.BtngAtas
Text2.text = Dimensi3.HtngAtas
Text3.text = Dimensi3.TtngAtas

Text4.text = AppliedStress3.XbaseCondition

Text5.text = Dimensi3.BrskWAtas
Text6.text = Dimensi3.HrskWAtas
Text7.text = Dimensi3.PrskWAtas
Text8.text = Dimensi3.BrskEAtas
Text9.text = Dimensi3.HrskEAtas
Text10.text = Dimensi3.PrskEAtas
Text11.text = Dimensi3.BrskSAtas
Text12.text = Dimensi3.HrskSAtas
Text13.text = Dimensi3.PrskSAtas
Text14.text = Dimensi3.BrskNAtas
Text15.text = Dimensi3.HrskNAtas
Text16.text = Dimensi3.PrskNAtas

Text17.text = Dimensi3.BtianG
Text18.text = Dimensi3.HtianG
Text19.text = Dimensi3.TtianG

Text20.text = AppliedStress3.MomentX22
Text21.text = AppliedStress3.MomentY22
Text22.text = AppliedStress3.MomentX11
Text23.text = AppliedStress3.MomentY11

Text24.text = Dimensi3.BrskNBawah
Text25.text = Dimensi3.HrskNBawah
Text26.text = Dimensi3.PrskNBawah
Text27.text = Dimensi3.BrskSBawah
Text28.text = Dimensi3.HrskSBawah
Text29.text = Dimensi3.PrskSBawah

Text30.text = Dimensi3.BrskEBawah
Text31.text = Dimensi3.HrskEBawah
Text32.text = Dimensi3.PrskEBawah
Text33.text = Dimensi3.BrskWBawah
Text34.text = Dimensi3.HrskWBawah
Text35.text = Dimensi3.PrskWBawah

Text36.text = Dimensi3.BtngBawah
Text37.text = Dimensi3.HtngBawah
Text38.text = Dimensi3.TtngBawah

Text39.text = AppliedStress3.YbaseCondition
Text40.text = AppliedStress3.AxiaLoaD

Text41.text = Tetulang3.BarOnYWno
Text42.text = Tetulang3.BarOnYWdia
Text43.text = Tetulang3.BarOnYEno
Text44.text = Tetulang3.BarOnYEdia

Text45.text = Tetulang3.BarOnXNno
Text46.text = Tetulang3.BarOnXNdia
Text47.text = Tetulang3.BarOnXSno
Text48.text = Tetulang3.BarOnXSdia

Text49.text = Dimensi3.NamaTiang
Text50.text = Dimensi3.Cover

Text61.text = AppliedStress3.XBracedFrame
Text62.text = AppliedStress3.YBracedFrame

''************************************************************

End Sub

Private Function GraphicOne( _
ByVal sizeB As Double, _
ByVal sizeH As Double, _
ByVal BarOnB As Double, _
ByVal BarOnH As Double) As Double

If sizeB = sizeH Then
Shape1.Left = 8100
Shape1.Top = 4180
Shape3.Left = 9230
Shape3.Top = 4180
Shape4.Left = 8100
Shape4.Top = 5290
Shape6.Left = 9230
Shape6.Top = 5290

    Shape2.Left = 8640
    Shape2.Top = 4160
    Shape5.Left = 8640
    Shape5.Top = 5310
    Shape7.Left = 8060
    Shape7.Top = 4696
    Shape8.Left = 9260
    Shape8.Top = 4696
    
Shape9.Left = 7920
Shape9.Width = 1700
Shape9.Top = 4025
Shape9.Height = 1600
Shape10.Left = 8040
Shape10.Width = 1450
Shape10.Top = 4137
Shape10.Height = 1400
End If
''''''''''''''''''''''''
If sizeB < sizeH Then
Shape1.Left = 8320
Shape1.Top = 4160
Shape3.Left = 9000
Shape3.Top = 4160
Shape4.Left = 8320
Shape4.Top = 5300
Shape6.Left = 9000
Shape6.Top = 5300

    Shape2.Left = 8640
    Shape2.Top = 4160
    Shape5.Left = 8640
    Shape5.Top = 5310
    Shape7.Left = 8320
    Shape7.Top = 4696
    Shape8.Left = 9000
    Shape8.Top = 4696
    
 
Shape9.Left = 8160
Shape9.Width = 1200
Shape9.Top = 4025
Shape9.Height = 1600
Shape10.Left = 8290
Shape10.Width = 950
Shape10.Top = 4137
Shape10.Height = 1400
End If
'''''''''''''''''''''''''''''
If sizeB > sizeH Then
Shape1.Left = 8070
Shape1.Top = 4490
Shape3.Left = 9250
Shape3.Top = 4490
Shape4.Left = 8070
Shape4.Top = 4950
Shape6.Left = 9250
Shape6.Top = 4950
    Shape2.Left = 8640
    Shape2.Top = 4490
    Shape5.Left = 8640
    Shape5.Top = 4950
    Shape7.Left = 8060
    Shape7.Top = 4696
    Shape8.Left = 9260
    Shape8.Top = 4696
Shape9.Left = 7920
Shape9.Width = 1700
Shape9.Top = 4361
Shape9.Height = 894
Shape10.Left = 8040
Shape10.Width = 1450
Shape10.Top = 4465
Shape10.Height = 694
End If

End Function

Private Sub EnableText1To62()
Text1.Enabled = True   ''dimensions
 
   
Text2.Enabled = True
Text3.Enabled = True

Text4.Enabled = True   ''base cond.x

Text5.Enabled = True   ''dimensions
Text6.Enabled = True
Text7.Enabled = True
Text8.Enabled = True
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

Text20.Enabled = True  ''moments
Text21.Enabled = True
Text22.Enabled = True
Text23.Enabled = True

Text24.Enabled = True  ''dimensions
Text25.Enabled = True
Text26.Enabled = True
Text27.Enabled = True
Text28.Enabled = True
Text29.Enabled = True
Text30.Enabled = True
Text31.Enabled = True
Text32.Enabled = True
Text33.Enabled = True
Text34.Enabled = True
Text35.Enabled = True
Text36.Enabled = True
Text37.Enabled = True
Text38.Enabled = True

Text39.Enabled = True   ''base cond.y

Text40.Enabled = True   '' axial load

Text41.Enabled = True   ''reinfocements
Text42.Enabled = True
Text43.Enabled = True
Text44.Enabled = True
Text45.Enabled = True
Text46.Enabled = True
Text47.Enabled = True
Text48.Enabled = True

Text49.Enabled = True   ''grid tiang
Text50.Enabled = True   ''cover

Text61.Enabled = True   ''frame braced
Text62.Enabled = True
'''''''''''**********''''''''''''''
End Sub
Private Sub DisableText1To62()
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
Text7.Enabled = False
Text8.Enabled = False
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
Text20.Enabled = False
Text21.Enabled = False
Text22.Enabled = False
Text23.Enabled = False   ''''
Text24.Enabled = False
Text25.Enabled = False
Text26.Enabled = False   ''''
Text27.Enabled = False
Text28.Enabled = False
Text29.Enabled = False
Text30.Enabled = False
Text31.Enabled = False
Text32.Enabled = False
Text33.Enabled = False
Text34.Enabled = False
Text35.Enabled = False
Text36.Enabled = False
Text37.Enabled = False
Text38.Enabled = False
Text39.Enabled = False
Text40.Enabled = False
Text41.Enabled = False
Text42.Enabled = False
Text43.Enabled = False
Text44.Enabled = False
Text45.Enabled = False
Text46.Enabled = False
Text47.Enabled = False
Text48.Enabled = False
Text49.Enabled = False
Text50.Enabled = False


Text51.Enabled = False
Text52.Enabled = False
Text53.Enabled = False
Text54.Enabled = False
Text55.Enabled = False
Text56.Enabled = False
Text57.Enabled = False
Text58.Enabled = False
Text59.Enabled = False
Text60.Enabled = False
Text61.Enabled = False
Text62.Enabled = False

End Sub
Public Function TransferIData(ByVal Nom As Integer) As Integer

UppCoB(Nom) = Text1.text
UppCoH(Nom) = Text2.text
UppCoHgt(Nom) = Text3.text
xBASEresistmnt(Nom) = Text4.text
yUppLeBb(Nom) = Text5.text
yUppLeBh(Nom) = Text6.text
yUppLeBl(Nom) = Text7.text
yUppRiBb(Nom) = Text8.text
yUppRiBh(Nom) = Text9.text
yUppRiBl(Nom) = Text10.text

xUppRiBb(Nom) = Text11.text
xUppRiBh(Nom) = Text12.text
xUppRiBl(Nom) = Text13.text
xUppLeBb(Nom) = Text14.text
xUppLeBh(Nom) = Text15.text
xUppLeBl(Nom) = Text16.text

bTng(Nom) = Text17.text
hTng(Nom) = Text18.text
FloorHgt(Nom) = Text19.text

Mx2(Nom) = Text20.text
My2(Nom) = Text21.text
Mx1(Nom) = Text22.text
My1(Nom) = Text23.text

xLowLeBb(Nom) = Text24.text
xLowLeBh(Nom) = Text25.text
xLowLeBl(Nom) = Text26.text
xLowRiBb(Nom) = Text27.text
xLowRiBh(Nom) = Text28.text
xLowRiBl(Nom) = Text29.text

yLowRiBb(Nom) = Text30.text
yLowRiBh(Nom) = Text31.text
yLowRiBl(Nom) = Text32.text
yLowLeBb(Nom) = Text33.text
yLowLeBh(Nom) = Text34.text
yLowLeBl(Nom) = Text35.text

LowCoB(Nom) = Text36.text
LowCoH(Nom) = Text37.text
LowCoHgt(Nom) = Text38.text

yBASEresistmnt(Nom) = Text39.text
DesgAxial(Nom) = Text40.text

BarYLefN(Nom) = Text41.text
BarYLefD(Nom) = Text42.text
BarYRigN(Nom) = Text43.text
BarYRigD(Nom) = Text44.text

BarXTopN(Nom) = Text45.text
BarXTopD(Nom) = Text46.text
BarXBotN(Nom) = Text47.text
BarXBotD(Nom) = Text48.text

GridTng(Nom) = Text49.text
CoverTng(Nom) = Text50.text

xBracedCol(Nom) = Text61.text
yBracedCol(Nom) = Text62.text
''''''''''''''''''''''''''''''''


End Function

Private Sub Option4_Click()
List1.Clear
Command1.Visible = False
Form1.Picture = LoadPicture(NamaFolder & "icon\datam.ico")
Image1.Picture = LoadPicture(NamaFolder & "icon\ukad3.ico")
Picture1.Enabled = False
  Shape2.FillStyle = 0
  Shape5.FillStyle = 0
  Shape7.FillStyle = 0
  Shape8.FillStyle = 0
        Picture1.Top = 900
        Picture1.Left = 0
        Picture1.Height = 50
        Picture1.Width = 50
        Picture1.Visible = False
        Picture1.Cls
        
bilJenisTng = 4
ReaDFileTiang
ShowAllVisibility
Shape25.Shape = 3
Call GraphicOne(Val(bTng(bilJenisTng)), Val(hTng(bilJenisTng)), _
       Val(BarXBotN(bilJenisTng)) * Val(BarXBotD(bilJenisTng)) ^ 2 * 3.14 / 4, _
       Val(BarXBotD(bilJenisTng)) ^ 2 * 3.14 / 2 + Val(BarYLefN(bilJenisTng)) * Val(BarYLefD(bilJenisTng)) ^ 2 * 3.14 / 4)
         
Dim fnum  As Integer
Dim txtFile, Temp As String
fnum = FreeFile
txtFile = NamaFolder & "tiang\datainput\ColFourGET.txt"
Command2.Enabled = True
Command2.Left = 3110
Command2.Caption = "Tiang - 4"
Command4.Enabled = False
Command5.Enabled = False
If Command1.Enabled = True Then
   Command1.Enabled = False
      End If
      
Option1.Enabled = False
Option2.Enabled = False
Option3.Enabled = False
Option4.Enabled = True
Option5.Enabled = False
Option6.Enabled = False
Option7.Enabled = False
Option8.Enabled = False
Option9.Enabled = False

EnableText1To62


''************************************''


Text1.text = UppCoB(4)
Text2.text = UppCoH(4)
Text3.text = UppCoHgt(4)

Text4.text = xBASEresistmnt(4)

Text5.text = yUppLeBb(4)
Text6.text = yUppLeBh(4)
Text7.text = yUppLeBl(4)
Text8.text = yUppRiBb(4)
Text9.text = yUppRiBh(4)
Text10.text = yUppRiBl(4)
Text11.text = xUppRiBb(4)
Text12.text = xUppRiBh(4)
Text13.text = xUppRiBl(4)
Text14.text = xUppLeBb(4)
Text15.text = xUppLeBh(4)
Text16.text = xUppLeBl(4)

Text17.text = bTng(4)
Text18.text = hTng(4)
Text19.text = FloorHgt(4)

Text20.text = Mx2(4)
Text21.text = My2(4)
Text22.text = Mx1(4)
Text23.text = My1(4)

Text24.text = xLowLeBb(4)
Text25.text = xLowLeBh(4)
Text26.text = xLowLeBl(4)
Text27.text = xLowRiBb(4)
Text28.text = xLowRiBh(4)
Text29.text = xLowRiBl(4)

Text30.text = yLowRiBb(4)
Text31.text = yLowRiBh(4)
Text32.text = yLowRiBl(4)
Text33.text = yLowLeBb(4)
Text34.text = yLowLeBh(4)
Text35.text = yLowLeBl(4)

Text36.text = LowCoB(4)
Text37.text = LowCoH(4)
Text38.text = LowCoHgt(4)

Text39.text = yBASEresistmnt(4)
Text40.text = DesgAxial(4)

Text41.text = BarYLefN(4)
Text42.text = BarYLefD(4)
Text43.text = BarYRigN(4)
Text44.text = BarYRigD(4)

Text45.text = BarXTopN(4)
Text46.text = BarXTopD(4)
Text47.text = BarXBotN(4)
Text48.text = BarXBotD(4)

Text49.text = GridTng(4)
Text50.text = CoverTng(4)

Text61.text = xBracedCol(4)
Text62.text = yBracedCol(4)



''*************************************************************''

Dimensi4.Value1 = Text1.text
Dimensi4.Value2 = Text2.text
Dimensi4.Value3 = Text3.text

AppliedStress4.Value4 = Text4.text

Dimensi4.Value5 = Text5.text
Dimensi4.Value6 = Text6.text
Dimensi4.Value7 = Text7.text
Dimensi4.Value8 = Text8.text
Dimensi4.Value9 = Text9.text
Dimensi4.Value10 = Text10.text
Dimensi4.Value11 = Text11.text
Dimensi4.Value12 = Text12.text
Dimensi4.Value13 = Text13.text
Dimensi4.Value14 = Text14.text
Dimensi4.Value15 = Text15.text
Dimensi4.Value16 = Text16.text

Dimensi4.Value17 = Text17.text
Dimensi4.Value18 = Text18.text
Dimensi4.Value19 = Text19.text

AppliedStress4.Value20 = Text20.text
AppliedStress4.Value21 = Text21.text
AppliedStress4.Value22 = Text22.text
AppliedStress4.Value23 = Text23.text

Dimensi4.Value24 = Text24.text
Dimensi4.Value25 = Text25.text
Dimensi4.Value26 = Text26.text
Dimensi4.Value27 = Text27.text
Dimensi4.Value28 = Text28.text
Dimensi4.Value29 = Text29.text

Dimensi4.Value30 = Text30.text
Dimensi4.Value31 = Text31.text
Dimensi4.Value32 = Text32.text
Dimensi4.Value33 = Text33.text
Dimensi4.Value34 = Text34.text
Dimensi4.Value35 = Text35.text

Dimensi4.Value36 = Text36.text
Dimensi4.Value37 = Text37.text
Dimensi4.Value38 = Text38.text

AppliedStress4.Value39 = Text39.text
AppliedStress4.Value40 = Text40.text

Tetulang4.Value41 = Text41.text
Tetulang4.Value42 = Text42.text
Tetulang4.Value43 = Text43.text
Tetulang4.Value44 = Text44.text

Tetulang4.Value45 = Text45.text
Tetulang4.Value46 = Text46.text
Tetulang4.Value47 = Text47.text
Tetulang4.Value48 = Text48.text

Dimensi4.Value49 = Text49.text
Dimensi4.Value50 = Text50.text

AppliedStress4.Value61 = Text61.text
AppliedStress4.Value62 = Text62.text

''''**********************************************************''''



Text1.text = Dimensi4.BtngAtas
Text2.text = Dimensi4.HtngAtas
Text3.text = Dimensi4.TtngAtas

Text4.text = AppliedStress4.XbaseCondition

Text5.text = Dimensi4.BrskWAtas
Text6.text = Dimensi4.HrskWAtas
Text7.text = Dimensi4.PrskWAtas
Text8.text = Dimensi4.BrskEAtas
Text9.text = Dimensi4.HrskEAtas
Text10.text = Dimensi4.PrskEAtas
Text11.text = Dimensi4.BrskSAtas
Text12.text = Dimensi4.HrskSAtas
Text13.text = Dimensi4.PrskSAtas
Text14.text = Dimensi4.BrskNAtas
Text15.text = Dimensi4.HrskNAtas
Text16.text = Dimensi4.PrskNAtas

Text17.text = Dimensi4.BtianG
Text18.text = Dimensi4.HtianG
Text19.text = Dimensi4.TtianG

Text20.text = AppliedStress4.MomentX22
Text21.text = AppliedStress4.MomentY22
Text22.text = AppliedStress4.MomentX11
Text23.text = AppliedStress4.MomentY11

Text24.text = Dimensi4.BrskNBawah
Text25.text = Dimensi4.HrskNBawah
Text26.text = Dimensi4.PrskNBawah
Text27.text = Dimensi4.BrskSBawah
Text28.text = Dimensi4.HrskSBawah
Text29.text = Dimensi4.PrskSBawah

Text30.text = Dimensi4.BrskEBawah
Text31.text = Dimensi4.HrskEBawah
Text32.text = Dimensi4.PrskEBawah
Text33.text = Dimensi4.BrskWBawah
Text34.text = Dimensi4.HrskWBawah
Text35.text = Dimensi4.PrskWBawah

Text36.text = Dimensi4.BtngBawah
Text37.text = Dimensi4.HtngBawah
Text38.text = Dimensi4.TtngBawah

Text39.text = AppliedStress4.YbaseCondition
Text40.text = AppliedStress4.AxiaLoaD

Text41.text = Tetulang4.BarOnYWno
Text42.text = Tetulang4.BarOnYWdia
Text43.text = Tetulang4.BarOnYEno
Text44.text = Tetulang4.BarOnYEdia

Text45.text = Tetulang4.BarOnXNno
Text46.text = Tetulang4.BarOnXNdia
Text47.text = Tetulang4.BarOnXSno
Text48.text = Tetulang4.BarOnXSdia

Text49.text = Dimensi4.NamaTiang
Text50.text = Dimensi4.Cover

Text61.text = AppliedStress4.XBracedFrame
Text62.text = AppliedStress4.YBracedFrame

''************************************************************

End Sub


Private Sub Option5_Click()
List1.Clear
Command1.Visible = False
Form1.Picture = LoadPicture(NamaFolder & "icon\datam.ico")
Image1.Picture = LoadPicture(NamaFolder & "icon\cskp.ico")
Picture1.Enabled = False
  Shape2.FillStyle = 0
  Shape5.FillStyle = 0
  Shape7.FillStyle = 0
  Shape8.FillStyle = 0
        Picture1.Top = 900
        Picture1.Left = 0
        Picture1.Height = 50
        Picture1.Width = 50
        Picture1.Visible = False
        Picture1.Cls
        
bilJenisTng = 5
ReaDFileTiang
ShowAllVisibility
Shape25.Shape = 3
Call GraphicOne(Val(bTng(bilJenisTng)), Val(hTng(bilJenisTng)), _
       Val(BarXBotN(bilJenisTng)) * Val(BarXBotD(bilJenisTng)) ^ 2 * 3.14 / 4, _
       Val(BarXBotD(bilJenisTng)) ^ 2 * 3.14 / 2 + Val(BarYLefN(bilJenisTng)) * Val(BarYLefD(bilJenisTng)) ^ 2 * 3.14 / 4)
         
Dim fnum As Integer
Dim txtFile, Temp As String
fnum = FreeFile
txtFile = NamaFolder & "tiang\datainput\ColFiveGET.txt"
Command2.Enabled = True
Command2.Left = 4110
Command2.Caption = "Tiang - 5"
Command4.Enabled = False
Command5.Enabled = False
If Command1.Enabled = True Then
   Command1.Enabled = False
      End If

Option1.Enabled = False
Option2.Enabled = False
Option3.Enabled = False
Option4.Enabled = False
Option5.Enabled = True
Option6.Enabled = False
Option7.Enabled = False
Option8.Enabled = False
Option9.Enabled = False

EnableText1To62


''************************************''

Text1.text = UppCoB(5)
Text2.text = UppCoH(5)
Text3.text = UppCoHgt(5)

Text4.text = xBASEresistmnt(5)

Text5.text = yUppLeBb(5)
Text6.text = yUppLeBh(5)
Text7.text = yUppLeBl(5)
Text8.text = yUppRiBb(5)
Text9.text = yUppRiBh(5)
Text10.text = yUppRiBl(5)
Text11.text = xUppRiBb(5)
Text12.text = xUppRiBh(5)
Text13.text = xUppRiBl(5)
Text14.text = xUppLeBb(5)
Text15.text = xUppLeBh(5)
Text16.text = xUppLeBl(5)

Text17.text = bTng(5)
Text18.text = hTng(5)
Text19.text = FloorHgt(5)

Text20.text = Mx2(5)
Text21.text = My2(5)
Text22.text = Mx1(5)
Text23.text = My1(5)

Text24.text = xLowLeBb(5)
Text25.text = xLowLeBh(5)
Text26.text = xLowLeBl(5)
Text27.text = xLowRiBb(5)
Text28.text = xLowRiBh(5)
Text29.text = xLowRiBl(5)

Text30.text = yLowRiBb(5)
Text31.text = yLowRiBh(5)
Text32.text = yLowRiBl(5)
Text33.text = yLowLeBb(5)
Text34.text = yLowLeBh(5)
Text35.text = yLowLeBl(5)

Text36.text = LowCoB(5)
Text37.text = LowCoH(5)
Text38.text = LowCoHgt(5)

Text39.text = yBASEresistmnt(5)
Text40.text = DesgAxial(5)

Text41.text = BarYLefN(5)
Text42.text = BarYLefD(5)
Text43.text = BarYRigN(5)
Text44.text = BarYRigD(5)

Text45.text = BarXTopN(5)
Text46.text = BarXTopD(5)
Text47.text = BarXBotN(5)
Text48.text = BarXBotD(5)

Text49.text = GridTng(5)
Text50.text = CoverTng(5)

Text61.text = xBracedCol(5)
Text62.text = yBracedCol(5)




''*************************************************************''

Dimensi5.Value1 = Text1.text
Dimensi5.Value2 = Text2.text
Dimensi5.Value3 = Text3.text

AppliedStress5.Value4 = Text4.text

Dimensi5.Value5 = Text5.text
Dimensi5.Value6 = Text6.text
Dimensi5.Value7 = Text7.text
Dimensi5.Value8 = Text8.text
Dimensi5.Value9 = Text9.text
Dimensi5.Value10 = Text10.text
Dimensi5.Value11 = Text11.text
Dimensi5.Value12 = Text12.text
Dimensi5.Value13 = Text13.text
Dimensi5.Value14 = Text14.text
Dimensi5.Value15 = Text15.text
Dimensi5.Value16 = Text16.text

Dimensi5.Value17 = Text17.text
Dimensi5.Value18 = Text18.text
Dimensi5.Value19 = Text19.text

AppliedStress5.Value20 = Text20.text
AppliedStress5.Value21 = Text21.text
AppliedStress5.Value22 = Text22.text
AppliedStress5.Value23 = Text23.text

Dimensi5.Value24 = Text24.text
Dimensi5.Value25 = Text25.text
Dimensi5.Value26 = Text26.text
Dimensi5.Value27 = Text27.text
Dimensi5.Value28 = Text28.text
Dimensi5.Value29 = Text29.text

Dimensi5.Value30 = Text30.text
Dimensi5.Value31 = Text31.text
Dimensi5.Value32 = Text32.text
Dimensi5.Value33 = Text33.text
Dimensi5.Value34 = Text34.text
Dimensi5.Value35 = Text35.text

Dimensi5.Value36 = Text36.text
Dimensi5.Value37 = Text37.text
Dimensi5.Value38 = Text38.text

AppliedStress5.Value39 = Text39.text
AppliedStress5.Value40 = Text40.text

Tetulang5.Value41 = Text41.text
Tetulang5.Value42 = Text42.text
Tetulang5.Value43 = Text43.text
Tetulang5.Value44 = Text44.text

Tetulang5.Value45 = Text45.text
Tetulang5.Value46 = Text46.text
Tetulang5.Value47 = Text47.text
Tetulang5.Value48 = Text48.text

Dimensi5.Value49 = Text49.text
Dimensi5.Value50 = Text50.text

AppliedStress5.Value61 = Text61.text
AppliedStress5.Value62 = Text62.text

''''**********************************************************''''



Text1.text = Dimensi5.BtngAtas
Text2.text = Dimensi5.HtngAtas
Text3.text = Dimensi5.TtngAtas

Text4.text = AppliedStress5.XbaseCondition

Text5.text = Dimensi5.BrskWAtas
Text6.text = Dimensi5.HrskWAtas
Text7.text = Dimensi5.PrskWAtas
Text8.text = Dimensi5.BrskEAtas
Text9.text = Dimensi5.HrskEAtas
Text10.text = Dimensi5.PrskEAtas
Text11.text = Dimensi5.BrskSAtas
Text12.text = Dimensi5.HrskSAtas
Text13.text = Dimensi5.PrskSAtas
Text14.text = Dimensi5.BrskNAtas
Text15.text = Dimensi5.HrskNAtas
Text16.text = Dimensi5.PrskNAtas

Text17.text = Dimensi5.BtianG
Text18.text = Dimensi5.HtianG
Text19.text = Dimensi5.TtianG

Text20.text = AppliedStress5.MomentX22
Text21.text = AppliedStress5.MomentY22
Text22.text = AppliedStress5.MomentX11
Text23.text = AppliedStress5.MomentY11

Text24.text = Dimensi5.BrskNBawah
Text25.text = Dimensi5.HrskNBawah
Text26.text = Dimensi5.PrskNBawah
Text27.text = Dimensi5.BrskSBawah
Text28.text = Dimensi5.HrskSBawah
Text29.text = Dimensi5.PrskSBawah

Text30.text = Dimensi5.BrskEBawah
Text31.text = Dimensi5.HrskEBawah
Text32.text = Dimensi5.PrskEBawah
Text33.text = Dimensi5.BrskWBawah
Text34.text = Dimensi5.HrskWBawah
Text35.text = Dimensi5.PrskWBawah

Text36.text = Dimensi5.BtngBawah
Text37.text = Dimensi5.HtngBawah
Text38.text = Dimensi5.TtngBawah

Text39.text = AppliedStress5.YbaseCondition
Text40.text = AppliedStress5.AxiaLoaD

Text41.text = Tetulang5.BarOnYWno
Text42.text = Tetulang5.BarOnYWdia
Text43.text = Tetulang5.BarOnYEno
Text44.text = Tetulang5.BarOnYEdia

Text45.text = Tetulang5.BarOnXNno
Text46.text = Tetulang5.BarOnXNdia
Text47.text = Tetulang5.BarOnXSno
Text48.text = Tetulang5.BarOnXSdia

Text49.text = Dimensi5.NamaTiang
Text50.text = Dimensi5.Cover

Text61.text = AppliedStress5.XBracedFrame
Text62.text = AppliedStress5.YBracedFrame

''************************************************************
End Sub

Private Sub Option6_Click()
List1.Clear
Command1.Visible = False
Form1.Picture = LoadPicture(NamaFolder & "icon\datam.ico")
Image1.Picture = LoadPicture(NamaFolder & "icon\ukad4.ico")
Picture1.Enabled = False
  Shape2.FillStyle = 0
  Shape5.FillStyle = 0
  Shape7.FillStyle = 0
  Shape8.FillStyle = 0
        Picture1.Top = 900
        Picture1.Left = 0
        Picture1.Height = 50
        Picture1.Width = 50
        Picture1.Visible = False
        Picture1.Cls
        
bilJenisTng = 6
ReaDFileTiang
ShowAllVisibility
Shape25.Shape = 3
Call GraphicOne(Val(bTng(bilJenisTng)), Val(hTng(bilJenisTng)), _
       Val(BarXBotN(bilJenisTng)) * Val(BarXBotD(bilJenisTng)) ^ 2 * 3.14 / 4, _
       Val(BarXBotD(bilJenisTng)) ^ 2 * 3.14 / 2 + Val(BarYLefN(bilJenisTng)) * Val(BarYLefD(bilJenisTng)) ^ 2 * 3.14 / 4)
         
Dim fnum As Integer
Dim txtFile, Temp As String
fnum = FreeFile
txtFile = NamaFolder & "tiang\datainput\ColSixGET.txt"
Command2.Enabled = True
Command2.Left = 5110
Command2.Caption = "Tiang - 6"
Command4.Enabled = False
Command5.Enabled = False
If Command1.Enabled = True Then
   Command1.Enabled = False
      End If
      
Option1.Enabled = False
Option2.Enabled = False
Option3.Enabled = False
Option4.Enabled = False
Option5.Enabled = False
Option6.Enabled = True
Option7.Enabled = False
Option8.Enabled = False
Option9.Enabled = False

EnableText1To62


''************************************''


Text1.text = UppCoB(6)
Text2.text = UppCoH(6)
Text3.text = UppCoHgt(6)

Text4.text = xBASEresistmnt(6)

Text5.text = yUppLeBb(6)
Text6.text = yUppLeBh(6)
Text7.text = yUppLeBl(6)
Text8.text = yUppRiBb(6)
Text9.text = yUppRiBh(6)
Text10.text = yUppRiBl(6)
Text11.text = xUppRiBb(6)
Text12.text = xUppRiBh(6)
Text13.text = xUppRiBl(6)
Text14.text = xUppLeBb(6)
Text15.text = xUppLeBh(6)
Text16.text = xUppLeBl(6)

Text17.text = bTng(6)
Text18.text = hTng(6)
Text19.text = FloorHgt(6)

Text20.text = Mx2(6)
Text21.text = My2(6)
Text22.text = Mx1(6)
Text23.text = My1(6)

Text24.text = xLowLeBb(6)
Text25.text = xLowLeBh(6)
Text26.text = xLowLeBl(6)
Text27.text = xLowRiBb(6)
Text28.text = xLowRiBh(6)
Text29.text = xLowRiBl(6)

Text30.text = yLowRiBb(6)
Text31.text = yLowRiBh(6)
Text32.text = yLowRiBl(6)
Text33.text = yLowLeBb(6)
Text34.text = yLowLeBh(6)
Text35.text = yLowLeBl(6)

Text36.text = LowCoB(6)
Text37.text = LowCoH(6)
Text38.text = LowCoHgt(6)

Text39.text = yBASEresistmnt(6)
Text40.text = DesgAxial(6)

Text41.text = BarYLefN(6)
Text42.text = BarYLefD(6)
Text43.text = BarYRigN(6)
Text44.text = BarYRigD(6)

Text45.text = BarXTopN(6)
Text46.text = BarXTopD(6)
Text47.text = BarXBotN(6)
Text48.text = BarXBotD(6)

Text49.text = GridTng(6)
Text50.text = CoverTng(6)

Text61.text = xBracedCol(6)
Text62.text = yBracedCol(6)


''*************************************************************''

Dimensi6.Value1 = Text1.text
Dimensi6.Value2 = Text2.text
Dimensi6.Value3 = Text3.text

AppliedStress6.Value4 = Text4.text

Dimensi6.Value5 = Text5.text
Dimensi6.Value6 = Text6.text
Dimensi6.Value7 = Text7.text
Dimensi6.Value8 = Text8.text
Dimensi6.Value9 = Text9.text
Dimensi6.Value10 = Text10.text
Dimensi6.Value11 = Text11.text
Dimensi6.Value12 = Text12.text
Dimensi6.Value13 = Text13.text
Dimensi6.Value14 = Text14.text
Dimensi6.Value15 = Text15.text
Dimensi6.Value16 = Text16.text

Dimensi6.Value17 = Text17.text
Dimensi6.Value18 = Text18.text
Dimensi6.Value19 = Text19.text

AppliedStress6.Value20 = Text20.text
AppliedStress6.Value21 = Text21.text
AppliedStress6.Value22 = Text22.text
AppliedStress6.Value23 = Text23.text

Dimensi6.Value24 = Text24.text
Dimensi6.Value25 = Text25.text
Dimensi6.Value26 = Text26.text
Dimensi6.Value27 = Text27.text
Dimensi6.Value28 = Text28.text
Dimensi6.Value29 = Text29.text

Dimensi6.Value30 = Text30.text
Dimensi6.Value31 = Text31.text
Dimensi6.Value32 = Text32.text
Dimensi6.Value33 = Text33.text
Dimensi6.Value34 = Text34.text
Dimensi6.Value35 = Text35.text

Dimensi6.Value36 = Text36.text
Dimensi6.Value37 = Text37.text
Dimensi6.Value38 = Text38.text

AppliedStress6.Value39 = Text39.text
AppliedStress6.Value40 = Text40.text

Tetulang6.Value41 = Text41.text
Tetulang6.Value42 = Text42.text
Tetulang6.Value43 = Text43.text
Tetulang6.Value44 = Text44.text

Tetulang6.Value45 = Text45.text
Tetulang6.Value46 = Text46.text
Tetulang6.Value47 = Text47.text
Tetulang6.Value48 = Text48.text

Dimensi6.Value49 = Text49.text
Dimensi6.Value50 = Text50.text

AppliedStress6.Value61 = Text61.text
AppliedStress6.Value62 = Text62.text

''''**********************************************************''''



Text1.text = Dimensi6.BtngAtas
Text2.text = Dimensi6.HtngAtas
Text3.text = Dimensi6.TtngAtas

Text4.text = AppliedStress6.XbaseCondition

Text5.text = Dimensi6.BrskWAtas
Text6.text = Dimensi6.HrskWAtas
Text7.text = Dimensi6.PrskWAtas
Text8.text = Dimensi6.BrskEAtas
Text9.text = Dimensi6.HrskEAtas
Text10.text = Dimensi6.PrskEAtas
Text11.text = Dimensi6.BrskSAtas
Text12.text = Dimensi6.HrskSAtas
Text13.text = Dimensi6.PrskSAtas
Text14.text = Dimensi6.BrskNAtas
Text15.text = Dimensi6.HrskNAtas
Text16.text = Dimensi6.PrskNAtas

Text17.text = Dimensi6.BtianG
Text18.text = Dimensi6.HtianG
Text19.text = Dimensi6.TtianG

Text20.text = AppliedStress6.MomentX22
Text21.text = AppliedStress6.MomentY22
Text22.text = AppliedStress6.MomentX11
Text23.text = AppliedStress6.MomentY11

Text24.text = Dimensi6.BrskNBawah
Text25.text = Dimensi6.HrskNBawah
Text26.text = Dimensi6.PrskNBawah
Text27.text = Dimensi6.BrskSBawah
Text28.text = Dimensi6.HrskSBawah
Text29.text = Dimensi6.PrskSBawah

Text30.text = Dimensi6.BrskEBawah
Text31.text = Dimensi6.HrskEBawah
Text32.text = Dimensi6.PrskEBawah
Text33.text = Dimensi6.BrskWBawah
Text34.text = Dimensi6.HrskWBawah
Text35.text = Dimensi6.PrskWBawah

Text36.text = Dimensi6.BtngBawah
Text37.text = Dimensi6.HtngBawah
Text38.text = Dimensi6.TtngBawah

Text39.text = AppliedStress6.YbaseCondition
Text40.text = AppliedStress6.AxiaLoaD

Text41.text = Tetulang6.BarOnYWno
Text42.text = Tetulang6.BarOnYWdia
Text43.text = Tetulang6.BarOnYEno
Text44.text = Tetulang6.BarOnYEdia

Text45.text = Tetulang6.BarOnXNno
Text46.text = Tetulang6.BarOnXNdia
Text47.text = Tetulang6.BarOnXSno
Text48.text = Tetulang6.BarOnXSdia

Text49.text = Dimensi6.NamaTiang
Text50.text = Dimensi6.Cover

Text61.text = AppliedStress6.XBracedFrame
Text62.text = AppliedStress6.YBracedFrame

''************************************************************

End Sub


Private Sub Option7_Click()
List1.Clear
Command1.Visible = False
Form1.Picture = LoadPicture(NamaFolder & "icon\datam.ico")
Image1.Picture = LoadPicture(NamaFolder & "icon\ukad3.ico")
Picture1.Enabled = False
  Shape2.FillStyle = 0
  Shape5.FillStyle = 0
  Shape7.FillStyle = 0
  Shape8.FillStyle = 0
        Picture1.Top = 900
        Picture1.Left = 0
        Picture1.Height = 50
        Picture1.Width = 50
        Picture1.Visible = False
        Picture1.Cls
        
bilJenisTng = 7
ReaDFileTiang
ShowAllVisibility
Shape25.Shape = 3
Call GraphicOne(Val(bTng(bilJenisTng)), Val(hTng(bilJenisTng)), _
       Val(BarXBotN(bilJenisTng)) * Val(BarXBotD(bilJenisTng)) ^ 2 * 3.14 / 4, _
       Val(BarXBotD(bilJenisTng)) ^ 2 * 3.14 / 2 + Val(BarYLefN(bilJenisTng)) * Val(BarYLefD(bilJenisTng)) ^ 2 * 3.14 / 4)
         
Dim fnum As Integer
Dim txtFile, Temp As String
fnum = FreeFile
txtFile = NamaFolder & "tiang\datainput\ColSevenGET.txt"
Command2.Enabled = True
Command2.Left = 6110
Command2.Caption = "Tiang - 7"
Command4.Enabled = False
Command5.Enabled = False
If Command1.Enabled = True Then
   Command1.Enabled = False
      End If

Option1.Enabled = False
Option2.Enabled = False
Option3.Enabled = False
Option4.Enabled = False
Option5.Enabled = False
Option6.Enabled = False
Option7.Enabled = True
Option8.Enabled = False
Option9.Enabled = False

EnableText1To62


''************************************''


Text1.text = UppCoB(7)
Text2.text = UppCoH(7)
Text3.text = UppCoHgt(7)

Text4.text = xBASEresistmnt(7)

Text5.text = yUppLeBb(7)
Text6.text = yUppLeBh(7)
Text7.text = yUppLeBl(7)
Text8.text = yUppRiBb(7)
Text9.text = yUppRiBh(7)
Text10.text = yUppRiBl(7)
Text11.text = xUppRiBb(7)
Text12.text = xUppRiBh(7)
Text13.text = xUppRiBl(7)
Text14.text = xUppLeBb(7)
Text15.text = xUppLeBh(7)
Text16.text = xUppLeBl(7)

Text17.text = bTng(7)
Text18.text = hTng(7)
Text19.text = FloorHgt(7)

Text20.text = Mx2(7)
Text21.text = My2(7)
Text22.text = Mx1(7)
Text23.text = My1(7)

Text24.text = xLowLeBb(7)
Text25.text = xLowLeBh(7)
Text26.text = xLowLeBl(7)
Text27.text = xLowRiBb(7)
Text28.text = xLowRiBh(7)
Text29.text = xLowRiBl(7)

Text30.text = yLowRiBb(7)
Text31.text = yLowRiBh(7)
Text32.text = yLowRiBl(7)
Text33.text = yLowLeBb(7)
Text34.text = yLowLeBh(7)
Text35.text = yLowLeBl(7)

Text36.text = LowCoB(7)
Text37.text = LowCoH(7)
Text38.text = LowCoHgt(7)

Text39.text = yBASEresistmnt(7)
Text40.text = DesgAxial(7)

Text41.text = BarYLefN(7)
Text42.text = BarYLefD(7)
Text43.text = BarYRigN(7)
Text44.text = BarYRigD(7)

Text45.text = BarXTopN(7)
Text46.text = BarXTopD(7)
Text47.text = BarXBotN(7)
Text48.text = BarXBotD(7)

Text49.text = GridTng(7)
Text50.text = CoverTng(7)

Text61.text = xBracedCol(7)
Text62.text = yBracedCol(7)


''*************************************************************''

Dimensi7.Value1 = Text1.text
Dimensi7.Value2 = Text2.text
Dimensi7.Value3 = Text3.text

AppliedStress7.Value4 = Text4.text

Dimensi7.Value5 = Text5.text
Dimensi7.Value6 = Text6.text
Dimensi7.Value7 = Text7.text
Dimensi7.Value8 = Text8.text
Dimensi7.Value9 = Text9.text
Dimensi7.Value10 = Text10.text
Dimensi7.Value11 = Text11.text
Dimensi7.Value12 = Text12.text
Dimensi7.Value13 = Text13.text
Dimensi7.Value14 = Text14.text
Dimensi7.Value15 = Text15.text
Dimensi7.Value16 = Text16.text

Dimensi7.Value17 = Text17.text
Dimensi7.Value18 = Text18.text
Dimensi7.Value19 = Text19.text

AppliedStress7.Value20 = Text20.text
AppliedStress7.Value21 = Text21.text
AppliedStress7.Value22 = Text22.text
AppliedStress7.Value23 = Text23.text

Dimensi7.Value24 = Text24.text
Dimensi7.Value25 = Text25.text
Dimensi7.Value26 = Text26.text
Dimensi7.Value27 = Text27.text
Dimensi7.Value28 = Text28.text
Dimensi7.Value29 = Text29.text

Dimensi7.Value30 = Text30.text
Dimensi7.Value31 = Text31.text
Dimensi7.Value32 = Text32.text
Dimensi7.Value33 = Text33.text
Dimensi7.Value34 = Text34.text
Dimensi7.Value35 = Text35.text

Dimensi7.Value36 = Text36.text
Dimensi7.Value37 = Text37.text
Dimensi7.Value38 = Text38.text

AppliedStress7.Value39 = Text39.text
AppliedStress7.Value40 = Text40.text

Tetulang7.Value41 = Text41.text
Tetulang7.Value42 = Text42.text
Tetulang7.Value43 = Text43.text
Tetulang7.Value44 = Text44.text

Tetulang7.Value45 = Text45.text
Tetulang7.Value46 = Text46.text
Tetulang7.Value47 = Text47.text
Tetulang7.Value48 = Text48.text

Dimensi7.Value49 = Text49.text
Dimensi7.Value50 = Text50.text

AppliedStress7.Value61 = Text61.text
AppliedStress7.Value62 = Text62.text

''''**********************************************************''''



Text1.text = Dimensi7.BtngAtas
Text2.text = Dimensi7.HtngAtas
Text3.text = Dimensi7.TtngAtas

Text4.text = AppliedStress7.XbaseCondition

Text5.text = Dimensi7.BrskWAtas
Text6.text = Dimensi7.HrskWAtas
Text7.text = Dimensi7.PrskWAtas
Text8.text = Dimensi7.BrskEAtas
Text9.text = Dimensi7.HrskEAtas
Text10.text = Dimensi7.PrskEAtas
Text11.text = Dimensi7.BrskSAtas
Text12.text = Dimensi7.HrskSAtas
Text13.text = Dimensi7.PrskSAtas
Text14.text = Dimensi7.BrskNAtas
Text15.text = Dimensi7.HrskNAtas
Text16.text = Dimensi7.PrskNAtas

Text17.text = Dimensi7.BtianG
Text18.text = Dimensi7.HtianG
Text19.text = Dimensi7.TtianG

Text20.text = AppliedStress7.MomentX22
Text21.text = AppliedStress7.MomentY22
Text22.text = AppliedStress7.MomentX11
Text23.text = AppliedStress7.MomentY11

Text24.text = Dimensi7.BrskNBawah
Text25.text = Dimensi7.HrskNBawah
Text26.text = Dimensi7.PrskNBawah
Text27.text = Dimensi7.BrskSBawah
Text28.text = Dimensi7.HrskSBawah
Text29.text = Dimensi7.PrskSBawah

Text30.text = Dimensi7.BrskEBawah
Text31.text = Dimensi7.HrskEBawah
Text32.text = Dimensi7.PrskEBawah
Text33.text = Dimensi7.BrskWBawah
Text34.text = Dimensi7.HrskWBawah
Text35.text = Dimensi7.PrskWBawah

Text36.text = Dimensi7.BtngBawah
Text37.text = Dimensi7.HtngBawah
Text38.text = Dimensi7.TtngBawah

Text39.text = AppliedStress7.YbaseCondition
Text40.text = AppliedStress7.AxiaLoaD

Text41.text = Tetulang7.BarOnYWno
Text42.text = Tetulang7.BarOnYWdia
Text43.text = Tetulang7.BarOnYEno
Text44.text = Tetulang7.BarOnYEdia

Text45.text = Tetulang7.BarOnXNno
Text46.text = Tetulang7.BarOnXNdia
Text47.text = Tetulang7.BarOnXSno
Text48.text = Tetulang7.BarOnXSdia

Text49.text = Dimensi7.NamaTiang
Text50.text = Dimensi7.Cover

Text61.text = AppliedStress7.XBracedFrame
Text62.text = AppliedStress7.YBracedFrame

''************************************************************
End Sub

Private Sub Option8_Click()
List1.Clear
Command1.Visible = False
Form1.Picture = LoadPicture(NamaFolder & "icon\datam.ico")
Image1.Picture = LoadPicture(NamaFolder & "icon\cskp.ico")
Picture1.Enabled = False
  Shape2.FillStyle = 0
  Shape5.FillStyle = 0
  Shape7.FillStyle = 0
  Shape8.FillStyle = 0
        Picture1.Top = 900
        Picture1.Left = 0
        Picture1.Height = 50
        Picture1.Width = 50
        Picture1.Visible = False
        Picture1.Cls
        
bilJenisTng = 8
ReaDFileTiang
ShowAllVisibility
Shape25.Shape = 3
Call GraphicOne(Val(bTng(bilJenisTng)), Val(hTng(bilJenisTng)), _
       Val(BarXBotN(bilJenisTng)) * Val(BarXBotD(bilJenisTng)) ^ 2 * 3.14 / 4, _
       Val(BarXBotD(bilJenisTng)) ^ 2 * 3.14 / 2 + Val(BarYLefN(bilJenisTng)) * Val(BarYLefD(bilJenisTng)) ^ 2 * 3.14 / 4)
         
Dim fnum As Integer
Dim txtFile, Temp As String
fnum = FreeFile
txtFile = NamaFolder & "tiang\datainput\ColEightGET.txt"
Command2.Enabled = True
Command2.Left = 7110
Command2.Caption = "Tiang - 8"
Command4.Enabled = False
Command5.Enabled = False
If Command1.Enabled = True Then
   Command1.Enabled = False
      End If

Option1.Enabled = False
Option2.Enabled = False
Option3.Enabled = False
Option4.Enabled = False
Option5.Enabled = False
Option6.Enabled = False
Option7.Enabled = False
Option8.Enabled = True
Option9.Enabled = False

EnableText1To62


''************************************''


Text1.text = UppCoB(8)
Text2.text = UppCoH(8)
Text3.text = UppCoHgt(8)

Text4.text = xBASEresistmnt(8)

Text5.text = yUppLeBb(8)
Text6.text = yUppLeBh(8)
Text7.text = yUppLeBl(8)
Text8.text = yUppRiBb(8)
Text9.text = yUppRiBh(8)
Text10.text = yUppRiBl(8)
Text11.text = xUppRiBb(8)
Text12.text = xUppRiBh(8)
Text13.text = xUppRiBl(8)
Text14.text = xUppLeBb(8)
Text15.text = xUppLeBh(8)
Text16.text = xUppLeBl(8)

Text17.text = bTng(8)
Text18.text = hTng(8)
Text19.text = FloorHgt(8)

Text20.text = Mx2(8)
Text21.text = My2(8)
Text22.text = Mx1(8)
Text23.text = My1(8)

Text24.text = xLowLeBb(8)
Text25.text = xLowLeBh(8)
Text26.text = xLowLeBl(8)
Text27.text = xLowRiBb(8)
Text28.text = xLowRiBh(8)
Text29.text = xLowRiBl(8)

Text30.text = yLowRiBb(8)
Text31.text = yLowRiBh(8)
Text32.text = yLowRiBl(8)
Text33.text = yLowLeBb(8)
Text34.text = yLowLeBh(8)
Text35.text = yLowLeBl(8)

Text36.text = LowCoB(8)
Text37.text = LowCoH(8)
Text38.text = LowCoHgt(8)

Text39.text = yBASEresistmnt(8)
Text40.text = DesgAxial(8)

Text41.text = BarYLefN(8)
Text42.text = BarYLefD(8)
Text43.text = BarYRigN(8)
Text44.text = BarYRigD(8)

Text45.text = BarXTopN(8)
Text46.text = BarXTopD(8)
Text47.text = BarXBotN(8)
Text48.text = BarXBotD(8)

Text49.text = GridTng(8)
Text50.text = CoverTng(8)

Text61.text = xBracedCol(8)
Text62.text = yBracedCol(8)


''*************************************************************''

Dimensi8.Value1 = Text1.text
Dimensi8.Value2 = Text2.text
Dimensi8.Value3 = Text3.text

AppliedStress8.Value4 = Text4.text

Dimensi8.Value5 = Text5.text
Dimensi8.Value6 = Text6.text
Dimensi8.Value7 = Text7.text
Dimensi8.Value8 = Text8.text
Dimensi8.Value9 = Text9.text
Dimensi8.Value10 = Text10.text
Dimensi8.Value11 = Text11.text
Dimensi8.Value12 = Text12.text
Dimensi8.Value13 = Text13.text
Dimensi8.Value14 = Text14.text
Dimensi8.Value15 = Text15.text
Dimensi8.Value16 = Text16.text

Dimensi8.Value17 = Text17.text
Dimensi8.Value18 = Text18.text
Dimensi8.Value19 = Text19.text

AppliedStress8.Value20 = Text20.text
AppliedStress8.Value21 = Text21.text
AppliedStress8.Value22 = Text22.text
AppliedStress8.Value23 = Text23.text

Dimensi8.Value24 = Text24.text
Dimensi8.Value25 = Text25.text
Dimensi8.Value26 = Text26.text
Dimensi8.Value27 = Text27.text
Dimensi8.Value28 = Text28.text
Dimensi8.Value29 = Text29.text

Dimensi8.Value30 = Text30.text
Dimensi8.Value31 = Text31.text
Dimensi8.Value32 = Text32.text
Dimensi8.Value33 = Text33.text
Dimensi8.Value34 = Text34.text
Dimensi8.Value35 = Text35.text

Dimensi8.Value36 = Text36.text
Dimensi8.Value37 = Text37.text
Dimensi8.Value38 = Text38.text

AppliedStress8.Value39 = Text39.text
AppliedStress8.Value40 = Text40.text

Tetulang8.Value41 = Text41.text
Tetulang8.Value42 = Text42.text
Tetulang8.Value43 = Text43.text
Tetulang8.Value44 = Text44.text

Tetulang8.Value45 = Text45.text
Tetulang8.Value46 = Text46.text
Tetulang8.Value47 = Text47.text
Tetulang8.Value48 = Text48.text

Dimensi8.Value49 = Text49.text
Dimensi8.Value50 = Text50.text

AppliedStress8.Value61 = Text61.text
AppliedStress8.Value62 = Text62.text

''''**********************************************************''''



Text1.text = Dimensi8.BtngAtas
Text2.text = Dimensi8.HtngAtas
Text3.text = Dimensi8.TtngAtas

Text4.text = AppliedStress8.XbaseCondition

Text5.text = Dimensi8.BrskWAtas
Text6.text = Dimensi8.HrskWAtas
Text7.text = Dimensi8.PrskWAtas
Text8.text = Dimensi8.BrskEAtas
Text9.text = Dimensi8.HrskEAtas
Text10.text = Dimensi8.PrskEAtas
Text11.text = Dimensi8.BrskSAtas
Text12.text = Dimensi8.HrskSAtas
Text13.text = Dimensi8.PrskSAtas
Text14.text = Dimensi8.BrskNAtas
Text15.text = Dimensi8.HrskNAtas
Text16.text = Dimensi8.PrskNAtas

Text17.text = Dimensi8.BtianG
Text18.text = Dimensi8.HtianG
Text19.text = Dimensi8.TtianG

Text20.text = AppliedStress8.MomentX22
Text21.text = AppliedStress8.MomentY22
Text22.text = AppliedStress8.MomentX11
Text23.text = AppliedStress8.MomentY11

Text24.text = Dimensi8.BrskNBawah
Text25.text = Dimensi8.HrskNBawah
Text26.text = Dimensi8.PrskNBawah
Text27.text = Dimensi8.BrskSBawah
Text28.text = Dimensi8.HrskSBawah
Text29.text = Dimensi8.PrskSBawah

Text30.text = Dimensi8.BrskEBawah
Text31.text = Dimensi8.HrskEBawah
Text32.text = Dimensi8.PrskEBawah
Text33.text = Dimensi8.BrskWBawah
Text34.text = Dimensi8.HrskWBawah
Text35.text = Dimensi8.PrskWBawah

Text36.text = Dimensi8.BtngBawah
Text37.text = Dimensi8.HtngBawah
Text38.text = Dimensi8.TtngBawah

Text39.text = AppliedStress8.YbaseCondition
Text40.text = AppliedStress8.AxiaLoaD

Text41.text = Tetulang8.BarOnYWno
Text42.text = Tetulang8.BarOnYWdia
Text43.text = Tetulang8.BarOnYEno
Text44.text = Tetulang8.BarOnYEdia

Text45.text = Tetulang8.BarOnXNno
Text46.text = Tetulang8.BarOnXNdia
Text47.text = Tetulang8.BarOnXSno
Text48.text = Tetulang8.BarOnXSdia

Text49.text = Dimensi8.NamaTiang
Text50.text = Dimensi8.Cover

Text61.text = AppliedStress8.XBracedFrame
Text62.text = AppliedStress8.YBracedFrame

''***************************************
End Sub

Private Sub Option9_Click()
List1.Clear
Command1.Visible = False
Form1.Picture = LoadPicture(NamaFolder & "icon\datam.ico")
Image1.Picture = LoadPicture(NamaFolder & "icon\ukad4.ico")
Picture1.Enabled = False
  Shape2.FillStyle = 0
  Shape5.FillStyle = 0
  Shape7.FillStyle = 0
  Shape8.FillStyle = 0
        Picture1.Top = 900
        Picture1.Left = 0
        Picture1.Height = 50
        Picture1.Width = 50
        Picture1.Visible = False
        Picture1.Cls
        
bilJenisTng = 9
ReaDFileTiang
ShowAllVisibility
Shape25.Shape = 3
Call GraphicOne(Val(bTng(bilJenisTng)), Val(hTng(bilJenisTng)), _
       Val(BarXBotN(bilJenisTng)) * Val(BarXBotD(bilJenisTng)) ^ 2 * 3.14 / 4, _
       Val(BarXBotD(bilJenisTng)) ^ 2 * 3.14 / 2 + Val(BarYLefN(bilJenisTng)) * Val(BarYLefD(bilJenisTng)) ^ 2 * 3.14 / 4)
         
Dim fnum As Integer
Dim txtFile, Temp As String
fnum = FreeFile
txtFile = NamaFolder & "tiang\datainput\ColNineGET.txt"
Command2.Enabled = True
Command2.Left = 8110
Command2.Caption = "Tiang - 9"
Command4.Enabled = False
Command5.Enabled = False
If Command1.Enabled = True Then
   Command1.Enabled = False
      End If

Option1.Enabled = False
Option2.Enabled = False
Option3.Enabled = False
Option4.Enabled = False
Option5.Enabled = False
Option6.Enabled = False
Option7.Enabled = False
Option8.Enabled = False
Option9.Enabled = True

EnableText1To62


''************************************''


Text1.text = UppCoB(9)
Text2.text = UppCoH(9)
Text3.text = UppCoHgt(9)

Text4.text = xBASEresistmnt(9)

Text5.text = yUppLeBb(9)
Text6.text = yUppLeBh(9)
Text7.text = yUppLeBl(9)
Text8.text = yUppRiBb(9)
Text9.text = yUppRiBh(9)
Text10.text = yUppRiBl(9)
Text11.text = xUppRiBb(9)
Text12.text = xUppRiBh(9)
Text13.text = xUppRiBl(9)
Text14.text = xUppLeBb(9)
Text15.text = xUppLeBh(9)
Text16.text = xUppLeBl(9)

Text17.text = bTng(9)
Text18.text = hTng(9)
Text19.text = FloorHgt(9)

Text20.text = Mx2(9)
Text21.text = My2(9)
Text22.text = Mx1(9)
Text23.text = My1(9)

Text24.text = xLowLeBb(9)
Text25.text = xLowLeBh(9)
Text26.text = xLowLeBl(9)
Text27.text = xLowRiBb(9)
Text28.text = xLowRiBh(9)
Text29.text = xLowRiBl(9)

Text30.text = yLowRiBb(9)
Text31.text = yLowRiBh(9)
Text32.text = yLowRiBl(9)
Text33.text = yLowLeBb(9)
Text34.text = yLowLeBh(9)
Text35.text = yLowLeBl(9)

Text36.text = LowCoB(9)
Text37.text = LowCoH(9)
Text38.text = LowCoHgt(9)

Text39.text = yBASEresistmnt(9)
Text40.text = DesgAxial(9)

Text41.text = BarYLefN(9)
Text42.text = BarYLefD(9)
Text43.text = BarYRigN(9)
Text44.text = BarYRigD(9)

Text45.text = BarXTopN(9)
Text46.text = BarXTopD(9)
Text47.text = BarXBotN(9)
Text48.text = BarXBotD(9)

Text49.text = GridTng(9)
Text50.text = CoverTng(9)

Text61.text = xBracedCol(9)
Text62.text = yBracedCol(9)


''*************************************************************''

Dimensi9.Value1 = Text1.text
Dimensi9.Value2 = Text2.text
Dimensi9.Value3 = Text3.text

AppliedStress9.Value4 = Text4.text

Dimensi9.Value5 = Text5.text
Dimensi9.Value6 = Text6.text
Dimensi9.Value7 = Text7.text
Dimensi9.Value8 = Text8.text
Dimensi9.Value9 = Text9.text
Dimensi9.Value10 = Text10.text
Dimensi9.Value11 = Text11.text
Dimensi9.Value12 = Text12.text
Dimensi9.Value13 = Text13.text
Dimensi9.Value14 = Text14.text
Dimensi9.Value15 = Text15.text
Dimensi9.Value16 = Text16.text

Dimensi9.Value17 = Text17.text
Dimensi9.Value18 = Text18.text
Dimensi9.Value19 = Text19.text

AppliedStress9.Value20 = Text20.text
AppliedStress9.Value21 = Text21.text
AppliedStress9.Value22 = Text22.text
AppliedStress9.Value23 = Text23.text

Dimensi9.Value24 = Text24.text
Dimensi9.Value25 = Text25.text
Dimensi9.Value26 = Text26.text
Dimensi9.Value27 = Text27.text
Dimensi9.Value28 = Text28.text
Dimensi9.Value29 = Text29.text

Dimensi9.Value30 = Text30.text
Dimensi9.Value31 = Text31.text
Dimensi9.Value32 = Text32.text
Dimensi9.Value33 = Text33.text
Dimensi9.Value34 = Text34.text
Dimensi9.Value35 = Text35.text

Dimensi9.Value36 = Text36.text
Dimensi9.Value37 = Text37.text
Dimensi9.Value38 = Text38.text

AppliedStress9.Value39 = Text39.text
AppliedStress9.Value40 = Text40.text

Tetulang9.Value41 = Text41.text
Tetulang9.Value42 = Text42.text
Tetulang9.Value43 = Text43.text
Tetulang9.Value44 = Text44.text

Tetulang9.Value45 = Text45.text
Tetulang9.Value46 = Text46.text
Tetulang9.Value47 = Text47.text
Tetulang9.Value48 = Text48.text

Dimensi9.Value49 = Text49.text
Dimensi9.Value50 = Text50.text

AppliedStress9.Value61 = Text61.text
AppliedStress9.Value62 = Text62.text

''''**********************************************************''''



Text1.text = Dimensi9.BtngAtas
Text2.text = Dimensi9.HtngAtas
Text3.text = Dimensi9.TtngAtas

Text4.text = AppliedStress9.XbaseCondition

Text5.text = Dimensi9.BrskWAtas
Text6.text = Dimensi9.HrskWAtas
Text7.text = Dimensi9.PrskWAtas
Text8.text = Dimensi9.BrskEAtas
Text9.text = Dimensi9.HrskEAtas
Text10.text = Dimensi9.PrskEAtas
Text11.text = Dimensi9.BrskSAtas
Text12.text = Dimensi9.HrskSAtas
Text13.text = Dimensi9.PrskSAtas
Text14.text = Dimensi9.BrskNAtas
Text15.text = Dimensi9.HrskNAtas
Text16.text = Dimensi9.PrskNAtas

Text17.text = Dimensi9.BtianG
Text18.text = Dimensi9.HtianG
Text19.text = Dimensi9.TtianG

Text20.text = AppliedStress9.MomentX22
Text21.text = AppliedStress9.MomentY22
Text22.text = AppliedStress9.MomentX11
Text23.text = AppliedStress9.MomentY11

Text24.text = Dimensi9.BrskNBawah
Text25.text = Dimensi9.HrskNBawah
Text26.text = Dimensi9.PrskNBawah
Text27.text = Dimensi9.BrskSBawah
Text28.text = Dimensi9.HrskSBawah
Text29.text = Dimensi9.PrskSBawah

Text30.text = Dimensi9.BrskEBawah
Text31.text = Dimensi9.HrskEBawah
Text32.text = Dimensi9.PrskEBawah
Text33.text = Dimensi9.BrskWBawah
Text34.text = Dimensi9.HrskWBawah
Text35.text = Dimensi9.PrskWBawah

Text36.text = Dimensi9.BtngBawah
Text37.text = Dimensi9.HtngBawah
Text38.text = Dimensi9.TtngBawah

Text39.text = AppliedStress9.YbaseCondition
Text40.text = AppliedStress9.AxiaLoaD

Text41.text = Tetulang9.BarOnYWno
Text42.text = Tetulang9.BarOnYWdia
Text43.text = Tetulang9.BarOnYEno
Text44.text = Tetulang9.BarOnYEdia

Text45.text = Tetulang9.BarOnXNno
Text46.text = Tetulang9.BarOnXNdia
Text47.text = Tetulang9.BarOnXSno
Text48.text = Tetulang9.BarOnXSdia

Text49.text = Dimensi9.NamaTiang
Text50.text = Dimensi9.Cover

Text61.text = AppliedStress9.XBracedFrame
Text62.text = AppliedStress9.YBracedFrame

''***************************************
End Sub

Private Sub Text1_Change()
Dim Txt As String
If Left(Text1.text, 1) = "A" Or Left(Text1.text, 1) = "a" Or _
Left(Text1.text, 1) = "S" Or Left(Text1.text, 1) = "s" Or _
Left(Text1.text, 1) = "D" Or Left(Text1.text, 1) = "d" Or _
Left(Text1.text, 1) = "F" Or Left(Text1.text, 1) = "f" Or _
Left(Text1.text, 1) = "G" Or Left(Text1.text, 1) = "g" Or _
Left(Text1.text, 1) = "H" Or Left(Text1.text, 1) = "h" Or _
Left(Text1.text, 1) = "J" Or Left(Text1.text, 1) = "j" Or _
Left(Text1.text, 1) = "K" Or Left(Text1.text, 1) = "k" Or _
Left(Text1.text, 1) = "L" Or Left(Text1.text, 1) = "l" Then
  Txt = SetStandardColumnDepth(Left(Text1.text, 1))
    Text1.text = Txt
      End If
      Text1.text = Val((Text1.text))
If Val(Text1.text) <= 0 Then
   Text1.text = 0
   End If
End Sub

Private Sub Text1_GotFocus()
Shape23.BorderColor = vbGreen
   Shape23.BorderWidth = 2
   
End Sub

Private Sub Text1_LostFocus()
Shape23.BorderColor = vbBlack
   Shape23.BorderWidth = 1
   
End Sub

Private Sub Text10_Change()
Dim Txt As String
If Left(Text10.text, 1) = "A" Or Left(Text10.text, 1) = "a" Or _
Left(Text10.text, 1) = "S" Or Left(Text10.text, 1) = "s" Or _
Left(Text10.text, 1) = "D" Or Left(Text10.text, 1) = "d" Or _
Left(Text10.text, 1) = "F" Or Left(Text10.text, 1) = "f" Or _
Left(Text10.text, 1) = "G" Or Left(Text10.text, 1) = "g" Or _
Left(Text10.text, 1) = "H" Or Left(Text10.text, 1) = "h" Or _
Left(Text10.text, 1) = "J" Or Left(Text10.text, 1) = "j" Or _
Left(Text10.text, 1) = "K" Or Left(Text10.text, 1) = "k" Or _
Left(Text10.text, 1) = "L" Or Left(Text10.text, 1) = "l" Then
  Txt = SetStandardBeamLength(Left(Text10.text, 1))
    Text10.text = Txt
      End If
      Text10.text = Val((Text10.text))
If Val(Text10.text) <= 0 Then
   Text10.text = 0
   End If
End Sub

Private Sub Text10_GotFocus()
Line18.BorderColor = vbCyan
   Line18.BorderWidth = 2
   Shape19.BorderColor = vbCyan
   Shape19.FillStyle = 0
   Shape19.BorderWidth = 2
End Sub

Private Sub Text10_LostFocus()
Line18.BorderColor = vbBlack
   Line18.BorderWidth = 1
   Shape19.BorderColor = vbBlack
   Shape19.FillStyle = 1
   Shape19.BorderWidth = 1
End Sub

Private Sub Text11_Change()
Dim Txt As String
If Left(Text11.text, 1) = "A" Or Left(Text11.text, 1) = "a" Or _
Left(Text11.text, 1) = "S" Or Left(Text11.text, 1) = "s" Or _
Left(Text11.text, 1) = "D" Or Left(Text11.text, 1) = "d" Or _
Left(Text11.text, 1) = "F" Or Left(Text11.text, 1) = "f" Or _
Left(Text11.text, 1) = "G" Or Left(Text11.text, 1) = "g" Or _
Left(Text11.text, 1) = "H" Or Left(Text11.text, 1) = "h" Or _
Left(Text11.text, 1) = "J" Or Left(Text11.text, 1) = "j" Or _
Left(Text11.text, 1) = "K" Or Left(Text11.text, 1) = "k" Or _
Left(Text11.text, 1) = "L" Or Left(Text11.text, 1) = "l" Then
  Txt = SetStandardBeamBreadth(Left(Text11.text, 1))
    Text11.text = Txt
      End If
      Text11.text = Val((Text11.text))
If Val(Text11.text) <= 0 Then
   Text11.text = 0
   End If
End Sub

Private Sub Text11_GotFocus()
Line16.BorderColor = vbYellow
   Line16.BorderWidth = 2
   Shape17.BorderColor = vbYellow
   Shape17.FillStyle = 0
   Shape17.BorderWidth = 2
End Sub

Private Sub Text11_LostFocus()
Line16.BorderColor = vbBlack
   Line16.BorderWidth = 1
   Shape17.BorderColor = vbBlack
   Shape17.FillStyle = 1
   Shape17.BorderWidth = 1
End Sub

Private Sub Text12_Change()
Dim Txt As String
If Left(Text12.text, 1) = "A" Or Left(Text12.text, 1) = "a" Or _
Left(Text12.text, 1) = "S" Or Left(Text12.text, 1) = "s" Or _
Left(Text12.text, 1) = "D" Or Left(Text12.text, 1) = "d" Or _
Left(Text12.text, 1) = "F" Or Left(Text12.text, 1) = "f" Or _
Left(Text12.text, 1) = "G" Or Left(Text12.text, 1) = "g" Or _
Left(Text12.text, 1) = "H" Or Left(Text12.text, 1) = "h" Or _
Left(Text12.text, 1) = "J" Or Left(Text12.text, 1) = "j" Or _
Left(Text12.text, 1) = "K" Or Left(Text12.text, 1) = "k" Or _
Left(Text12.text, 1) = "L" Or Left(Text12.text, 1) = "l" Then
  Txt = SetStandardBeamDepth(Left(Text12.text, 1))
    Text12.text = Txt
      End If
      Text12.text = Val((Text12.text))
If Val(Text12.text) <= 0 Then
   Text12.text = 0
   End If
End Sub

Private Sub Text12_GotFocus()
Line16.BorderColor = vbYellow
   Line16.BorderWidth = 2
   Shape17.BorderColor = vbYellow
   Shape17.FillStyle = 0
   Shape17.BorderWidth = 2
End Sub

Private Sub Text12_LostFocus()
Line16.BorderColor = vbBlack
   Line16.BorderWidth = 1
   Shape17.BorderColor = vbBlack
   Shape17.FillStyle = 1
   Shape17.BorderWidth = 1
End Sub

Private Sub Text13_Change()
Dim Txt As String
If Left(Text13.text, 1) = "A" Or Left(Text13.text, 1) = "a" Or _
Left(Text13.text, 1) = "S" Or Left(Text13.text, 1) = "s" Or _
Left(Text13.text, 1) = "D" Or Left(Text13.text, 1) = "d" Or _
Left(Text13.text, 1) = "F" Or Left(Text13.text, 1) = "f" Or _
Left(Text13.text, 1) = "G" Or Left(Text13.text, 1) = "g" Or _
Left(Text13.text, 1) = "H" Or Left(Text13.text, 1) = "h" Or _
Left(Text13.text, 1) = "J" Or Left(Text13.text, 1) = "j" Or _
Left(Text13.text, 1) = "K" Or Left(Text13.text, 1) = "k" Or _
Left(Text13.text, 1) = "L" Or Left(Text13.text, 1) = "l" Then
  Txt = SetStandardBeamLength(Left(Text13.text, 1))
    Text13.text = Txt
      End If
      Text13.text = Val((Text13.text))
If Val(Text13.text) <= 0 Then
   Text13.text = 0
   End If
End Sub

Private Sub Text13_GotFocus()
Line16.BorderColor = vbYellow
   Line16.BorderWidth = 2
   Shape17.BorderColor = vbYellow
   Shape17.FillStyle = 0
   Shape17.BorderWidth = 2
End Sub

Private Sub Text13_LostFocus()
Line16.BorderColor = vbBlack
   Line16.BorderWidth = 1
   Shape17.BorderColor = vbBlack
   Shape17.FillStyle = 1
   Shape17.BorderWidth = 1
End Sub

Private Sub Text14_Change()
Dim Txt As String
If Left(Text14.text, 1) = "A" Or Left(Text14.text, 1) = "a" Or _
Left(Text14.text, 1) = "S" Or Left(Text14.text, 1) = "s" Or _
Left(Text14.text, 1) = "D" Or Left(Text14.text, 1) = "d" Or _
Left(Text14.text, 1) = "F" Or Left(Text14.text, 1) = "f" Or _
Left(Text14.text, 1) = "G" Or Left(Text14.text, 1) = "g" Or _
Left(Text14.text, 1) = "H" Or Left(Text14.text, 1) = "h" Or _
Left(Text14.text, 1) = "J" Or Left(Text14.text, 1) = "j" Or _
Left(Text14.text, 1) = "K" Or Left(Text14.text, 1) = "k" Or _
Left(Text14.text, 1) = "L" Or Left(Text14.text, 1) = "l" Then
  Txt = SetStandardBeamBreadth(Left(Text14.text, 1))
    Text14.text = Txt
      End If
      Text14.text = Val((Text14.text))
If Val(Text14.text) <= 0 Then
   Text14.text = 0
   End If
End Sub

Private Sub Text14_GotFocus()
 Line17.BorderColor = vbYellow
   Line17.BorderWidth = 2
   Shape18.BorderColor = vbYellow
   Shape18.FillStyle = 0
   Shape18.BorderWidth = 2
End Sub

Private Sub Text14_LostFocus()
 Line17.BorderColor = vbBlack
   Line17.BorderWidth = 1
   Shape18.BorderColor = vbBlack
   Shape18.FillStyle = 1
   Shape18.BorderWidth = 1
End Sub

Private Sub Text15_Change()
Dim Txt As String
If Left(Text15.text, 1) = "A" Or Left(Text15.text, 1) = "a" Or _
Left(Text15.text, 1) = "S" Or Left(Text15.text, 1) = "s" Or _
Left(Text15.text, 1) = "D" Or Left(Text15.text, 1) = "d" Or _
Left(Text15.text, 1) = "F" Or Left(Text15.text, 1) = "f" Or _
Left(Text15.text, 1) = "G" Or Left(Text15.text, 1) = "g" Or _
Left(Text15.text, 1) = "H" Or Left(Text15.text, 1) = "h" Or _
Left(Text15.text, 1) = "J" Or Left(Text15.text, 1) = "j" Or _
Left(Text15.text, 1) = "K" Or Left(Text15.text, 1) = "k" Or _
Left(Text15.text, 1) = "L" Or Left(Text15.text, 1) = "l" Then
  Txt = SetStandardBeamDepth(Left(Text15.text, 1))
    Text15.text = Txt
      End If
      Text15.text = Val((Text15.text))
If Val(Text15.text) <= 0 Then
   Text15.text = 0
   End If
End Sub

Private Sub Text15_GotFocus()
 Line17.BorderColor = vbYellow
   Line17.BorderWidth = 2
   Shape18.BorderColor = vbYellow
   Shape18.FillStyle = 0
   Shape18.BorderWidth = 2
End Sub

Private Sub Text15_LostFocus()
 Line17.BorderColor = vbBlack
   Line17.BorderWidth = 1
   Shape18.BorderColor = vbBlack
   Shape18.FillStyle = 1
   Shape18.BorderWidth = 1
End Sub

Private Sub Text16_Change()
Dim Txt As String
If Left(Text16.text, 1) = "A" Or Left(Text16.text, 1) = "a" Or _
Left(Text16.text, 1) = "S" Or Left(Text16.text, 1) = "s" Or _
Left(Text16.text, 1) = "D" Or Left(Text16.text, 1) = "d" Or _
Left(Text16.text, 1) = "F" Or Left(Text16.text, 1) = "f" Or _
Left(Text16.text, 1) = "G" Or Left(Text16.text, 1) = "g" Or _
Left(Text16.text, 1) = "H" Or Left(Text16.text, 1) = "h" Or _
Left(Text16.text, 1) = "J" Or Left(Text16.text, 1) = "j" Or _
Left(Text16.text, 1) = "K" Or Left(Text16.text, 1) = "k" Or _
Left(Text16.text, 1) = "L" Or Left(Text16.text, 1) = "l" Then
  Txt = SetStandardBeamLength(Left(Text16.text, 1))
    Text16.text = Txt
      End If
      Text16.text = Val((Text16.text))
If Val(Text16.text) <= 0 Then
   Text16.text = 0
   End If
End Sub

Private Sub Text16_GotFocus()
 Line17.BorderColor = vbYellow
   Line17.BorderWidth = 2
   Shape18.BorderColor = vbYellow
   Shape18.FillStyle = 0
   Shape18.BorderWidth = 2
End Sub

Private Sub Text16_LostFocus()
 Line17.BorderColor = vbBlack
   Line17.BorderWidth = 1
   Shape18.BorderColor = vbBlack
   Shape18.FillStyle = 1
   Shape18.BorderWidth = 1
End Sub

Private Sub Text17_Change()
  Dim Txt As String
If Left(Text17.text, 1) = "A" Or Left(Text17.text, 1) = "a" Or _
Left(Text17.text, 1) = "S" Or Left(Text17.text, 1) = "s" Or _
Left(Text17.text, 1) = "D" Or Left(Text17.text, 1) = "d" Or _
Left(Text17.text, 1) = "F" Or Left(Text17.text, 1) = "f" Or _
Left(Text17.text, 1) = "G" Or Left(Text17.text, 1) = "g" Or _
Left(Text17.text, 1) = "H" Or Left(Text17.text, 1) = "h" Or _
Left(Text17.text, 1) = "J" Or Left(Text17.text, 1) = "j" Or _
Left(Text17.text, 1) = "K" Or Left(Text17.text, 1) = "k" Or _
Left(Text17.text, 1) = "L" Or Left(Text17.text, 1) = "l" Then
 Txt = SetStandardColumnDepth(Left(Text17.text, 1))
    Text17.text = Txt
      End If
      Text17.text = Val((Text17.text))

If Val(Text17.text) <= 0 Then
   Text17.text = 50
   End If
End Sub

Private Sub Text17_GotFocus()
Shape22.FillColor = vbGreen
   Shape22.FillStyle = 0
   Shape22.BorderColor = vbCyan
   Shape22.BorderWidth = 2
   Line9.BorderColor = vbRed
   Line9.BorderWidth = 2
   Line8.BorderColor = vbRed
   Line8.BorderWidth = 2
   
End Sub

Private Sub Text17_LostFocus()
Shape22.FillColor = vbBlack
   Shape22.FillStyle = 1
   Shape22.BorderColor = vbBlack
   Shape22.BorderWidth = 1
   Line9.BorderColor = vbWhite
   Line9.BorderWidth = 1
   Line8.BorderColor = vbWhite
   Line8.BorderWidth = 1
   Call GraphicOne(Val(Text17.text), Val(Text18.text), _
       Val(BarXBotN(bilJenisTng)) * Val(BarXBotD(bilJenisTng)) ^ 2 * 3.14 / 4, _
       Val(BarXBotD(bilJenisTng)) ^ 2 * 3.14 / 2 + Val(BarYLefN(bilJenisTng)) * Val(BarYLefD(bilJenisTng)) ^ 2 * 3.14 / 4)
         
End Sub

Private Sub Text18_Change()
Dim Txt As String
If Left(Text18.text, 1) = "A" Or Left(Text18.text, 1) = "a" Or _
Left(Text18.text, 1) = "S" Or Left(Text18.text, 1) = "s" Or _
Left(Text18.text, 1) = "D" Or Left(Text18.text, 1) = "d" Or _
Left(Text18.text, 1) = "F" Or Left(Text18.text, 1) = "f" Or _
Left(Text18.text, 1) = "G" Or Left(Text18.text, 1) = "g" Or _
Left(Text18.text, 1) = "H" Or Left(Text18.text, 1) = "h" Or _
Left(Text18.text, 1) = "J" Or Left(Text18.text, 1) = "j" Or _
Left(Text18.text, 1) = "K" Or Left(Text18.text, 1) = "k" Or _
Left(Text18.text, 1) = "L" Or Left(Text18.text, 1) = "l" Then
 Txt = SetStandardColumnDepth(Left(Text18.text, 1))
    Text18.text = Txt
      End If
      Text18.text = Val((Text18.text))
If Val(Text18.text) <= 0 Then
  Text18.text = 50
  End If
End Sub

Private Sub Text18_GotFocus()
Shape22.FillColor = vbGreen
   Shape22.FillStyle = 0
   Shape22.BorderColor = vbCyan
   Shape22.BorderWidth = 2
   Line7.BorderColor = vbRed
   Line7.BorderWidth = 2
   Line1.BorderColor = vbRed
   Line1.BorderWidth = 2
End Sub

Private Sub Text18_LostFocus()
Shape22.FillColor = vbBlack
   Shape22.FillStyle = 1
   Shape22.BorderColor = vbBlack
   Shape22.BorderWidth = 1
   Line7.BorderColor = vbWhite
   Line7.BorderWidth = 1
   Line1.BorderColor = vbWhite
   Line1.BorderWidth = 1
   Call GraphicOne(Val(Text17.text), Val(Text18.text), _
       Val(BarXBotN(bilJenisTng)) * Val(BarXBotD(bilJenisTng)) ^ 2 * 3.14 / 4, _
       Val(BarXBotD(bilJenisTng)) ^ 2 * 3.14 / 2 + Val(BarYLefN(bilJenisTng)) * Val(BarYLefD(bilJenisTng)) ^ 2 * 3.14 / 4)
         
End Sub

Private Sub Text19_Change()
Dim Txt As String
If Left(Text19.text, 1) = "A" Or Left(Text19.text, 1) = "a" Or _
Left(Text19.text, 1) = "S" Or Left(Text19.text, 1) = "s" Or _
Left(Text19.text, 1) = "D" Or Left(Text19.text, 1) = "d" Or _
Left(Text19.text, 1) = "F" Or Left(Text19.text, 1) = "f" Or _
Left(Text19.text, 1) = "G" Or Left(Text19.text, 1) = "g" Or _
Left(Text19.text, 1) = "H" Or Left(Text19.text, 1) = "h" Or _
Left(Text19.text, 1) = "J" Or Left(Text19.text, 1) = "j" Or _
Left(Text19.text, 1) = "K" Or Left(Text19.text, 1) = "k" Or _
Left(Text19.text, 1) = "L" Or Left(Text19.text, 1) = "l" Then
  Txt = SetStandardColumnHeight(Left(Text19.text, 1))
    Text19.text = Txt
      End If
      Text19.text = Val((Text19.text))
If Val(Text19.text) <= 0 Then
   Text19.text = 1000
   End If
End Sub

Private Sub Text19_GotFocus()
Shape22.FillColor = vbGreen
   Shape22.FillStyle = 0
   Shape22.BorderColor = vbCyan
   Shape22.BorderWidth = 2
End Sub

Private Sub Text19_LostFocus()
Shape22.FillColor = vbBlack
   Shape22.FillStyle = 1
   Shape22.BorderColor = vbBlack
   Shape22.BorderWidth = 1
End Sub

Private Sub Text2_Change()
Dim Txt As String
If Left(Text2.text, 1) = "A" Or Left(Text2.text, 1) = "a" Or _
Left(Text2.text, 1) = "S" Or Left(Text2.text, 1) = "s" Or _
Left(Text2.text, 1) = "D" Or Left(Text2.text, 1) = "d" Or _
Left(Text2.text, 1) = "F" Or Left(Text2.text, 1) = "f" Or _
Left(Text2.text, 1) = "G" Or Left(Text2.text, 1) = "g" Or _
Left(Text2.text, 1) = "H" Or Left(Text2.text, 1) = "h" Or _
Left(Text2.text, 1) = "J" Or Left(Text2.text, 1) = "j" Or _
Left(Text2.text, 1) = "K" Or Left(Text2.text, 1) = "k" Or _
Left(Text2.text, 1) = "L" Or Left(Text2.text, 1) = "l" Then
  Txt = SetStandardColumnDepth(Left(Text2.text, 1))
    Text2.text = Txt
      End If
      Text2.text = Val((Text2.text))
If Val(Text2.text) <= 0 Then
   Text2.text = 0
   End If
End Sub

Private Sub Text2_GotFocus()
Shape23.BorderColor = vbGreen
   Shape23.BorderWidth = 2
End Sub

Private Sub Text2_LostFocus()
Shape23.BorderColor = vbBlack
   Shape23.BorderWidth = 1
End Sub

Private Sub Text20_Change()
Dim Txt As String
If Left(Text20.text, 1) = "A" Or Left(Text20.text, 1) = "a" Or _
Left(Text20.text, 1) = "S" Or Left(Text20.text, 1) = "s" Or _
Left(Text20.text, 1) = "D" Or Left(Text20.text, 1) = "d" Or _
Left(Text20.text, 1) = "F" Or Left(Text20.text, 1) = "f" Or _
Left(Text20.text, 1) = "G" Or Left(Text20.text, 1) = "g" Or _
Left(Text20.text, 1) = "H" Or Left(Text20.text, 1) = "h" Or _
Left(Text20.text, 1) = "J" Or Left(Text20.text, 1) = "j" Or _
Left(Text20.text, 1) = "K" Or Left(Text20.text, 1) = "k" Or _
Left(Text20.text, 1) = "L" Or Left(Text20.text, 1) = "l" Then
  Txt = SetStandardMoment(Left(Text20.text, 1))
    Text20.text = Txt
      End If
      Text20.text = Val((Text20.text))
If Val(Text20.text) = 0 And Left(Text20.text, 1) <> "-" Then
  Text20.text = 0
  End If
End Sub

Private Sub Text20_GotFocus()
Shape26.BorderColor = vbGreen
   Shape26.BorderWidth = 2
End Sub

Private Sub Text20_LostFocus()
Shape26.BorderColor = vbBlack
   Shape26.BorderWidth = 1
End Sub

Private Sub Text21_Change()
Dim Txt As String
If Left(Text21.text, 1) = "A" Or Left(Text21.text, 1) = "a" Or _
Left(Text21.text, 1) = "S" Or Left(Text21.text, 1) = "s" Or _
Left(Text21.text, 1) = "D" Or Left(Text21.text, 1) = "d" Or _
Left(Text21.text, 1) = "F" Or Left(Text21.text, 1) = "f" Or _
Left(Text21.text, 1) = "G" Or Left(Text21.text, 1) = "g" Or _
Left(Text21.text, 1) = "H" Or Left(Text21.text, 1) = "h" Or _
Left(Text21.text, 1) = "J" Or Left(Text21.text, 1) = "j" Or _
Left(Text21.text, 1) = "K" Or Left(Text21.text, 1) = "k" Or _
Left(Text21.text, 1) = "L" Or Left(Text21.text, 1) = "l" Then
  Txt = SetStandardMoment(Left(Text21.text, 1))
    Text21.text = Txt
      End If
      Text21.text = Val((Text21.text))
If Val(Text21.text) = 0 And Left(Text21.text, 1) <> "-" Then
  Text21.text = 0
  End If
End Sub

Private Sub Text21_GotFocus()
Shape26.BorderColor = vbYellow
   Shape26.BorderWidth = 2
End Sub

Private Sub Text21_LostFocus()
Shape26.BorderColor = vbBlack
   Shape26.BorderWidth = 1
End Sub

Private Sub Text22_Change()
 Dim Txt As String
If Left(Text22.text, 1) = "A" Or Left(Text22.text, 1) = "a" Or _
Left(Text22.text, 1) = "S" Or Left(Text22.text, 1) = "s" Or _
Left(Text22.text, 1) = "D" Or Left(Text22.text, 1) = "d" Or _
Left(Text22.text, 1) = "F" Or Left(Text22.text, 1) = "f" Or _
Left(Text22.text, 1) = "G" Or Left(Text22.text, 1) = "g" Or _
Left(Text22.text, 1) = "H" Or Left(Text22.text, 1) = "h" Or _
Left(Text22.text, 1) = "J" Or Left(Text22.text, 1) = "j" Or _
Left(Text22.text, 1) = "K" Or Left(Text22.text, 1) = "k" Or _
Left(Text22.text, 1) = "L" Or Left(Text22.text, 1) = "l" Then
  Txt = SetStandardMoment(Left(Text22.text, 1))
    Text22.text = Txt
      End If
      Text22.text = Val((Text22.text))
If Val(Text22.text) = 0 And Left(Text22.text, 1) <> "-" Then
  Text22.text = 0
  End If
End Sub

Private Sub Text22_GotFocus()
Shape25.BorderColor = vbGreen
   Shape25.BorderWidth = 2
End Sub

Private Sub Text22_LostFocus()
Shape25.BorderColor = vbBlack
   Shape25.BorderWidth = 1
End Sub

Private Sub Text23_Change()
Dim Txt As String
If Left(Text23.text, 1) = "A" Or Left(Text23.text, 1) = "a" Or _
Left(Text23.text, 1) = "S" Or Left(Text23.text, 1) = "s" Or _
Left(Text23.text, 1) = "D" Or Left(Text23.text, 1) = "d" Or _
Left(Text23.text, 1) = "F" Or Left(Text23.text, 1) = "f" Or _
Left(Text23.text, 1) = "G" Or Left(Text23.text, 1) = "g" Or _
Left(Text23.text, 1) = "H" Or Left(Text23.text, 1) = "h" Or _
Left(Text23.text, 1) = "J" Or Left(Text23.text, 1) = "j" Or _
Left(Text23.text, 1) = "K" Or Left(Text23.text, 1) = "k" Or _
Left(Text23.text, 1) = "L" Or Left(Text23.text, 1) = "l" Then
  Txt = SetStandardMoment(Left(Text23.text, 1))
    Text23.text = Txt
      End If
      Text23.text = Val((Text23.text))
If Val(Text23.text) = 0 And Left(Text23.text, 1) <> "-" Then
  Text23.text = 0
  End If
End Sub

Private Sub Text23_GotFocus()
Shape25.BorderColor = vbYellow
   Shape25.BorderWidth = 2
End Sub

Private Sub Text23_LostFocus()
Shape25.BorderColor = vbBlack
   Shape25.BorderWidth = 1
End Sub

Private Sub Text24_Change()
 Dim Txt As String
If Left(Text24.text, 1) = "A" Or Left(Text24.text, 1) = "a" Or _
Left(Text24.text, 1) = "S" Or Left(Text24.text, 1) = "s" Or _
Left(Text24.text, 1) = "D" Or Left(Text24.text, 1) = "d" Or _
Left(Text24.text, 1) = "F" Or Left(Text24.text, 1) = "f" Or _
Left(Text24.text, 1) = "G" Or Left(Text24.text, 1) = "g" Or _
Left(Text24.text, 1) = "H" Or Left(Text24.text, 1) = "h" Or _
Left(Text24.text, 1) = "J" Or Left(Text24.text, 1) = "j" Or _
Left(Text24.text, 1) = "K" Or Left(Text24.text, 1) = "k" Or _
Left(Text24.text, 1) = "L" Or Left(Text24.text, 1) = "l" Then
  Txt = SetStandardBeamBreadth(Left(Text24.text, 1))
    Text24.text = Txt
      End If
      Text24.text = Val((Text24.text))
If Val(Text24.text) <= 0 Then
   Text24.text = 0
   End If
End Sub

Private Sub Text24_GotFocus()
Line13.BorderColor = vbYellow
   Line13.BorderWidth = 2
   Shape14.BorderColor = vbYellow
   Shape14.FillStyle = 0
   Shape14.BorderWidth = 2
End Sub

Private Sub Text24_LostFocus()
Line13.BorderColor = vbBlack
   Line13.BorderWidth = 1
   Shape14.BorderColor = vbBlack
   Shape14.FillStyle = 1
   Shape14.BorderWidth = 1
End Sub

Private Sub Text25_Change()
Dim Txt As String
If Left(Text25.text, 1) = "A" Or Left(Text25.text, 1) = "a" Or _
Left(Text25.text, 1) = "S" Or Left(Text25.text, 1) = "s" Or _
Left(Text25.text, 1) = "D" Or Left(Text25.text, 1) = "d" Or _
Left(Text25.text, 1) = "F" Or Left(Text25.text, 1) = "f" Or _
Left(Text25.text, 1) = "G" Or Left(Text25.text, 1) = "g" Or _
Left(Text25.text, 1) = "H" Or Left(Text25.text, 1) = "h" Or _
Left(Text25.text, 1) = "J" Or Left(Text25.text, 1) = "j" Or _
Left(Text25.text, 1) = "K" Or Left(Text25.text, 1) = "k" Or _
Left(Text25.text, 1) = "L" Or Left(Text25.text, 1) = "l" Then
  Txt = SetStandardBeamDepth(Left(Text25.text, 1))
    Text25.text = Txt
      End If
      Text25.text = Val((Text25.text))
If Val(Text25.text) <= 0 Then
   Text25.text = 0
   End If
End Sub

Private Sub Text25_GotFocus()
Line13.BorderColor = vbYellow
   Line13.BorderWidth = 2
   Shape14.BorderColor = vbYellow
   Shape14.FillStyle = 0
   Shape14.BorderWidth = 2
End Sub

Private Sub Text25_LostFocus()
Line13.BorderColor = vbBlack
   Line13.BorderWidth = 1
   Shape14.BorderColor = vbBlack
   Shape14.FillStyle = 1
   Shape14.BorderWidth = 1
End Sub

Private Sub Text26_Change()
Dim Txt As String
If Left(Text26.text, 1) = "A" Or Left(Text26.text, 1) = "a" Or _
Left(Text26.text, 1) = "S" Or Left(Text26.text, 1) = "s" Or _
Left(Text26.text, 1) = "D" Or Left(Text26.text, 1) = "d" Or _
Left(Text26.text, 1) = "F" Or Left(Text26.text, 1) = "f" Or _
Left(Text26.text, 1) = "G" Or Left(Text26.text, 1) = "g" Or _
Left(Text26.text, 1) = "H" Or Left(Text26.text, 1) = "h" Or _
Left(Text26.text, 1) = "J" Or Left(Text26.text, 1) = "j" Or _
Left(Text26.text, 1) = "K" Or Left(Text26.text, 1) = "k" Or _
Left(Text26.text, 1) = "L" Or Left(Text26.text, 1) = "l" Then
  Txt = SetStandardBeamLength(Left(Text26.text, 1))
    Text26.text = Txt
      End If
      Text26.text = Val((Text26.text))
If Val(Text26.text) <= 0 Then
   Text26.text = 0
   End If
End Sub

Private Sub Text26_GotFocus()
Line13.BorderColor = vbYellow
   Line13.BorderWidth = 2
   Shape14.BorderColor = vbYellow
   Shape14.FillStyle = 0
   Shape14.BorderWidth = 2
End Sub

Private Sub Text26_LostFocus()
Line13.BorderColor = vbBlack
   Line13.BorderWidth = 1
   Shape14.BorderColor = vbBlack
   Shape14.FillStyle = 1
   Shape14.BorderWidth = 1
End Sub

Private Sub Text27_Change()
   Dim Txt As String
If Left(Text27.text, 1) = "A" Or Left(Text27.text, 1) = "a" Or _
Left(Text27.text, 1) = "S" Or Left(Text27.text, 1) = "s" Or _
Left(Text27.text, 1) = "D" Or Left(Text27.text, 1) = "d" Or _
Left(Text27.text, 1) = "F" Or Left(Text27.text, 1) = "f" Or _
Left(Text27.text, 1) = "G" Or Left(Text27.text, 1) = "g" Or _
Left(Text27.text, 1) = "H" Or Left(Text27.text, 1) = "h" Or _
Left(Text27.text, 1) = "J" Or Left(Text27.text, 1) = "j" Or _
Left(Text27.text, 1) = "K" Or Left(Text27.text, 1) = "k" Or _
Left(Text27.text, 1) = "L" Or Left(Text27.text, 1) = "l" Then
  Txt = SetStandardBeamBreadth(Left(Text27.text, 1))
    Text27.text = Txt
      End If
      Text27.text = Val((Text27.text))
If Val(Text27.text) <= 0 Then
   Text27.text = 0
   End If
End Sub

Private Sub Text27_GotFocus()
 Line12.BorderColor = vbYellow
   Line12.BorderWidth = 2
   Shape13.BorderColor = vbYellow
   Shape13.FillStyle = 0
   Shape13.BorderWidth = 2
End Sub

Private Sub Text27_LostFocus()
 Line12.BorderColor = vbBlack
   Line12.BorderWidth = 1
   Shape13.BorderColor = vbBlack
   Shape13.FillStyle = 1
   Shape13.BorderWidth = 1
End Sub

Private Sub Text28_Change()
Dim Txt As String
If Left(Text28.text, 1) = "A" Or Left(Text28.text, 1) = "a" Or _
Left(Text28.text, 1) = "S" Or Left(Text28.text, 1) = "s" Or _
Left(Text28.text, 1) = "D" Or Left(Text28.text, 1) = "d" Or _
Left(Text28.text, 1) = "F" Or Left(Text28.text, 1) = "f" Or _
Left(Text28.text, 1) = "G" Or Left(Text28.text, 1) = "g" Or _
Left(Text28.text, 1) = "H" Or Left(Text28.text, 1) = "h" Or _
Left(Text28.text, 1) = "J" Or Left(Text28.text, 1) = "j" Or _
Left(Text28.text, 1) = "K" Or Left(Text28.text, 1) = "k" Or _
Left(Text28.text, 1) = "L" Or Left(Text28.text, 1) = "l" Then
  Txt = SetStandardBeamDepth(Left(Text28.text, 1))
    Text28.text = Txt
      End If
      Text28.text = Val((Text28.text))
If Val(Text28.text) <= 0 Then
   Text28.text = 0
   End If
End Sub

Private Sub Text28_GotFocus()
 Line12.BorderColor = vbYellow
   Line12.BorderWidth = 2
   Shape13.BorderColor = vbYellow
   Shape13.FillStyle = 0
   Shape13.BorderWidth = 2
End Sub

Private Sub Text28_LostFocus()
 Line12.BorderColor = vbBlack
   Line12.BorderWidth = 1
   Shape13.BorderColor = vbBlack
   Shape13.FillStyle = 1
   Shape13.BorderWidth = 1
End Sub

Private Sub Text29_Change()
Dim Txt As String
If Left(Text29.text, 1) = "A" Or Left(Text29.text, 1) = "a" Or _
Left(Text29.text, 1) = "S" Or Left(Text29.text, 1) = "s" Or _
Left(Text29.text, 1) = "D" Or Left(Text29.text, 1) = "d" Or _
Left(Text29.text, 1) = "F" Or Left(Text29.text, 1) = "f" Or _
Left(Text29.text, 1) = "G" Or Left(Text29.text, 1) = "g" Or _
Left(Text29.text, 1) = "H" Or Left(Text29.text, 1) = "h" Or _
Left(Text29.text, 1) = "J" Or Left(Text29.text, 1) = "j" Or _
Left(Text29.text, 1) = "K" Or Left(Text29.text, 1) = "k" Or _
Left(Text29.text, 1) = "L" Or Left(Text29.text, 1) = "l" Then
  Txt = SetStandardBeamLength(Left(Text29.text, 1))
    Text29.text = Txt
      End If
      Text29.text = Val((Text29.text))
If Val(Text29.text) <= 0 Then
   Text29.text = 0
   End If
End Sub

Private Sub Text29_GotFocus()
 Line12.BorderColor = vbYellow
   Line12.BorderWidth = 2
   Shape13.BorderColor = vbYellow
   Shape13.FillStyle = 0
   Shape13.BorderWidth = 2
End Sub

Private Sub Text29_LostFocus()
 Line12.BorderColor = vbBlack
   Line12.BorderWidth = 1
   Shape13.BorderColor = vbBlack
   Shape13.FillStyle = 1
   Shape13.BorderWidth = 1
End Sub

Private Sub Text3_Change()
Dim Txt As String
If Left(Text3.text, 1) = "A" Or Left(Text3.text, 1) = "a" Or _
Left(Text3.text, 1) = "S" Or Left(Text3.text, 1) = "s" Or _
Left(Text3.text, 1) = "D" Or Left(Text3.text, 1) = "d" Or _
Left(Text3.text, 1) = "F" Or Left(Text3.text, 1) = "f" Or _
Left(Text3.text, 1) = "G" Or Left(Text3.text, 1) = "g" Or _
Left(Text3.text, 1) = "H" Or Left(Text3.text, 1) = "h" Or _
Left(Text3.text, 1) = "J" Or Left(Text3.text, 1) = "j" Or _
Left(Text3.text, 1) = "K" Or Left(Text3.text, 1) = "k" Or _
Left(Text3.text, 1) = "L" Or Left(Text3.text, 1) = "l" Then
  Txt = SetStandardColumnHeight(Left(Text3.text, 1))
    Text3.text = Txt
      End If
      Text3.text = Val((Text3.text))
If Val(Text3.text) <= 0 Then
   Text3.text = 0
   End If
End Sub

Private Sub Text3_GotFocus()
Shape23.BorderColor = vbGreen
   Shape23.BorderWidth = 2
  
End Sub

Private Sub Text3_LostFocus()
Shape23.BorderColor = vbBlack
   Shape23.BorderWidth = 1
  
End Sub

Private Sub Text30_Change()
Dim Txt As String
If Left(Text30.text, 1) = "A" Or Left(Text30.text, 1) = "a" Or _
Left(Text30.text, 1) = "S" Or Left(Text30.text, 1) = "s" Or _
Left(Text30.text, 1) = "D" Or Left(Text30.text, 1) = "d" Or _
Left(Text30.text, 1) = "F" Or Left(Text30.text, 1) = "f" Or _
Left(Text30.text, 1) = "G" Or Left(Text30.text, 1) = "g" Or _
Left(Text30.text, 1) = "H" Or Left(Text30.text, 1) = "h" Or _
Left(Text30.text, 1) = "J" Or Left(Text30.text, 1) = "j" Or _
Left(Text30.text, 1) = "K" Or Left(Text30.text, 1) = "k" Or _
Left(Text30.text, 1) = "L" Or Left(Text30.text, 1) = "l" Then
  Txt = SetStandardBeamBreadth(Left(Text30.text, 1))
    Text30.text = Txt
      End If
      Text30.text = Val((Text30.text))
If Val(Text30.text) <= 0 Then
   Text30.text = 0
   End If
End Sub

Private Sub Text30_GotFocus()
Line14.BorderColor = vbCyan
   Line14.BorderWidth = 2
   Shape15.BorderColor = vbCyan
   Shape15.FillStyle = 0
   Shape15.BorderWidth = 2
End Sub

Private Sub Text30_LostFocus()
Line14.BorderColor = vbBlack
   Line14.BorderWidth = 1
   Shape15.BorderColor = vbBlack
   Shape15.FillStyle = 1
   Shape15.BorderWidth = 1
End Sub

Private Sub Text31_Change()
Dim Txt As String
If Left(Text31.text, 1) = "A" Or Left(Text31.text, 1) = "a" Or _
Left(Text31.text, 1) = "S" Or Left(Text31.text, 1) = "s" Or _
Left(Text31.text, 1) = "D" Or Left(Text31.text, 1) = "d" Or _
Left(Text31.text, 1) = "F" Or Left(Text31.text, 1) = "f" Or _
Left(Text31.text, 1) = "G" Or Left(Text31.text, 1) = "g" Or _
Left(Text31.text, 1) = "H" Or Left(Text31.text, 1) = "h" Or _
Left(Text31.text, 1) = "J" Or Left(Text31.text, 1) = "j" Or _
Left(Text31.text, 1) = "K" Or Left(Text31.text, 1) = "k" Or _
Left(Text31.text, 1) = "L" Or Left(Text31.text, 1) = "l" Then
  Txt = SetStandardBeamDepth(Left(Text31.text, 1))
    Text31.text = Txt
      End If
      Text31.text = Val((Text31.text))
If Val(Text31.text) <= 0 Then
   Text31.text = 0
   End If
End Sub

Private Sub Text31_GotFocus()
Line14.BorderColor = vbCyan
   Line14.BorderWidth = 2
   Shape15.BorderColor = vbCyan
   Shape15.FillStyle = 0
   Shape15.BorderWidth = 2
End Sub

Private Sub Text31_LostFocus()
Line14.BorderColor = vbBlack
   Line14.BorderWidth = 1
   Shape15.BorderColor = vbBlack
   Shape15.FillStyle = 1
   Shape15.BorderWidth = 1
End Sub

Private Sub Text32_Change()
Dim Txt As String
If Left(Text32.text, 1) = "A" Or Left(Text32.text, 1) = "a" Or _
Left(Text32.text, 1) = "S" Or Left(Text32.text, 1) = "s" Or _
Left(Text32.text, 1) = "D" Or Left(Text32.text, 1) = "d" Or _
Left(Text32.text, 1) = "F" Or Left(Text32.text, 1) = "f" Or _
Left(Text32.text, 1) = "G" Or Left(Text32.text, 1) = "g" Or _
Left(Text32.text, 1) = "H" Or Left(Text32.text, 1) = "h" Or _
Left(Text32.text, 1) = "J" Or Left(Text32.text, 1) = "j" Or _
Left(Text32.text, 1) = "K" Or Left(Text32.text, 1) = "k" Or _
Left(Text32.text, 1) = "L" Or Left(Text32.text, 1) = "l" Then
  Txt = SetStandardBeamLength(Left(Text32.text, 1))
    Text32.text = Txt
      End If
      Text32.text = Val((Text32.text))
If Val(Text32.text) <= 0 Then
   Text32.text = 0
   End If
End Sub

Private Sub Text32_GotFocus()
Line14.BorderColor = vbCyan
   Line14.BorderWidth = 2
   Shape15.BorderColor = vbCyan
   Shape15.FillStyle = 0
   Shape15.BorderWidth = 2
End Sub

Private Sub Text32_LostFocus()
Line14.BorderColor = vbBlack
   Line14.BorderWidth = 1
   Shape15.BorderColor = vbBlack
   Shape15.FillStyle = 1
   Shape15.BorderWidth = 1
End Sub

Private Sub Text33_Change()
Dim Txt As String
If Left(Text33.text, 1) = "A" Or Left(Text33.text, 1) = "a" Or _
Left(Text33.text, 1) = "S" Or Left(Text33.text, 1) = "s" Or _
Left(Text33.text, 1) = "D" Or Left(Text33.text, 1) = "d" Or _
Left(Text33.text, 1) = "F" Or Left(Text33.text, 1) = "f" Or _
Left(Text33.text, 1) = "G" Or Left(Text33.text, 1) = "g" Or _
Left(Text33.text, 1) = "H" Or Left(Text33.text, 1) = "h" Or _
Left(Text33.text, 1) = "J" Or Left(Text33.text, 1) = "j" Or _
Left(Text33.text, 1) = "K" Or Left(Text33.text, 1) = "k" Or _
Left(Text33.text, 1) = "L" Or Left(Text33.text, 1) = "l" Then
  Txt = SetStandardBeamBreadth(Left(Text33.text, 1))
    Text33.text = Txt
      End If
      Text33.text = Val((Text33.text))
If Val(Text33.text) <= 0 Then
   Text33.text = 0
   End If
End Sub

Private Sub Text33_GotFocus()
Line15.BorderColor = vbCyan
   Line15.BorderWidth = 2
   Shape16.BorderColor = vbCyan
   Shape16.FillStyle = 0
   Shape16.BorderWidth = 2
End Sub

Private Sub Text33_LostFocus()
Line15.BorderColor = vbBlack
   Line15.BorderWidth = 1
   Shape16.BorderColor = vbBlack
   Shape16.FillStyle = 1
   Shape16.BorderWidth = 1
End Sub

Private Sub Text34_Change()
   Dim Txt As String
If Left(Text34.text, 1) = "A" Or Left(Text34.text, 1) = "a" Or _
Left(Text34.text, 1) = "S" Or Left(Text34.text, 1) = "s" Or _
Left(Text34.text, 1) = "D" Or Left(Text34.text, 1) = "d" Or _
Left(Text34.text, 1) = "F" Or Left(Text34.text, 1) = "f" Or _
Left(Text34.text, 1) = "G" Or Left(Text34.text, 1) = "g" Or _
Left(Text34.text, 1) = "H" Or Left(Text34.text, 1) = "h" Or _
Left(Text34.text, 1) = "J" Or Left(Text34.text, 1) = "j" Or _
Left(Text34.text, 1) = "K" Or Left(Text34.text, 1) = "k" Or _
Left(Text34.text, 1) = "L" Or Left(Text34.text, 1) = "l" Then
  Txt = SetStandardBeamDepth(Left(Text34.text, 1))
    Text34.text = Txt
      End If
      Text34.text = Val((Text34.text))
If Val(Text34.text) <= 0 Then
   Text34.text = 0
   End If
End Sub

Private Sub Text34_GotFocus()
Line15.BorderColor = vbCyan
   Line15.BorderWidth = 2
   Shape16.BorderColor = vbCyan
   Shape16.FillStyle = 0
   Shape16.BorderWidth = 2
End Sub

Private Sub Text34_LostFocus()
Line15.BorderColor = vbBlack
   Line15.BorderWidth = 1
   Shape16.BorderColor = vbBlack
   Shape16.FillStyle = 1
   Shape16.BorderWidth = 1
End Sub

Private Sub Text35_Change()
Dim Txt As String
If Left(Text35.text, 1) = "A" Or Left(Text35.text, 1) = "a" Or _
Left(Text35.text, 1) = "S" Or Left(Text35.text, 1) = "s" Or _
Left(Text35.text, 1) = "D" Or Left(Text35.text, 1) = "d" Or _
Left(Text35.text, 1) = "F" Or Left(Text35.text, 1) = "f" Or _
Left(Text35.text, 1) = "G" Or Left(Text35.text, 1) = "g" Or _
Left(Text35.text, 1) = "H" Or Left(Text35.text, 1) = "h" Or _
Left(Text35.text, 1) = "J" Or Left(Text35.text, 1) = "j" Or _
Left(Text35.text, 1) = "K" Or Left(Text35.text, 1) = "k" Or _
Left(Text35.text, 1) = "L" Or Left(Text35.text, 1) = "l" Then
  Txt = SetStandardBeamLength(Left(Text35.text, 1))
    Text35.text = Txt
      End If
      Text35.text = Val((Text35.text))
If Val(Text35.text) <= 0 Then
   Text35.text = 0
   End If
End Sub

Private Sub Text35_GotFocus()
Line15.BorderColor = vbCyan
   Line15.BorderWidth = 2
   Shape16.BorderColor = vbCyan
   Shape16.FillStyle = 0
   Shape16.BorderWidth = 2
End Sub

Private Sub Text35_LostFocus()
Line15.BorderColor = vbBlack
   Line15.BorderWidth = 1
   Shape16.BorderColor = vbBlack
   Shape16.FillStyle = 1
   Shape16.BorderWidth = 1
End Sub

Private Sub Text36_Change()
    Dim Txt As String
If Left(Text36.text, 1) = "A" Or Left(Text36.text, 1) = "a" Or _
Left(Text36.text, 1) = "S" Or Left(Text36.text, 1) = "s" Or _
Left(Text36.text, 1) = "D" Or Left(Text36.text, 1) = "d" Or _
Left(Text36.text, 1) = "F" Or Left(Text36.text, 1) = "f" Or _
Left(Text36.text, 1) = "G" Or Left(Text36.text, 1) = "g" Or _
Left(Text36.text, 1) = "H" Or Left(Text36.text, 1) = "h" Or _
Left(Text36.text, 1) = "J" Or Left(Text36.text, 1) = "j" Or _
Left(Text36.text, 1) = "K" Or Left(Text36.text, 1) = "k" Or _
Left(Text36.text, 1) = "L" Or Left(Text36.text, 1) = "l" Then
  Txt = SetStandardColumnDepth(Left(Text36.text, 1))
    Text36.text = Txt
      End If
      Text36.text = Val((Text36.text))
If Val(Text36.text) <= 0 Then
   Text36.text = 0
   End If
End Sub

Private Sub Text36_GotFocus()
Shape21.BorderColor = vbGreen
   Shape21.BorderWidth = 2
End Sub

Private Sub Text36_LostFocus()
Shape21.BorderColor = vbBlack
   Shape21.BorderWidth = 1
End Sub

Private Sub Text37_Change()
Dim Txt As String
If Left(Text37.text, 1) = "A" Or Left(Text37.text, 1) = "a" Or _
Left(Text37.text, 1) = "S" Or Left(Text37.text, 1) = "s" Or _
Left(Text37.text, 1) = "D" Or Left(Text37.text, 1) = "d" Or _
Left(Text37.text, 1) = "F" Or Left(Text37.text, 1) = "f" Or _
Left(Text37.text, 1) = "G" Or Left(Text37.text, 1) = "g" Or _
Left(Text37.text, 1) = "H" Or Left(Text37.text, 1) = "h" Or _
Left(Text37.text, 1) = "J" Or Left(Text37.text, 1) = "j" Or _
Left(Text37.text, 1) = "K" Or Left(Text37.text, 1) = "k" Or _
Left(Text37.text, 1) = "L" Or Left(Text37.text, 1) = "l" Then
  Txt = SetStandardColumnDepth(Left(Text37.text, 1))
    Text37.text = Txt
      End If
      Text37.text = Val((Text37.text))
If Val(Text37.text) <= 0 Then
   Text37.text = 0
   End If
End Sub

Private Sub Text37_GotFocus()
Shape21.BorderColor = vbGreen
   Shape21.BorderWidth = 2
End Sub

Private Sub Text37_LostFocus()
Shape21.BorderColor = vbBlack
   Shape21.BorderWidth = 1
End Sub

Private Sub Text38_Change()
Dim Txt As String
If Left(Text38.text, 1) = "A" Or Left(Text38.text, 1) = "a" Or _
Left(Text38.text, 1) = "S" Or Left(Text38.text, 1) = "s" Or _
Left(Text38.text, 1) = "D" Or Left(Text38.text, 1) = "d" Or _
Left(Text38.text, 1) = "F" Or Left(Text38.text, 1) = "f" Or _
Left(Text38.text, 1) = "G" Or Left(Text38.text, 1) = "g" Or _
Left(Text38.text, 1) = "H" Or Left(Text38.text, 1) = "h" Or _
Left(Text38.text, 1) = "J" Or Left(Text38.text, 1) = "j" Or _
Left(Text38.text, 1) = "K" Or Left(Text38.text, 1) = "k" Or _
Left(Text38.text, 1) = "L" Or Left(Text38.text, 1) = "l" Then
  Txt = SetStandardColumnHeight(Left(Text38.text, 1))
    Text38.text = Txt
      End If
      Text38.text = Val((Text38.text))
If Val(Text38.text) <= 0 Then
   Text38.text = 0
   End If
End Sub

Private Sub Text38_GotFocus()
Shape21.BorderColor = vbGreen
   Shape21.BorderWidth = 2
End Sub

Private Sub Text38_LostFocus()
Shape21.BorderColor = vbBlack
   Shape21.BorderWidth = 1
End Sub

Private Sub Text39_Change()

''text fixity yy
End Sub

Private Sub Text39_Click()
Text39.text = "FREE"
End Sub

Private Sub Text39_DblClick()
Text39.text = "FIXED"
End Sub

Private Sub Text39_GotFocus()
Shape24.BorderColor = vbYellow
   Shape24.BorderWidth = 2
  
End Sub

Private Sub Text39_LostFocus()
Shape24.BorderColor = vbBlack
   Shape24.BorderWidth = 1
  
End Sub

Private Sub Text4_Change()

'''fixity xx
End Sub

Private Sub Text4_Click()
Text4.text = "FREE"
End Sub

Private Sub Text4_DblClick()
Text4.text = "FIXED"
End Sub

Private Sub Text4_GotFocus()
Shape24.BorderColor = vbGreen
   Shape24.BorderWidth = 2
  
End Sub

Private Sub Text4_LostFocus()
Shape24.BorderColor = vbBlack
   Shape24.BorderWidth = 1
  
End Sub

Private Sub Text40_Change()
Dim Txt As String
If Left(Text40.text, 1) = "A" Or Left(Text40.text, 1) = "a" Or _
Left(Text40.text, 1) = "S" Or Left(Text40.text, 1) = "s" Or _
Left(Text40.text, 1) = "D" Or Left(Text40.text, 1) = "d" Or _
Left(Text40.text, 1) = "F" Or Left(Text40.text, 1) = "f" Or _
Left(Text40.text, 1) = "G" Or Left(Text40.text, 1) = "g" Or _
Left(Text40.text, 1) = "H" Or Left(Text40.text, 1) = "h" Or _
Left(Text40.text, 1) = "J" Or Left(Text40.text, 1) = "j" Or _
Left(Text40.text, 1) = "K" Or Left(Text40.text, 1) = "k" Or _
Left(Text40.text, 1) = "L" Or Left(Text40.text, 1) = "l" Then
  Txt = SetStandardColumnLoad(Left(Text40.text, 1))
    Text40.text = Txt
      End If
      Text40.text = Val((Text40.text))
If Val(Text40.text) <= 0 Then
   Text40.text = 0
   End If
End Sub

Private Sub Text40_GotFocus()
Line2.BorderColor = vbRed
   Line17.BorderColor = vbRed
   Line17.BorderWidth = 2
   Line17.X1 = 6200
   Line19.BorderColor = vbRed
   Line19.BorderWidth = 2
   Line19.X1 = 6040
End Sub

Private Sub Text40_LostFocus()
Line2.BorderColor = vbWhite
   Line17.BorderColor = vbBlack
   Line17.BorderWidth = 1
   Line17.X1 = 6600
   Line19.BorderColor = vbBlack
   Line19.BorderWidth = 1
   Line19.X1 = 5640
End Sub

Private Sub Text41_Change()
Dim Txt As String
If Left(Text41.text, 1) = "A" Or Left(Text41.text, 1) = "a" Or _
Left(Text41.text, 1) = "S" Or Left(Text41.text, 1) = "s" Or _
Left(Text41.text, 1) = "D" Or Left(Text41.text, 1) = "d" Or _
Left(Text41.text, 1) = "F" Or Left(Text41.text, 1) = "f" Or _
Left(Text41.text, 1) = "G" Or Left(Text41.text, 1) = "g" Or _
Left(Text41.text, 1) = "H" Or Left(Text41.text, 1) = "h" Or _
Left(Text41.text, 1) = "J" Or Left(Text41.text, 1) = "j" Or _
Left(Text41.text, 1) = "K" Or Left(Text41.text, 1) = "k" Or _
Left(Text41.text, 1) = "L" Or Left(Text41.text, 1) = "l" Then
  Txt = SetStandardBarNumber(Left(Text41.text, 1))
    Text41.text = Txt
      End If
      Text41.text = Val((Text41.text))
If Val(Text41.text) <= 0 Then
   Text41.text = 0
   Text42.text = 0
   End If
End Sub

Private Sub Text42_Change()
Dim Txt As String
If Left(Text42.text, 1) = "A" Or Left(Text42.text, 1) = "a" Or _
Left(Text42.text, 1) = "S" Or Left(Text42.text, 1) = "s" Or _
Left(Text42.text, 1) = "D" Or Left(Text42.text, 1) = "d" Or _
Left(Text42.text, 1) = "F" Or Left(Text42.text, 1) = "f" Or _
Left(Text42.text, 1) = "G" Or Left(Text42.text, 1) = "g" Or _
Left(Text42.text, 1) = "H" Or Left(Text42.text, 1) = "h" Or _
Left(Text42.text, 1) = "J" Or Left(Text42.text, 1) = "j" Or _
Left(Text42.text, 1) = "K" Or Left(Text42.text, 1) = "k" Or _
Left(Text42.text, 1) = "L" Or Left(Text42.text, 1) = "l" Then
  Txt = SetStandardBarSize(Left(Text42.text, 1))
    Text42.text = Txt
      End If
      Text42.text = Val((Text42.text))
If Val(Text42.text) <= 0 Then
   Text42.text = 0
   End If
End Sub

Private Sub Text43_Change()
Dim Txt As String
If Left(Text43.text, 1) = "A" Or Left(Text43.text, 1) = "a" Or _
Left(Text43.text, 1) = "S" Or Left(Text43.text, 1) = "s" Or _
Left(Text43.text, 1) = "D" Or Left(Text43.text, 1) = "d" Or _
Left(Text43.text, 1) = "F" Or Left(Text43.text, 1) = "f" Or _
Left(Text43.text, 1) = "G" Or Left(Text43.text, 1) = "g" Or _
Left(Text43.text, 1) = "H" Or Left(Text43.text, 1) = "h" Or _
Left(Text43.text, 1) = "J" Or Left(Text43.text, 1) = "j" Or _
Left(Text43.text, 1) = "K" Or Left(Text43.text, 1) = "k" Or _
Left(Text43.text, 1) = "L" Or Left(Text43.text, 1) = "l" Then
  Txt = SetStandardBarNumber(Left(Text43.text, 1))
    Text43.text = Txt
      End If
      Text43.text = Val((Text43.text))
If Val(Text43.text) <= 0 Then
   Text43.text = 0
   Text44.text = 0
   End If
End Sub

Private Sub Text44_Change()
Dim Txt As String
If Left(Text44.text, 1) = "A" Or Left(Text44.text, 1) = "a" Or _
Left(Text44.text, 1) = "S" Or Left(Text44.text, 1) = "s" Or _
Left(Text44.text, 1) = "D" Or Left(Text44.text, 1) = "d" Or _
Left(Text44.text, 1) = "F" Or Left(Text44.text, 1) = "f" Or _
Left(Text44.text, 1) = "G" Or Left(Text44.text, 1) = "g" Or _
Left(Text44.text, 1) = "H" Or Left(Text44.text, 1) = "h" Or _
Left(Text44.text, 1) = "J" Or Left(Text44.text, 1) = "j" Or _
Left(Text44.text, 1) = "K" Or Left(Text44.text, 1) = "k" Or _
Left(Text44.text, 1) = "L" Or Left(Text44.text, 1) = "l" Then
  Txt = SetStandardBarSize(Left(Text44.text, 1))
    Text44.text = Txt
      End If
      Text44.text = Val((Text44.text))
If Val(Text44.text) <= 0 Then
   Text44.text = 0
   End If
End Sub

Private Sub Text45_Change()
Dim Txt As String
If Left(Text45.text, 1) = "A" Or Left(Text45.text, 1) = "a" Or _
Left(Text45.text, 1) = "S" Or Left(Text45.text, 1) = "s" Or _
Left(Text45.text, 1) = "D" Or Left(Text45.text, 1) = "d" Or _
Left(Text45.text, 1) = "F" Or Left(Text45.text, 1) = "f" Or _
Left(Text45.text, 1) = "G" Or Left(Text45.text, 1) = "g" Or _
Left(Text45.text, 1) = "H" Or Left(Text45.text, 1) = "h" Or _
Left(Text45.text, 1) = "J" Or Left(Text45.text, 1) = "j" Or _
Left(Text45.text, 1) = "K" Or Left(Text45.text, 1) = "k" Or _
Left(Text45.text, 1) = "L" Or Left(Text45.text, 1) = "l" Then
  Txt = SetStandardBarNumber(Left(Text45.text, 1))
    Text45.text = Txt
      End If
      Text45.text = Val((Text45.text))
If Val(Text45.text) <= 0 Then
   Text45.text = 0
   Text46.text = 0
   End If
End Sub

Private Sub Text46_Change()
Dim Txt As String
If Left(Text46.text, 1) = "A" Or Left(Text46.text, 1) = "a" Or _
Left(Text46.text, 1) = "S" Or Left(Text46.text, 1) = "s" Or _
Left(Text46.text, 1) = "D" Or Left(Text46.text, 1) = "d" Or _
Left(Text46.text, 1) = "F" Or Left(Text46.text, 1) = "f" Or _
Left(Text46.text, 1) = "G" Or Left(Text46.text, 1) = "g" Or _
Left(Text46.text, 1) = "H" Or Left(Text46.text, 1) = "h" Or _
Left(Text46.text, 1) = "J" Or Left(Text46.text, 1) = "j" Or _
Left(Text46.text, 1) = "K" Or Left(Text46.text, 1) = "k" Or _
Left(Text46.text, 1) = "L" Or Left(Text46.text, 1) = "l" Then
  Txt = SetStandardBarSize(Left(Text46.text, 1))
    Text46.text = Txt
      End If
      Text46.text = Val((Text46.text))
If Val(Text46.text) <= 0 Then
   Text46.text = 0
   End If
   If Val(Text46.text) >= 40 Then
   Text46.text = 40
   End If
End Sub

Private Sub Text47_Change()
Dim Txt As String
If Left(Text47.text, 1) = "A" Or Left(Text47.text, 1) = "a" Or _
Left(Text47.text, 1) = "S" Or Left(Text47.text, 1) = "s" Or _
Left(Text47.text, 1) = "D" Or Left(Text47.text, 1) = "d" Or _
Left(Text47.text, 1) = "F" Or Left(Text47.text, 1) = "f" Or _
Left(Text47.text, 1) = "G" Or Left(Text47.text, 1) = "g" Or _
Left(Text47.text, 1) = "H" Or Left(Text47.text, 1) = "h" Or _
Left(Text47.text, 1) = "J" Or Left(Text47.text, 1) = "j" Or _
Left(Text47.text, 1) = "K" Or Left(Text47.text, 1) = "k" Or _
Left(Text47.text, 1) = "L" Or Left(Text47.text, 1) = "l" Then
  Txt = SetStandardBarNumber(Left(Text47.text, 1))
    Text47.text = Txt
      End If
      Text47.text = Val((Text47.text))
If Val(Text47.text) <= 0 Then
   Text47.text = 0
   Text48.text = 0
   End If
End Sub

Private Sub Text48_Change()
Dim Txt As String
If Left(Text48.text, 1) = "A" Or Left(Text48.text, 1) = "a" Or _
Left(Text48.text, 1) = "S" Or Left(Text48.text, 1) = "s" Or _
Left(Text48.text, 1) = "D" Or Left(Text48.text, 1) = "d" Or _
Left(Text48.text, 1) = "F" Or Left(Text48.text, 1) = "f" Or _
Left(Text48.text, 1) = "G" Or Left(Text48.text, 1) = "g" Or _
Left(Text48.text, 1) = "H" Or Left(Text48.text, 1) = "h" Or _
Left(Text48.text, 1) = "J" Or Left(Text48.text, 1) = "j" Or _
Left(Text48.text, 1) = "K" Or Left(Text48.text, 1) = "k" Or _
Left(Text48.text, 1) = "L" Or Left(Text48.text, 1) = "l" Then
  Txt = SetStandardBarSize(Left(Text48.text, 1))
    Text48.text = Txt
      End If
      Text48.text = Val((Text48.text))
If Val(Text48.text) <= 0 Then
   Text48.text = 0
   End If
   If Val(Text48.text) >= 40 Then
   Text48.text = 40
   End If
End Sub

Private Sub Text49_Change()

''namatiang
End Sub

Private Sub Text5_Change()
Dim Txt As String
If Left(Text5.text, 1) = "A" Or Left(Text5.text, 1) = "a" Or _
Left(Text5.text, 1) = "S" Or Left(Text5.text, 1) = "s" Or _
Left(Text5.text, 1) = "D" Or Left(Text5.text, 1) = "d" Or _
Left(Text5.text, 1) = "F" Or Left(Text5.text, 1) = "f" Or _
Left(Text5.text, 1) = "G" Or Left(Text5.text, 1) = "g" Or _
Left(Text5.text, 1) = "H" Or Left(Text5.text, 1) = "h" Or _
Left(Text5.text, 1) = "J" Or Left(Text5.text, 1) = "j" Or _
Left(Text5.text, 1) = "K" Or Left(Text5.text, 1) = "k" Or _
Left(Text5.text, 1) = "L" Or Left(Text5.text, 1) = "l" Then
  Txt = SetStandardBeamBreadth(Left(Text5.text, 1))
    Text5.text = Txt
      End If
      Text5.text = Val((Text5.text))
If Val(Text5.text) <= 0 Then
   Text5.text = 0
   End If
End Sub

Private Sub Text5_GotFocus()
Line19.BorderColor = vbCyan
   Line19.BorderWidth = 2
   Shape20.BorderColor = vbCyan
   Shape20.FillStyle = 0
   Shape20.BorderWidth = 2
End Sub

Private Sub Text5_LostFocus()
Line19.BorderColor = vbBlack
   Line19.BorderWidth = 1
   Shape20.BorderColor = vbBlack
   Shape20.FillStyle = 1
   Shape20.BorderWidth = 1
End Sub

Private Sub Text50_Change()
Dim Txt As String
If Left(Text50.text, 1) = "A" Or Left(Text50.text, 1) = "a" Or _
Left(Text50.text, 1) = "S" Or Left(Text50.text, 1) = "s" Or _
Left(Text50.text, 1) = "D" Or Left(Text50.text, 1) = "d" Or _
Left(Text50.text, 1) = "F" Or Left(Text50.text, 1) = "f" Or _
Left(Text50.text, 1) = "G" Or Left(Text50.text, 1) = "g" Or _
Left(Text50.text, 1) = "H" Or Left(Text50.text, 1) = "h" Or _
Left(Text50.text, 1) = "J" Or Left(Text50.text, 1) = "j" Or _
Left(Text50.text, 1) = "K" Or Left(Text50.text, 1) = "k" Or _
Left(Text50.text, 1) = "L" Or Left(Text50.text, 1) = "l" Then
  Txt = SetStandardCover(Left(Text50.text, 1))
    Text50.text = Txt
      End If
      Text50.text = Val((Text50.text))
If Val(Text50.text) <= 0 Then
   Text50.text = 0
   End If
End Sub

Private Sub Text51_Change()

If Val(Text51.text) <= 0 And Left(Text51.text, 1) <> "-" Then
   Text51.text = 0
   End If
End Sub

Private Sub Text52_Change()

If Val(Text52.text) <= 0 And Left(Text52.text, 1) <> "-" Then
   Text52.text = 0
   End If
End Sub

Private Sub Text53_Change()
Dim Txt As String
If Left(Text53.text, 1) = "A" Or Left(Text53.text, 1) = "a" Or _
Left(Text53.text, 1) = "S" Or Left(Text53.text, 1) = "s" Or _
Left(Text53.text, 1) = "D" Or Left(Text53.text, 1) = "d" Or _
Left(Text53.text, 1) = "F" Or Left(Text53.text, 1) = "f" Or _
Left(Text53.text, 1) = "G" Or Left(Text53.text, 1) = "g" Or _
Left(Text53.text, 1) = "H" Or Left(Text53.text, 1) = "h" Or _
Left(Text53.text, 1) = "J" Or Left(Text53.text, 1) = "j" Or _
Left(Text53.text, 1) = "K" Or Left(Text53.text, 1) = "k" Or _
Left(Text53.text, 1) = "L" Or Left(Text53.text, 1) = "l" Then
  Txt = SetStandardConcreteFcu(Left(Text53.text, 1))
    Text53.text = Txt
      End If
      Text53.text = Val((Text53.text))
If Val(Text53.text) <= 0 Then
   Text53.text = 30
   End If
End Sub

Private Sub Text54_Change()

If Val(Text54.text) <= 0 Then
   Text54.text = 460
   End If
End Sub

Private Sub Text54_Click()
Text54.text = 460
End Sub

Private Sub Text54_DblClick()
Text54.text = 250
End Sub

Private Sub Text55_Change()

If Val(Text55.text) <= 0 Then
   Text55.text = 250
   End If
End Sub

Private Sub Text55_Click()
Text55.text = 250
End Sub

Private Sub Text55_DblClick()
Text55.text = 460
End Sub

Private Sub Text56_Change()
Dim Txt As String
If Left(Text56.text, 1) = "A" Or Left(Text56.text, 1) = "a" Or _
Left(Text56.text, 1) = "S" Or Left(Text56.text, 1) = "s" Or _
Left(Text56.text, 1) = "D" Or Left(Text56.text, 1) = "d" Or _
Left(Text56.text, 1) = "F" Or Left(Text56.text, 1) = "f" Or _
Left(Text56.text, 1) = "G" Or Left(Text56.text, 1) = "g" Or _
Left(Text56.text, 1) = "H" Or Left(Text56.text, 1) = "h" Or _
Left(Text56.text, 1) = "J" Or Left(Text56.text, 1) = "j" Or _
Left(Text56.text, 1) = "K" Or Left(Text56.text, 1) = "k" Or _
Left(Text56.text, 1) = "L" Or Left(Text56.text, 1) = "l" Then
  Txt = SetStandardShrink(Left(Text56.text, 1))
    Text56.text = Txt
      End If
      Text56.text = Val((Text56.text))
If Val(Text56.text) <= 0 Then
   Text56.text = 0.0003
   End If
End Sub

Private Sub Text57_Change()
Dim Txt As String
If Left(Text57.text, 1) = "A" Or Left(Text57.text, 1) = "a" Or _
Left(Text57.text, 1) = "S" Or Left(Text57.text, 1) = "s" Or _
Left(Text57.text, 1) = "D" Or Left(Text57.text, 1) = "d" Or _
Left(Text57.text, 1) = "F" Or Left(Text57.text, 1) = "f" Or _
Left(Text57.text, 1) = "G" Or Left(Text57.text, 1) = "g" Or _
Left(Text57.text, 1) = "H" Or Left(Text57.text, 1) = "h" Or _
Left(Text57.text, 1) = "J" Or Left(Text57.text, 1) = "j" Or _
Left(Text57.text, 1) = "K" Or Left(Text57.text, 1) = "k" Or _
Left(Text57.text, 1) = "L" Or Left(Text57.text, 1) = "l" Then
  Txt = SetStandardCreep(Left(Text57.text, 1))
    Text57.text = Txt
      End If
      Text57.text = Val((Text57.text))
If Val(Text57.text) <= 0 Then
   Text57.text = 2.5
   End If
End Sub

Private Sub Text58_Change()
Dim Txt As String
If Left(Text58.text, 1) = "A" Or Left(Text58.text, 1) = "a" Or _
Left(Text58.text, 1) = "S" Or Left(Text58.text, 1) = "s" Or _
Left(Text58.text, 1) = "D" Or Left(Text58.text, 1) = "d" Or _
Left(Text58.text, 1) = "F" Or Left(Text58.text, 1) = "f" Or _
Left(Text58.text, 1) = "G" Or Left(Text58.text, 1) = "g" Or _
Left(Text58.text, 1) = "H" Or Left(Text58.text, 1) = "h" Or _
Left(Text58.text, 1) = "J" Or Left(Text58.text, 1) = "j" Or _
Left(Text58.text, 1) = "K" Or Left(Text58.text, 1) = "k" Or _
Left(Text58.text, 1) = "L" Or Left(Text58.text, 1) = "l" Then
  Txt = SetStandardBarSize(Left(Text58.text, 1))
    Text58.text = Txt
      End If
      Text58.text = Val((Text58.text))
If Val(Text58.text) <= 0 Then
   Text58.text = 8
   End If
End Sub

Private Sub Text59_Change()
Dim Txt As String
If Left(Text59.text, 1) = "A" Or Left(Text59.text, 1) = "a" Or _
Left(Text59.text, 1) = "S" Or Left(Text59.text, 1) = "s" Or _
Left(Text59.text, 1) = "D" Or Left(Text59.text, 1) = "d" Or _
Left(Text59.text, 1) = "F" Or Left(Text59.text, 1) = "f" Or _
Left(Text59.text, 1) = "G" Or Left(Text59.text, 1) = "g" Or _
Left(Text59.text, 1) = "H" Or Left(Text59.text, 1) = "h" Or _
Left(Text59.text, 1) = "J" Or Left(Text59.text, 1) = "j" Or _
Left(Text59.text, 1) = "K" Or Left(Text59.text, 1) = "k" Or _
Left(Text59.text, 1) = "L" Or Left(Text59.text, 1) = "l" Then
  Txt = SetStandardLinkSpacing(Left(Text59.text, 1))
    Text59.text = Txt
      End If
      Text59.text = Val((Text59.text))
If Val(Text59.text) <= 0 Then
   Text59.text = 200
   End If
End Sub

Private Sub Text6_Change()
Dim Txt As String
If Left(Text6.text, 1) = "A" Or Left(Text6.text, 1) = "a" Or _
Left(Text6.text, 1) = "S" Or Left(Text6.text, 1) = "s" Or _
Left(Text6.text, 1) = "D" Or Left(Text6.text, 1) = "d" Or _
Left(Text6.text, 1) = "F" Or Left(Text6.text, 1) = "f" Or _
Left(Text6.text, 1) = "G" Or Left(Text6.text, 1) = "g" Or _
Left(Text6.text, 1) = "H" Or Left(Text6.text, 1) = "h" Or _
Left(Text6.text, 1) = "J" Or Left(Text6.text, 1) = "j" Or _
Left(Text6.text, 1) = "K" Or Left(Text6.text, 1) = "k" Or _
Left(Text6.text, 1) = "L" Or Left(Text6.text, 1) = "l" Then
  Txt = SetStandardBeamDepth(Left(Text6.text, 1))
    Text6.text = Txt
      End If
      Text6.text = Val((Text6.text))
If Val(Text6.text) <= 0 Then
   Text6.text = 0
   End If
End Sub

Private Sub Text6_GotFocus()
Line19.BorderColor = vbCyan
   Line19.BorderWidth = 2
   Shape20.BorderColor = vbCyan
   Shape20.FillStyle = 0
   Shape20.BorderWidth = 2
End Sub

Private Sub Text6_LostFocus()
Line19.BorderColor = vbBlack
   Line19.BorderWidth = 1
   Shape20.BorderColor = vbBlack
   Shape20.FillStyle = 1
   Shape20.BorderWidth = 1
End Sub

Private Sub Text60_Change()

If Val(Text60.text) <= 0 Then
   Text60.text = 0
   End If
End Sub

Private Sub Text60_GotFocus()
Label61.Visible = True
End Sub

Private Sub Text60_LostFocus()
Label61.Visible = False
End Sub

Private Sub Text61_Change()
  
'''brace xx
End Sub

Private Sub Text61_Click()
Text61.text = "UNBRACED"
End Sub

Private Sub Text61_DblClick()
Text61.text = "BRACED"
End Sub

Private Sub Text61_GotFocus()
 Line21.BorderWidth = 2
   Line23.BorderWidth = 2
   Line25.BorderWidth = 2
   Line21.BorderColor = vbGreen
   Line23.BorderColor = vbCyan
   Line25.BorderColor = vbGreen
  
End Sub

Private Sub Text61_LostFocus()
 Line21.BorderWidth = 1
   Line23.BorderWidth = 1
   Line25.BorderWidth = 1
   Line21.BorderColor = vbBlack
   Line23.BorderColor = vbBlack
   Line25.BorderColor = vbBlack
  
End Sub

Private Sub Text62_Change()
   
'''brace yy
End Sub

Private Sub Text62_Click()
Text62.text = "UNBRACED"
End Sub

Private Sub Text62_DblClick()
Text62.text = "BRACED"
End Sub

Private Sub Text62_GotFocus()
Line22.BorderWidth = 2
   Line24.BorderWidth = 2
   Line26.BorderWidth = 2
   Line22.BorderColor = vbYellow
   Line24.BorderColor = vbCyan
   Line26.BorderColor = vbYellow

End Sub

Private Sub Text62_LostFocus()
Line22.BorderWidth = 1
   Line24.BorderWidth = 1
   Line26.BorderWidth = 1
   Line22.BorderColor = vbBlack
   Line24.BorderColor = vbBlack
   Line26.BorderColor = vbBlack
End Sub

Private Sub Text7_Change()
Dim Txt As String
If Left(Text7.text, 1) = "A" Or Left(Text7.text, 1) = "a" Or _
Left(Text7.text, 1) = "S" Or Left(Text7.text, 1) = "s" Or _
Left(Text7.text, 1) = "D" Or Left(Text7.text, 1) = "d" Or _
Left(Text7.text, 1) = "F" Or Left(Text7.text, 1) = "f" Or _
Left(Text7.text, 1) = "G" Or Left(Text7.text, 1) = "g" Or _
Left(Text7.text, 1) = "H" Or Left(Text7.text, 1) = "h" Or _
Left(Text7.text, 1) = "J" Or Left(Text7.text, 1) = "j" Or _
Left(Text7.text, 1) = "K" Or Left(Text7.text, 1) = "k" Or _
Left(Text7.text, 1) = "L" Or Left(Text7.text, 1) = "l" Then
  Txt = SetStandardBeamLength(Left(Text7.text, 1))
    Text7.text = Txt
      End If
      Text7.text = Val((Text7.text))
If Val(Text7.text) <= 0 Then
   Text7.text = 0
   End If
End Sub

Private Sub Text7_GotFocus()
Line19.BorderColor = vbCyan
   Line19.BorderWidth = 2
   Shape20.BorderColor = vbCyan
   Shape20.FillStyle = 0
   Shape20.BorderWidth = 2
End Sub

Private Sub Text7_LostFocus()
Line19.BorderColor = vbBlack
   Line19.BorderWidth = 1
   Shape20.BorderColor = vbBlack
   Shape20.FillStyle = 1
   Shape20.BorderWidth = 1
End Sub

Private Sub Text8_Change()
 Dim Txt As String
If Left(Text8.text, 1) = "A" Or Left(Text8.text, 1) = "a" Or _
Left(Text8.text, 1) = "S" Or Left(Text8.text, 1) = "s" Or _
Left(Text8.text, 1) = "D" Or Left(Text8.text, 1) = "d" Or _
Left(Text8.text, 1) = "F" Or Left(Text8.text, 1) = "f" Or _
Left(Text8.text, 1) = "G" Or Left(Text8.text, 1) = "g" Or _
Left(Text8.text, 1) = "H" Or Left(Text8.text, 1) = "h" Or _
Left(Text8.text, 1) = "J" Or Left(Text8.text, 1) = "j" Or _
Left(Text8.text, 1) = "K" Or Left(Text8.text, 1) = "k" Or _
Left(Text8.text, 1) = "L" Or Left(Text8.text, 1) = "l" Then
  Txt = SetStandardBeamBreadth(Left(Text8.text, 1))
    Text8.text = Txt
      End If
      Text8.text = Val((Text8.text))
If Val(Text8.text) <= 0 Then
   Text8.text = 0
   End If
End Sub

Private Sub Text8_GotFocus()
Line18.BorderColor = vbCyan
   Line18.BorderWidth = 2
   Shape19.BorderColor = vbCyan
   Shape19.FillStyle = 0
   Shape19.BorderWidth = 2
End Sub

Private Sub Text8_LostFocus()
Line18.BorderColor = vbBlack
   Line18.BorderWidth = 1
   Shape19.BorderColor = vbBlack
   Shape19.FillStyle = 1
   Shape19.BorderWidth = 1
End Sub

Private Sub Text9_Change()
Dim Txt As String
If Left(Text9.text, 1) = "A" Or Left(Text9.text, 1) = "a" Or _
Left(Text9.text, 1) = "S" Or Left(Text9.text, 1) = "s" Or _
Left(Text9.text, 1) = "D" Or Left(Text9.text, 1) = "d" Or _
Left(Text9.text, 1) = "F" Or Left(Text9.text, 1) = "f" Or _
Left(Text9.text, 1) = "G" Or Left(Text9.text, 1) = "g" Or _
Left(Text9.text, 1) = "H" Or Left(Text9.text, 1) = "h" Or _
Left(Text9.text, 1) = "J" Or Left(Text9.text, 1) = "j" Or _
Left(Text9.text, 1) = "K" Or Left(Text9.text, 1) = "k" Or _
Left(Text9.text, 1) = "L" Or Left(Text9.text, 1) = "l" Then
  Txt = SetStandardBeamDepth(Left(Text9.text, 1))
    Text9.text = Txt
      End If
      Text9.text = Val((Text9.text))
If Val(Text9.text) <= 0 Then
   Text9.text = 0
   End If
End Sub

Private Function HighlightMember(ByVal N As Integer)

  
End Function

Private Sub ShowExistingMember(ByVal N As Integer)


If UppCoB(N) = 0 Or UppCoH(N) = 0 Or UppCoHgt(N) = 0 Then
   Shape23.Visible = False
   Else
   Shape23.Visible = True
   End If
   
'''xBASEresistmnt(N)
  
If yUppLeBb(N) = 0 Or yUppLeBh(N) = 0 Or yUppLeBl(N) = 0 Then
   Line19.Visible = False
   Shape20.Visible = False
   Else
   Line19.Visible = True
   Shape20.Visible = True
   End If
   
If yUppRiBb(N) = 0 Or yUppRiBh(N) = 0 Or yUppRiBl(N) = 0 Then
   Line18.Visible = False
   Shape19.Visible = False
   Else
   Line18.Visible = True
   Shape19.Visible = True
   End If

If xUppRiBb(N) = 0 Or xUppRiBh(N) = 0 Or xUppRiBl(N) = 0 Then
   Line16.Visible = False
   Shape17.Visible = False
   Else
   Line16.Visible = True
   Shape17.Visible = True
   End If

If xUppLeBb(N) = 0 Or xUppLeBh(N) = 0 Or xUppLeBl(N) = 0 Then
   Line17.Visible = False
   Shape18.Visible = False
   Else
   Line17.Visible = True
   Shape18.Visible = True
   End If

'''bTng(1), hTng(1), FloorHgt(1)

''' Mx2(1), My2(1)
''' Mx1(1), My1(1)

If xLowLeBb(N) = 0 Or xLowLeBh(N) = 0 Or xLowLeBl(N) = 0 Then
   Line13.Visible = False
   Shape14.Visible = False
   Else
   Line13.Visible = True
   Shape14.Visible = True
   End If

If xLowRiBb(N) = 0 Or xLowRiBh(N) = 0 Or xLowRiBl(N) = 0 Then
   Line12.Visible = False
   Shape13.Visible = False
   Else
   Line12.Visible = True
   Shape13.Visible = True
   End If

If yLowRiBb(N) = 0 Or yLowRiBh(N) = 0 Or yLowRiBl(N) = 0 Then
   Line14.Visible = False
   Shape15.Visible = False
   Else
   Line14.Visible = True
   Shape15.Visible = True
   End If

If yLowLeBb(N) = 0 Or yLowLeBh(N) = 0 Or yLowLeBl(N) = 0 Then
   Line15.Visible = False
   Shape16.Visible = False
   Else
   Line15.Visible = True
   Shape16.Visible = True
   End If

If LowCoB(N) = 0 Or LowCoH(N) = 0 Or LowCoHgt(N) = 0 Then
   Shape21.Visible = False
   Shape24.Visible = False
   Shape25.Shape = 0
   Else
   Shape21.Visible = True
   Shape24.Visible = True
   Shape25.Shape = 3
   End If

'''yBASEresistmnt(N)
      
'''DesgAxial(1)

If BarYLefN(N) = 0 Or BarYLefD(N) = 0 Then
   Shape7.Visible = False
   Else
   Shape7.Visible = True
   End If
   
   If BarYLefN(N) = 1 And BarYLefD(N) <> 0 Then
   ''Shape1.Visible = False
   Shape7.Visible = True
   ''Shape4.Visible = False
   End If
   
   If BarYLefN(N) = 2 And BarYLefD(N) <> 0 Then
   Shape1.Visible = True
   Shape7.Visible = True
   Shape4.Visible = True
   End If
   
If BarYRigN(N) = 0 Or BarYRigD(N) = 0 Then
   Shape8.Visible = False
   Else
   Shape8.Visible = True
   End If

   If BarYRigN(N) = 1 And BarYRigD(N) <> 0 Then
   ''Shape3.Visible = False
   Shape8.Visible = True
   ''Shape6.Visible = False
   End If
   
   If BarYRigN(N) = 2 And BarYRigD(N) <> 0 Then
   Shape3.Visible = True
   Shape8.Visible = True
   Shape6.Visible = True
   End If

If BarXTopN(N) = 0 Or BarXTopD(N) = 0 Then
   Shape1.Visible = False
   Shape2.Visible = False
   Shape3.Visible = False
   End If
   
   If BarXTopN(N) = 1 And BarXTopD(N) <> 0 Then
   Shape1.Visible = False
   Shape2.Visible = True
   Shape3.Visible = False
   End If
     
   If BarXTopN(N) = 2 And BarXTopD(N) <> 0 Then
   Shape1.Visible = True
   Shape2.Visible = False
   Shape3.Visible = True
   End If
   
   If BarXTopN(N) > 2 And BarXTopD(N) <> 0 Then
   Shape1.Visible = True
   Shape2.Visible = True
   Shape3.Visible = True
   End If

If BarXBotN(N) = 0 Or BarXBotD(N) = 0 Then
   Shape4.Visible = False
   Shape5.Visible = False
   Shape6.Visible = False
   End If

   If BarXBotN(N) = 1 And BarXBotD(N) <> 0 Then
   Shape4.Visible = False
   Shape5.Visible = True
   Shape6.Visible = False
   End If
   
   If BarXBotN(N) = 2 And BarXBotD(N) <> 0 Then
   Shape4.Visible = True
   Shape5.Visible = False
   Shape6.Visible = True
   End If
   
   If BarXBotN(N) > 2 And BarXBotD(N) <> 0 Then
   Shape4.Visible = True
   Shape5.Visible = True
   Shape6.Visible = True
   End If
  
'''GridTng(1)
'''CoverTng(1)

If xBracedCol(N) = "UNBRACED" Then
   Line21.Visible = False
   Line23.Visible = False
   Line25.Visible = False
     Else
   Line21.Visible = True
   Line23.Visible = True
   Line25.Visible = True
   End If
   
If yBracedCol(N) = "UNBRACED" Then
   Line22.Visible = False
   Line24.Visible = False
   Line26.Visible = False
     Else
   Line22.Visible = True
   Line24.Visible = True
   Line26.Visible = True
   End If


End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ShowAllVisibility()

 
  If Shape23.Visible = False Then
     Shape23.Visible = True
     End If
      
  If Line19.Visible = False Then
     Line19.Visible = True
     End If
   
 If Shape20.Visible = False Then
    Shape20.Visible = True
    End If
 
 If Line18.Visible = False Then
    Line18.Visible = True
 End If
 
 If Shape19.Visible = False Then
    Shape19.Visible = True
 End If
  
 If Line16.Visible = False Then
    Line16.Visible = True
 End If
 
If Shape17.Visible = False Then
   Shape17.Visible = True
   End If
   
 
  If Line17.Visible = False Then
     Line17.Visible = True
  End If
  
 If Shape18.Visible = False Then
     Shape18.Visible = True
     End If
     
     
 If Line13.Visible = False Then
    Line13.Visible = True
    End If
    
    
    
If Shape14.Visible = False Then
    Shape14.Visible = True
End If

   
 If Line12.Visible = False Then
     Line12.Visible = True
     End If
 
 If Shape13.Visible = False Then
    Shape13.Visible = True
    End If
 
 
 If Line14.Visible = False Then
     Line14.Visible = True
     End If
 
 If Shape15.Visible = False Then
     Shape15.Visible = True
 End If
  
 If Line15.Visible = False Then
     Line15.Visible = True
     End If
     
 
If Shape16.Visible = False Then
   Shape16.Visible = True
   End If
  
   If Shape21.Visible = False Then
    Shape21.Visible = True
   End If
   
   
   If Shape24.Visible = False Then
     Shape24.Visible = True
  End If
   
 ''Shape25.Shape = 3 ''
  
  If Shape1.Visible = False Then
   Shape1.Visible = True
 End If
  
If Shape2.Visible = False Then
Shape2.Visible = True
End If

  If Shape3.Visible = False Then
  Shape3.Visible = True
  End If
  
 If Shape4.Visible = False Then
 Shape4.Visible = True
 End If
 
 
  If Shape5.Visible = False Then
  Shape5.Visible = True
 End If
  
  
  
If Shape6.Visible = False Then
Shape6.Visible = True
End If




  If Shape7.Visible = False Then
  Shape7.Visible = True
  End If
  
  If Shape8.Visible = False Then
  Shape8.Visible = True
  End If
  
  
  If Line21.Visible = False Then
   Line21.Visible = True
   End If
   
   
 If Line23.Visible = False Then
  Line23.Visible = True
  
  End If
  
  
 If Line25.Visible = False Then
  Line25.Visible = True
 End If
  
  
If Line22.Visible = False Then
 Line22.Visible = True
 End If
 
 
 If Line24.Visible = False Then
 Line24.Visible = True
 End If
 
If Line26.Visible = False Then
Line26.Visible = True
End If

Line4.BorderColor = vbGreen
Line4.BorderWidth = 3

End Sub

Private Sub CloseAllVisibility()

   Line4.BorderColor = vbWhite

   Shape23.Visible = False
      
   Line19.Visible = False
   Shape20.Visible = False
 
   Line18.Visible = False
   Shape19.Visible = False
  
   Line16.Visible = False
   Shape17.Visible = False
 
   Line17.Visible = False
   Shape18.Visible = False
   
   Line13.Visible = False
   Shape14.Visible = False
   
   Line12.Visible = False
   Shape13.Visible = False
 
   Line14.Visible = False
   Shape15.Visible = False
  
   Line15.Visible = False
   Shape16.Visible = False
  
   Shape21.Visible = False
   Shape24.Visible = False
   Shape25.Shape = 3
  
   Shape1.Visible = False
   Shape2.Visible = False
   Shape3.Visible = False
   Shape4.Visible = False
   Shape5.Visible = False
   Shape6.Visible = False
   Shape7.Visible = False
   Shape8.Visible = False
  
   Line21.Visible = False
   Line23.Visible = False
   Line25.Visible = False
  
   Line22.Visible = False
   Line24.Visible = False
   Line26.Visible = False

End Sub


Private Sub Text9_GotFocus()
Line18.BorderColor = vbCyan
   Line18.BorderWidth = 2
   Shape19.BorderColor = vbCyan
   Shape19.FillStyle = 0
   Shape19.BorderWidth = 2
End Sub

Private Sub Text9_LostFocus()
Line18.BorderColor = vbBlack
   Line18.BorderWidth = 1
   Shape19.BorderColor = vbBlack
   Shape19.FillStyle = 1
   Shape19.BorderWidth = 1
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function SetStandardBarNumber(ByVal SetBar As String)

 If SetBar = "A" Or SetBar = "a" Then
    SetStandardBarNumber = "1"
    End If
 If SetBar = "S" Or SetBar = "s" Then
    SetStandardBarNumber = "2"
    End If
 If SetBar = "D" Or SetBar = "d" Then
    SetStandardBarNumber = "3"
    End If
 If SetBar = "F" Or SetBar = "f" Then
    SetStandardBarNumber = "4"
    End If
 If SetBar = "G" Or SetBar = "g" Then
    SetStandardBarNumber = "5"
    End If
 If SetBar = "H" Or SetBar = "h" Then
    SetStandardBarNumber = "6"
    End If
 If SetBar = "J" Or SetBar = "j" Then
    SetStandardBarNumber = "7"
    End If
 If SetBar = "K" Or SetBar = "k" Then
    SetStandardBarNumber = "8"
    End If
If SetBar = "L" Or SetBar = "l" Then
    SetStandardBarNumber = "9"
    End If

End Function

Private Function SetStandardMoment(ByVal SetBar As String)

 If SetBar = "A" Or SetBar = "a" Then
    SetStandardMoment = "50"
    End If
 If SetBar = "S" Or SetBar = "s" Then
    SetStandardMoment = "75"
    End If
 If SetBar = "D" Or SetBar = "d" Then
    SetStandardMoment = "100"
    End If
 If SetBar = "F" Or SetBar = "f" Then
    SetStandardMoment = "125"
    End If
 If SetBar = "G" Or SetBar = "g" Then
    SetStandardMoment = "150"
    End If
 If SetBar = "H" Or SetBar = "h" Then
    SetStandardMoment = "175"
    End If
 If SetBar = "J" Or SetBar = "j" Then
    SetStandardMoment = "200"
    End If
 If SetBar = "K" Or SetBar = "k" Then
    SetStandardMoment = "225"
    End If
If SetBar = "L" Or SetBar = "l" Then
    SetStandardMoment = "250"
    End If
If SetBar = "P" Or SetBar = "p" Then
    SetStandardMoment = "275"
    End If
If SetBar = "O" Or SetBar = "o" Then
    SetStandardMoment = "300"
    End If
    
End Function

Private Function SetStandardColumnDepth(ByVal SetBar As String)

 If SetBar = "A" Or SetBar = "a" Then
    SetStandardColumnDepth = "200"
    End If
 If SetBar = "S" Or SetBar = "s" Then
    SetStandardColumnDepth = "250"
    End If
 If SetBar = "D" Or SetBar = "d" Then
    SetStandardColumnDepth = "300"
    End If
 If SetBar = "F" Or SetBar = "f" Then
    SetStandardColumnDepth = "350"
    End If
 If SetBar = "G" Or SetBar = "g" Then
    SetStandardColumnDepth = "400"
    End If
 If SetBar = "H" Or SetBar = "h" Then
    SetStandardColumnDepth = "450"
    End If
 If SetBar = "J" Or SetBar = "j" Then
    SetStandardColumnDepth = "500"
    End If
 If SetBar = "K" Or SetBar = "k" Then
    SetStandardColumnDepth = "550"
    End If
If SetBar = "L" Or SetBar = "l" Then
    SetStandardColumnDepth = "600"
    End If

End Function

Private Function SetStandardBeamBreadth(ByVal SetBar As String)

 If SetBar = "A" Or SetBar = "a" Then
    SetStandardBeamBreadth = "200"
    End If
 If SetBar = "S" Or SetBar = "s" Then
    SetStandardBeamBreadth = "225"
    End If
 If SetBar = "D" Or SetBar = "d" Then
    SetStandardBeamBreadth = "250"
    End If
 If SetBar = "F" Or SetBar = "f" Then
    SetStandardBeamBreadth = "275"
    End If
 If SetBar = "G" Or SetBar = "g" Then
    SetStandardBeamBreadth = "300"
    End If
 If SetBar = "H" Or SetBar = "h" Then
    SetStandardBeamBreadth = "325"
    End If
 If SetBar = "J" Or SetBar = "j" Then
    SetStandardBeamBreadth = "350"
    End If
 If SetBar = "K" Or SetBar = "k" Then
    SetStandardBeamBreadth = "375"
    End If
If SetBar = "L" Or SetBar = "l" Then
    SetStandardBeamBreadth = "400"
    End If

End Function


Private Function SetStandardBeamDepth(ByVal SetBar As String)

 If SetBar = "A" Or SetBar = "a" Then
    SetStandardBeamDepth = "400"
    End If
 If SetBar = "S" Or SetBar = "s" Then
    SetStandardBeamDepth = "450"
    End If
 If SetBar = "D" Or SetBar = "d" Then
    SetStandardBeamDepth = "500"
    End If
 If SetBar = "F" Or SetBar = "f" Then
    SetStandardBeamDepth = "550"
    End If
 If SetBar = "G" Or SetBar = "g" Then
    SetStandardBeamDepth = "600"
    End If
 If SetBar = "H" Or SetBar = "h" Then
    SetStandardBeamDepth = "650"
    End If
 If SetBar = "J" Or SetBar = "j" Then
    SetStandardBeamDepth = "700"
    End If
 If SetBar = "K" Or SetBar = "k" Then
    SetStandardBeamDepth = "750"
    End If
If SetBar = "L" Or SetBar = "l" Then
    SetStandardBeamDepth = "800"
    End If

End Function


Private Function SetStandardConcreteFcu(ByVal SetBar As String)

 If SetBar = "A" Or SetBar = "a" Then
    SetStandardConcreteFcu = "20"
    End If
 If SetBar = "S" Or SetBar = "s" Then
    SetStandardConcreteFcu = "25"
    End If
 If SetBar = "D" Or SetBar = "d" Then
    SetStandardConcreteFcu = "30"
    End If
 If SetBar = "F" Or SetBar = "f" Then
    SetStandardConcreteFcu = "35"
    End If
 If SetBar = "G" Or SetBar = "g" Then
    SetStandardConcreteFcu = "40"
    End If
 If SetBar = "H" Or SetBar = "h" Then
    SetStandardConcreteFcu = "45"
    End If
 If SetBar = "J" Or SetBar = "j" Then
    SetStandardConcreteFcu = "50"
    End If
 If SetBar = "K" Or SetBar = "k" Then
    SetStandardConcreteFcu = "55"
    End If
If SetBar = "L" Or SetBar = "l" Then
    SetStandardConcreteFcu = "60"
    End If

End Function



Private Function SetStandardBarSize(ByVal SetBar As String)

 If SetBar = "A" Or SetBar = "a" Then
    SetStandardBarSize = "6"
    End If
 If SetBar = "S" Or SetBar = "s" Then
    SetStandardBarSize = "8"
    End If
 If SetBar = "D" Or SetBar = "d" Then
    SetStandardBarSize = "10"
    End If
 If SetBar = "F" Or SetBar = "f" Then
    SetStandardBarSize = "12"
    End If
 If SetBar = "G" Or SetBar = "g" Then
    SetStandardBarSize = "16"
    End If
 If SetBar = "H" Or SetBar = "h" Then
    SetStandardBarSize = "20"
    End If
 If SetBar = "J" Or SetBar = "j" Then
    SetStandardBarSize = "25"
    End If
 If SetBar = "K" Or SetBar = "k" Then
    SetStandardBarSize = "32"
    End If
If SetBar = "L" Or SetBar = "l" Then
    SetStandardBarSize = "40"
    End If

End Function


Private Function SetStandardCover(ByVal SetBar As String)

 If SetBar = "A" Or SetBar = "a" Then
    SetStandardCover = "15"
    End If
 If SetBar = "S" Or SetBar = "s" Then
    SetStandardCover = "20"
    End If
 If SetBar = "D" Or SetBar = "d" Then
    SetStandardCover = "25"
    End If
 If SetBar = "F" Or SetBar = "f" Then
    SetStandardCover = "30"
    End If
 If SetBar = "G" Or SetBar = "g" Then
    SetStandardCover = "35"
    End If
 If SetBar = "H" Or SetBar = "h" Then
    SetStandardCover = "40"
    End If
 If SetBar = "J" Or SetBar = "j" Then
    SetStandardCover = "45"
    End If
    
If SetBar = "K" Or SetBar = "k" Then
    SetStandardCover = "50"
    End If
If SetBar = "L" Or SetBar = "l" Then
    SetStandardCover = "55"
    End If
    
End Function

Private Function SetStandardCreep(ByVal SetBar As String)

 If SetBar = "A" Or SetBar = "a" Then
    SetStandardCreep = "1.0"
    End If
 If SetBar = "S" Or SetBar = "s" Then
    SetStandardCreep = "1.5"
    End If
 If SetBar = "D" Or SetBar = "d" Then
    SetStandardCreep = "2.0"
    End If
 If SetBar = "F" Or SetBar = "f" Then
    SetStandardCreep = "2.5"
    End If
 If SetBar = "G" Or SetBar = "g" Then
    SetStandardCreep = "3.0"
    End If
 If SetBar = "H" Or SetBar = "h" Then
    SetStandardCreep = "3.5"
    End If
 If SetBar = "J" Or SetBar = "j" Then
    SetStandardCreep = "4.0"
    End If
 
If SetBar = "K" Or SetBar = "k" Then
    SetStandardCreep = "2.7"
    End If
If SetBar = "L" Or SetBar = "l" Then
    SetStandardCreep = "3.3"
    End If
    
End Function



Private Function SetStandardShrink(ByVal SetBar As String)

 If SetBar = "A" Or SetBar = "a" Then
    SetStandardShrink = "0.0001"
    End If
 If SetBar = "S" Or SetBar = "s" Then
    SetStandardShrink = "0.0002"
    End If
 If SetBar = "D" Or SetBar = "d" Then
    SetStandardShrink = "0.0003"
    End If
 If SetBar = "F" Or SetBar = "f" Then
    SetStandardShrink = "0.0004"
    End If
 
 If SetBar = "G" Or SetBar = "g" Then
    SetStandardShrink = "0.00015"
    End If
 If SetBar = "H" Or SetBar = "h" Then
    SetStandardShrink = "0.00025"
    End If
 If SetBar = "J" Or SetBar = "j" Then
    SetStandardShrink = "0.00035"
    End If
 If SetBar = "K" Or SetBar = "k" Then
    SetStandardShrink = "0.00042"
    End If
If SetBar = "L" Or SetBar = "l" Then
    SetStandardShrink = "0.00045"
    End If
    
End Function

Private Function SetStandardColumnHeight(ByVal SetBar As String)

 If SetBar = "A" Or SetBar = "a" Then
    SetStandardColumnHeight = "1500"
    End If
 If SetBar = "S" Or SetBar = "s" Then
    SetStandardColumnHeight = "3200"
    End If
 If SetBar = "D" Or SetBar = "d" Then
    SetStandardColumnHeight = "3600"
    End If
 If SetBar = "F" Or SetBar = "f" Then
    SetStandardColumnHeight = "3900"
    End If
 If SetBar = "G" Or SetBar = "g" Then
    SetStandardColumnHeight = "4000"
    End If
 If SetBar = "H" Or SetBar = "h" Then
    SetStandardColumnHeight = "4200"
    End If
 If SetBar = "J" Or SetBar = "j" Then
    SetStandardColumnHeight = "5000"
    End If
 If SetBar = "K" Or SetBar = "k" Then
    SetStandardColumnHeight = "7200"
    End If
If SetBar = "L" Or SetBar = "l" Then
    SetStandardColumnHeight = "8000"
    End If

End Function



Private Function SetStandardBeamLength(ByVal SetBar As String)

 If SetBar = "A" Or SetBar = "a" Then
    SetStandardBeamLength = "3000"
    End If
 If SetBar = "S" Or SetBar = "s" Then
    SetStandardBeamLength = "4500"
    End If
 If SetBar = "D" Or SetBar = "d" Then
    SetStandardBeamLength = "6000"
    End If
 If SetBar = "F" Or SetBar = "f" Then
    SetStandardBeamLength = "7500"
    End If
 If SetBar = "G" Or SetBar = "g" Then
    SetStandardBeamLength = "7500"
    End If
 If SetBar = "H" Or SetBar = "h" Then
    SetStandardBeamLength = "7800"
    End If
 If SetBar = "J" Or SetBar = "j" Then
    SetStandardBeamLength = "9500"
    End If
 If SetBar = "K" Or SetBar = "k" Then
    SetStandardBeamLength = "9800"
    End If
If SetBar = "L" Or SetBar = "l" Then
    SetStandardBeamLength = "10000"
    End If

End Function
Private Function SetStandardColumnLoad(ByVal SetBar As String)

 If SetBar = "A" Or SetBar = "a" Then
    SetStandardColumnLoad = "225"
    End If
 If SetBar = "S" Or SetBar = "s" Then
    SetStandardColumnLoad = "500"
    End If
 If SetBar = "D" Or SetBar = "d" Then
    SetStandardColumnLoad = "750"
    End If
 If SetBar = "F" Or SetBar = "f" Then
    SetStandardColumnLoad = "1000"
    End If
 If SetBar = "G" Or SetBar = "g" Then
    SetStandardColumnLoad = "1250"
    End If
 If SetBar = "H" Or SetBar = "h" Then
    SetStandardColumnLoad = "1500"
    End If
 If SetBar = "J" Or SetBar = "j" Then
    SetStandardColumnLoad = "1750"
    End If
 If SetBar = "K" Or SetBar = "k" Then
    SetStandardColumnLoad = "2000"
    End If
If SetBar = "L" Or SetBar = "l" Then
    SetStandardColumnLoad = "2250"
    End If

End Function


Private Function SetStandardLinkSpacing(ByVal SetBar As String)

 If SetBar = "A" Or SetBar = "a" Then
    SetStandardLinkSpacing = "50"
    End If
 If SetBar = "S" Or SetBar = "s" Then
    SetStandardLinkSpacing = "75"
    End If
 If SetBar = "D" Or SetBar = "d" Then
    SetStandardLinkSpacing = "100"
    End If
 If SetBar = "F" Or SetBar = "f" Then
    SetStandardLinkSpacing = "125"
    End If
 If SetBar = "G" Or SetBar = "g" Then
    SetStandardLinkSpacing = "150"
    End If
 If SetBar = "H" Or SetBar = "h" Then
    SetStandardLinkSpacing = "175"
    End If
 If SetBar = "J" Or SetBar = "j" Then
    SetStandardLinkSpacing = "200"
    End If
 If SetBar = "K" Or SetBar = "k" Then
    SetStandardLinkSpacing = "225"
    End If
If SetBar = "L" Or SetBar = "l" Then
    SetStandardLinkSpacing = "250"
    End If
If SetBar = "P" Or SetBar = "p" Then
    SetStandardLinkSpacing = "275"
    End If
If SetBar = "O" Or SetBar = "o" Then
    SetStandardLinkSpacing = "300"
    End If
End Function
