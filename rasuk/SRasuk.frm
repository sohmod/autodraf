VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reka & Lukis Rasuk (versi:04/2011)"
   ClientHeight    =   8235
   ClientLeft      =   345
   ClientTop       =   810
   ClientWidth     =   8130
   FillColor       =   &H000080FF&
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000009&
   Icon            =   "SRasuk.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "SRasuk.frx":030A
   ScaleHeight     =   9000
   ScaleMode       =   0  'User
   ScaleWidth      =   8130
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "Kemaskini data"
      Height          =   274
      Left            =   2400
      TabIndex        =   103
      Top             =   595
      Width           =   3255
   End
   Begin VB.TextBox Text57 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   315
      Left            =   3600
      TabIndex        =   100
      Text            =   "50"
      Top             =   1200
      Width           =   700
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C000C0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   420
      Left            =   600
      MultiSelect     =   1  'Simple
      OLEDropMode     =   1  'Manual
      TabIndex        =   99
      Top             =   4200
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Calc.Strength"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   98
      Top             =   3720
      Width           =   1335
   End
   Begin VB.OptionButton Option9 
      BackColor       =   &H0080C0FF&
      Caption         =   "Option9"
      Enabled         =   0   'False
      Height          =   255
      Left            =   5040
      TabIndex        =   26
      Top             =   3960
      Width           =   255
   End
   Begin VB.TextBox Text56 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6120
      MaxLength       =   5
      TabIndex        =   97
      Text            =   "700"
      Top             =   6960
      Width           =   615
   End
   Begin VB.TextBox Text55 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6720
      MaxLength       =   3
      TabIndex        =   96
      Text            =   "20"
      Top             =   6960
      Width           =   615
   End
   Begin VB.TextBox Text54 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7320
      MaxLength       =   2
      TabIndex        =   95
      Text            =   "2"
      Top             =   6960
      Width           =   600
   End
   Begin VB.TextBox Text53 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1440
      MaxLength       =   5
      TabIndex        =   94
      Text            =   "700"
      Top             =   6960
      Width           =   600
   End
   Begin VB.TextBox Text52 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   840
      MaxLength       =   3
      TabIndex        =   93
      Text            =   "20"
      Top             =   6960
      Width           =   600
   End
   Begin VB.TextBox Text51 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      MaxLength       =   2
      TabIndex        =   92
      Text            =   "2"
      Top             =   6960
      Width           =   600
   End
   Begin VB.TextBox Text50 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4080
      MaxLength       =   3
      TabIndex        =   91
      Text            =   "16"
      Top             =   5760
      Width           =   600
   End
   Begin VB.TextBox Text49 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3480
      MaxLength       =   2
      TabIndex        =   90
      Text            =   "2"
      Top             =   5760
      Width           =   600
   End
   Begin VB.TextBox Text48 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000014&
      Enabled         =   0   'False
      ForeColor       =   &H80000002&
      Height          =   329
      Left            =   7200
      TabIndex        =   89
      Text            =   "0"
      ToolTipText     =   "Nilai numerikal +ve sahaja !"
      Top             =   3240
      Width           =   800
   End
   Begin VB.TextBox Text47 
      BackColor       =   &H80000014&
      Enabled         =   0   'False
      ForeColor       =   &H80000010&
      Height          =   329
      Left            =   80
      TabIndex        =   88
      Text            =   "0"
      ToolTipText     =   "Nilai numerikal +ve sahaja !"
      Top             =   3240
      Width           =   800
   End
   Begin VB.OptionButton Option8 
      BackColor       =   &H0080C0FF&
      Caption         =   "Option8"
      Enabled         =   0   'False
      Height          =   255
      Left            =   4440
      TabIndex        =   25
      Top             =   3960
      Width           =   240
   End
   Begin VB.OptionButton Option7 
      BackColor       =   &H0080C0FF&
      Caption         =   "Option7"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3840
      TabIndex        =   24
      Top             =   3960
      Width           =   240
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ok!"
      Enabled         =   0   'False
      Height          =   1275
      Left            =   6000
      TabIndex        =   15
      Top             =   1200
      Width           =   225
   End
   Begin VB.TextBox Text46 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
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
      Height          =   315
      Left            =   4320
      TabIndex        =   85
      Text            =   "101"
      Top             =   2160
      Width           =   1545
   End
   Begin VB.TextBox Text45 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
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
      Height          =   315
      Left            =   3600
      MaxLength       =   3
      TabIndex        =   84
      Text            =   "10"
      Top             =   2160
      Width           =   700
   End
   Begin VB.TextBox Text44 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
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
      Height          =   315
      Left            =   2880
      MaxLength       =   4
      TabIndex        =   83
      Text            =   "150"
      Top             =   2160
      Width           =   700
   End
   Begin VB.TextBox Text43 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
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
      Height          =   315
      Left            =   2160
      MaxLength       =   3
      TabIndex        =   82
      Text            =   "25"
      Top             =   2160
      Width           =   700
   End
   Begin VB.TextBox Text42 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
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
      Height          =   315
      Left            =   4320
      MaxLength       =   4
      TabIndex        =   81
      Text            =   "2.5"
      Top             =   1800
      Width           =   700
   End
   Begin VB.TextBox Text41 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
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
      Height          =   315
      Left            =   5040
      MaxLength       =   7
      TabIndex        =   80
      Text            =   "0.0003"
      Top             =   1800
      Width           =   825
   End
   Begin VB.TextBox Text40 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
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
      Height          =   315
      Left            =   3600
      MaxLength       =   4
      TabIndex        =   79
      Text            =   "250"
      Top             =   1800
      Width           =   700
   End
   Begin VB.TextBox Text39 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
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
      Height          =   315
      Left            =   2880
      MaxLength       =   4
      TabIndex        =   78
      Text            =   "460"
      Top             =   1800
      Width           =   700
   End
   Begin VB.TextBox Text38 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
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
      Height          =   315
      Left            =   2160
      MaxLength       =   3
      TabIndex        =   77
      Text            =   "30"
      Top             =   1800
      Width           =   700
   End
   Begin VB.TextBox Text37 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   315
      Left            =   4320
      MultiLine       =   -1  'True
      TabIndex        =   76
      Text            =   "SRasuk.frx":0614
      Top             =   1200
      Width           =   1560
   End
   Begin VB.TextBox Text36 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
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
      Height          =   315
      Left            =   2880
      TabIndex        =   75
      Text            =   "0"
      Top             =   1200
      Width           =   700
   End
   Begin VB.TextBox Text35 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
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
      Height          =   315
      Left            =   2160
      TabIndex        =   70
      Text            =   "0"
      Top             =   1200
      Width           =   700
   End
   Begin VB.TextBox Text34 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000004&
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
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   330
      Left            =   6720
      MaxLength       =   4
      TabIndex        =   68
      Text            =   "100"
      Top             =   5760
      Width           =   600
   End
   Begin VB.TextBox Text33 
      Alignment       =   1  'Right Justify
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   330
      Left            =   6120
      MaxLength       =   5
      TabIndex        =   67
      Text            =   "800"
      Top             =   6480
      Width           =   600
   End
   Begin VB.TextBox Text32 
      Alignment       =   1  'Right Justify
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   330
      Left            =   6720
      MaxLength       =   3
      TabIndex        =   66
      Text            =   "16"
      Top             =   6480
      Width           =   600
   End
   Begin VB.TextBox Text31 
      Alignment       =   1  'Right Justify
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   330
      Left            =   7320
      MaxLength       =   2
      TabIndex        =   65
      Text            =   "2"
      Top             =   6480
      Width           =   600
   End
   Begin VB.TextBox Text30 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000004&
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
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   6120
      MaxLength       =   5
      TabIndex        =   64
      Text            =   "1000"
      Top             =   6120
      Width           =   600
   End
   Begin VB.TextBox Text29 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000004&
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
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   6720
      MaxLength       =   3
      MousePointer    =   1  'Arrow
      TabIndex        =   63
      Text            =   "20"
      Top             =   6120
      Width           =   600
   End
   Begin VB.TextBox Text28 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000004&
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
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   7320
      MaxLength       =   2
      TabIndex        =   62
      Text            =   "2"
      Top             =   6120
      Width           =   600
   End
   Begin VB.TextBox Text27 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
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
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   330
      Left            =   3720
      MaxLength       =   4
      TabIndex        =   61
      Text            =   "200"
      Top             =   7320
      Width           =   600
   End
   Begin VB.TextBox Text26 
      Alignment       =   2  'Center
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   315
      Left            =   4680
      MaxLength       =   5
      TabIndex        =   60
      Text            =   "800"
      Top             =   6600
      Width           =   600
   End
   Begin VB.TextBox Text25 
      Alignment       =   2  'Center
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   330
      Left            =   4080
      MaxLength       =   3
      TabIndex        =   59
      Text            =   "16"
      Top             =   6600
      Width           =   600
   End
   Begin VB.TextBox Text24 
      Alignment       =   2  'Center
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   330
      Left            =   3480
      MaxLength       =   2
      TabIndex        =   58
      Text            =   "2"
      Top             =   6600
      Width           =   600
   End
   Begin VB.TextBox Text23 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   315
      Left            =   2880
      MaxLength       =   5
      TabIndex        =   57
      Text            =   "800"
      Top             =   6600
      Width           =   600
   End
   Begin VB.TextBox Text22 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
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
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   4080
      MaxLength       =   3
      TabIndex        =   56
      Text            =   "20"
      Top             =   6960
      Width           =   600
   End
   Begin VB.TextBox Text21 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
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
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   3480
      MaxLength       =   2
      TabIndex        =   55
      Text            =   "3"
      Top             =   6960
      Width           =   600
   End
   Begin VB.TextBox Text20 
      BackColor       =   &H80000004&
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
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   330
      Left            =   840
      MaxLength       =   4
      TabIndex        =   54
      Text            =   "100"
      Top             =   5760
      Width           =   600
   End
   Begin VB.TextBox Text19 
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   330
      Left            =   1440
      MaxLength       =   5
      TabIndex        =   53
      Text            =   "800"
      Top             =   6480
      Width           =   600
   End
   Begin VB.TextBox Text18 
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   330
      Left            =   840
      MaxLength       =   3
      TabIndex        =   52
      Text            =   "16"
      Top             =   6480
      Width           =   600
   End
   Begin VB.TextBox Text17 
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   330
      Left            =   240
      MaxLength       =   2
      TabIndex        =   51
      Text            =   "2"
      Top             =   6480
      Width           =   600
   End
   Begin VB.TextBox Text16 
      BackColor       =   &H80000004&
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
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   1440
      MaxLength       =   5
      TabIndex        =   50
      Text            =   "1000"
      Top             =   6120
      Width           =   600
   End
   Begin VB.TextBox Text15 
      BackColor       =   &H80000004&
      DataField       =   " "
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
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   840
      MaxLength       =   3
      MousePointer    =   1  'Arrow
      TabIndex        =   49
      Tag             =   " "
      Text            =   "20"
      Top             =   6120
      Width           =   600
   End
   Begin VB.TextBox Text14 
      BackColor       =   &H80000004&
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
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   240
      MaxLength       =   2
      TabIndex        =   48
      Text            =   "2"
      Top             =   6120
      Width           =   600
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Span 1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   268
      Left            =   80
      MaskColor       =   &H000000FF&
      TabIndex        =   27
      Top             =   3691
      Width           =   600
   End
   Begin VB.OptionButton Option6 
      BackColor       =   &H0080C0FF&
      Caption         =   "Option6"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3260
      TabIndex        =   23
      Top             =   3960
      Width           =   240
   End
   Begin VB.OptionButton Option5 
      BackColor       =   &H0080C0FF&
      Caption         =   "Option5"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2660
      TabIndex        =   22
      Top             =   3960
      Width           =   240
   End
   Begin VB.OptionButton Option4 
      BackColor       =   &H0080C0FF&
      Caption         =   "Option4"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2060
      TabIndex        =   21
      Top             =   3960
      Width           =   240
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H0080C0FF&
      Caption         =   "Option3"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1460
      TabIndex        =   20
      Top             =   3960
      Width           =   240
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H0080C0FF&
      Caption         =   "Option2"
      Enabled         =   0   'False
      Height          =   255
      Left            =   860
      TabIndex        =   19
      Top             =   3960
      Width           =   240
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Option1"
      Enabled         =   0   'False
      Height          =   255
      Left            =   260
      TabIndex        =   18
      Top             =   3960
      Width           =   240
   End
   Begin VB.TextBox Text13 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000014&
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
      ForeColor       =   &H00000080&
      Height          =   405
      Left            =   7200
      TabIndex        =   17
      Tag             =   "b"
      Text            =   "B"
      Top             =   955
      Width           =   800
   End
   Begin VB.TextBox Text12 
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   330
      Left            =   7200
      MaxLength       =   4
      TabIndex        =   16
      Tag             =   "75"
      Text            =   "300"
      ToolTipText     =   "Nilai numerikal sahaja !"
      Top             =   1610
      Width           =   800
   End
   Begin VB.TextBox Text11 
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   330
      Left            =   7200
      MaxLength       =   4
      TabIndex        =   14
      Tag             =   "300"
      Text            =   "500"
      ToolTipText     =   "Nilai numerikal sahaja !"
      Top             =   2704
      Width           =   800
   End
   Begin VB.TextBox Text10 
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   330
      Left            =   7200
      MaxLength       =   4
      TabIndex        =   13
      Tag             =   "75"
      Text            =   "200"
      ToolTipText     =   "Nilai numerikal sahaja !"
      Top             =   2157
      Width           =   800
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
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
      DragIcon        =   "SRasuk.frx":0620
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   5640
      TabIndex        =   12
      Tag             =   "0"
      Text            =   "0"
      ToolTipText     =   "Nilai numerikal sahaja !"
      Top             =   3240
      Width           =   800
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   1800
      TabIndex        =   11
      Tag             =   "0"
      Text            =   "0"
      ToolTipText     =   "Nilai numerikal sahaja !"
      Top             =   3240
      Width           =   800
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   4680
      MaxLength       =   5
      TabIndex        =   10
      Tag             =   "300"
      Text            =   "550"
      ToolTipText     =   "Nilai numerikal sahaja !"
      Top             =   3240
      Width           =   800
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   3720
      MaxLength       =   5
      TabIndex        =   9
      Tag             =   "150"
      Text            =   "200"
      ToolTipText     =   "Nilai numerikal sahaja !"
      Top             =   3240
      Width           =   800
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   2760
      MaxLength       =   6
      TabIndex        =   8
      Tag             =   "3000"
      Text            =   "7500"
      ToolTipText     =   "Nilai numerikal sahaja !"
      Top             =   3240
      Width           =   800
   End
   Begin VB.TextBox Text4 
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   330
      Left            =   80
      MaxLength       =   4
      TabIndex        =   7
      Tag             =   "300"
      Text            =   "500"
      ToolTipText     =   "Nilai numerikal sahaja !"
      Top             =   2704
      Width           =   800
   End
   Begin VB.TextBox Text3 
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   330
      Left            =   80
      MaxLength       =   4
      TabIndex        =   6
      Tag             =   "75"
      Text            =   "200"
      ToolTipText     =   "Nilai numerikal sahaja !"
      Top             =   2157
      Width           =   800
   End
   Begin VB.TextBox Text2 
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   330
      Left            =   80
      MaxLength       =   4
      TabIndex        =   5
      Tag             =   "75"
      Text            =   "300"
      ToolTipText     =   "Nilai numerikal sahaja !"
      Top             =   1610
      Width           =   800
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000014&
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
      ForeColor       =   &H00000080&
      Height          =   405
      Left            =   80
      TabIndex        =   4
      Tag             =   "a"
      Text            =   "A"
      Top             =   955
      Width           =   800
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6360
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command6 
      Caption         =   ">>LUKIS<<"
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
      Height          =   495
      Left            =   7080
      TabIndex        =   1
      Top             =   3720
      Width           =   980
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080C0FF&
      Caption         =   "http://www.wanluqman.com/"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5040
      TabIndex        =   102
      Top             =   4200
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      Caption         =   "Saiz Font"
      Height          =   255
      Left            =   3600
      TabIndex        =   101
      Top             =   960
      Width           =   735
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   5
      Height          =   2055
      Left            =   2040
      Shape           =   4  'Rounded Rectangle
      Top             =   840
      Width           =   3975
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   345
      Left            =   855
      Top             =   60
      Width           =   2805
   End
   Begin VB.Line Line18 
      BorderColor     =   &H000000FF&
      X1              =   620
      X2              =   7500
      Y1              =   5065.574
      Y2              =   5065.574
   End
   Begin VB.Line Line17 
      BorderColor     =   &H0000FFFF&
      X1              =   7440
      X2              =   7440
      Y1              =   5049.18
      Y2              =   5963.934
   End
   Begin VB.Line Line16 
      BorderColor     =   &H0000FFFF&
      X1              =   5520
      X2              =   5520
      Y1              =   5049.18
      Y2              =   5963.934
   End
   Begin VB.Line Line15 
      BorderColor     =   &H0000FFFF&
      X1              =   5280
      X2              =   5280
      Y1              =   5049.18
      Y2              =   5963.934
   End
   Begin VB.Line Line14 
      BorderColor     =   &H0000FFFF&
      X1              =   2760
      X2              =   2760
      Y1              =   5049.18
      Y2              =   5963.934
   End
   Begin VB.Line Line13 
      BorderColor     =   &H0000FFFF&
      X1              =   2520
      X2              =   2520
      Y1              =   5049.18
      Y2              =   5963.934
   End
   Begin VB.Line Line12 
      BorderColor     =   &H0000FFFF&
      X1              =   650
      X2              =   650
      Y1              =   5042.623
      Y2              =   5963.934
   End
   Begin VB.Line Line7 
      BorderColor     =   &H000000FF&
      X1              =   45
      X2              =   2160
      Y1              =   5195.628
      Y2              =   5195.628
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      X1              =   50
      X2              =   1700
      Y1              =   5899.454
      Y2              =   5899.454
   End
   Begin VB.Line Line11 
      BorderColor     =   &H000000FF&
      X1              =   1300
      X2              =   6700
      Y1              =   5840.437
      Y2              =   5840.437
   End
   Begin VB.Line Line10 
      BorderColor     =   &H000000FF&
      X1              =   620
      X2              =   7500
      Y1              =   5957.377
      Y2              =   5957.377
   End
   Begin VB.Line Line9 
      BorderColor     =   &H000000FF&
      X1              =   6100
      X2              =   7995
      Y1              =   5195.628
      Y2              =   5195.628
   End
   Begin VB.Line Line8 
      BorderColor     =   &H000000FF&
      X1              =   5400
      X2              =   7995
      Y1              =   5136.612
      Y2              =   5136.612
   End
   Begin VB.Line Line6 
      BorderColor     =   &H000000FF&
      X1              =   45
      X2              =   2640
      Y1              =   5136.612
      Y2              =   5136.612
   End
   Begin VB.Line Line5 
      BorderColor     =   &H000000FF&
      X1              =   6500
      X2              =   8000
      Y1              =   5899.454
      Y2              =   5899.454
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00008000&
      Height          =   1455
      Left            =   7560
      Top             =   4320
      Width           =   300
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00008000&
      Height          =   1455
      Left            =   240
      Top             =   4320
      Width           =   300
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00008000&
      Height          =   975
      Left            =   0
      Top             =   4560
      Width           =   8115
   End
   Begin VB.Image Image1 
      Height          =   1245
      Left            =   960
      Picture         =   "SRasuk.frx":092A
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label44 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080C0FF&
      Caption         =   "Back slab lvl."
      Height          =   215
      Left            =   6960
      TabIndex        =   87
      Top             =   3058
      Width           =   1000
   End
   Begin VB.Label Label43 
      BackColor       =   &H0080C0FF&
      Caption         =   "Front slab lvl."
      Height          =   215
      Left            =   80
      TabIndex        =   86
      Top             =   3058
      Width           =   1000
   End
   Begin VB.Label Label38 
      BackColor       =   &H0080C0FF&
      Caption         =   "Cover        Sab thk.    Link dia.                     Barmark  "
      Height          =   255
      Left            =   2160
      TabIndex        =   74
      Top             =   2468
      Width           =   3705
   End
   Begin VB.Label Label33 
      BackColor       =   &H0080C0FF&
      Caption         =   "fcu             fy             fyv          creep         shrink                           "
      Height          =   255
      Left            =   2280
      TabIndex        =   73
      Top             =   1588
      Width           =   3465
   End
   Begin VB.Label Label32 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080C0FF&
      Caption         =   "Nama Rsk.  "
      Height          =   255
      Left            =   4920
      TabIndex        =   72
      Top             =   960
      Width           =   900
   End
   Begin VB.Label Label31 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080C0FF&
      Caption         =   "Y-inst.  "
      Height          =   255
      Left            =   3000
      TabIndex        =   71
      Top             =   960
      Width           =   540
   End
   Begin VB.Label Label30 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080C0FF&
      Caption         =   "X-inst."
      Height          =   255
      Left            =   2205
      TabIndex        =   69
      Top             =   960
      Width           =   540
   End
   Begin VB.Label Label29 
      BackColor       =   &H0080C0FF&
      Caption         =   "     Curt.            Dia.             No."
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   135
      Left            =   6120
      TabIndex        =   47
      Top             =   7297
      Width           =   1845
   End
   Begin VB.Label Label27 
      BackColor       =   &H0080C0FF&
      Caption         =   "Link Spacing"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   255
      Left            =   3700
      TabIndex        =   46
      Top             =   7680
      Width           =   795
   End
   Begin VB.Label Label26 
      BackColor       =   &H0080C0FF&
      Caption         =   "Link Spacing"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   830
      TabIndex        =   45
      Top             =   5582
      Width           =   780
   End
   Begin VB.Label Label25 
      BackColor       =   &H0080C0FF&
      Caption         =   "Left Curt.                                          Right Curt."
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   2880
      TabIndex        =   44
      Top             =   6428
      Width           =   2460
   End
   Begin VB.Label Label24 
      BackColor       =   &H0080C0FF&
      Caption         =   "Link Spacing"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6700
      TabIndex        =   43
      Top             =   5582
      Width           =   900
   End
   Begin VB.Label Label23 
      BackColor       =   &H0080C0FF&
      Caption         =   "No.               Dia."
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   42
      Top             =   5612
      Width           =   1020
   End
   Begin VB.Label Label4 
      BackColor       =   &H0080C0FF&
      Caption         =   "      No.            Dia.            Curt. "
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   135
      Left            =   240
      TabIndex        =   41
      Top             =   7297
      Width           =   1740
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080C0FF&
      Caption         =   "X-beam H."
      Height          =   215
      Left            =   7080
      TabIndex        =   40
      Top             =   2522
      Width           =   900
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080C0FF&
      Caption         =   "X-beam B/2."
      Height          =   215
      Left            =   7080
      TabIndex        =   39
      Top             =   1985
      Width           =   900
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080C0FF&
      Caption         =   "Col.size H/2."
      Height          =   215
      Left            =   7080
      TabIndex        =   38
      Top             =   1406
      Width           =   900
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080C0FF&
      Caption         =   "Grid label."
      Height          =   215
      Left            =   7080
      TabIndex        =   37
      Top             =   751
      Width           =   900
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      Caption         =   "Beam soffit."
      Height          =   210
      Left            =   5640
      TabIndex        =   36
      Top             =   3000
      Width           =   1005
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      Caption         =   "Beam H."
      Height          =   210
      Left            =   4560
      TabIndex        =   35
      Top             =   3000
      Width           =   1125
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      Caption         =   "Beam B."
      Height          =   210
      Left            =   3600
      TabIndex        =   34
      Top             =   3000
      Width           =   1005
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      Caption         =   "Beam length L."
      Height          =   210
      Left            =   2640
      TabIndex        =   33
      Top             =   3000
      Width           =   1125
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      Caption         =   "Beam top."
      Height          =   210
      Left            =   1680
      TabIndex        =   32
      Top             =   3000
      Width           =   1005
   End
   Begin VB.Label Label10 
      BackColor       =   &H0080C0FF&
      Caption         =   "X-beam H."
      Height          =   215
      Left            =   80
      TabIndex        =   31
      Top             =   2522
      Width           =   900
   End
   Begin VB.Label Label9 
      BackColor       =   &H0080C0FF&
      Caption         =   "X-beam B/2."
      Height          =   215
      Left            =   80
      TabIndex        =   30
      Top             =   1985
      Width           =   900
   End
   Begin VB.Label Label8 
      BackColor       =   &H0080C0FF&
      Caption         =   "Col.size H/2."
      Height          =   215
      Left            =   80
      TabIndex        =   29
      Top             =   1406
      Width           =   900
   End
   Begin VB.Label Label7 
      BackColor       =   &H0080C0FF&
      Caption         =   "Grid label."
      Height          =   215
      Left            =   80
      TabIndex        =   28
      Top             =   751
      Width           =   900
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
      TabIndex        =   3
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
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
      ForeColor       =   &H00008000&
      Height          =   195
      Left            =   3840
      TabIndex        =   2
      Top             =   120
      Width           =   4155
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "RASUK MS1195."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   255
      Left            =   960
      TabIndex        =   0
      Top             =   90
      Width           =   2340
   End
   Begin VB.Menu mnuItemFile 
      Caption         =   "&Fail_dwg"
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
'''THE PROGRAM IS AN INPUT TEMPLATE FOR STRUCTURAL REINFORCED CONCRETE BEAM,  '''
'''TO INTERFACE INTO AUTOCAD ENVIRONMENT WITH PRE-DESIGNED ALGORITHMS THAT    '''
'''ENABLE TO CONVERT THE INPUT DATA AUTOMATICALLY INTO STRUCT. R.C. DRAWINGS. '''
'''CREATED IN 2001 BY : WAN SOHAIMI BIN WAN MOHAMED.                          '''
'''(LATEST REVISION FEB 2002)- [butiran rasuk sahaja]                         '''
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
Dim Span1 As New Dimensi
Dim Span2 As New Dimensi
Dim Span3 As New Dimensi
Dim Span4 As New Dimensi
Dim Span5 As New Dimensi
Dim Span6 As New Dimensi
Dim Span7 As New Dimensi
Dim Span8 As New Dimensi
Dim Span9 As New Dimensi

Dim Tetulang1 As New Tetulang
Dim Tetulang2 As New Tetulang
Dim Tetulang3 As New Tetulang
Dim Tetulang4 As New Tetulang
Dim Tetulang5 As New Tetulang
Dim Tetulang6 As New Tetulang
Dim Tetulang7 As New Tetulang
Dim Tetulang8 As New Tetulang
Dim Tetulang9 As New Tetulang

Dim Stresses1 As New Stresses

Dim Moment As New CalcMoment
Dim Shear As New CalcShear
Dim Curve1 As New ACurvature
Dim Curve2 As New ACurvature
Dim Deflection As New CalcDeflection
Dim CrackWidth As New CalcCrackWidth



Private Sub Check1_Click()
OpenDataFile
'''''''''''''''
NoOfSpan = Int(Val(Right(Command4.Caption, 1))) ''
i = NoOfSpan ''
'''''''''''''''
Form1.Picture = LoadPicture("C:\autodraf\icon\datam.ico")

List1.Clear
List1.Visible = False
'If Picture1.Visible = True Then
'   Picture1.Visible = False
'   End If

'Dim fnum As Integer
'Dim txtFile, Temp As String
'fnum = FreeFile
'txtFile = "C:\autodraf\rasuk\input_data\DefaultStress.txt"
Command2.Enabled = True
Command3.Enabled = False
Command4.Enabled = False
Command6.Enabled = False
Command5.Enabled = False


Option1.Enabled = False
Option2.Enabled = False
Option3.Enabled = False
Option4.Enabled = False
Option5.Enabled = False
Option6.Enabled = False
Option7.Enabled = False
Option8.Enabled = False
Option9.Enabled = False

Txt_StatusTwo
Text8.Enabled = False
Text9.Enabled = False



Text35.text = Xinsertion
Text36.text = Yinsertion
Text57.text = FontSz
Text37.text = NamaRasuk
Text38.text = fcu
Text39.text = fy
Text40.text = fyv
Text41.text = Shrink
Text42.text = Creep
Text43.text = cVr
Text44.text = slabThick
Text45.text = stirupD
Text46.text = BarMark

Stresses1.Value35 = Text35.text
Stresses1.Value36 = Text36.text
Stresses1.Value57 = Text57.text
Stresses1.Value37 = Text37.text
Stresses1.Value38 = Text38.text
Stresses1.Value39 = Text39.text
Stresses1.Value40 = Text40.text
Stresses1.Value41 = Text41.text
Stresses1.Value42 = Text42.text
Stresses1.Value43 = Text43.text
Stresses1.Value44 = Text44.text
Stresses1.Value45 = Text45.text
Stresses1.Value46 = Text46.text

Text35.text = Stresses1.XinsertPt
Text36.text = Stresses1.YinsertPt
Text57.text = Stresses1.SetFontSize
Text37.text = Stresses1.NamaRasuk
Text38.text = Stresses1.fcu
Text39.text = Stresses1.fy
Text40.text = Stresses1.fyv
Text41.text = Stresses1.Shrink
Text42.text = Stresses1.Creep
Text43.text = Stresses1.Cover
Text44.text = Stresses1.SlabThk
Text45.text = Stresses1.LinkD
Text46.text = Stresses1.BarMark


End Sub



''''AUTOCAD''''
Private Sub Command1_Click()

Form1.Picture = LoadPicture("C:\autodraf\icon\ukad3.ico")
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim SetLuput As Date
SetLuput = DateValue("6/15/3002")
If Date >= SetLuput Then
 MsgBox ":::Sila hubungi Wan Sohaimi Wan Mohamed @ http//www.wanluqman.com/", , "To reinstall"
 Exit Sub
End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

If dwgName = "" Then
MsgBox "Sila pilih fail dwg untuk kerja.", , "NOTA AM:"
Exit Sub
End If

StartAutoCAD
SetLayer


Option1.Enabled = True
Option2.Enabled = True
Option3.Enabled = True
Option4.Enabled = True
Option5.Enabled = True
Option6.Enabled = True
Option7.Enabled = True
Option8.Enabled = True
Option9.Enabled = True
''''Command4.Enabled = True
 mnuItemExit.Enabled = True
 mnuItemOpenDwg.Enabled = False
 
 mnuItemFile.Enabled = True
 mnuItemFile.Visible = True
 mnuItemFile.Caption = "Klik di sini!"

End Sub

Private Sub Command2_Click()
Dim fnum As Integer
Dim txtFile As String
''NoOfSpan = Int(Val(Right(Command4.Caption, 1)))
''i = NoOfSpan

fnum = FreeFile
Form1.Picture = LoadPicture("C:\autodraf\icon\pilihR.ico")
Image1.Picture = LoadPicture("C:\autodraf\icon\cskp.ico")

Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command6.Enabled = False
Command5.Enabled = False
Command5.Visible = False

Option1.Enabled = True
Option2.Enabled = True
Option3.Enabled = True
Option4.Enabled = True
Option5.Enabled = True
Option6.Enabled = True
Option7.Enabled = True
Option8.Enabled = True
Option9.Enabled = True


Text35.Enabled = False
Text36.Enabled = False
Text57.Enabled = False
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

Xinsertion = Text35.text
Yinsertion = Text36.text
FontSz = Text57.text
NamaRasuk = Text37.text
fcu = Text38.text
fy = Text39.text
fyv = Text40.text
Shrink = Text41.text
Creep = Text42.text
cVr = Text43.text
slabThick = Text44.text
stirupD = Text45.text
BarMark = Text46.text


txtFile = "C:\autodraf\rasuk\input_data\DefaultStress.txt"
Open txtFile For Output As #fnum
Print #fnum, Xinsertion
Print #fnum, Yinsertion
Print #fnum, FontSz
Print #fnum, NamaRasuk
Print #fnum, fcu
Print #fnum, fy
Print #fnum, fyv
Print #fnum, Shrink
Print #fnum, Creep
Print #fnum, cVr
Print #fnum, slabThick
Print #fnum, stirupD
Print #fnum, BarMark
Close #fnum

End Sub

Private Sub Command3_Click()
Dim SetLuput As Date
SetLuput = DateValue("12/15/3013")
If Date >= SetLuput Then
 MsgBox ":::Sila hubungi Wan Sohaimi Wan Mohamed @ 603-61574717::: ", , "To reinstall"
End If

List1.Clear
Command5.Enabled = True
Command5.Visible = True

Option1.Enabled = True
Option2.Enabled = True
Option3.Enabled = True
Option4.Enabled = True
Option5.Enabled = True
Option6.Enabled = True
Option7.Enabled = True
Option8.Enabled = True
Option9.Enabled = True
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
'''Command6.Enabled = False
NoOfSpan = Int(Val(Right(Command4.Caption, 1)))
i = NoOfSpan

CalculateStrength

''''Picture1.Visible = True

End Sub

Private Sub Command4_Click()
Dim fnum As Integer
Dim txtFile As String
Dim N As Integer

Option1.Enabled = True
Option2.Enabled = True
Option3.Enabled = True
Option4.Enabled = True
Option5.Enabled = True
Option6.Enabled = True
Option7.Enabled = True
Option8.Enabled = True
Option9.Enabled = True

Command3.Enabled = True
Command4.Enabled = False
NoOfSpan = Int(Val(Right(Command4.Caption, 1)))
i = NoOfSpan
fnum = FreeFile
Form1.Picture = LoadPicture("C:\autodraf\icon\ukad4.ico")
'''''''''''''''''''''''''
 
If mnuItemOpenDwg.Enabled = True Then
Command6.Enabled = False
Else
Command6.Enabled = True
End If

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
Text47.Enabled = False
Text48.Enabled = False

''''''''''''''''''''''''''
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
Text57.Enabled = False
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
'''''>>>>'''''''
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


''''''''''''''''''''***********************************************
If Command4.Left = 80 Then

txtFile = "C:\autodraf\rasuk\input_data\SpanOneGET.txt"
Option1.Enabled = False
''''''''''''''''''
Transfer_TxtToData (1)
'''''''''''''''''''''
N = 1
''''''''''''''''''
Open txtFile For Output As #fnum

Print #fnum, GridNameL(N)
Print #fnum, scL(N), sbL(N), shL(N)
Print #fnum, beamL(N), beamB(N), beamH(N)

Print #fnum, sbR(N), shR(N), scR(N)
Print #fnum, slabDrop(N), beamUplift(N)
Print #fnum, FrontSlabLvl(N), BackSlabLvl(N)
Print #fnum, GridNameR(N)
''''''''''''
Print #fnum, RbarTL1no(N), RbarTL1dia(N), RbarTL1curE(N)
Print #fnum, RbarTL2no(N), RbarTL2dia(N), RbarTL2curE(N)
Print #fnum, LinkLSpace(N)
'''''
Print #fnum, RbarMS1no(N), RbarMS1dia(N)
Print #fnum, RbarMS2curS(N), RbarMS2no(N), RbarMS2dia(N), RbarMS2curE(N)   ''''
Print #fnum, LinkMSpace(N)
'''''
Print #fnum, RbarTR1no(N), RbarTR1dia(N), RbarTR1curS(N)
Print #fnum, RbarTR2no(N), RbarTR2dia(N), RbarTR2curS(N)
Print #fnum, LinkRSpace(N)

Print #fnum, RbarCLfno(N), RbarCLfdia(N), RbarCLfcurE(N) '''''''''
Print #fnum, RbarLCno(N), RbarLCdia(N)                 ''''''''''link
Print #fnum, RbarCRtno(N), RbarCRtdia(N), RbarCRtcurS(N)  '''''''''
Close #fnum
'''''BEFORE
'''''nil

''''AFTER
Open "C:\autodraf\rasuk\input_data\SpanTwoGET.txt" For Output As #fnum

Print #fnum, GridNameR(N)
Print #fnum, scR(N), sbR(N), shR(N)
Print #fnum, beamL(N + 1), beamB(N + 1), beamH(N + 1)

Print #fnum, sbR(N + 1), shR(N + 1), scR(N + 1)
Print #fnum, slabDrop(N + 1), beamUplift(N + 1)
Print #fnum, FrontSlabLvl(N + 1), BackSlabLvl(N + 1)
Print #fnum, GridNameR(N + 1)
''''''''''''
Print #fnum, RbarTR1no(N), RbarTR1dia(N), RbarTL1curE(N + 1)
Print #fnum, RbarTR2no(N), RbarTR2dia(N), RbarTL2curE(N + 1)
Print #fnum, LinkLSpace(N + 1)
'''''
Print #fnum, RbarMS1no(N + 1), RbarMS1dia(N + 1)
Print #fnum, RbarMS2curS(N + 1), RbarMS2no(N + 1), RbarMS2dia(N + 1), RbarMS2curE(N + 1) ''''
Print #fnum, LinkMSpace(N + 1)
'''''
Print #fnum, RbarTR1no(N + 1), RbarTR1dia(N + 1), RbarTR1curS(N + 1)
Print #fnum, RbarTR2no(N + 1), RbarTR2dia(N + 1), RbarTR2curS(N + 1)
Print #fnum, LinkRSpace(N + 1)

Print #fnum, RbarCRtno(N), RbarCRtdia(N), RbarCLfcurE(N + 1)   '''''''''
Print #fnum, RbarLCno(N + 1), RbarLCdia(N + 1)              ''''''''''link
Print #fnum, RbarCRtno(N + 1), RbarCRtdia(N + 1), RbarCRtcurS(N + 1) '''''''''
Close #fnum

End If


'''''''''''''''2222'''''**********************************************
If Command4.Left = 680 Then
txtFile = "C:\autodraf\rasuk\input_data\SpanTwoGET.txt"
Option2.Enabled = False
''''''''''''''''''''''
Transfer_TxtToData (2)
'''''''''''''''''''''
N = 2

''''''''''''''''''''''''''''
Open txtFile For Output As #fnum
Print #fnum, GridNameL(N)
Print #fnum, scL(N), sbL(N), shL(N)
Print #fnum, beamL(N), beamB(N), beamH(N)

Print #fnum, sbR(N), shR(N), scR(N)
Print #fnum, slabDrop(N), beamUplift(N)
Print #fnum, FrontSlabLvl(N), BackSlabLvl(N)
Print #fnum, GridNameR(N)
''''''''''''
Print #fnum, RbarTL1no(N), RbarTL1dia(N), RbarTL1curE(N)
Print #fnum, RbarTL2no(N), RbarTL2dia(N), RbarTL2curE(N)
Print #fnum, LinkLSpace(N)
'''''
Print #fnum, RbarMS1no(N), RbarMS1dia(N)
Print #fnum, RbarMS2curS(N), RbarMS2no(N), RbarMS2dia(N), RbarMS2curE(N)   ''''
Print #fnum, LinkMSpace(N)
'''''
Print #fnum, RbarTR1no(N), RbarTR1dia(N), RbarTR1curS(N)
Print #fnum, RbarTR2no(N), RbarTR2dia(N), RbarTR2curS(N)
Print #fnum, LinkRSpace(N)

Print #fnum, RbarCLfno(N), RbarCLfdia(N), RbarCLfcurE(N) '''''''''
Print #fnum, RbarLCno(N), RbarLCdia(N)                 ''''''''''link
Print #fnum, RbarCRtno(N), RbarCRtdia(N), RbarCRtcurS(N)  '''''''''
Close #fnum


''''BEFORE
Open "C:\autodraf\rasuk\input_data\SpanOneGET.txt" For Output As #fnum
Print #fnum, GridNameL(N - 1)
Print #fnum, scL(N - 1), sbL(N - 1), shL(N - 1)
Print #fnum, beamL(N - 1), beamB(N - 1), beamH(N - 1)

Print #fnum, sbL(N), shL(N), scL(N)
Print #fnum, slabDrop(N - 1), beamUplift(N - 1)
Print #fnum, FrontSlabLvl(N - 1), BackSlabLvl(N - 1)
Print #fnum, GridNameL(N)
''''''''''''
Print #fnum, RbarTL1no(N - 1), RbarTL1dia(N - 1), RbarTL1curE(N - 1)
Print #fnum, RbarTL2no(N - 1), RbarTL2dia(N - 1), RbarTL2curE(N - 1)
Print #fnum, LinkLSpace(N - 1)
'''''
Print #fnum, RbarMS1no(N - 1), RbarMS1dia(N - 1) ''''
Print #fnum, RbarMS2curS(N - 1); RbarMS2no(N - 1), RbarMS2dia(N - 1), RbarMS2curE(N - 1) ''''
Print #fnum, LinkMSpace(N - 1)
'''''
Print #fnum, RbarTL1no(N), RbarTL1dia(N), RbarTR1curS(N - 1)
Print #fnum, RbarTL2no(N), RbarTL2dia(N), RbarTR2curS(N - 1)
Print #fnum, LinkRSpace(N - 1)

Print #fnum, RbarCLfno(N - 1), RbarCLfdia(N - 1), RbarCLfcurE(N - 1) '''''''''
Print #fnum, RbarLCno(N - 1), RbarLCdia(N - 1) ''''''''''link
Print #fnum, RbarCLfno(N), RbarCLfdia(N), RbarCRtcurS(N - 1)  '''''''''
Close #fnum

'''AFTER

Open "C:\autodraf\rasuk\input_data\SpanThreeGET.txt" For Output As #fnum

Print #fnum, GridNameR(N)
Print #fnum, scR(N), sbR(N), shR(N)
Print #fnum, beamL(N + 1), beamB(N + 1), beamH(N + 1)

Print #fnum, sbR(N + 1), shR(N + 1), scR(N + 1)
Print #fnum, slabDrop(N + 1), beamUplift(N + 1)
Print #fnum, FrontSlabLvl(N + 1), BackSlabLvl(N + 1)
Print #fnum, GridNameR(N + 1)
''''''''''''
Print #fnum, RbarTR1no(N), RbarTR1dia(N), RbarTL1curE(N + 1)
Print #fnum, RbarTR2no(N), RbarTR2dia(N), RbarTL2curE(N + 1)
Print #fnum, LinkLSpace(N + 1)
'''''
Print #fnum, RbarMS1no(N + 1), RbarMS1dia(N + 1)
Print #fnum, RbarMS2curS(N + 1), RbarMS2no(N + 1), RbarMS2dia(N + 1), RbarMS2curE(N + 1) ''''
Print #fnum, LinkMSpace(N + 1)
'''''
Print #fnum, RbarTR1no(N + 1), RbarTR1dia(N + 1), RbarTR1curS(N + 1)
Print #fnum, RbarTR2no(N + 1), RbarTR2dia(N + 1), RbarTR2curS(N + 1)
Print #fnum, LinkRSpace(N + 1)

Print #fnum, RbarCRtno(N), RbarCRtdia(N), RbarCLfcurE(N + 1)   '''''''''
Print #fnum, RbarLCno(N + 1), RbarLCdia(N + 1)              ''''''''''link
Print #fnum, RbarCRtno(N + 1), RbarCRtdia(N + 1), RbarCRtcurS(N + 1) '''''''''
Close #fnum

End If


''''''''''''''333333''''*************************************************
If Command4.Left = 1280 Then
txtFile = "C:\autodraf\rasuk\input_data\SpanThreeGET.txt"
Option3.Enabled = False
'''''''''''''''''''''
Transfer_TxtToData (3)
'''''''''''''''''''''
N = 3

''''''''''''''''''''''''''''
Open txtFile For Output As #fnum
Print #fnum, GridNameL(N)
Print #fnum, scL(N), sbL(N), shL(N)
Print #fnum, beamL(N), beamB(N), beamH(N)

Print #fnum, sbR(N), shR(N), scR(N)
Print #fnum, slabDrop(N), beamUplift(N)
Print #fnum, FrontSlabLvl(N), BackSlabLvl(N)
Print #fnum, GridNameR(N)
''''''''''''
Print #fnum, RbarTL1no(N), RbarTL1dia(N), RbarTL1curE(N)
Print #fnum, RbarTL2no(N), RbarTL2dia(N), RbarTL2curE(N)
Print #fnum, LinkLSpace(N)
'''''
Print #fnum, RbarMS1no(N), RbarMS1dia(N)
Print #fnum, RbarMS2curS(N), RbarMS2no(N), RbarMS2dia(N), RbarMS2curE(N)   ''''
Print #fnum, LinkMSpace(N)
'''''
Print #fnum, RbarTR1no(N), RbarTR1dia(N), RbarTR1curS(N)
Print #fnum, RbarTR2no(N), RbarTR2dia(N), RbarTR2curS(N)
Print #fnum, LinkRSpace(N)

Print #fnum, RbarCLfno(N), RbarCLfdia(N), RbarCLfcurE(N) '''''''''
Print #fnum, RbarLCno(N), RbarLCdia(N)                 ''''''''''link
Print #fnum, RbarCRtno(N), RbarCRtdia(N), RbarCRtcurS(N)  '''''''''
Close #fnum
'''''''''''''''''''''''''''''''''''''''''''^^^^^^^^^^^^^^^^^

''''BEFORE
Open "C:\autodraf\rasuk\input_data\SpanTwoGET.txt" For Output As #fnum
Print #fnum, GridNameL(N - 1)
Print #fnum, scL(N - 1), sbL(N - 1), shL(N - 1)
Print #fnum, beamL(N - 1), beamB(N - 1), beamH(N - 1)

Print #fnum, sbL(N), shL(N), scL(N)
Print #fnum, slabDrop(N - 1), beamUplift(N - 1)
Print #fnum, FrontSlabLvl(N - 1), BackSlabLvl(N - 1)
Print #fnum, GridNameL(N)
''''''''''''
Print #fnum, RbarTL1no(N - 1), RbarTL1dia(N - 1), RbarTL1curE(N - 1)
Print #fnum, RbarTL2no(N - 1), RbarTL2dia(N - 1), RbarTL2curE(N - 1)
Print #fnum, LinkLSpace(N - 1)
'''''
Print #fnum, RbarMS1no(N - 1), RbarMS1dia(N - 1) ''''
Print #fnum, RbarMS2curS(N - 1), RbarMS2no(N - 1), RbarMS2dia(N - 1), RbarMS2curE(N - 1) ''''
Print #fnum, LinkMSpace(N - 1)
'''''
Print #fnum, RbarTL1no(N), RbarTL1dia(N), RbarTR1curS(N - 1)
Print #fnum, RbarTL2no(N), RbarTL2dia(N), RbarTR2curS(N - 1)
Print #fnum, LinkRSpace(N - 1)

Print #fnum, RbarCLfno(N - 1), RbarCLfdia(N - 1), RbarCLfcurE(N - 1) '''''''''
Print #fnum, RbarLCno(N - 1), RbarLCdia(N - 1) ''''''''''link
Print #fnum, RbarCLfno(N), RbarCLfdia(N), RbarCRtcurS(N - 1)  '''''''''
Close #fnum

'''AFTER

Open "C:\autodraf\rasuk\input_data\SpanFourGET.txt" For Output As #fnum

Print #fnum, GridNameR(N)
Print #fnum, scR(N), sbR(N), shR(N)
Print #fnum, beamL(N + 1), beamB(N + 1), beamH(N + 1)

Print #fnum, sbR(N + 1), shR(N + 1), scR(N + 1)
Print #fnum, slabDrop(N + 1), beamUplift(N + 1)
Print #fnum, FrontSlabLvl(N + 1), BackSlabLvl(N + 1)
Print #fnum, GridNameR(N + 1)
''''''''''''
Print #fnum, RbarTR1no(N), RbarTR1dia(N), RbarTL1curE(N + 1)
Print #fnum, RbarTR2no(N), RbarTR2dia(N), RbarTL2curE(N + 1)
Print #fnum, LinkLSpace(N + 1)
'''''
Print #fnum, RbarMS1no(N + 1), RbarMS1dia(N + 1)
Print #fnum, RbarMS2curS(N + 1), RbarMS2no(N + 1), RbarMS2dia(N + 1), RbarMS2curE(N + 1) ''''
Print #fnum, LinkMSpace(N + 1)
'''''
Print #fnum, RbarTR1no(N + 1), RbarTR1dia(N + 1), RbarTR1curS(N + 1)
Print #fnum, RbarTR2no(N + 1), RbarTR2dia(N + 1), RbarTR2curS(N + 1)
Print #fnum, LinkRSpace(N + 1)

Print #fnum, RbarCRtno(N), RbarCRtdia(N), RbarCLfcurE(N + 1)   '''''''''
Print #fnum, RbarLCno(N + 1), RbarLCdia(N + 1)              ''''''''''link
Print #fnum, RbarCRtno(N + 1), RbarCRtdia(N + 1), RbarCRtcurS(N + 1) '''''''''
Close #fnum

End If


'''''''''''''444444'''''''************************************************
If Command4.Left = 1880 Then
txtFile = "C:\autodraf\rasuk\input_data\SpanFourGET.txt"
Option4.Enabled = False
'''''''''''''''''''''''

Transfer_TxtToData (4)
'''''''''''''''''''''
N = 4

''''''''''''''''''''''''''''
Open txtFile For Output As #fnum
Print #fnum, GridNameL(N)
Print #fnum, scL(N), sbL(N), shL(N)
Print #fnum, beamL(N), beamB(N), beamH(N)

Print #fnum, sbR(N), shR(N), scR(N)
Print #fnum, slabDrop(N), beamUplift(N)
Print #fnum, FrontSlabLvl(N), BackSlabLvl(N)
Print #fnum, GridNameR(N)
''''''''''''
Print #fnum, RbarTL1no(N), RbarTL1dia(N), RbarTL1curE(N)
Print #fnum, RbarTL2no(N), RbarTL2dia(N), RbarTL2curE(N)
Print #fnum, LinkLSpace(N)
'''''
Print #fnum, RbarMS1no(N), RbarMS1dia(N)
Print #fnum, RbarMS2curS(N), RbarMS2no(N), RbarMS2dia(N), RbarMS2curE(N)   ''''
Print #fnum, LinkMSpace(N)
'''''
Print #fnum, RbarTR1no(N), RbarTR1dia(N), RbarTR1curS(N)
Print #fnum, RbarTR2no(N), RbarTR2dia(N), RbarTR2curS(N)
Print #fnum, LinkRSpace(N)

Print #fnum, RbarCLfno(N), RbarCLfdia(N), RbarCLfcurE(N) '''''''''
Print #fnum, RbarLCno(N), RbarLCdia(N)                 ''''''''''link
Print #fnum, RbarCRtno(N), RbarCRtdia(N), RbarCRtcurS(N)  '''''''''
Close #fnum


''''BEFORE
Open "C:\autodraf\rasuk\input_data\SpanThreeGET.txt" For Output As #fnum
Print #fnum, GridNameL(N - 1)
Print #fnum, scL(N - 1), sbL(N - 1), shL(N - 1)
Print #fnum, beamL(N - 1), beamB(N - 1), beamH(N - 1)

Print #fnum, sbL(N), shL(N), scL(N)
Print #fnum, slabDrop(N - 1), beamUplift(N - 1)
Print #fnum, FrontSlabLvl(N - 1), BackSlabLvl(N - 1)
Print #fnum, GridNameL(N)
''''''''''''
Print #fnum, RbarTL1no(N - 1), RbarTL1dia(N - 1), RbarTL1curE(N - 1)
Print #fnum, RbarTL2no(N - 1), RbarTL2dia(N - 1), RbarTL2curE(N - 1)
Print #fnum, LinkLSpace(N - 1)
'''''
Print #fnum, RbarMS1no(N - 1), RbarMS1dia(N - 1) ''''
Print #fnum, RbarMS2curS(N - 1), RbarMS2no(N - 1), RbarMS2dia(N - 1), RbarMS2curE(N - 1) ''''
Print #fnum, LinkMSpace(N - 1)
'''''
Print #fnum, RbarTL1no(N), RbarTL1dia(N), RbarTR1curS(N - 1)
Print #fnum, RbarTL2no(N), RbarTL2dia(N), RbarTR2curS(N - 1)
Print #fnum, LinkRSpace(N - 1)

Print #fnum, RbarCLfno(N - 1), RbarCLfdia(N - 1), RbarCLfcurE(N - 1) '''''''''
Print #fnum, RbarLCno(N - 1), RbarLCdia(N - 1) ''''''''''link
Print #fnum, RbarCLfno(N), RbarCLfdia(N), RbarCRtcurS(N - 1)  '''''''''
Close #fnum

'''AFTER

Open "C:\autodraf\rasuk\input_data\SpanFiveGET.txt" For Output As #fnum

Print #fnum, GridNameR(N)
Print #fnum, scR(N), sbR(N), shR(N)
Print #fnum, beamL(N + 1), beamB(N + 1), beamH(N + 1)

Print #fnum, sbR(N + 1), shR(N + 1), scR(N + 1)
Print #fnum, slabDrop(N + 1), beamUplift(N + 1)
Print #fnum, FrontSlabLvl(N + 1), BackSlabLvl(N + 1)
Print #fnum, GridNameR(N + 1)
''''''''''''
Print #fnum, RbarTR1no(N), RbarTR1dia(N), RbarTL1curE(N + 1)
Print #fnum, RbarTR2no(N), RbarTR2dia(N), RbarTL2curE(N + 1)
Print #fnum, LinkLSpace(N + 1)
'''''
Print #fnum, RbarMS1no(N + 1), RbarMS1dia(N + 1)
Print #fnum, RbarMS2curS(N + 1), RbarMS2no(N + 1), RbarMS2dia(N + 1), RbarMS2curE(N + 1) ''''
Print #fnum, LinkMSpace(N + 1)
'''''
Print #fnum, RbarTR1no(N + 1), RbarTR1dia(N + 1), RbarTR1curS(N + 1)
Print #fnum, RbarTR2no(N + 1), RbarTR2dia(N + 1), RbarTR2curS(N + 1)
Print #fnum, LinkRSpace(N + 1)

Print #fnum, RbarCRtno(N), RbarCRtdia(N), RbarCLfcurE(N + 1)   '''''''''
Print #fnum, RbarLCno(N + 1), RbarLCdia(N + 1)              ''''''''''link
Print #fnum, RbarCRtno(N + 1), RbarCRtdia(N + 1), RbarCRtcurS(N + 1) '''''''''
Close #fnum

End If

''''''''555555''''''''''************************************************
If Command4.Left = 2480 Then
txtFile = "C:\autodraf\rasuk\input_data\SpanFiveGET.txt"
Option5.Enabled = False

Transfer_TxtToData (5)
'''''''''''''''''''''
N = 5

''''''''''''''''''''''''''''
Open txtFile For Output As #fnum
Print #fnum, GridNameL(N)
Print #fnum, scL(N), sbL(N), shL(N)
Print #fnum, beamL(N), beamB(N), beamH(N)

Print #fnum, sbR(N), shR(N), scR(N)
Print #fnum, slabDrop(N), beamUplift(N)
Print #fnum, FrontSlabLvl(N), BackSlabLvl(N)
Print #fnum, GridNameR(N)
''''''''''''
Print #fnum, RbarTL1no(N), RbarTL1dia(N), RbarTL1curE(N)
Print #fnum, RbarTL2no(N), RbarTL2dia(N), RbarTL2curE(N)
Print #fnum, LinkLSpace(N)
'''''
Print #fnum, RbarMS1no(N), RbarMS1dia(N)
Print #fnum, RbarMS2curS(N), RbarMS2no(N), RbarMS2dia(N), RbarMS2curE(N)   ''''
Print #fnum, LinkMSpace(N)
'''''
Print #fnum, RbarTR1no(N), RbarTR1dia(N), RbarTR1curS(N)
Print #fnum, RbarTR2no(N), RbarTR2dia(N), RbarTR2curS(N)
Print #fnum, LinkRSpace(N)

Print #fnum, RbarCLfno(N), RbarCLfdia(N), RbarCLfcurE(N) '''''''''
Print #fnum, RbarLCno(N), RbarLCdia(N)                 ''''''''''link
Print #fnum, RbarCRtno(N), RbarCRtdia(N), RbarCRtcurS(N)  '''''''''
Close #fnum


''''BEFORE
Open "C:\autodraf\rasuk\input_data\SpanFourGET.txt" For Output As #fnum
Print #fnum, GridNameL(N - 1)
Print #fnum, scL(N - 1), sbL(N - 1), shL(N - 1)
Print #fnum, beamL(N - 1), beamB(N - 1), beamH(N - 1)

Print #fnum, sbL(N), shL(N), scL(N)
Print #fnum, slabDrop(N - 1), beamUplift(N - 1)
Print #fnum, FrontSlabLvl(N - 1), BackSlabLvl(N - 1)
Print #fnum, GridNameL(N)
''''''''''''
Print #fnum, RbarTL1no(N - 1), RbarTL1dia(N - 1), RbarTL1curE(N - 1)
Print #fnum, RbarTL2no(N - 1), RbarTL2dia(N - 1), RbarTL2curE(N - 1)
Print #fnum, LinkLSpace(N - 1)
'''''
Print #fnum, RbarMS1no(N - 1), RbarMS1dia(N - 1) ''''
Print #fnum, RbarMS2curS(N - 1), RbarMS2no(N - 1), RbarMS2dia(N - 1), RbarMS2curE(N - 1) ''''
Print #fnum, LinkMSpace(N - 1)
'''''
Print #fnum, RbarTL1no(N), RbarTL1dia(N), RbarTR1curS(N - 1)
Print #fnum, RbarTL2no(N), RbarTL2dia(N), RbarTR2curS(N - 1)
Print #fnum, LinkRSpace(N - 1)

Print #fnum, RbarCLfno(N - 1), RbarCLfdia(N - 1), RbarCLfcurE(N - 1) '''''''''
Print #fnum, RbarLCno(N - 1), RbarLCdia(N - 1) ''''''''''link
Print #fnum, RbarCLfno(N), RbarCLfdia(N), RbarCRtcurS(N - 1)  '''''''''
Close #fnum

'''AFTER

Open "C:\autodraf\rasuk\input_data\SpanSixGET.txt" For Output As #fnum

Print #fnum, GridNameR(N)
Print #fnum, scR(N), sbR(N), shR(N)
Print #fnum, beamL(N + 1), beamB(N + 1), beamH(N + 1)

Print #fnum, sbR(N + 1), shR(N + 1), scR(N + 1)
Print #fnum, slabDrop(N + 1), beamUplift(N + 1)
Print #fnum, FrontSlabLvl(N + 1), BackSlabLvl(N + 1)
Print #fnum, GridNameR(N + 1)
''''''''''''
Print #fnum, RbarTR1no(N), RbarTR1dia(N), RbarTL1curE(N + 1)
Print #fnum, RbarTR2no(N), RbarTR2dia(N), RbarTL2curE(N + 1)
Print #fnum, LinkLSpace(N + 1)
'''''
Print #fnum, RbarMS1no(N + 1), RbarMS1dia(N + 1)
Print #fnum, RbarMS2curS(N + 1), RbarMS2no(N + 1), RbarMS2dia(N + 1), RbarMS2curE(N + 1) ''''
Print #fnum, LinkMSpace(N + 1)
'''''
Print #fnum, RbarTR1no(N + 1), RbarTR1dia(N + 1), RbarTR1curS(N + 1)
Print #fnum, RbarTR2no(N + 1), RbarTR2dia(N + 1), RbarTR2curS(N + 1)
Print #fnum, LinkRSpace(N + 1)

Print #fnum, RbarCRtno(N), RbarCRtdia(N), RbarCLfcurE(N + 1)   '''''''''
Print #fnum, RbarLCno(N + 1), RbarLCdia(N + 1)              ''''''''''link
Print #fnum, RbarCRtno(N + 1), RbarCRtdia(N + 1), RbarCRtcurS(N + 1) '''''''''
Close #fnum

End If


'''''''''''666666'''''''************************************************
If Command4.Left = 3080 Then
txtFile = "C:\autodraf\rasuk\input_data\SpanSixGET.txt"
Option6.Enabled = False

Transfer_TxtToData (6)
'''''''''''''''''''''
N = 6

''''''''''''''''''''''''''''
Open txtFile For Output As #fnum
Print #fnum, GridNameL(N)
Print #fnum, scL(N), sbL(N), shL(N)
Print #fnum, beamL(N), beamB(N), beamH(N)

Print #fnum, sbR(N), shR(N), scR(N)
Print #fnum, slabDrop(N), beamUplift(N)
Print #fnum, FrontSlabLvl(N), BackSlabLvl(N)
Print #fnum, GridNameR(N)
''''''''''''
Print #fnum, RbarTL1no(N), RbarTL1dia(N), RbarTL1curE(N)
Print #fnum, RbarTL2no(N), RbarTL2dia(N), RbarTL2curE(N)
Print #fnum, LinkLSpace(N)
'''''
Print #fnum, RbarMS1no(N), RbarMS1dia(N)
Print #fnum, RbarMS2curS(N), RbarMS2no(N), RbarMS2dia(N), RbarMS2curE(N)   ''''
Print #fnum, LinkMSpace(N)
'''''
Print #fnum, RbarTR1no(N), RbarTR1dia(N), RbarTR1curS(N)
Print #fnum, RbarTR2no(N), RbarTR2dia(N), RbarTR2curS(N)
Print #fnum, LinkRSpace(N)

Print #fnum, RbarCLfno(N), RbarCLfdia(N), RbarCLfcurE(N) '''''''''
Print #fnum, RbarLCno(N), RbarLCdia(N)                 ''''''''''link
Print #fnum, RbarCRtno(N), RbarCRtdia(N), RbarCRtcurS(N)  '''''''''
Close #fnum


''''BEFORE
Open "C:\autodraf\rasuk\input_data\SpanFiveGET.txt" For Output As #fnum
Print #fnum, GridNameL(N - 1)
Print #fnum, scL(N - 1), sbL(N - 1), shL(N - 1)
Print #fnum, beamL(N - 1), beamB(N - 1), beamH(N - 1)

Print #fnum, sbL(N), shL(N), scL(N)
Print #fnum, slabDrop(N - 1), beamUplift(N - 1)
Print #fnum, FrontSlabLvl(N - 1), BackSlabLvl(N - 1)
Print #fnum, GridNameL(N)
''''''''''''
Print #fnum, RbarTL1no(N - 1), RbarTL1dia(N - 1), RbarTL1curE(N - 1)
Print #fnum, RbarTL2no(N - 1), RbarTL2dia(N - 1), RbarTL2curE(N - 1)
Print #fnum, LinkLSpace(N - 1)
'''''
Print #fnum, RbarMS1no(N - 1), RbarMS1dia(N - 1) ''''
Print #fnum, RbarMS2curS(N - 1), RbarMS2no(N - 1), RbarMS2dia(N - 1), RbarMS2curE(N - 1) ''''
Print #fnum, LinkMSpace(N - 1)
'''''
Print #fnum, RbarTL1no(N), RbarTL1dia(N), RbarTR1curS(N - 1)
Print #fnum, RbarTL2no(N), RbarTL2dia(N), RbarTR2curS(N - 1)
Print #fnum, LinkRSpace(N - 1)

Print #fnum, RbarCLfno(N - 1), RbarCLfdia(N - 1), RbarCLfcurE(N - 1) '''''''''
Print #fnum, RbarLCno(N - 1), RbarLCdia(N - 1) ''''''''''link
Print #fnum, RbarCLfno(N), RbarCLfdia(N), RbarCRtcurS(N - 1)  '''''''''
Close #fnum

'''AFTER

Open "C:\autodraf\rasuk\input_data\SpanSevenGET.txt" For Output As #fnum

Print #fnum, GridNameR(N)
Print #fnum, scR(N), sbR(N), shR(N)
Print #fnum, beamL(N + 1), beamB(N + 1), beamH(N + 1)

Print #fnum, sbR(N + 1), shR(N + 1), scR(N + 1)
Print #fnum, slabDrop(N + 1), beamUplift(N + 1)
Print #fnum, FrontSlabLvl(N + 1), BackSlabLvl(N + 1)
Print #fnum, GridNameR(N + 1)
''''''''''''
Print #fnum, RbarTR1no(N), RbarTR1dia(N), RbarTL1curE(N + 1)
Print #fnum, RbarTR2no(N), RbarTR2dia(N), RbarTL2curE(N + 1)
Print #fnum, LinkLSpace(N + 1)
'''''
Print #fnum, RbarMS1no(N + 1), RbarMS1dia(N + 1)
Print #fnum, RbarMS2curS(N + 1), RbarMS2no(N + 1), RbarMS2dia(N + 1), RbarMS2curE(N + 1) ''''
Print #fnum, LinkMSpace(N + 1)
'''''
Print #fnum, RbarTR1no(N + 1), RbarTR1dia(N + 1), RbarTR1curS(N + 1)
Print #fnum, RbarTR2no(N + 1), RbarTR2dia(N + 1), RbarTR2curS(N + 1)
Print #fnum, LinkRSpace(N + 1)

Print #fnum, RbarCRtno(N), RbarCRtdia(N), RbarCLfcurE(N + 1)   '''''''''
Print #fnum, RbarLCno(N + 1), RbarLCdia(N + 1)              ''''''''''link
Print #fnum, RbarCRtno(N + 1), RbarCRtdia(N + 1), RbarCRtcurS(N + 1) '''''''''
Close #fnum

End If


'''''7777''''''''''*********************************************
If Command4.Left = 3680 Then
txtFile = "C:\autodraf\rasuk\input_data\SpanSevenGET.txt"
Option7.Enabled = False

Transfer_TxtToData (7)
'''''''''''''''''''''
N = 7

''''''''''''''''''''''''''''
Open txtFile For Output As #fnum
Print #fnum, GridNameL(N)
Print #fnum, scL(N), sbL(N), shL(N)
Print #fnum, beamL(N), beamB(N), beamH(N)

Print #fnum, sbR(N), shR(N), scR(N)
Print #fnum, slabDrop(N), beamUplift(N)
Print #fnum, FrontSlabLvl(N), BackSlabLvl(N)
Print #fnum, GridNameR(N)
''''''''''''
Print #fnum, RbarTL1no(N), RbarTL1dia(N), RbarTL1curE(N)
Print #fnum, RbarTL2no(N), RbarTL2dia(N), RbarTL2curE(N)
Print #fnum, LinkLSpace(N)
'''''
Print #fnum, RbarMS1no(N), RbarMS1dia(N)
Print #fnum, RbarMS2curS(N), RbarMS2no(N), RbarMS2dia(N), RbarMS2curE(N)   ''''
Print #fnum, LinkMSpace(N)
'''''
Print #fnum, RbarTR1no(N), RbarTR1dia(N), RbarTR1curS(N)
Print #fnum, RbarTR2no(N), RbarTR2dia(N), RbarTR2curS(N)
Print #fnum, LinkRSpace(N)

Print #fnum, RbarCLfno(N), RbarCLfdia(N), RbarCLfcurE(N) '''''''''
Print #fnum, RbarLCno(N), RbarLCdia(N)                 ''''''''''link
Print #fnum, RbarCRtno(N), RbarCRtdia(N), RbarCRtcurS(N)  '''''''''
Close #fnum


''''BEFORE
Open "C:\autodraf\rasuk\input_data\SpanSixGET.txt" For Output As #fnum
Print #fnum, GridNameL(N - 1)
Print #fnum, scL(N - 1), sbL(N - 1), shL(N - 1)
Print #fnum, beamL(N - 1), beamB(N - 1), beamH(N - 1)

Print #fnum, sbL(N), shL(N), scL(N)
Print #fnum, slabDrop(N - 1), beamUplift(N - 1)
Print #fnum, FrontSlabLvl(N - 1), BackSlabLvl(N - 1)
Print #fnum, GridNameL(N)
''''''''''''
Print #fnum, RbarTL1no(N - 1), RbarTL1dia(N - 1), RbarTL1curE(N - 1)
Print #fnum, RbarTL2no(N - 1), RbarTL2dia(N - 1), RbarTL2curE(N - 1)
Print #fnum, LinkLSpace(N - 1)
'''''
Print #fnum, RbarMS1no(N - 1), RbarMS1dia(N - 1) ''''
Print #fnum, RbarMS2curS(N - 1), RbarMS2no(N - 1), RbarMS2dia(N - 1), RbarMS2curE(N - 1) ''''
Print #fnum, LinkMSpace(N - 1)
'''''
Print #fnum, RbarTL1no(N), RbarTL1dia(N), RbarTR1curS(N - 1)
Print #fnum, RbarTL2no(N), RbarTL2dia(N), RbarTR2curS(N - 1)
Print #fnum, LinkRSpace(N - 1)

Print #fnum, RbarCLfno(N - 1), RbarCLfdia(N - 1), RbarCLfcurE(N - 1) '''''''''
Print #fnum, RbarLCno(N - 1), RbarLCdia(N - 1) ''''''''''link
Print #fnum, RbarCLfno(N), RbarCLfdia(N), RbarCRtcurS(N - 1)  '''''''''
Close #fnum

'''AFTER

Open "C:\autodraf\rasuk\input_data\SpanEightGET.txt" For Output As #fnum

Print #fnum, GridNameR(N)
Print #fnum, scR(N), sbR(N), shR(N)
Print #fnum, beamL(N + 1), beamB(N + 1), beamH(N + 1)

Print #fnum, sbR(N + 1), shR(N + 1), scR(N + 1)
Print #fnum, slabDrop(N + 1), beamUplift(N + 1)
Print #fnum, FrontSlabLvl(N + 1), BackSlabLvl(N + 1)
Print #fnum, GridNameR(N + 1)
''''''''''''
Print #fnum, RbarTR1no(N), RbarTR1dia(N), RbarTL1curE(N + 1)
Print #fnum, RbarTR2no(N), RbarTR2dia(N), RbarTL2curE(N + 1)
Print #fnum, LinkLSpace(N + 1)
'''''
Print #fnum, RbarMS1no(N + 1), RbarMS1dia(N + 1)
Print #fnum, RbarMS2curS(N + 1), RbarMS2no(N + 1), RbarMS2dia(N + 1), RbarMS2curE(N + 1) ''''
Print #fnum, LinkMSpace(N + 1)
'''''
Print #fnum, RbarTR1no(N + 1), RbarTR1dia(N + 1), RbarTR1curS(N + 1)
Print #fnum, RbarTR2no(N + 1), RbarTR2dia(N + 1), RbarTR2curS(N + 1)
Print #fnum, LinkRSpace(N + 1)

Print #fnum, RbarCRtno(N), RbarCRtdia(N), RbarCLfcurE(N + 1)   '''''''''
Print #fnum, RbarLCno(N + 1), RbarLCdia(N + 1)              ''''''''''link
Print #fnum, RbarCRtno(N + 1), RbarCRtdia(N + 1), RbarCRtcurS(N + 1) '''''''''
Close #fnum

End If





''''88888''''''''''*********************************************
If Command4.Left = 4280 Then
txtFile = "C:\autodraf\rasuk\input_data\SpanEightGET.txt"
Option8.Enabled = False

Transfer_TxtToData (8)
'''''''''''''''''''''
N = 8

''''''''''''''''''''''''''''
Open txtFile For Output As #fnum
Print #fnum, GridNameL(N)
Print #fnum, scL(N), sbL(N), shL(N)
Print #fnum, beamL(N), beamB(N), beamH(N)

Print #fnum, sbR(N), shR(N), scR(N)
Print #fnum, slabDrop(N), beamUplift(N)
Print #fnum, FrontSlabLvl(N), BackSlabLvl(N)
Print #fnum, GridNameR(N)
''''''''''''
Print #fnum, RbarTL1no(N), RbarTL1dia(N), RbarTL1curE(N)
Print #fnum, RbarTL2no(N), RbarTL2dia(N), RbarTL2curE(N)
Print #fnum, LinkLSpace(N)
'''''
Print #fnum, RbarMS1no(N), RbarMS1dia(N)
Print #fnum, RbarMS2curS(N), RbarMS2no(N), RbarMS2dia(N), RbarMS2curE(N)   ''''
Print #fnum, LinkMSpace(N)
'''''
Print #fnum, RbarTR1no(N), RbarTR1dia(N), RbarTR1curS(N)
Print #fnum, RbarTR2no(N), RbarTR2dia(N), RbarTR2curS(N)
Print #fnum, LinkRSpace(N)

Print #fnum, RbarCLfno(N), RbarCLfdia(N), RbarCLfcurE(N) '''''''''
Print #fnum, RbarLCno(N), RbarLCdia(N)                 ''''''''''link
Print #fnum, RbarCRtno(N), RbarCRtdia(N), RbarCRtcurS(N)  '''''''''
Close #fnum


''''BEFORE
Open "C:\autodraf\rasuk\input_data\SpanSevenGET.txt" For Output As #fnum
Print #fnum, GridNameL(N - 1)
Print #fnum, scL(N - 1), sbL(N - 1), shL(N - 1)
Print #fnum, beamL(N - 1), beamB(N - 1), beamH(N - 1)

Print #fnum, sbL(N), shL(N), scL(N)
Print #fnum, slabDrop(N - 1), beamUplift(N - 1)
Print #fnum, FrontSlabLvl(N - 1), BackSlabLvl(N - 1)
Print #fnum, GridNameL(N)
''''''''''''
Print #fnum, RbarTL1no(N - 1), RbarTL1dia(N - 1), RbarTL1curE(N - 1)
Print #fnum, RbarTL2no(N - 1), RbarTL2dia(N - 1), RbarTL2curE(N - 1)
Print #fnum, LinkLSpace(N - 1)
'''''
Print #fnum, RbarMS1no(N - 1), RbarMS1dia(N - 1) ''''
Print #fnum, RbarMS2curS(N - 1), RbarMS2no(N - 1), RbarMS2dia(N - 1), RbarMS2curE(N - 1) ''''
Print #fnum, LinkMSpace(N - 1)
'''''
Print #fnum, RbarTL1no(N), RbarTL1dia(N), RbarTR1curS(N - 1)
Print #fnum, RbarTL2no(N), RbarTL2dia(N), RbarTR2curS(N - 1)
Print #fnum, LinkRSpace(N - 1)

Print #fnum, RbarCLfno(N - 1), RbarCLfdia(N - 1), RbarCLfcurE(N - 1) '''''''''
Print #fnum, RbarLCno(N - 1), RbarLCdia(N - 1) ''''''''''link
Print #fnum, RbarCLfno(N), RbarCLfdia(N), RbarCRtcurS(N - 1)  '''''''''
Close #fnum

'''AFTER

Open "C:\autodraf\rasuk\input_data\SpanNineGET.txt" For Output As #fnum

Print #fnum, GridNameR(N)
Print #fnum, scR(N), sbR(N), shR(N)
Print #fnum, beamL(N + 1), beamB(N + 1), beamH(N + 1)

Print #fnum, sbR(N + 1), shR(N + 1), scR(N + 1)
Print #fnum, slabDrop(N + 1), beamUplift(N + 1)
Print #fnum, FrontSlabLvl(N + 1), BackSlabLvl(N + 1)
Print #fnum, GridNameR(N + 1)
''''''''''''
Print #fnum, RbarTR1no(N), RbarTR1dia(N), RbarTL1curE(N + 1)
Print #fnum, RbarTR2no(N), RbarTR2dia(N), RbarTL2curE(N + 1)
Print #fnum, LinkLSpace(N + 1)
'''''
Print #fnum, RbarMS1no(N + 1), RbarMS1dia(N + 1)
Print #fnum, RbarMS2curS(N + 1), RbarMS2no(N + 1), RbarMS2dia(N + 1), RbarMS2curE(N + 1) ''''
Print #fnum, LinkMSpace(N + 1)
'''''
Print #fnum, RbarTR1no(N + 1), RbarTR1dia(N + 1), RbarTR1curS(N + 1)
Print #fnum, RbarTR2no(N + 1), RbarTR2dia(N + 1), RbarTR2curS(N + 1)
Print #fnum, LinkRSpace(N + 1)

Print #fnum, RbarCRtno(N), RbarCRtdia(N), RbarCLfcurE(N + 1)   '''''''''
Print #fnum, RbarLCno(N + 1), RbarLCdia(N + 1)              ''''''''''link
Print #fnum, RbarCRtno(N + 1), RbarCRtdia(N + 1), RbarCRtcurS(N + 1) '''''''''
Close #fnum

End If



'''''99999'''''''''''*********************************************
If Command4.Left = 4880 Then
txtFile = "C:\autodraf\rasuk\input_data\SpanNineGET.txt"
Option9.Enabled = False

Transfer_TxtToData (9)
'''''''''''''''''''''
N = 9

''''''''''''''''''''''''''''
Open txtFile For Output As #fnum
Print #fnum, GridNameL(N)
Print #fnum, scL(N), sbL(N), shL(N)
Print #fnum, beamL(N), beamB(N), beamH(N)

Print #fnum, sbR(N), shR(N), scR(N)
Print #fnum, slabDrop(N), beamUplift(N)
Print #fnum, FrontSlabLvl(N), BackSlabLvl(N)
Print #fnum, GridNameR(N)
''''''''''''
Print #fnum, RbarTL1no(N), RbarTL1dia(N), RbarTL1curE(N)
Print #fnum, RbarTL2no(N), RbarTL2dia(N), RbarTL2curE(N)
Print #fnum, LinkLSpace(N)
'''''
Print #fnum, RbarMS1no(N), RbarMS1dia(N)
Print #fnum, RbarMS2curS(N), RbarMS2no(N), RbarMS2dia(N), RbarMS2curE(N)   ''''
Print #fnum, LinkMSpace(N)
'''''
Print #fnum, RbarTR1no(N), RbarTR1dia(N), RbarTR1curS(N)
Print #fnum, RbarTR2no(N), RbarTR2dia(N), RbarTR2curS(N)
Print #fnum, LinkRSpace(N)

Print #fnum, RbarCLfno(N), RbarCLfdia(N), RbarCLfcurE(N) '''''''''
Print #fnum, RbarLCno(N), RbarLCdia(N)                 ''''''''''link
Print #fnum, RbarCRtno(N), RbarCRtdia(N), RbarCRtcurS(N)  '''''''''
Close #fnum


''''BEFORE
Open "C:\autodraf\rasuk\input_data\SpanEightGET.txt" For Output As #fnum
Print #fnum, GridNameL(N - 1)
Print #fnum, scL(N - 1), sbL(N - 1), shL(N - 1)
Print #fnum, beamL(N - 1), beamB(N - 1), beamH(N - 1)

Print #fnum, sbL(N), shL(N), scL(N)
Print #fnum, slabDrop(N - 1), beamUplift(N - 1)
Print #fnum, FrontSlabLvl(N - 1), BackSlabLvl(N - 1)
Print #fnum, GridNameL(N)
''''''''''''
Print #fnum, RbarTL1no(N - 1), RbarTL1dia(N - 1), RbarTL1curE(N - 1)
Print #fnum, RbarTL2no(N - 1), RbarTL2dia(N - 1), RbarTL2curE(N - 1)
Print #fnum, LinkLSpace(N - 1)
'''''
Print #fnum, RbarMS1no(N - 1), RbarMS1dia(N - 1) ''''
Print #fnum, RbarMS2curS(N - 1), RbarMS2no(N - 1), RbarMS2dia(N - 1), RbarMS2curE(N - 1) ''''
Print #fnum, LinkMSpace(N - 1)
'''''
Print #fnum, RbarTL1no(N), RbarTL1dia(N), RbarTR1curS(N - 1)
Print #fnum, RbarTL2no(N), RbarTL2dia(N), RbarTR2curS(N - 1)
Print #fnum, LinkRSpace(N - 1)

Print #fnum, RbarCLfno(N - 1), RbarCLfdia(N - 1), RbarCLfcurE(N - 1) '''''''''
Print #fnum, RbarLCno(N - 1), RbarLCdia(N - 1) ''''''''''link
Print #fnum, RbarCLfno(N), RbarCLfdia(N), RbarCRtcurS(N - 1)  '''''''''
Close #fnum

'''AFTER
'''nil
End If

Form1.Picture = LoadPicture("C:\autodraf\icon\datas.ico")

End Sub

Private Sub Command5_Click()
OpenDataFile
'''''''''''''''
NoOfSpan = Int(Val(Right(Command4.Caption, 1))) ''
i = NoOfSpan ''
'''''''''''''''
Form1.Picture = LoadPicture("C:\autodraf\icon\datam.ico")

List1.Clear
List1.Visible = False
'If Picture1.Visible = True Then
'   Picture1.Visible = False
'   End If

'Dim fnum As Integer
'Dim txtFile, Temp As String
'fnum = FreeFile
'txtFile = "C:\autodraf\rasuk\input_data\DefaultStress.txt"
Command2.Enabled = True
Command3.Enabled = False
Command4.Enabled = False
Command6.Enabled = False
Command5.Enabled = False


Option1.Enabled = False
Option2.Enabled = False
Option3.Enabled = False
Option4.Enabled = False
Option5.Enabled = False
Option6.Enabled = False
Option7.Enabled = False
Option8.Enabled = False
Option9.Enabled = False

Txt_StatusTwo
Text8.Enabled = False
Text9.Enabled = False



Text35.text = Xinsertion
Text36.text = Yinsertion
Text57.text = FontSz
Text37.text = NamaRasuk
Text38.text = fcu
Text39.text = fy
Text40.text = fyv
Text41.text = Shrink
Text42.text = Creep
Text43.text = cVr
Text44.text = slabThick
Text45.text = stirupD
Text46.text = BarMark

Stresses1.Value35 = Text35.text
Stresses1.Value36 = Text36.text
Stresses1.Value57 = Text57.text
Stresses1.Value37 = Text37.text
Stresses1.Value38 = Text38.text
Stresses1.Value39 = Text39.text
Stresses1.Value40 = Text40.text
Stresses1.Value41 = Text41.text
Stresses1.Value42 = Text42.text
Stresses1.Value43 = Text43.text
Stresses1.Value44 = Text44.text
Stresses1.Value45 = Text45.text
Stresses1.Value46 = Text46.text

Text35.text = Stresses1.XinsertPt
Text36.text = Stresses1.YinsertPt
Text57.text = Stresses1.SetFontSize
Text37.text = Stresses1.NamaRasuk
Text38.text = Stresses1.fcu
Text39.text = Stresses1.fy
Text40.text = Stresses1.fyv
Text41.text = Stresses1.Shrink
Text42.text = Stresses1.Creep
Text43.text = Stresses1.Cover
Text44.text = Stresses1.SlabThk
Text45.text = Stresses1.LinkD
Text46.text = Stresses1.BarMark

End Sub

''''RASUK''''
Private Sub Command6_Click()
Command3.Enabled = False
Command5.Enabled = True
Command5.Visible = True
Option1.Enabled = True
Option2.Enabled = True
Option3.Enabled = True
Option4.Enabled = True
Option5.Enabled = True
Option6.Enabled = True
Option7.Enabled = True
Option8.Enabled = True
Option9.Enabled = True

OpenDataFile
  
Form1.Picture = LoadPicture("C:\autodraf\icon\ukad1.ico")
'''Label5.Caption = "..." & Right(dwgName, 50)
Components = 6
FWcoordinate
 
Command6.Enabled = False
                DrawFWRasuk
                
                DrawTetulangSatu
                
                '''DrawTetulangDua
                
                DrawColumn
                DrawGrid
                DrawDimension
                DrawNamaRasuk
                DrawStrength
                

Form1.Picture = LoadPicture("C:\autodraf\icon\rasuk1.ico")
End Sub

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




Private Sub OpenDataFile()
Form1.Picture = LoadPicture("C:\autodraf\icon\datam.ico")
Dim fnum As Integer
Dim txtFile, Temp As String
fnum = FreeFile
Dim N As Integer


txtFile = "C:\autodraf\rasuk\input_data\DefaultStress.txt"
Open txtFile For Input As #fnum
Input #fnum, Xinsertion
Input #fnum, Yinsertion
Input #fnum, FontSz
Input #fnum, NamaRasuk
Input #fnum, fcu
Input #fnum, fy
Input #fnum, fyv
Input #fnum, Shrink
Input #fnum, Creep
Input #fnum, cVr
Input #fnum, slabThick
Input #fnum, stirupD
Input #fnum, BarMark
Close #fnum

txtFile = "C:\autodraf\rasuk\input_data\SpanOneGET.txt"
''''''''''''''''''''
N = 1
''''''''''''''''''
Open txtFile For Input As #fnum

Input #fnum, GridNameL(N)
Input #fnum, scL(N), sbL(N), shL(N)
Input #fnum, beamL(N), beamB(N), beamH(N)

Input #fnum, sbR(N), shR(N), scR(N)
Input #fnum, slabDrop(N), beamUplift(N)
Input #fnum, FrontSlabLvl(N), BackSlabLvl(N)
Input #fnum, GridNameR(N)
''''''''''''
Input #fnum, RbarTL1no(N), RbarTL1dia(N), RbarTL1curE(N)
Input #fnum, RbarTL2no(N), RbarTL2dia(N), RbarTL2curE(N)
Input #fnum, LinkLSpace(N)
'''''
Input #fnum, RbarMS1no(N), RbarMS1dia(N)
Input #fnum, RbarMS2curS(N), RbarMS2no(N), RbarMS2dia(N), RbarMS2curE(N)   ''''
Input #fnum, LinkMSpace(N)
'''''
Input #fnum, RbarTR1no(N), RbarTR1dia(N), RbarTR1curS(N)
Input #fnum, RbarTR2no(N), RbarTR2dia(N), RbarTR2curS(N)
Input #fnum, LinkRSpace(N)

Input #fnum, RbarCLfno(N), RbarCLfdia(N), RbarCLfcurE(N) '''''''''
Input #fnum, RbarLCno(N), RbarLCdia(N)                 ''''''''''link
Input #fnum, RbarCRtno(N), RbarCRtdia(N), RbarCRtcurS(N)  '''''''''
Close #fnum



txtFile = "C:\autodraf\rasuk\input_data\SpanTwoGET.txt"
''''''''''''''''''''
N = 2
''''''''''''''''''
Open txtFile For Input As #fnum

Input #fnum, GridNameL(N)
Input #fnum, scL(N), sbL(N), shL(N)
Input #fnum, beamL(N), beamB(N), beamH(N)

Input #fnum, sbR(N), shR(N), scR(N)
Input #fnum, slabDrop(N), beamUplift(N)
Input #fnum, FrontSlabLvl(N), BackSlabLvl(N)
Input #fnum, GridNameR(N)
''''''''''''
Input #fnum, RbarTL1no(N), RbarTL1dia(N), RbarTL1curE(N)
Input #fnum, RbarTL2no(N), RbarTL2dia(N), RbarTL2curE(N)
Input #fnum, LinkLSpace(N)
'''''
Input #fnum, RbarMS1no(N), RbarMS1dia(N)
Input #fnum, RbarMS2curS(N), RbarMS2no(N), RbarMS2dia(N), RbarMS2curE(N)   ''''
Input #fnum, LinkMSpace(N)
'''''
Input #fnum, RbarTR1no(N), RbarTR1dia(N), RbarTR1curS(N)
Input #fnum, RbarTR2no(N), RbarTR2dia(N), RbarTR2curS(N)
Input #fnum, LinkRSpace(N)

Input #fnum, RbarCLfno(N), RbarCLfdia(N), RbarCLfcurE(N) '''''''''
Input #fnum, RbarLCno(N), RbarLCdia(N)                 ''''''''''link
Input #fnum, RbarCRtno(N), RbarCRtdia(N), RbarCRtcurS(N)  '''''''''
Close #fnum

txtFile = "C:\autodraf\rasuk\input_data\SpanThreeGET.txt"
''''''''''''''''''''
N = 3
''''''''''''''''''
Open txtFile For Input As #fnum

Input #fnum, GridNameL(N)
Input #fnum, scL(N), sbL(N), shL(N)
Input #fnum, beamL(N), beamB(N), beamH(N)

Input #fnum, sbR(N), shR(N), scR(N)
Input #fnum, slabDrop(N), beamUplift(N)
Input #fnum, FrontSlabLvl(N), BackSlabLvl(N)
Input #fnum, GridNameR(N)
''''''''''''
Input #fnum, RbarTL1no(N), RbarTL1dia(N), RbarTL1curE(N)
Input #fnum, RbarTL2no(N), RbarTL2dia(N), RbarTL2curE(N)
Input #fnum, LinkLSpace(N)
'''''
Input #fnum, RbarMS1no(N), RbarMS1dia(N)
Input #fnum, RbarMS2curS(N), RbarMS2no(N), RbarMS2dia(N), RbarMS2curE(N)   ''''
Input #fnum, LinkMSpace(N)
'''''
Input #fnum, RbarTR1no(N), RbarTR1dia(N), RbarTR1curS(N)
Input #fnum, RbarTR2no(N), RbarTR2dia(N), RbarTR2curS(N)
Input #fnum, LinkRSpace(N)

Input #fnum, RbarCLfno(N), RbarCLfdia(N), RbarCLfcurE(N) '''''''''
Input #fnum, RbarLCno(N), RbarLCdia(N)                 ''''''''''link
Input #fnum, RbarCRtno(N), RbarCRtdia(N), RbarCRtcurS(N)  '''''''''
Close #fnum


txtFile = "C:\autodraf\rasuk\input_data\SpanFourGET.txt"
''''''''''''''''''''
N = 4
''''''''''''''''''
Open txtFile For Input As #fnum

Input #fnum, GridNameL(N)
Input #fnum, scL(N), sbL(N), shL(N)
Input #fnum, beamL(N), beamB(N), beamH(N)

Input #fnum, sbR(N), shR(N), scR(N)
Input #fnum, slabDrop(N), beamUplift(N)
Input #fnum, FrontSlabLvl(N), BackSlabLvl(N)
Input #fnum, GridNameR(N)
''''''''''''
Input #fnum, RbarTL1no(N), RbarTL1dia(N), RbarTL1curE(N)
Input #fnum, RbarTL2no(N), RbarTL2dia(N), RbarTL2curE(N)
Input #fnum, LinkLSpace(N)
'''''
Input #fnum, RbarMS1no(N), RbarMS1dia(N)
Input #fnum, RbarMS2curS(N), RbarMS2no(N), RbarMS2dia(N), RbarMS2curE(N)   ''''
Input #fnum, LinkMSpace(N)
'''''
Input #fnum, RbarTR1no(N), RbarTR1dia(N), RbarTR1curS(N)
Input #fnum, RbarTR2no(N), RbarTR2dia(N), RbarTR2curS(N)
Input #fnum, LinkRSpace(N)

Input #fnum, RbarCLfno(N), RbarCLfdia(N), RbarCLfcurE(N) '''''''''
Input #fnum, RbarLCno(N), RbarLCdia(N)                 ''''''''''link
Input #fnum, RbarCRtno(N), RbarCRtdia(N), RbarCRtcurS(N)  '''''''''
Close #fnum


txtFile = "C:\autodraf\rasuk\input_data\SpanFiveGET.txt"
N = 5
''''''''''''''''''
Open txtFile For Input As #fnum

Input #fnum, GridNameL(N)
Input #fnum, scL(N), sbL(N), shL(N)
Input #fnum, beamL(N), beamB(N), beamH(N)

Input #fnum, sbR(N), shR(N), scR(N)
Input #fnum, slabDrop(N), beamUplift(N)
Input #fnum, FrontSlabLvl(N), BackSlabLvl(N)
Input #fnum, GridNameR(N)
''''''''''''
Input #fnum, RbarTL1no(N), RbarTL1dia(N), RbarTL1curE(N)
Input #fnum, RbarTL2no(N), RbarTL2dia(N), RbarTL2curE(N)
Input #fnum, LinkLSpace(N)
'''''
Input #fnum, RbarMS1no(N), RbarMS1dia(N)
Input #fnum, RbarMS2curS(N), RbarMS2no(N), RbarMS2dia(N), RbarMS2curE(N)   ''''
Input #fnum, LinkMSpace(N)
'''''
Input #fnum, RbarTR1no(N), RbarTR1dia(N), RbarTR1curS(N)
Input #fnum, RbarTR2no(N), RbarTR2dia(N), RbarTR2curS(N)
Input #fnum, LinkRSpace(N)

Input #fnum, RbarCLfno(N), RbarCLfdia(N), RbarCLfcurE(N) '''''''''
Input #fnum, RbarLCno(N), RbarLCdia(N)                 ''''''''''link
Input #fnum, RbarCRtno(N), RbarCRtdia(N), RbarCRtcurS(N)  '''''''''
Close #fnum

txtFile = "C:\autodraf\rasuk\input_data\SpanSixGET.txt"
N = 6
''''''''''''''''''
Open txtFile For Input As #fnum

Input #fnum, GridNameL(N)
Input #fnum, scL(N), sbL(N), shL(N)
Input #fnum, beamL(N), beamB(N), beamH(N)

Input #fnum, sbR(N), shR(N), scR(N)
Input #fnum, slabDrop(N), beamUplift(N)
Input #fnum, FrontSlabLvl(N), BackSlabLvl(N)
Input #fnum, GridNameR(N)
''''''''''''
Input #fnum, RbarTL1no(N), RbarTL1dia(N), RbarTL1curE(N)
Input #fnum, RbarTL2no(N), RbarTL2dia(N), RbarTL2curE(N)
Input #fnum, LinkLSpace(N)
'''''
Input #fnum, RbarMS1no(N), RbarMS1dia(N)
Input #fnum, RbarMS2curS(N), RbarMS2no(N), RbarMS2dia(N), RbarMS2curE(N)   ''''
Input #fnum, LinkMSpace(N)
'''''
Input #fnum, RbarTR1no(N), RbarTR1dia(N), RbarTR1curS(N)
Input #fnum, RbarTR2no(N), RbarTR2dia(N), RbarTR2curS(N)
Input #fnum, LinkRSpace(N)

Input #fnum, RbarCLfno(N), RbarCLfdia(N), RbarCLfcurE(N) '''''''''
Input #fnum, RbarLCno(N), RbarLCdia(N)                 ''''''''''link
Input #fnum, RbarCRtno(N), RbarCRtdia(N), RbarCRtcurS(N)  '''''''''
Close #fnum


txtFile = "C:\autodraf\rasuk\input_data\SpanSevenGET.txt"
N = 7
''''''''''''''''''
Open txtFile For Input As #fnum

Input #fnum, GridNameL(N)
Input #fnum, scL(N), sbL(N), shL(N)
Input #fnum, beamL(N), beamB(N), beamH(N)

Input #fnum, sbR(N), shR(N), scR(N)
Input #fnum, slabDrop(N), beamUplift(N)
Input #fnum, FrontSlabLvl(N), BackSlabLvl(N)
Input #fnum, GridNameR(N)
''''''''''''
Input #fnum, RbarTL1no(N), RbarTL1dia(N), RbarTL1curE(N)
Input #fnum, RbarTL2no(N), RbarTL2dia(N), RbarTL2curE(N)
Input #fnum, LinkLSpace(N)
'''''
Input #fnum, RbarMS1no(N), RbarMS1dia(N)
Input #fnum, RbarMS2curS(N), RbarMS2no(N), RbarMS2dia(N), RbarMS2curE(N)   ''''
Input #fnum, LinkMSpace(N)
'''''
Input #fnum, RbarTR1no(N), RbarTR1dia(N), RbarTR1curS(N)
Input #fnum, RbarTR2no(N), RbarTR2dia(N), RbarTR2curS(N)
Input #fnum, LinkRSpace(N)

Input #fnum, RbarCLfno(N), RbarCLfdia(N), RbarCLfcurE(N) '''''''''
Input #fnum, RbarLCno(N), RbarLCdia(N)                 ''''''''''link
Input #fnum, RbarCRtno(N), RbarCRtdia(N), RbarCRtcurS(N)  '''''''''
Close #fnum

txtFile = "C:\autodraf\rasuk\input_data\SpanEightGET.txt"
N = 8
''''''''''''''''''
Open txtFile For Input As #fnum

Input #fnum, GridNameL(N)
Input #fnum, scL(N), sbL(N), shL(N)
Input #fnum, beamL(N), beamB(N), beamH(N)

Input #fnum, sbR(N), shR(N), scR(N)
Input #fnum, slabDrop(N), beamUplift(N)
Input #fnum, FrontSlabLvl(N), BackSlabLvl(N)
Input #fnum, GridNameR(N)
''''''''''''
Input #fnum, RbarTL1no(N), RbarTL1dia(N), RbarTL1curE(N)
Input #fnum, RbarTL2no(N), RbarTL2dia(N), RbarTL2curE(N)
Input #fnum, LinkLSpace(N)
'''''
Input #fnum, RbarMS1no(N), RbarMS1dia(N)
Input #fnum, RbarMS2curS(N), RbarMS2no(N), RbarMS2dia(N), RbarMS2curE(N)   ''''
Input #fnum, LinkMSpace(N)
'''''
Input #fnum, RbarTR1no(N), RbarTR1dia(N), RbarTR1curS(N)
Input #fnum, RbarTR2no(N), RbarTR2dia(N), RbarTR2curS(N)
Input #fnum, LinkRSpace(N)

Input #fnum, RbarCLfno(N), RbarCLfdia(N), RbarCLfcurE(N) '''''''''
Input #fnum, RbarLCno(N), RbarLCdia(N)                 ''''''''''link
Input #fnum, RbarCRtno(N), RbarCRtdia(N), RbarCRtcurS(N)  '''''''''
Close #fnum


txtFile = "C:\autodraf\rasuk\input_data\SpanNineGET.txt"
N = 9
''''''''''''''''''
Open txtFile For Input As #fnum

Input #fnum, GridNameL(N)
Input #fnum, scL(N), sbL(N), shL(N)
Input #fnum, beamL(N), beamB(N), beamH(N)

Input #fnum, sbR(N), shR(N), scR(N)
Input #fnum, slabDrop(N), beamUplift(N)
Input #fnum, FrontSlabLvl(N), BackSlabLvl(N)
Input #fnum, GridNameR(N)
''''''''''''
Input #fnum, RbarTL1no(N), RbarTL1dia(N), RbarTL1curE(N)
Input #fnum, RbarTL2no(N), RbarTL2dia(N), RbarTL2curE(N)
Input #fnum, LinkLSpace(N)
'''''
Input #fnum, RbarMS1no(N), RbarMS1dia(N)
Input #fnum, RbarMS2curS(N), RbarMS2no(N), RbarMS2dia(N), RbarMS2curE(N)   ''''
Input #fnum, LinkMSpace(N)
'''''
Input #fnum, RbarTR1no(N), RbarTR1dia(N), RbarTR1curS(N)
Input #fnum, RbarTR2no(N), RbarTR2dia(N), RbarTR2curS(N)
Input #fnum, LinkRSpace(N)

Input #fnum, RbarCLfno(N), RbarCLfdia(N), RbarCLfcurE(N) '''''''''
Input #fnum, RbarLCno(N), RbarLCdia(N)                 ''''''''''link
Input #fnum, RbarCRtno(N), RbarCRtdia(N), RbarCRtcurS(N)  '''''''''
Close #fnum


  For N = 1 To 9
  LinkDia(N) = stirupD
  Next N

End Sub
Public Sub Transfer_TxtToData(ByVal N As Integer)

GridNameL(N) = Text1.text
scL(N) = Text2.text
sbL(N) = Text3.text
shL(N) = Text4.text
beamL(N) = Text5.text
beamB(N) = Text6.text
beamH(N) = Text7.text
slabDrop(N) = Text8.text
beamUplift(N) = Text9.text
sbR(N) = Text10.text
shR(N) = Text11.text
scR(N) = Text12.text
GridNameR(N) = Text13.text
FrontSlabLvl(N) = Text47.text
BackSlabLvl(N) = Text48.text
'''''''''''
RbarTL1no(N) = Text14.text
RbarTL1dia(N) = Text15.text
RbarTL1curE(N) = Text16.text
RbarTL2no(N) = Text17.text
RbarTL2dia(N) = Text18.text
RbarTL2curE(N) = Text19.text
LinkLSpace(N) = Text20.text
RbarMS1no(N) = Text21.text
RbarMS1dia(N) = Text22.text
RbarMS2curS(N) = Text23.text  ''''
RbarMS2no(N) = Text24.text
RbarMS2dia(N) = Text25.text
RbarMS2curE(N) = Text26.text ''''
LinkMSpace(N) = Text27.text
RbarTR1no(N) = Text28.text
RbarTR1dia(N) = Text29.text
RbarTR1curS(N) = Text30.text
RbarTR2no(N) = Text31.text
RbarTR2dia(N) = Text32.text
RbarTR2curS(N) = Text33.text
LinkRSpace(N) = Text34.text
RbarCLfno(N) = Text51.text    ''''''''''''
RbarCLfdia(N) = Text52.text    ''''''''''
RbarCLfcurE(N) = Text53.text   '''''''''
RbarLCno(N) = Text49.text    ''''''''''''link
RbarLCdia(N) = Text50.text    ''''''''''link
RbarCRtno(N) = Text54.text    ''''''''''''
RbarCRtdia(N) = Text55.text    ''''''''''
RbarCRtcurS(N) = Text56.text   '''''''''

End Sub
Public Sub Transfer_DataToTxt(ByVal N As Integer)

Text1.text = GridNameL(N)
Text2.text = scL(N)
Text3.text = sbL(N)
Text4.text = shL(N)
 
Text5.text = beamL(N)
Text6.text = beamB(N)
Text7.text = beamH(N)
Text8.text = slabDrop(N)
Text9.text = beamUplift(N)


Text10.text = sbR(N)
Text11.text = shR(N)
Text12.text = scR(N)
Text13.text = GridNameR(N)
  
Text47.text = FrontSlabLvl(N)
Text48.text = BackSlabLvl(N)
''''''''''''''''''''''''''
Text14.text = RbarTL1no(N)
Text15.text = RbarTL1dia(N)
Text16.text = RbarTL1curE(N)
Text17.text = RbarTL2no(N)
Text18.text = RbarTL2dia(N)
Text19.text = RbarTL2curE(N)
Text20.text = LinkLSpace(N)
Text21.text = RbarMS1no(N)
Text22.text = RbarMS1dia(N)
Text23.text = RbarMS2curS(N)  ''''
Text24.text = RbarMS2no(N)
Text25.text = RbarMS2dia(N)
Text26.text = RbarMS2curE(N)  ''''
Text27.text = LinkMSpace(N)
Text28.text = RbarTR1no(N)
Text29.text = RbarTR1dia(N)
Text30.text = RbarTR1curS(N)
Text31.text = RbarTR2no(N)
Text32.text = RbarTR2dia(N)
Text33.text = RbarTR2curS(N)
Text34.text = LinkRSpace(N)

Text51.text = RbarCLfno(N)   ''''''''''''
Text52.text = RbarCLfdia(N)  ''''''''''
Text53.text = RbarCLfcurE(N) '''''''''
Text49.text = RbarLCno(N)   ''''''''''''link
Text50.text = RbarLCdia(N)    ''''''''''link
Text54.text = RbarCRtno(N)   ''''''''''''
Text55.text = RbarCRtdia(N)   ''''''''''
Text56.text = RbarCRtcurS(N)  '''''''''


End Sub




Private Sub Txt_StatusOne()

Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
'''''Text8.Enabled = False
'''''Text9.Enabled = False
Text10.Enabled = True
Text11.Enabled = True
Text12.Enabled = True
Text13.Enabled = True
''''''''''''''''''''''''''
Text14.Enabled = True
Text15.Enabled = True
Text16.Enabled = True
Text17.Enabled = True
Text18.Enabled = True
Text19.Enabled = True
Text20.Enabled = True
Text21.Enabled = True
Text22.Enabled = True
Text23.Enabled = True   ''''
Text24.Enabled = True
Text25.Enabled = True
Text26.Enabled = True  ''''
Text27.Enabled = True
Text28.Enabled = True
Text29.Enabled = True
Text30.Enabled = True
Text31.Enabled = True
Text32.Enabled = True
Text33.Enabled = True
Text34.Enabled = True

Text35.Enabled = False
Text36.Enabled = False
Text57.Enabled = False
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
Text47.Enabled = True
Text48.Enabled = True

Text49.Enabled = True
Text50.Enabled = True
Text51.Enabled = True
Text52.Enabled = True
Text53.Enabled = True
Text54.Enabled = True
Text55.Enabled = True
Text56.Enabled = True


End Sub
Private Sub Txt_StatusTwo()
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
Text7.Enabled = False
'''''Text8.Enabled = False
'''''Text9.Enabled = False
Text10.Enabled = False
Text11.Enabled = False
Text12.Enabled = False
Text13.Enabled = False
''''''''''''''''''''''''''
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
Text26.Enabled = False  ''''
Text27.Enabled = False
Text28.Enabled = False
Text29.Enabled = False
Text30.Enabled = False
Text31.Enabled = False
Text32.Enabled = False
Text33.Enabled = False
Text34.Enabled = False

Text35.Enabled = True
Text36.Enabled = True
Text57.Enabled = True
Text37.Enabled = True
Text38.Enabled = True
Text39.Enabled = True
Text40.Enabled = True
Text41.Enabled = True
Text42.Enabled = True
Text43.Enabled = True
Text44.Enabled = True
Text45.Enabled = True
Text46.Enabled = True

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

End Sub




''''FOR RASUK''''
Private Sub FWcoordinate()
''Form1.Picture = LoadPicture("C:\autodraf\icon\coordinate.ico")
Dim j, N, n1, n2 As Integer
Dim h, Drop, delX As Double

N = NoOfSpan
h = beamH(1)
pbx(0) = Xinsertion
For j = 1 To N
n1 = 8 * j
If j > 1 Then
    n2 = 8 * (j - 1)
    delX = sbR(j - 1)
   Else
    n2 = 2
    delX = sbR(j)
End If
pbx(n1 - 8) = pbx(n2 - 2) + sbL(j) + delX
If j = 1 Then
   pbx(n1 - 8) = pbx(n1 - 8) + Xinsertion
End If
'MsgBox " 1st:distance ", , n1 - 8 & "   " & pbx(n1 - 8)
pbx(n1 - 6) = pbx(n1 - 8)
'MsgBox " 2nd:  ", , n1 - 6 & "   " & pbx(n1 - 6)
pbx(n1 - 4) = pbx(n1 - 6) + beamL(j) - sbL(j) - sbR(j)
'MsgBox " 3rd:  ", , n1 - 4 & "   " & pbx(n1 - 4)
pbx(n1 - 2) = pbx(n1 - 4)
'MsgBox " 4th:  ", , n1 - 2 & "   " & pbx(n1 - 2)
Next j

'BEAM SURFACE
N = NoOfSpan
Dim dropL, dropR, tempY As Double
tempY = beamH(1) + Yinsertion
For j = 1 To N
If j = 1 Then
   dropL = slabDrop(j)
   Else
   dropL = slabDrop(j - 1)
End If
If j = N Then
   dropR = slabDrop(j)
   Else
   dropR = slabDrop(j + 1)
End If
If slabDrop(j) - dropL <= 0 Then
  pbt(8 * j - 7) = tempY - slabDrop(j)
Else
  pbt(8 * j - 7) = tempY - dropL
End If
pbt(8 * j - 5) = tempY - slabDrop(j)
pbt(8 * j - 3) = tempY - slabDrop(j)
If slabDrop(j) - dropR <= 0 Then
  pbt(8 * j - 1) = tempY - slabDrop(j)
Else
  pbt(8 * j - 1) = tempY - dropR
End If
Next j
  
'SLAB SOFFIT
N = NoOfSpan
Dim slabCounter, k As Integer
Dim Slab As Double
Slab = slabThick
slabCounter = 0
n1 = 1
For j = 1 To 8 * N - 1 Step 2
Select Case slabCounter

   Case 0: pbs(j) = pbt(j) - shL(n1)
           slabCounter = slabCounter + 1
   Case 1: pbs(j) = pbt(j) - Slab
           slabCounter = slabCounter + 1
   Case 2: pbs(j) = pbt(j) - Slab
           slabCounter = slabCounter + 1
   Case 3: pbs(j) = pbt(j) - shR(n1)
           slabCounter = 0
           n1 = n1 + 1
End Select
'MsgBox " slabsoffit: ", , j & "   " & pbt(j) & "  " & pbs(j)
Next


'BEAM BOTTOM
N = NoOfSpan
Dim UpliftL, UpliftR As Double
tempY = Yinsertion
For j = 1 To N
If j = 1 Then
   UpliftL = beamUplift(j)
   Else
   UpliftL = beamUplift(j - 1)
End If
If j = N Then
   UpliftR = beamUplift(j)
   Else
   UpliftR = beamUplift(j + 1)
End If
If UpliftL >= beamUplift(j) Then
      pbb(8 * j - 7) = tempY + beamUplift(j)
    Else
      pbb(8 * j - 7) = tempY + UpliftL
End If
pbb(8 * j - 5) = tempY + beamUplift(j)
pbb(8 * j - 3) = tempY + beamUplift(j)
If UpliftR >= 0 Then
  pbb(8 * j - 1) = tempY + beamUplift(j)
Else
  pbb(8 * j - 1) = tempY + UpliftR
End If
Next j
''''Form1.Picture = LoadPicture("C:\autodraf\icon\ukad1.ico")

End Sub


''''FOR RASUK''''
Private Sub DrawFWRasuk()
Form1.Picture = LoadPicture("C:\autodraf\icon\ukad5.ico")
Set moSpace = acadDoc.ModelSpace
ReDim pt(0 To 8 * NoOfSpan - 1) As Double

For i = 0 To 8 * NoOfSpan - 1 Step 2
pt(i) = pbx(i)
pt(i + 1) = pbt(i + 1)
Next i
Dim PolyPt As Object

   Set PolyPt = moSpace.AddLightWeightPolyline(pt)
   PolyPt.layer = "FormWork"
   PolyPt.Update
   
pbx(NoOfSpan * 8) = pbb(NoOfSpan * 8 - 2)
pbb(NoOfSpan * 8 + 1) = pbb(NoOfSpan * 8 - 1)
For i = 0 To 8 * NoOfSpan - 1 Step 2
pt(i) = pbx(i)
pt(i + 1) = pbs(i + 1)
Next i
   Set PolyPt = moSpace.AddLightWeightPolyline(pt)
   PolyPt.layer = "Slab"
   PolyPt.Update

For i = 0 To 8 * NoOfSpan - 1 Step 2
pt(i) = pbx(i)
pt(i + 1) = pbb(i + 1)
If pt(i + 1) > pbs(i + 1) Then
   pt(i + 1) = pbs(i + 1)
End If

   If pt(i + 1) > pbb(i + 3) And pt(i) <> pbx(i + 2) Then ''''additional condition
      pt(i + 1) = pbb(i + 3)
   End If

Next i
   Set PolyPt = moSpace.AddLightWeightPolyline(pt)
   PolyPt.layer = "FormWork"
   PolyPt.Update

Dim pt1(0 To 7) As Double
pt1(0) = pbx(0)
pt1(1) = pbt(1)
pt1(2) = pt1(0) - 2 * sbL(1)
pt1(3) = pt1(1)
pt1(4) = pt1(2)
pt1(5) = pbb(1)
If pt1(5) > pbs(1) Then
   pt1(5) = pbs(1)
End If
pt1(6) = pt1(4) + 2 * sbL(1)
pt1(7) = pt1(5)
Dim Polypt1 As Object

   Set Polypt1 = moSpace.AddLightWeightPolyline(pt1)
   Polypt1.layer = "FormWork"
   Polypt1.Update
   
pt1(0) = pbx(NoOfSpan * 8 - 2)
pt1(1) = pbt(NoOfSpan * 8 - 1)
pt1(2) = pt1(0) + 2 * sbR(NoOfSpan)
pt1(3) = pt1(1)
pt1(4) = pt1(2)
pt1(5) = pbb(NoOfSpan * 8 - 1)
If pt1(5) > pbs(NoOfSpan * 8 - 1) Then
   pt1(5) = pbs(NoOfSpan * 8 - 1)
End If
pt1(6) = pt1(4) - 2 * sbR(NoOfSpan)
pt1(7) = pt1(5)
   Set Polypt1 = moSpace.AddLightWeightPolyline(pt1)
   Polypt1.layer = "FormWork"
   Polypt1.Update
   
 
Dim pt2(0 To 3) As Double
pt2(0) = pbx(0)
pt2(1) = pbs(1)
pt2(2) = pt2(0) - 2 * sbL(1)
pt2(3) = pt2(1)
Dim Polypt2 As Object
   Set Polypt2 = moSpace.AddLightWeightPolyline(pt2)
   Polypt2.layer = "Slab"
   Polypt2.Update
   
pt2(0) = pbx(NoOfSpan * 8 - 2)
pt2(1) = pbs(NoOfSpan * 8 - 1)
pt2(2) = pt2(0) + 2 * sbR(NoOfSpan)
pt2(3) = pt2(1)
   Set Polypt2 = moSpace.AddLightWeightPolyline(pt2)
   Polypt2.layer = "Slab"
   Polypt2.Update
 Form1.Picture = LoadPicture("C:\autodraf\icon\ukad3.ico")
 
End Sub
''''FOR RASUK''''
Private Sub DrawTetulangSatu()
'''Form1.Picture = LoadPicture("C:\autodraf\icon\tetulang.ico")
Set moSpace = acadDoc.ModelSpace
Dim xStat, yStat As Double
Dim caseI As Integer
Dim pt0(0 To 31) As Double
Dim pt0a(0 To 15) As Double
Dim pt1(0 To 7) As Double
Dim pt2(0 To 11) As Double
Dim polypt0 As Object

Dim polypt0a As Object

Dim Polypt1 As Object

Dim Polypt2 As Object

Cv1 = 0.309016
Cv2 = 0.278769
Cv3 = 0.221231
Cv4 = 0.142041
Cv5 = 0.048944

caseI = 0
For i = 1 To NoOfSpan
   caseI = 1
      If i = 1 And i <> NoOfSpan Then
        caseI = 0
          End If
       If i = NoOfSpan And i <> 1 Then
          caseI = 2
              End If
           If i = 1 And i = NoOfSpan Then
              caseI = 0
                  End If
SWdthLft = scL(i)
SWdthRght = scR(i)
     If scL(i) = 0 Then
        SWdthLft = sbL(i)
           End If
     If scR(i) = 0 Then
        SWdthRght = sbR(i)
           End If


Select Case caseI
Case 0:
      Call KesSifar
Case 1:
      Call KesSatu
Case 2:
      Call KesDua
End Select

  Call DrawLink(pbx(8 * i - 6) - sbL(i) + SWdthLft, _
              pbb(8 * i - 5) + cVr + LinkDia(i) / 2, _
              beamB(i), beamH(i), LinkDia(i), _
              RbarTL1curE(i) - SWdthLft, LinkLSpace(i), _
              beamL(i) - RbarTL1curE(i) - RbarTR1curS(i), _
              LinkMSpace(i), RbarTR1curS(i) - SWdthRght, _
              LinkRSpace(i))
                            
              If i = 1 Then
                   Bar6No = RbarTL1no(i)
                      Bar6Dia = RbarTL1dia(i)
                        Bar6BM = Bar1BM
                         Else
                      Bar6No = RbarCLfno(i)
                  Bar6Dia = RbarCLfdia(i)
                Bar6BM = Bar6BM
              End If
    Call DrawSection(pbx(8 * i - 8) + RbarTL1curE(i) / 2, _
              Yinsertion - 800, _
              LinkDia(i), beamB(i), beamH(i), slabThick, _
              FrontSlabLvl(i), BackSlabLvl(i), _
              RbarTL1no(i), RbarTL1dia(i), Bar1BM, _
              RbarTL2no(i), RbarTL2dia(i), Bar2BM, _
              RbarLCno(i), RbarLCdia(i), Bar3BM, _
              RbarMS1no(i), RbarMS1dia(i), Bar4BM, _
              RbarMS2no(i), RbarMS2dia(i), Bar5BM, _
              Bar6No, Bar6Dia, Bar6BM)
              
     Call DrawLinkSection(pbx(8 * i - 8) + RbarTL1curE(i) / 2 + _
              slabThick + cVr + 2.5 * LinkDia(i), _
              Yinsertion - 800 - cVr - LinkDia(i) / 2, _
              LinkDia(i), beamB(i), beamH(i), slabThick, _
              RbarTL1no(i), RbarTL1dia(i), Bar1BM, _
              RbarTL2no(i), RbarTL2dia(i), Bar2BM, _
              RbarLCno(i), RbarLCdia(i), Bar3BM, _
              RbarMS1no(i), RbarMS1dia(i), Bar4BM, _
              RbarMS2no(i), RbarMS2dia(i), Bar5BM, _
              Bar6No, Bar6Dia, Bar6BM)
       
     Call CircleBarMark(pbx(8 * i - 8) + RbarTL1curE(i) / 2 + _
              slabThick + cVr + 2.5 * LinkDia(i), _
              Yinsertion - 800 - cVr - LinkDia(i) / 2, _
              LinkDia(i), beamB(i), beamH(i), slabThick, _
              RbarTL1no(i), RbarTL1dia(i), Bar1BM, _
              RbarTL2no(i), RbarTL2dia(i), Bar2BM, _
              RbarLCno(i), RbarLCdia(i), Bar3BM, _
              RbarMS1no(i), RbarMS1dia(i), Bar4BM, _
              RbarMS2no(i), RbarMS2dia(i), Bar5BM, _
              Bar6No, Bar6Dia, Bar6BM)


  Next i
End Sub
'''''''tetulang jenis dua ''vvvvvvvvvv
''''FOR RASUK''''
Private Sub DrawTetulangDua()
'''Form1.Picture = LoadPicture("C:\autodraf\icon\tetulang.ico")
Set moSpace = acadDoc.ModelSpace
Dim xStat, yStat As Double
Dim caseI As Integer
Dim pt0(0 To 31) As Double
Dim pt0a(0 To 15) As Double
Dim pt1(0 To 7) As Double
Dim pt2(0 To 11) As Double
Dim polypt0 As Object

Dim polypt0a As Object

Dim Polypt1 As Object

Dim Polypt2 As Object

Cv1 = 0.309016
Cv2 = 0.278769
Cv3 = 0.221231
Cv4 = 0.142041
Cv5 = 0.048944

caseI = 0
For i = 1 To NoOfSpan
   caseI = 1
      If i = 1 And i <> NoOfSpan Then
        caseI = 0
          End If
       If i = NoOfSpan And i <> 1 Then
          caseI = 2
              End If
           If i = 1 And i = NoOfSpan Then
              caseI = 0
                  End If
SWdthLft = scL(i)
SWdthRght = scR(i)
     If scL(i) = 0 Then
        SWdthLft = sbL(i)
           End If
     If scR(i) = 0 Then
        SWdthRght = sbR(i)
           End If


Select Case caseI
Case 0:
      Call CaseZERO
Case 1:
      Call CaseONE
Case 2:
      Call CaseTWO
End Select

   Call DrawLink(pbx(8 * i - 6) - sbL(i) + SWdthLft, _
              pbb(8 * i - 5) + cVr + LinkDia(i) / 2, _
              beamB(i), beamH(i), LinkDia(i), _
              RbarTL1curE(i) - SWdthLft, LinkLSpace(i), _
              beamL(i) - RbarTL1curE(i) - RbarTR1curS(i), _
              LinkMSpace(i), RbarTR1curS(i) - SWdthRght, _
              LinkRSpace(i))
                            
              If i = 1 Then
                   Bar6No = RbarTL1no(i)
                      Bar6Dia = RbarTL1dia(i)
                        Bar6BM = Bar1BM
                         Else
                      Bar6No = RbarCLfno(i)
                  Bar6Dia = RbarCLfdia(i)
                Bar6BM = Bar6BM
              End If
    Call DrawSectionDua(pbx(8 * i - 8) + RbarTL1curE(i) / 2, _
              Yinsertion - 800, _
              LinkDia(i), beamB(i), beamH(i), slabThick, _
              FrontSlabLvl(i), BackSlabLvl(i), _
              RbarTL1no(i), RbarTL1dia(i), Bar1BM, _
              RbarTL2no(i), RbarTL2dia(i), Bar2BM, _
              RbarLCno(i), RbarLCdia(i), Bar3BM, _
              RbarMS1no(i), RbarMS1dia(i), Bar4BM, _
              RbarMS2no(i), RbarMS2dia(i), Bar5BM, _
              Bar6No, Bar6Dia, Bar6BM)
              
     Call DrawLinkSectionDua(pbx(8 * i - 8) + RbarTL1curE(i) / 2 + _
              slabThick + cVr + 2.5 * LinkDia(i), _
              Yinsertion - 800 - cVr - LinkDia(i) / 2, _
              LinkDia(i), beamB(i), beamH(i), slabThick, _
              RbarTL1no(i), RbarTL1dia(i), Bar1BM, _
              RbarTL2no(i), RbarTL2dia(i), Bar2BM, _
              RbarLCno(i), RbarLCdia(i), Bar3BM, _
              RbarMS1no(i), RbarMS1dia(i), Bar4BM, _
              RbarMS2no(i), RbarMS2dia(i), Bar5BM, _
              Bar6No, Bar6Dia, Bar6BM)
       
     Call CircleBarMarkDua(pbx(8 * i - 8) + RbarTL1curE(i) / 2 + _
              slabThick + cVr + 2.5 * LinkDia(i), _
              Yinsertion - 800 - cVr - LinkDia(i) / 2, _
              LinkDia(i), beamB(i), beamH(i), slabThick, _
              RbarTL1no(i), RbarTL1dia(i), Bar1BM, _
              RbarTL2no(i), RbarTL2dia(i), Bar2BM, _
              RbarLCno(i), RbarLCdia(i), Bar3BM, _
              RbarMS1no(i), RbarMS1dia(i), Bar4BM, _
              RbarMS2no(i), RbarMS2dia(i), Bar5BM, _
              Bar6No, Bar6Dia, Bar6BM)



  Next i
End Sub
'''^^^^^tetulang jenis dua ''''''''



''''FOR RASUK''''
Private Sub DrawColumn()
'''Form1.Picture = LoadPicture("C:\autodraf\icon\tiang.ico")
Set moSpace = acadDoc.ModelSpace
Dim delX As Double
Dim pt0(0 To 27) As Double
Dim polypt0 As Object

Dim pt1(0 To 15) As Double
Dim Polypt1 As Object


For i = 1 To NoOfSpan
'''''''''''''
If i = 1 Then '''
'''''''''''''
pt0(0) = pbx(0) + scL(i) - sbL(i)
pt0(1) = pbb(8 * i - 7)
pt0(2) = pt0(0)
pt0(3) = pt0(1) - 1 * beamB(1)  ''''scL(i)
delX = 2 * scL(i) / 3
pt0(4) = pt0(2) - delX
pt0(5) = pt0(3)
pt0(6) = pt0(4)
pt0(7) = pt0(5) + delX / 2
pt0(8) = pt0(6) - delX
pt0(9) = pt0(7) - delX
pt0(10) = pt0(8)
pt0(11) = pt0(9) + delX / 2
pt0(12) = pt0(10) - delX
pt0(13) = pt0(11)
pt0(14) = pt0(12)
pt0(15) = pbt(8 * i - 7) + beamB(1)  '''scL(i)
pt0(16) = pt0(14) + delX
pt0(17) = pt0(15)
pt0(18) = pt0(16)
pt0(19) = pt0(17) - delX / 2
pt0(20) = pt0(18) + delX
pt0(21) = pt0(19) + delX
pt0(22) = pt0(20)
pt0(23) = pt0(21) - delX / 2
pt0(24) = pt0(22) + delX
pt0(25) = pt0(23)
pt0(26) = pt0(24)
pt0(27) = pbt(8 * i - 7)
   Set polypt0 = moSpace.AddLightWeightPolyline(pt0)
   polypt0.layer = "Column"
   polypt0.Update
   ''Call polypt0.setwidth(pt0, 5, 10)
   
   
End If

'''''''''''''''''''''''''''''''''
If i <> 1 And i <> NoOfSpan Then ''
'''''''''''''''''''''''''''''''''
pt1(0) = pbx(8 * i - 12) - scR(i - 1) + sbR(i - 1)
pt1(1) = pbt(8 * i - 11)
pt1(2) = pt1(0)
pt1(3) = pt1(1) + beamB(1)  '''scR(i - 1)
delX = (scR(i - 1) + scL(i)) / 3
pt1(4) = pt1(2) + delX
pt1(5) = pt1(3)
pt1(6) = pt1(4)
pt1(7) = pt1(5) - delX / 2
pt1(8) = pt1(6) + delX
pt1(9) = pt1(7) + delX
pt1(10) = pt1(8)
pt1(11) = pt1(9) - delX / 2
pt1(12) = pt1(10) + delX
pt1(13) = pt1(11)
pt1(14) = pt1(12)
pt1(15) = pbt(8 * i - 5)
Set Polypt1 = moSpace.AddLightWeightPolyline(pt1)
   Polypt1.layer = "Column"
   Polypt1.Update
pt1(0) = pbx(8 * i - 12) - scR(i - 1) + sbR(i - 1)
pt1(1) = pbb(8 * i - 11)
pt1(2) = pt1(0)
pt1(3) = pt1(1) - 1 * beamB(1)  '''scR(i - 1)
delX = (scR(i - 1) + scL(i)) / 3
pt1(4) = pt1(2) + delX
pt1(5) = pt1(3)
pt1(6) = pt1(4)
pt1(7) = pt1(5) - delX / 2
pt1(8) = pt1(6) + delX
pt1(9) = pt1(7) + delX
pt1(10) = pt1(8)
pt1(11) = pt1(9) - delX / 2
pt1(12) = pt1(10) + delX
pt1(13) = pt1(11)
pt1(14) = pt1(12)
pt1(15) = pbb(8 * i - 5)
Set Polypt1 = moSpace.AddLightWeightPolyline(pt1)
   Polypt1.layer = "Column"
   Polypt1.Update
End If

''''''''''''''''''''''
If i = NoOfSpan Then  ''
''''''''''''''''''''''
    If i = 1 Then
        pt0(0) = pbx(6) - scR(1) + sbR(1)
        pt0(1) = pbb(7)
    Else
pt0(0) = pbx(8 * i - 2) - scR(i) + sbR(i)
pt0(1) = pbb(8 * i - 1)
    End If
pt0(2) = pt0(0)
pt0(3) = pt0(1) - 1 * beamB(1)  '''scR(i)
delX = 2 * scR(i) / 3
pt0(4) = pt0(2) + delX
pt0(5) = pt0(3)
pt0(6) = pt0(4)
pt0(7) = pt0(5) - delX / 2
pt0(8) = pt0(6) + delX
pt0(9) = pt0(7) + delX
pt0(10) = pt0(8)
pt0(11) = pt0(9) - delX / 2
pt0(12) = pt0(10) + delX
pt0(13) = pt0(11)
pt0(14) = pt0(12)
    If i = 1 Then
    pt0(15) = pbt(5) + beamB(1)  '''scR(1)
    Else
pt0(15) = pbt(8 * i - 1) + beamB(1)  '''scR(i)
    End If
pt0(16) = pt0(14) - delX
pt0(17) = pt0(15)
pt0(18) = pt0(16)
pt0(19) = pt0(17) + delX / 2
pt0(20) = pt0(18) - delX
pt0(21) = pt0(19) - delX
pt0(22) = pt0(20)
pt0(23) = pt0(21) + delX / 2
pt0(24) = pt0(22) - delX
pt0(25) = pt0(23)
pt0(26) = pt0(24)
If i = 1 Then
    pt0(27) = pbt(5)
    Else
pt0(27) = pbt(8 * i - 1)
    End If

   Set polypt0 = moSpace.AddLightWeightPolyline(pt0)
   polypt0.layer = "Column"
   polypt0.Update
End If

If i <> 1 Then
pt1(0) = pbx(8 * i - 12) - scR(i - 1) + sbR(i - 1)
pt1(1) = pbt(8 * i - 11)
pt1(2) = pt1(0)
pt1(3) = pt1(1) + beamB(1)  ''''scR(i - 1)
delX = (scR(i - 1) + scL(i)) / 3
pt1(4) = pt1(2) + delX
pt1(5) = pt1(3)
pt1(6) = pt1(4)
pt1(7) = pt1(5) - delX / 2
pt1(8) = pt1(6) + delX
pt1(9) = pt1(7) + delX
pt1(10) = pt1(8)
pt1(11) = pt1(9) - delX / 2
pt1(12) = pt1(10) + delX
pt1(13) = pt1(11)
pt1(14) = pt1(12)
pt1(15) = pbt(8 * i - 5)
Set Polypt1 = moSpace.AddLightWeightPolyline(pt1)
   Polypt1.layer = "Column"
   Polypt1.Update
pt1(0) = pbx(8 * i - 12) - scR(i - 1) + sbR(i - 1)
pt1(1) = pbb(8 * i - 11)
pt1(2) = pt1(0)
pt1(3) = pt1(1) - 1 * beamB(1) '''scR(i - 1)
delX = (scR(i - 1) + scL(i)) / 3
pt1(4) = pt1(2) + delX
pt1(5) = pt1(3)
pt1(6) = pt1(4)
pt1(7) = pt1(5) - delX / 2
pt1(8) = pt1(6) + delX
pt1(9) = pt1(7) + delX
pt1(10) = pt1(8)
pt1(11) = pt1(9) - delX / 2
pt1(12) = pt1(10) + delX
pt1(13) = pt1(11)
pt1(14) = pt1(12)
pt1(15) = pbb(8 * i - 5)
Set Polypt1 = moSpace.AddLightWeightPolyline(pt1)
   Polypt1.layer = "Column"
   Polypt1.Update
 End If
Next i

''''''''''''''''''''''
End Sub

''''FOR RASUK''''
Private Sub DrawGrid()
'''Form1.Picture = LoadPicture("C:\autodraf\icon\grid.ico")
Dim xStat, yStat, j, Radius, Length, textHgt, Rotate As Double
Set moSpace = acadDoc.ModelSpace
Dim pt(0 To 3) As Double
Dim PolyPt As Object
Dim center(0 To 2) As Double
Dim insPnt(0 To 2) As Double
Dim textStr As String
Dim circleObj As Object

Dim gridnameObj As Object

Length = 0
xStat = pbx(0) - sbL(1)
yStat = pbt(1) + beamB(1) + 25 '''scL(1)
For j = 1 To NoOfSpan
Length = Length + beamL(j)
pt(0) = xStat + Length
pt(1) = yStat
pt(2) = pt(0)
pt(3) = pt(1) + 700
Set PolyPt = moSpace.AddLightWeightPolyline(pt)
             PolyPt.layer = "Grid"
             PolyPt.Update
 
  center(0) = pt(2)
  center(1) = pt(3) + 4 * FontSz
  center(2) = 0
  Radius = 4 * FontSz
  Set circleObj = moSpace.AddCircle(center, Radius)
  circleObj.layer = "Grid"
  circleObj.Update
  
  insPnt(0) = center(0) - FontSz
  insPnt(1) = center(1) - FontSz
  insPnt(2) = 0
  textHgt = 3 * FontSz
  textStr = GridNameR(j)
  Set gridnameObj = moSpace.AddText(textStr, insPnt, textHgt)
   gridnameObj.layer = "Grid"
   gridnameObj.Update
   
  Next j
''''''''''''''''''''''''
 pt(0) = pbx(0) - sbL(1)
 pt(1) = pbt(1) + beamB(1) + FontSz / 2
 pt(2) = pt(0)
 pt(3) = pt(1) + 700
 Set PolyPt = moSpace.AddLightWeightPolyline(pt)
              PolyPt.layer = "Grid"
              PolyPt.Update
              
  center(0) = pt(2)
  center(1) = pt(3) + 4 * FontSz
  center(2) = 0
  Radius = 4 * FontSz
  Set circleObj = moSpace.AddCircle(center, Radius)
  circleObj.layer = "Grid"
  circleObj.Update
  
  insPnt(0) = center(0)
  insPnt(1) = center(1) - FontSz
  insPnt(2) = 0
  textHgt = 3 * FontSz
  textStr = GridNameL(1)
  Set gridnameObj = moSpace.AddText(textStr, insPnt, textHgt)
   gridnameObj.layer = "Grid"
   gridnameObj.Update
   
End Sub

''''FOR RASUK''''
Private Sub DrawDimension()
'''Form1.Picture = LoadPicture("C:\autodraf\icon\dimensi.ico")
Dim xLoc, yLoc, j, textHgt, Rotate As Double
Set moSpace = acadDoc.ModelSpace
Dim pt(0 To 3) As Double
Dim insPnt(0 To 2) As Double
Dim textStr As String
Dim LengthObj As Object


xLoc = pbx(0) - sbL(1)
yLoc = pbt(1) + 600
For j = 1 To NoOfSpan
Call ArrowHorizontal(xLoc, yLoc, beamL(j), 60, 30, "BeamDimension")
    
  insPnt(0) = xLoc + beamL(j) * 0.4
  insPnt(1) = yLoc + 40
  insPnt(2) = 0
  textHgt = 2 * FontSz
  textStr = Str(beamL(j))
  Set LengthObj = moSpace.AddText(textStr, insPnt, textHgt)
  LengthObj.layer = "BeamDimension"
  LengthObj.Update
  xLoc = xLoc + beamL(j)
  Next j
End Sub

''''FOR RASUK''''
Private Sub DrawNamaRasuk()
'''Form1.Picture = LoadPicture("C:\autodraf\icon\rasuk.ico")
Dim xLoc, yLoc, k, textHgt, Rotate As Double
Set moSpace = acadDoc.ModelSpace
Dim insPnt(0 To 2) As Double
Dim textStr, retTextFont As String
Dim NamaRasukObj As Object

xLoc = pbx(0) + beamL(1) / 5

yLoc = pbb(1) - 2000

  insPnt(0) = xLoc
  insPnt(1) = yLoc
  insPnt(2) = 0
  textHgt = 4 * FontSz
  textStr = NamaRasuk
  Set NamaRasukObj = moSpace.AddText(textStr, insPnt, textHgt)
  NamaRasukObj.layer = "BeamName"
  NamaRasukObj.Update
  
  'NamaRasukObj.fontfile = "Times New Roman"
  'retText = NamaRasukObj.fontfile
  
End Sub

''''FOR RASUK''''
Private Sub DrawStrength()
Form1.Picture = LoadPicture("C:\autodraf\icon\strength.ico")
Dim xLoc, yLoc, j, textHgt, Rotate As Double
Set moSpace = acadDoc.ModelSpace
Dim pt(0 To 3) As Double
Dim insPnt(0 To 2) As Double
Dim textStr As String
Dim shearL, shearM, shearR As Integer
Dim momentL, momentM, momentR As Integer
Dim cL1, cL2, cM1, cM2, cR1, cR2, Deflect, Pi As Double
Dim dL, dpL, asvL, astL, ascL As Double
Dim dM, dpM, asvM, astM, ascM As Double
Dim dR, dpR, asvR, astR, ascR As Double
Dim barGap, totalAst As Double
Dim Crack As Double
Dim LengthObj As Object

NoOfSpan = Val(Right(Command4.Caption, 1))
Pi = 3.141592654
For j = 1 To NoOfSpan
''''''''''''''''''
If RbarTL2dia(j) = 0 Or RbarTL2no(j) = 0 Then
   barGap = RbarTL1dia(j) / 2
   Else
   barGap = RbarTL1dia(j) + 10
   End If
totalAst = RbarTL1no(j) * Pi * RbarTL1dia(j) ^ 2 / 4 + _
           RbarTL2no(j) * Pi * RbarTL2dia(j) ^ 2 / 4

dL = beamH(j) - cVr - LinkDia(j) - RbarLCdia(j) - barGap
dpL = cVr + LinkDia(j) + RbarMS1dia(j) + RbarCLfdia(j) / 2
asvL = 2 * Pi * LinkDia(j) ^ 2 / 4
astL = totalAst
ascL = RbarCLfno(j) * Pi * RbarCLfdia(j) ^ 2 / 4
shearL = Shear.CalcShear(LinkDia(j), cVr, LinkLSpace(j), asvL, _
         0, 0, fcu, fy, fyv, beamB(j), beamH(j), dL, dpL, astL, ascL)
momentL = Moment.CalcMoment(fcu, fy, beamB(j), beamB(j), beamH(j), _
          dL, dpL, astL, ascL, 0, 0, 0, 0)
cL1 = Curve1.ACurvature(1, fcu, Shrink, Creep, momentL, momentL, _
                    beamB(j), beamH(j), dL, dpL, astL, ascL, "L")
cL2 = Curve2.ACurvature(2, fcu, Shrink, Creep, momentL, momentL, _
                    beamB(j), beamH(j), dL, dpL, astL, ascL, "L")
''''''''''''''''''''
If RbarMS2dia(j) = 0 Or RbarMS2no(j) = 0 Then
   barGap = RbarMS1dia(j) / 2
   Else
   barGap = RbarMS1dia(j) + 10
   End If
totalAst = RbarMS1no(j) * Pi * RbarMS1dia(j) ^ 2 / 4 + _
           RbarMS2no(j) * Pi * RbarMS2dia(j) ^ 2 / 4
             
dM = beamH(j) - cVr - LinkDia(j) - barGap
dpM = cVr + LinkDia(j) + RbarLCdia(j) / 2
asvM = 2 * Pi * LinkDia(j) ^ 2 / 4
astM = totalAst
ascM = RbarLCno(j) * Pi * RbarLCdia(j) ^ 2 / 4
shearM = Shear.CalcShear(LinkDia(j), cVr, LinkLSpace(j), asvL, _
         0, 0, fcu, fy, fyv, beamB(j), beamH(j), dM, dpM, astM, ascM)
momentM = Moment.CalcMoment(fcu, fy, beamB(j), beamB(j), beamH(j), _
          dM, dpM, astM, ascM, 0, 0, 0, 0)
cM1 = Curve1.ACurvature(1, fcu, Shrink, Creep, momentM, momentM, _
                    beamB(j), beamH(j), dM, dpM, astM, ascM, "L")
cM2 = Curve2.ACurvature(2, fcu, Shrink, Creep, momentM, momentM, _
                    beamB(j), beamH(j), dM, dpM, astM, ascM, "L")
''''''''''''''''''''
If RbarTR2dia(j) = 0 Or RbarTR2no(j) = 0 Then
   barGap = RbarTR1dia(j) / 2
   Else
   barGap = RbarTR1dia(j) + 10
   End If
totalAst = RbarTR1no(j) * Pi * RbarTR1dia(j) ^ 2 / 4 + _
           RbarTR2no(j) * Pi * RbarTR2dia(j) ^ 2 / 4
                   
dR = beamH(j) - cVr - LinkDia(j) - barGap
dpR = cVr + LinkDia(j) + RbarMS1dia(j) + RbarCRtdia(j) / 2
asvR = 2 * Pi * LinkDia(j) ^ 2 / 4
astR = totalAst
ascR = RbarCRtno(j) * Pi * RbarCRtdia(j) ^ 2 / 4
shearR = Shear.CalcShear(LinkDia(j), cVr, LinkLSpace(j), asvR, _
         0, 0, fcu, fy, fyv, beamB(j), beamH(j), dR, dpR, astR, ascR)
momentR = Moment.CalcMoment(fcu, fy, beamB(j), beamB(j), beamH(j), _
          dR, dpR, astR, ascR, 0, 0, 0, 0)
cR1 = Curve1.ACurvature(1, fcu, Shrink, Creep, momentR, momentR, _
                    beamB(j), beamH(j), dR, dpR, astR, ascR, "L")
cR2 = Curve2.ACurvature(2, fcu, Shrink, Creep, momentR, momentR, _
                    beamB(j), beamH(j), dR, dpR, astR, ascR, "L")
                                                       
Deflect = Deflection.CalcDeflection(beamL(j), cM1, cM2, -cL1, -cL2, -cR1, -cR2)
    
 
  insPnt(0) = pbx(8 * j - 8) + scL(j) + 400
  insPnt(1) = pbt(8 * j - 7) + 1000
  insPnt(2) = 0
  textHgt = 90
  textStr = "Mc = " & Str(momentL) & "kNm"
  Set LengthObj = moSpace.AddText(textStr, insPnt, textHgt)
  '''LengthObj.Color = 30
  LengthObj.layer = "Structural_Strength"
  LengthObj.Update
  insPnt(0) = pbx(8 * j - 8) + scL(j) + 400
  insPnt(1) = pbt(8 * j - 7) + 850
  insPnt(2) = 0
  textHgt = 90
  textStr = "Vc = " & Str(shearL) & "kN"
  Set LengthObj = moSpace.AddText(textStr, insPnt, textHgt)
  '''LengthObj.Color = 211
  LengthObj.layer = "Structural_Strength"
  LengthObj.Update
  ''''''''''''''''
    
  insPnt(0) = pbx(8 * j - 8) + beamL(j) / 2 - 400 ''''''
  insPnt(1) = pbb(8 * j - 7) - 600
  insPnt(2) = 0
  textHgt = 90
  textStr = "Mc = " & Str(momentM) & "kNm"
  Set LengthObj = moSpace.AddText(textStr, insPnt, textHgt)
  '''LengthObj.Color = 30
  LengthObj.layer = "Structural_Strength"
  LengthObj.Update
  insPnt(0) = pbx(8 * j - 8) + beamL(j) / 2 - 400 ''''''
  insPnt(1) = pbb(8 * j - 7) - 750
  insPnt(2) = 0
  textHgt = 90
  textStr = "Vc = " & Str(shearM) & "kN"
  Set LengthObj = moSpace.AddText(textStr, insPnt, textHgt)
  '''LengthObj.Color = 211
  LengthObj.layer = "Structural_Strength"
  LengthObj.Update
   
  insPnt(0) = pbx(8 * j - 2) - scL(j) - 1100  '''''
  insPnt(1) = pbt(8 * j - 1) + 1000
  insPnt(2) = 0
  textHgt = 90
  textStr = "Mc = " & Str(momentR) & "kNm"
  Set LengthObj = moSpace.AddText(textStr, insPnt, textHgt)
  '''LengthObj.Color = 30
  LengthObj.layer = "Structural_Strength"
  LengthObj.Update
  insPnt(0) = pbx(8 * j - 2) - scL(j) - 1100   '''''
  insPnt(1) = pbt(8 * j - 1) + 850
  insPnt(2) = 0
  textHgt = 90
  textStr = "Vc = " & Str(shearR) & "kN"
  Set LengthObj = moSpace.AddText(textStr, insPnt, textHgt)
  '''LengthObj.Color = 211
  LengthObj.layer = "Structural_Strength"
  LengthObj.Update
  Form1.Picture = LoadPicture("C:\autodraf\icon\ukad4.ico")
 
          
  Crack = CrackWidth.CalcCrackWidth("BotMiddleBar", momentM, fcu, fy, _
  RbarMS1dia(j), RbarMS1no(j), ascM, astM, beamB(j), beamH(j), stirupD + cVr, _
  stirupD + cVr)
           
  Crack = CrackWidth.CalcCrackWidth("Corner", momentM, fcu, fy, _
  RbarMS1dia(j), RbarMS1no(j), ascM, astM, beamB(j), beamH(j), stirupD + cVr, _
  stirupD + cVr)

  Next j

End Sub
Private Sub CalculateStrength()
Form1.Picture = LoadPicture("C:\autodraf\icon\strength.ico")

Dim shearL, shearM, shearR, j, q As Integer
Dim momentL, momentM, momentR As Integer
Dim cL1, cL2, cM1, cM2, cR1, cR2, Deflect, Pi As Double
Dim dL, dpL, asvL, astL, ascL As Double
Dim dM, dpM, asvM, astM, ascM As Double
Dim dR, dpR, asvR, astR, ascR As Double
Dim barGap, totalAst As Double
Dim CrackB, CrackC As Double


Pi = 3.141592654
j = Val(Right(Command4.Caption, 1))
''''''''''''''''''
If RbarTL2dia(j) = 0 Or RbarTL2no(j) = 0 Then
   barGap = RbarTL1dia(j) / 2
   Else
   barGap = RbarTL1dia(j) + 10
   End If
totalAst = RbarTL1no(j) * Pi * RbarTL1dia(j) ^ 2 / 4 + _
           RbarTL2no(j) * Pi * RbarTL2dia(j) ^ 2 / 4

dL = beamH(j) - cVr - LinkDia(j) - RbarLCdia(j) - barGap
dpL = cVr + LinkDia(j) + RbarMS1dia(j) + RbarCLfdia(j) / 2
asvL = 2 * Pi * LinkDia(j) ^ 2 / 4    ''to check*****************
astL = totalAst
ascL = RbarCLfno(j) * RbarCLfdia(j) ^ 2 * Pi / 4
shearL = Shear.CalcShear(LinkDia(j), cVr, LinkLSpace(j), asvL, _
         0, 0, fcu, fy, fyv, beamB(j), beamH(j), dL, dpL, astL, ascL)
momentL = Moment.CalcMoment(fcu, fy, beamB(j), beamB(j), beamH(j), _
          dL, dpL, astL, ascL, 0, 0, 0, 0)
cL1 = Curve1.ACurvature(1, fcu, Shrink, Creep, momentL, momentL, _
                    beamB(j), beamH(j), dL, dpL, astL, ascL, "L")
cL2 = Curve2.ACurvature(2, fcu, Shrink, Creep, momentL, momentL, _
                    beamB(j), beamH(j), dL, dpL, astL, ascL, "L")
''''''''''''''''''''

If RbarMS2dia(j) = 0 Or RbarMS2no(j) = 0 Then
   barGap = RbarMS1dia(j) / 2
   Else
   barGap = RbarMS1dia(j) + 10
   End If
totalAst = RbarMS1no(j) * Pi * RbarMS1dia(j) ^ 2 / 4 + _
           RbarMS2no(j) * Pi * RbarMS2dia(j) ^ 2 / 4
             
dM = beamH(j) - cVr - LinkDia(j) - barGap
dpM = cVr + LinkDia(j) + RbarLCdia(j) / 2
asvM = 2 * Pi * LinkDia(j) ^ 2 / 4
astM = totalAst
ascM = RbarLCno(j) * RbarLCdia(j) ^ 2 * Pi / 4
shearM = Shear.CalcShear(LinkDia(j), cVr, LinkLSpace(j), asvL, _
         0, 0, fcu, fy, fyv, beamB(j), beamH(j), dM, dpM, astM, ascM)
momentM = Moment.CalcMoment(fcu, fy, beamB(j), beamB(j), beamH(j), _
          dM, dpM, astM, ascM, 0, 0, 0, 0)
cM1 = Curve1.ACurvature(1, fcu, Shrink, Creep, momentM, momentM, _
                    beamB(j), beamH(j), dM, dpM, astM, ascM, "L")
cM2 = Curve2.ACurvature(2, fcu, Shrink, Creep, momentM, momentM, _
                    beamB(j), beamH(j), dM, dpM, astM, ascM, "L")
''''''''''''''''''''

If RbarTR2dia(j) = 0 Or RbarTR2no(j) = 0 Then
   barGap = RbarTR1dia(j) / 2
   Else
   barGap = RbarTR1dia(j) + 10
   End If
totalAst = RbarTR1no(j) * Pi * RbarTR1dia(j) ^ 2 / 4 + _
           RbarTR2no(j) * Pi * RbarTR2dia(j) ^ 2 / 4
                   
dR = beamH(j) - cVr - LinkDia(j) - RbarLCdia(j) - barGap
dpR = cVr + LinkDia(j) + RbarMS1dia(j) + RbarCRtdia(j) / 2
asvR = 2 * Pi * LinkDia(j) ^ 2 / 4
astR = totalAst
ascR = RbarCRtno(j) * RbarCRtdia(j) ^ 2 * Pi / 4
shearR = Shear.CalcShear(LinkDia(j), cVr, LinkLSpace(j), asvR, _
         0, 0, fcu, fy, fyv, beamB(j), beamH(j), dR, dpR, astR, ascR)
momentR = Moment.CalcMoment(fcu, fy, beamB(j), beamB(j), beamH(j), _
          dR, dpR, astR, ascR, 0, 0, 0, 0)
cR1 = Curve1.ACurvature(1, fcu, Shrink, Creep, momentR, momentR, _
                    beamB(j), beamH(j), dR, dpR, astR, ascR, "L")
cR2 = Curve2.ACurvature(2, fcu, Shrink, Creep, momentR, momentR, _
                    beamB(j), beamH(j), dR, dpR, astR, ascR, "L")

Deflect = Deflection.CalcDeflection(beamL(j), cM1, cM2, -cL1, -cL2, -cR1, -cR2)

Dim DefLimit As Double
DefLimit = beamL(j) / 500
If DefLimit > 20 Then
   DefLimit = 20
   End If

Dim allowMnt, allowDefl As Double
allowMnt = 1

For q = 1 To 200
cM1 = Curve1.ACurvature(1, fcu, Shrink, Creep, allowMnt, allowMnt, _
                    beamB(j), beamH(j), dM, dpM, astM, ascM, "L")
cM2 = Curve2.ACurvature(2, fcu, Shrink, Creep, momentM, momentM, _
                    beamB(j), beamH(j), dM, dpM, astM, ascM, "L")
                                                 
allowDefl = Deflection.CalcDeflection(beamL(j), cM1, cM2, -cL1, -cL2, -cR1, -cR2)
            If allowDefl > DefLimit Then
                GoTo 999
                    End If
allowMnt = allowMnt + 5
Next
999  ''jumper

CrackB = CrackWidth.CalcCrackWidth("BotMiddleBar", momentM, fcu, fy, _
  RbarMS1dia(j), RbarMS1no(j), ascM, astM, beamB(j), beamH(j), stirupD + cVr, _
  stirupD + cVr)
 
          
CrackC = CrackWidth.CalcCrackWidth("Corner", momentM, fcu, fy, _
  RbarMS1dia(j), RbarMS1no(j), ascM, astM, beamB(j), beamH(j), stirupD + cVr, _
  stirupD + cVr)
 
Dim EquivUltLoad  As Double
EquivUltLoad = (momentM + (momentL + momentR) / 2) * 8 / (beamL(j) / 1000) ^ 2
    
List1.Visible = True
List1.Height = 1900
List1.Top = 4985
List1.Width = 6900
''List1.BackColor = vbBlack
''List1.ForeColor = vbWhite
''List1.FontSize = 8
''List1.ForeColor = &HFF&
''List1.FontBold = True
'''''''''''''''''''''''''''(1)'''''''''''''''''''''''''''''''''''''
List1.AddItem "               M = " & Str(momentL) & _
              "                          <<ULTIMATE CAPACITY>>" & _
              "                    M = " & Str(momentR)
''''''''''''''''''''''''''''(2)''''''''''''''''''''''''''''''''''''
List1.AddItem "               V = " & Str(shearL) & "            " & _
              "                                                 " & _
              "                                 V = " & Str(shearR)
'''''''''''''''''''''''''''''(3)''''''''''''''''''''''''''''''''''
List1.AddItem "                                      " & _
              "                               M = " & Str(momentM)
'''''''''''''''''''''''''''''''(4)'''''''''''''''''''''''''''''''''''
List1.AddItem "                                      " & _
              "                               V = " & Str(shearM)
''''''''''''''''''''''''''''''''''/'''''''''''''''''''''''''''''''
List1.AddItem "               Deflection = " & Str(Int(Deflect * 100) / 100) & _
              "          Crack width:   @ soffit = " & Str(Int(CrackB * 100) / 100) & _
              "       @ corner = " & Str(Int(CrackC * 100) / 100)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
List1.AddItem "--------------------------------------------------" & _
Command4.Caption & "--------------------------------------------------------"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
List1.AddItem "                   Effective midspan Moment = " & Str(Int(allowMnt * 100) / 100) & _
              "kNm   upto the ( L/500 or 20mm ) "
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
List1.AddItem "     Equivalent Ultimate Load Carrying Capacity W = " & Str(Int(EquivUltLoad)) & _
               "kN/m   for continuous beam. "
              
End Sub



''''GENERAL''''
Public Sub SetLayer()
'''''Form1.Picture = LoadPicture("C:\autodraf\icon\layer.ico")
  
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
layerObj.Color = 1

Set layerObj = acadDoc.Layers.Add("LabelRebarSupt")
layerObj.Color = 7

Set layerObj = acadDoc.Layers.Add("RebarSpan")
layerObj.Color = 1

Set layerObj = acadDoc.Layers.Add("LabelRebarSpan")
layerObj.Color = 7

Set layerObj = acadDoc.Layers.Add("RebarLink")
layerObj.Color = 2

Set layerObj = acadDoc.Layers.Add("LabelRebarLink")
layerObj.Color = 7

Set layerObj = acadDoc.Layers.Add("Curtailment")
layerObj.Color = 7

Set layerObj = acadDoc.Layers.Add("BeamDimension")
layerObj.Color = 7

Set layerObj = acadDoc.Layers.Add("BeamSection")
layerObj.Color = 120   '' for formwork; 51 link; 30 rebar

Set layerObj = acadDoc.Layers.Add("BeamName")
layerObj.Color = 255

Set layerObj = acadDoc.Layers.Add("Structural_Strength")
layerObj.Color = 9

End Sub









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
    
'''''''''''''''''''''''''
OpenDataFile
NoOfSpan = Val(Right(Command4.Caption, 1))
i = NoOfSpan
'''''''''''''''''''''''''''
    
If dwgName = "" Then
 mnuItemExit.Enabled = True
 mnuItemOpenDwg.Enabled = True
 
 Form1.Picture = LoadPicture("C:\autodraf\icon\ukad.ico")
Else
 mnuItemExit.Enabled = True
 mnuItemOpenDwg.Enabled = False
 
 Form1.Picture = LoadPicture("C:\autodraf\icon\ukad1.ico")
 ''mnuItemFile.Enabled = False
 '' Command1.Enabled = True
 
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Form1.Picture = LoadPicture("C:\autodraf\icon\ukad3.ico")
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim SetLuput As Date
SetLuput = DateValue("6/15/3020")
If Date >= SetLuput Then
 MsgBox ":::Sila hubungi Wan Sohaimi Wan Mohamed @ 603-61574717::: ", , "To reinstall"
 Exit Sub
End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

If dwgName = "" Then
MsgBox "Sila pilih fail dwg.", , "NOTA AM:"
Exit Sub
End If

StartAutoCAD
SetLayer

Option1.Enabled = True
Option2.Enabled = True
Option3.Enabled = True
Option4.Enabled = True
Option5.Enabled = True
Option6.Enabled = True
Option7.Enabled = True
Option8.Enabled = True
Option9.Enabled = True
''''Command4.Enabled = True
 mnuItemExit.Enabled = True
 mnuItemOpenDwg.Enabled = False
 
 mnuItemFile.Enabled = True
 mnuItemFile.Visible = True
 mnuItemFile.Caption = "Jalan keluar!"
''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''
 
End If
    
       
End Sub


Private Sub mnuItemExit_Click()  'when user clicks Exit command
    End                          'quit program
End Sub




Private Sub Option1_Click()
Image1.Picture = LoadPicture("C:\autodraf\icon\ukad3.ico")
OpenDataFile
Command5.Visible = False

List1.Clear
List1.Visible = False

Dim fnum As Integer
Dim txtFile, Temp As String
fnum = FreeFile
txtFile = "C:\autodraf\rasuk\input_data\SpanOneGET.txt"
Command3.Enabled = False
Command4.Enabled = True
Command4.Left = 80
Command4.Caption = "Span 1"
Command6.Enabled = False
Command5.Enabled = False
NoOfSpan = 1
i = NoOfSpan

Option1.Enabled = True
Option2.Enabled = False
Option3.Enabled = False
Option4.Enabled = False
Option5.Enabled = False
Option6.Enabled = False
Option7.Enabled = False
Option8.Enabled = False
Option9.Enabled = False
Txt_StatusOne
Text7.Enabled = True
Text8.Enabled = False
Text9.Enabled = False

Transfer_DataToTxt (1)
''''''''''''''''''''''

Span1.Value1 = Text1.text
Span1.Value2 = Text2.text
Span1.Value3 = Text3.text
Span1.Value4 = Text4.text
Span1.Value5 = Text5.text
Span1.Value6 = Text6.text
Span1.Value7 = Text7.text
Span1.Value8 = Text8.text
Span1.Value9 = Text9.text
Span1.Value10 = Text10.text
Span1.Value11 = Text11.text
Span1.Value12 = Text12.text
Span1.Value13 = Text13.text
Span1.Value47 = Text47.text
Span1.Value48 = Text48.text
'''''''''''
Tetulang1.Value14 = Text14.text
Tetulang1.Value15 = Text15.text
Tetulang1.Value16 = Text16.text
Tetulang1.Value17 = Text17.text
Tetulang1.Value18 = Text18.text
Tetulang1.Value19 = Text19.text
Tetulang1.Value20 = Text20.text
Tetulang1.Value21 = Text21.text
Tetulang1.Value22 = Text22.text
Tetulang1.Value23 = Text23.text  ''''
Tetulang1.Value24 = Text24.text
Tetulang1.Value25 = Text25.text
Tetulang1.Value26 = Text26.text ''''
Tetulang1.Value27 = Text27.text
Tetulang1.Value28 = Text28.text
Tetulang1.Value29 = Text29.text
Tetulang1.Value30 = Text30.text
Tetulang1.Value31 = Text31.text
Tetulang1.Value32 = Text32.text
Tetulang1.Value33 = Text33.text
Tetulang1.Value34 = Text34.text

Tetulang1.Value51 = Text51.text   ''''''''''''
Tetulang1.Value52 = Text52.text  ''''''''''
Tetulang1.Value53 = Text53.text  '''''''''
Tetulang1.Value49 = Text49.text   ''''''''''''link
Tetulang1.Value50 = Text50.text    ''''''''''link
Tetulang1.Value54 = Text54.text   ''''''''''''
Tetulang1.Value55 = Text55.text   ''''''''''
Tetulang1.Value56 = Text56.text  '''''''''

Text1.text = Span1.LeftGrid
''Text1.FontBold = True
Text2.text = Span1.LeftColumnSize
Text3.text = Span1.LeftXbeamB
Text4.text = Span1.LeftXbeamH
Text5.text = Span1.BeamLength
Text6.text = Span1.BeamBreadth
Text7.text = Span1.BeamHeight
Text8.text = 0   '''Span1.BeamTopDrop
Text9.text = 0   '''Span1.BeamSoffitDrop
Text10.text = Span1.RightXbeamb
Text11.text = Span1.RightXbeamh
Text12.text = Span1.RightColumnSize
Text13.text = Span1.RightGrid
''Text13.FontBold = True
Text47.text = Span1.FrontSlabLv
Text48.text = Span1.BackSlabLv
''''''''''''''''''''''''''
Text14.text = Tetulang1.FirstLTBno
Text15.text = Tetulang1.FirstLTBdia
Text16.text = Tetulang1.FirstLTBcurt
Text17.text = Tetulang1.SecondLTBno
Text18.text = Tetulang1.SecondLTBdia
Text19.text = Tetulang1.SecondLTBcurt
Text20.text = Tetulang1.LinkSpacingLHS
Text21.text = Tetulang1.FirstMBBno
Text22.text = Tetulang1.FirstMBBdia
Text23.text = Tetulang1.SecondMBBcurtS ''''
Text24.text = Tetulang1.SecondMBBno
Text25.text = Tetulang1.SecondMBBdia
Text26.text = Tetulang1.SecondMBBcurtE ''''
Text27.text = Tetulang1.LinkSpacingMID
Text28.text = Tetulang1.FirstRTBno
Text29.text = Tetulang1.FirstRTBdia
Text30.text = Tetulang1.FirstRTBcurt
Text31.text = Tetulang1.SecondRTBno
Text32.text = Tetulang1.SecondRTBdia
Text33.text = Tetulang1.SecondRTBcurt
Text34.text = Tetulang1.LinkSpacingRHS

Text51.text = Tetulang1.LeftConnBarNo   ''''''''''''
Text52.text = Tetulang1.LeftConnBarDia  ''''''''''
Text53.text = Tetulang1.LeftConnBarCurtE '''''''''
Text49.text = Tetulang1.LinkCarrierNo  ''''''''''''link
Text50.text = Tetulang1.LinkCarrierDia   ''''''''''link
Text54.text = Tetulang1.RightConnBarNo ''''''''''''
Text55.text = Tetulang1.RightConnBarDia  ''''''''''
Text56.text = Tetulang1.RightConnBarCurtS '''''''''

''MsgBox "  " & Text21.text, , "text21"

End Sub

Private Sub Option2_Click()
Image1.Picture = LoadPicture("C:\autodraf\icon\cskp.ico")
OpenDataFile
Command5.Visible = False

List1.Clear
List1.Visible = False
Dim fnum As Integer
Dim txtFile, Temp As String
fnum = FreeFile
txtFile = "C:\autodraf\rasuk\input_data\SpanTwoGET.txt"
Command3.Enabled = False
Command4.Enabled = True
Command4.Left = 680
Command4.Caption = "Span 2"
Command6.Enabled = False
Command5.Enabled = False
NoOfSpan = 2
i = NoOfSpan

Option1.Enabled = False
Option2.Enabled = True
Option3.Enabled = False
Option4.Enabled = False
Option5.Enabled = False
Option6.Enabled = False
Option7.Enabled = False
Option8.Enabled = False
Option9.Enabled = False
Txt_StatusOne
Text7.Enabled = False
Text8.Enabled = True
Text9.Enabled = True

Transfer_DataToTxt (2)
''''''''''''''''''''''

Span2.Value1 = Text1.text
Span2.Value2 = Text2.text
Span2.Value3 = Text3.text
Span2.Value4 = Text4.text
Span2.Value5 = Text5.text
Span2.Value6 = Text6.text
Span2.Value7 = Text7.text
Span2.Value8 = Text8.text
Span2.Value9 = Text9.text
Span2.Value10 = Text10.text
Span2.Value11 = Text11.text
Span2.Value12 = Text12.text
Span2.Value13 = Text13.text
Span2.Value47 = Text47.text
Span2.Value48 = Text48.text
'''''''''''


'''''''''''
Tetulang2.Value14 = Text14.text
Tetulang2.Value15 = Text15.text
Tetulang2.Value16 = Text16.text
Tetulang2.Value17 = Text17.text
Tetulang2.Value18 = Text18.text
Tetulang2.Value19 = Text19.text
Tetulang2.Value20 = Text20.text
Tetulang2.Value21 = Text21.text
Tetulang2.Value22 = Text22.text
Tetulang2.Value23 = Text23.text  ''''
Tetulang2.Value24 = Text24.text
Tetulang2.Value25 = Text25.text
Tetulang2.Value26 = Text26.text ''''
Tetulang2.Value27 = Text27.text
Tetulang2.Value28 = Text28.text
Tetulang2.Value29 = Text29.text
Tetulang2.Value30 = Text30.text
Tetulang2.Value31 = Text31.text
Tetulang2.Value32 = Text32.text
Tetulang2.Value33 = Text33.text
Tetulang2.Value34 = Text34.text

Tetulang2.Value51 = Text51.text   ''''''''''''
Tetulang2.Value52 = Text52.text  ''''''''''
Tetulang2.Value53 = Text53.text  '''''''''
Tetulang2.Value49 = Text49.text   ''''''''''''link
Tetulang2.Value50 = Text50.text    ''''''''''link
Tetulang2.Value54 = Text54.text   ''''''''''''
Tetulang2.Value55 = Text55.text   ''''''''''
Tetulang2.Value56 = Text56.text  '''''''''

Text1.text = Span2.LeftGrid
Text1.FontBold = True
Text2.text = Span2.LeftColumnSize
Text3.text = Span2.LeftXbeamB
Text4.text = Span2.LeftXbeamH
Text5.text = Span2.BeamLength
Text6.text = Span2.BeamBreadth
Text7.text = Span2.BeamHeight
Text8.text = Span2.BeamTopDrop
Text9.text = Span2.BeamSoffitDrop
Text10.text = Span2.RightXbeamb
Text11.text = Span2.RightXbeamh
Text12.text = Span2.RightColumnSize
Text13.text = Span2.RightGrid
Text13.FontBold = True
Text47.text = Span2.FrontSlabLv
Text48.text = Span2.BackSlabLv

''''''''''''''''''''''''''
Text14.text = Tetulang2.FirstLTBno
Text15.text = Tetulang2.FirstLTBdia
Text16.text = Tetulang2.FirstLTBcurt
Text17.text = Tetulang2.SecondLTBno
Text18.text = Tetulang2.SecondLTBdia
Text19.text = Tetulang2.SecondLTBcurt
Text20.text = Tetulang2.LinkSpacingLHS
Text21.text = Tetulang2.FirstMBBno
Text22.text = Tetulang2.FirstMBBdia
Text23.text = Tetulang2.SecondMBBcurtS ''''
Text24.text = Tetulang2.SecondMBBno
Text25.text = Tetulang2.SecondMBBdia
Text26.text = Tetulang2.SecondMBBcurtE ''''
Text27.text = Tetulang2.LinkSpacingMID
Text28.text = Tetulang2.FirstRTBno
Text29.text = Tetulang2.FirstRTBdia
Text30.text = Tetulang2.FirstRTBcurt
Text31.text = Tetulang2.SecondRTBno
Text32.text = Tetulang2.SecondRTBdia
Text33.text = Tetulang2.SecondRTBcurt
Text34.text = Tetulang2.LinkSpacingRHS

Text51.text = Tetulang2.LeftConnBarNo   ''''''''''''
Text52.text = Tetulang2.LeftConnBarDia  ''''''''''
Text53.text = Tetulang2.LeftConnBarCurtE '''''''''
Text49.text = Tetulang2.LinkCarrierNo  ''''''''''''link
Text50.text = Tetulang2.LinkCarrierDia   ''''''''''link
Text54.text = Tetulang2.RightConnBarNo ''''''''''''
Text55.text = Tetulang2.RightConnBarDia  ''''''''''
Text56.text = Tetulang2.RightConnBarCurtS '''''''''

'''MsgBox "  " & Text21.text, , "text21"


End Sub

Private Sub Option3_Click()
Image1.Picture = LoadPicture("C:\autodraf\icon\ukad.ico")
OpenDataFile
Command5.Visible = False
 

List1.Clear
List1.Visible = False
Dim fnum As Integer
Dim txtFile, Temp As String
fnum = FreeFile
txtFile = "C:\autodraf\rasuk\input_data\SpanThreeGET.txt"
Command3.Enabled = False
Command4.Enabled = True
Command4.Left = 1280
Command4.Caption = "Span 3"
Command6.Enabled = False
Command5.Enabled = False
NoOfSpan = 3
i = NoOfSpan

Option1.Enabled = False
Option2.Enabled = False
Option3.Enabled = True
Option4.Enabled = False
Option5.Enabled = False
Option6.Enabled = False
Option7.Enabled = False
Option8.Enabled = False
Option9.Enabled = False
Txt_StatusOne
Text7.Enabled = False
Text8.Enabled = True
Text9.Enabled = True

Transfer_DataToTxt (3)
''''''''''''''''''''''


Span3.Value1 = Text1.text
Span3.Value2 = Text2.text
Span3.Value3 = Text3.text
Span3.Value4 = Text4.text
Span3.Value5 = Text5.text
Span3.Value6 = Text6.text
Span3.Value7 = Text7.text
Span3.Value8 = Text8.text
Span3.Value9 = Text9.text
Span3.Value10 = Text10.text
Span3.Value11 = Text11.text
Span3.Value12 = Text12.text
Span3.Value13 = Text13.text
Span3.Value47 = Text47.text
Span3.Value48 = Text48.text

'''''''''''
Tetulang3.Value14 = Text14.text
Tetulang3.Value15 = Text15.text
Tetulang3.Value16 = Text16.text
Tetulang3.Value17 = Text17.text
Tetulang3.Value18 = Text18.text
Tetulang3.Value19 = Text19.text
Tetulang3.Value20 = Text20.text
Tetulang3.Value21 = Text21.text
Tetulang3.Value22 = Text22.text
Tetulang3.Value23 = Text23.text  ''''
Tetulang3.Value24 = Text24.text
Tetulang3.Value25 = Text25.text
Tetulang3.Value26 = Text26.text ''''
Tetulang3.Value27 = Text27.text
Tetulang3.Value28 = Text28.text
Tetulang3.Value29 = Text29.text
Tetulang3.Value30 = Text30.text
Tetulang3.Value31 = Text31.text
Tetulang3.Value32 = Text32.text
Tetulang3.Value33 = Text33.text
Tetulang3.Value34 = Text34.text

Tetulang3.Value51 = Text51.text   ''''''''''''
Tetulang3.Value52 = Text52.text  ''''''''''
Tetulang3.Value53 = Text53.text  '''''''''
Tetulang3.Value49 = Text49.text   ''''''''''''link
Tetulang3.Value50 = Text50.text    ''''''''''link
Tetulang3.Value54 = Text54.text   ''''''''''''
Tetulang3.Value55 = Text55.text   ''''''''''
Tetulang3.Value56 = Text56.text  '''''''''

Text1.text = Span3.LeftGrid
Text1.FontBold = True
Text2.text = Span3.LeftColumnSize
Text3.text = Span3.LeftXbeamB
Text4.text = Span3.LeftXbeamH
Text5.text = Span3.BeamLength
Text6.text = Span3.BeamBreadth
Text7.text = Span3.BeamHeight
Text8.text = Span3.BeamTopDrop
Text9.text = Span3.BeamSoffitDrop
Text10.text = Span3.RightXbeamb
Text11.text = Span3.RightXbeamh
Text12.text = Span3.RightColumnSize
Text13.text = Span3.RightGrid
Text13.FontBold = True
Text47.text = Span3.FrontSlabLv
Text48.text = Span3.BackSlabLv
''''''''''''''''''''''''''
Text14.text = Tetulang3.FirstLTBno
Text15.text = Tetulang3.FirstLTBdia
Text16.text = Tetulang3.FirstLTBcurt
Text17.text = Tetulang3.SecondLTBno
Text18.text = Tetulang3.SecondLTBdia
Text19.text = Tetulang3.SecondLTBcurt
Text20.text = Tetulang3.LinkSpacingLHS
Text21.text = Tetulang3.FirstMBBno
Text22.text = Tetulang3.FirstMBBdia
Text23.text = Tetulang3.SecondMBBcurtS ''''
Text24.text = Tetulang3.SecondMBBno
Text25.text = Tetulang3.SecondMBBdia
Text26.text = Tetulang3.SecondMBBcurtE ''''
Text27.text = Tetulang3.LinkSpacingMID
Text28.text = Tetulang3.FirstRTBno
Text29.text = Tetulang3.FirstRTBdia
Text30.text = Tetulang3.FirstRTBcurt
Text31.text = Tetulang3.SecondRTBno
Text32.text = Tetulang3.SecondRTBdia
Text33.text = Tetulang3.SecondRTBcurt
Text34.text = Tetulang3.LinkSpacingRHS

Text51.text = Tetulang3.LeftConnBarNo   ''''''''''''
Text52.text = Tetulang3.LeftConnBarDia  ''''''''''
Text53.text = Tetulang3.LeftConnBarCurtE '''''''''
Text49.text = Tetulang3.LinkCarrierNo  ''''''''''''link
Text50.text = Tetulang3.LinkCarrierDia   ''''''''''link
Text54.text = Tetulang3.RightConnBarNo ''''''''''''
Text55.text = Tetulang3.RightConnBarDia  ''''''''''
Text56.text = Tetulang3.RightConnBarCurtS '''''''''



End Sub

Private Sub Option4_Click()
Image1.Picture = LoadPicture("C:\autodraf\icon\ukad3.ico")
OpenDataFile
Command5.Visible = False
 

List1.Clear
List1.Visible = False
Dim fnum As Integer
Dim txtFile, Temp As String
fnum = FreeFile
txtFile = "C:\autodraf\rasuk\input_data\SpanFourGET.txt"
Command3.Enabled = False
Command4.Enabled = True
Command4.Left = 1880
Command4.Caption = "Span 4"
Command6.Enabled = False
Command5.Enabled = False
NoOfSpan = 4
i = NoOfSpan

Option1.Enabled = False
Option2.Enabled = False
Option3.Enabled = False
Option4.Enabled = True
Option5.Enabled = False
Option6.Enabled = False
Option7.Enabled = False
Option8.Enabled = False
Option9.Enabled = False
Txt_StatusOne
Text7.Enabled = False
Text8.Enabled = True
Text9.Enabled = True

Transfer_DataToTxt (4)


Span4.Value1 = Text1.text
Span4.Value2 = Text2.text
Span4.Value3 = Text3.text
Span4.Value4 = Text4.text
Span4.Value5 = Text5.text
Span4.Value6 = Text6.text
Span4.Value7 = Text7.text
Span4.Value8 = Text8.text
Span4.Value9 = Text9.text
Span4.Value10 = Text10.text
Span4.Value11 = Text11.text
Span4.Value12 = Text12.text
Span4.Value13 = Text13.text
Span4.Value47 = Text47.text
Span4.Value48 = Text48.text
'''''''''''
Tetulang4.Value14 = Text14.text
Tetulang4.Value15 = Text15.text
Tetulang4.Value16 = Text16.text
Tetulang4.Value17 = Text17.text
Tetulang4.Value18 = Text18.text
Tetulang4.Value19 = Text19.text
Tetulang4.Value20 = Text20.text
Tetulang4.Value21 = Text21.text
Tetulang4.Value22 = Text22.text
Tetulang4.Value23 = Text23.text  ''''
Tetulang4.Value24 = Text24.text
Tetulang4.Value25 = Text25.text
Tetulang4.Value26 = Text26.text ''''
Tetulang4.Value27 = Text27.text
Tetulang4.Value28 = Text28.text
Tetulang4.Value29 = Text29.text
Tetulang4.Value30 = Text30.text
Tetulang4.Value31 = Text31.text
Tetulang4.Value32 = Text32.text
Tetulang4.Value33 = Text33.text
Tetulang4.Value34 = Text34.text

Tetulang4.Value51 = Text51.text   ''''''''''''
Tetulang4.Value52 = Text52.text  ''''''''''
Tetulang4.Value53 = Text53.text  '''''''''
Tetulang4.Value49 = Text49.text   ''''''''''''link
Tetulang4.Value50 = Text50.text    ''''''''''link
Tetulang4.Value54 = Text54.text   ''''''''''''
Tetulang4.Value55 = Text55.text   ''''''''''
Tetulang4.Value56 = Text56.text  '''''''''

Text1.text = Span4.LeftGrid
Text1.FontBold = True
Text2.text = Span4.LeftColumnSize
Text3.text = Span4.LeftXbeamB
Text4.text = Span4.LeftXbeamH
Text5.text = Span4.BeamLength
Text6.text = Span4.BeamBreadth
Text7.text = Span4.BeamHeight
Text8.text = Span4.BeamTopDrop
Text9.text = Span4.BeamSoffitDrop
Text10.text = Span4.RightXbeamb
Text11.text = Span4.RightXbeamh
Text12.text = Span4.RightColumnSize
Text13.text = Span4.RightGrid
Text13.FontBold = True
Text47.text = Span4.FrontSlabLv
Text48.text = Span4.BackSlabLv
''''''''''''''''''''''''''
Text14.text = Tetulang4.FirstLTBno
Text15.text = Tetulang4.FirstLTBdia
Text16.text = Tetulang4.FirstLTBcurt
Text17.text = Tetulang4.SecondLTBno
Text18.text = Tetulang4.SecondLTBdia
Text19.text = Tetulang4.SecondLTBcurt
Text20.text = Tetulang4.LinkSpacingLHS
Text21.text = Tetulang4.FirstMBBno
Text22.text = Tetulang4.FirstMBBdia
Text23.text = Tetulang4.SecondMBBcurtS ''''
Text24.text = Tetulang4.SecondMBBno
Text25.text = Tetulang4.SecondMBBdia
Text26.text = Tetulang4.SecondMBBcurtE ''''
Text27.text = Tetulang4.LinkSpacingMID
Text28.text = Tetulang4.FirstRTBno
Text29.text = Tetulang4.FirstRTBdia
Text30.text = Tetulang4.FirstRTBcurt
Text31.text = Tetulang4.SecondRTBno
Text32.text = Tetulang4.SecondRTBdia
Text33.text = Tetulang4.SecondRTBcurt
Text34.text = Tetulang4.LinkSpacingRHS

Text51.text = Tetulang4.LeftConnBarNo   ''''''''''''
Text52.text = Tetulang4.LeftConnBarDia  ''''''''''
Text53.text = Tetulang4.LeftConnBarCurtE '''''''''
Text49.text = Tetulang4.LinkCarrierNo  ''''''''''''link
Text50.text = Tetulang4.LinkCarrierDia   ''''''''''link
Text54.text = Tetulang4.RightConnBarNo ''''''''''''
Text55.text = Tetulang4.RightConnBarDia  ''''''''''
Text56.text = Tetulang4.RightConnBarCurtS '''''''''



End Sub

Private Sub Option5_Click()
Image1.Picture = LoadPicture("C:\autodraf\icon\cskp.ico")
OpenDataFile
Command5.Visible = False
 

List1.Clear
List1.Visible = False
Dim fnum As Integer
Dim txtFile, Temp As String
fnum = FreeFile
txtFile = "C:\autodraf\rasuk\input_data\SpanFiveGET.txt"
Command3.Enabled = False
Command4.Enabled = True
Command4.Left = 2480
Command4.Caption = "Span 5"
Command6.Enabled = False
Command5.Enabled = False
NoOfSpan = 5
i = NoOfSpan


Option1.Enabled = False
Option2.Enabled = False
Option3.Enabled = False
Option4.Enabled = False
Option5.Enabled = True
Option6.Enabled = False
Option7.Enabled = False
Option8.Enabled = False
Option9.Enabled = False
Txt_StatusOne
Text7.Enabled = False
Text8.Enabled = True
Text9.Enabled = True

Transfer_DataToTxt (5)



Span5.Value1 = Text1.text
Span5.Value2 = Text2.text
Span5.Value3 = Text3.text
Span5.Value4 = Text4.text
Span5.Value5 = Text5.text
Span5.Value6 = Text6.text
Span5.Value7 = Text7.text
Span5.Value8 = Text8.text
Span5.Value9 = Text9.text
Span5.Value10 = Text10.text
Span5.Value11 = Text11.text
Span5.Value12 = Text12.text
Span5.Value13 = Text13.text
Span5.Value47 = Text47.text
Span5.Value48 = Text48.text
'''''''''''
Tetulang5.Value14 = Text14.text
Tetulang5.Value15 = Text15.text
Tetulang5.Value16 = Text16.text
Tetulang5.Value17 = Text17.text
Tetulang5.Value18 = Text18.text
Tetulang5.Value19 = Text19.text
Tetulang5.Value20 = Text20.text
Tetulang5.Value21 = Text21.text
Tetulang5.Value22 = Text22.text
Tetulang5.Value23 = Text23.text  ''''
Tetulang5.Value24 = Text24.text
Tetulang5.Value25 = Text25.text
Tetulang5.Value26 = Text26.text ''''
Tetulang5.Value27 = Text27.text
Tetulang5.Value28 = Text28.text
Tetulang5.Value29 = Text29.text
Tetulang5.Value30 = Text30.text
Tetulang5.Value31 = Text31.text
Tetulang5.Value32 = Text32.text
Tetulang5.Value33 = Text33.text
Tetulang5.Value34 = Text34.text

Tetulang5.Value51 = Text51.text   ''''''''''''
Tetulang5.Value52 = Text52.text  ''''''''''
Tetulang5.Value53 = Text53.text  '''''''''
Tetulang5.Value49 = Text49.text   ''''''''''''link
Tetulang5.Value50 = Text50.text    ''''''''''link
Tetulang5.Value54 = Text54.text   ''''''''''''
Tetulang5.Value55 = Text55.text   ''''''''''
Tetulang5.Value56 = Text56.text  '''''''''

Text1.text = Span5.LeftGrid
Text1.FontBold = True
Text2.text = Span5.LeftColumnSize
Text3.text = Span5.LeftXbeamB
Text4.text = Span5.LeftXbeamH
Text5.text = Span5.BeamLength
Text6.text = Span5.BeamBreadth
Text7.text = Span5.BeamHeight
Text8.text = Span5.BeamTopDrop
Text9.text = Span5.BeamSoffitDrop
Text10.text = Span5.RightXbeamb
Text11.text = Span5.RightXbeamh
Text12.text = Span5.RightColumnSize
Text13.text = Span5.RightGrid
Text13.FontBold = True
Text47.text = Span5.FrontSlabLv
Text48.text = Span5.BackSlabLv
''''''''''''''''''''''''''
Text14.text = Tetulang5.FirstLTBno
Text15.text = Tetulang5.FirstLTBdia
Text16.text = Tetulang5.FirstLTBcurt
Text17.text = Tetulang5.SecondLTBno
Text18.text = Tetulang5.SecondLTBdia
Text19.text = Tetulang5.SecondLTBcurt
Text20.text = Tetulang5.LinkSpacingLHS
Text21.text = Tetulang5.FirstMBBno
Text22.text = Tetulang5.FirstMBBdia
Text23.text = Tetulang5.SecondMBBcurtS ''''
Text24.text = Tetulang5.SecondMBBno
Text25.text = Tetulang5.SecondMBBdia
Text26.text = Tetulang5.SecondMBBcurtE ''''
Text27.text = Tetulang5.LinkSpacingMID
Text28.text = Tetulang5.FirstRTBno
Text29.text = Tetulang5.FirstRTBdia
Text30.text = Tetulang5.FirstRTBcurt
Text31.text = Tetulang5.SecondRTBno
Text32.text = Tetulang5.SecondRTBdia
Text33.text = Tetulang5.SecondRTBcurt
Text34.text = Tetulang5.LinkSpacingRHS

Text51.text = Tetulang5.LeftConnBarNo   ''''''''''''
Text52.text = Tetulang5.LeftConnBarDia  ''''''''''
Text53.text = Tetulang5.LeftConnBarCurtE '''''''''
Text49.text = Tetulang5.LinkCarrierNo  ''''''''''''link
Text50.text = Tetulang5.LinkCarrierDia   ''''''''''link
Text54.text = Tetulang5.RightConnBarNo ''''''''''''
Text55.text = Tetulang5.RightConnBarDia  ''''''''''
Text56.text = Tetulang5.RightConnBarCurtS '''''''''




End Sub

Private Sub Option6_Click()
Image1.Picture = LoadPicture("C:\autodraf\icon\ukad.ico")
OpenDataFile
Command5.Visible = False
 

List1.Clear
List1.Visible = False
Dim fnum As Integer
Dim txtFile, Temp As String
fnum = FreeFile
txtFile = "C:\autodraf\rasuk\input_data\SpanSixGET.txt"
Command3.Enabled = False
Command4.Enabled = True
Command4.Left = 3080
Command4.Caption = "Span 6"
Command6.Enabled = False
Command5.Enabled = False
NoOfSpan = 6
i = NoOfSpan

Option1.Enabled = False
Option2.Enabled = False
Option3.Enabled = False
Option4.Enabled = False
Option5.Enabled = False
Option6.Enabled = True
Option7.Enabled = False
Option8.Enabled = False
Option9.Enabled = False
Txt_StatusOne
Text7.Enabled = False
Text8.Enabled = True
Text9.Enabled = True

Transfer_DataToTxt (6)


Span6.Value1 = Text1.text
Span6.Value2 = Text2.text
Span6.Value3 = Text3.text
Span6.Value4 = Text4.text
Span6.Value5 = Text5.text
Span6.Value6 = Text6.text
Span6.Value7 = Text7.text
Span6.Value8 = Text8.text
Span6.Value9 = Text9.text
Span6.Value10 = Text10.text
Span6.Value11 = Text11.text
Span6.Value12 = Text12.text
Span6.Value13 = Text13.text
Span6.Value47 = Text47.text
Span6.Value48 = Text48.text
'''''''''''
Tetulang6.Value14 = Text14.text
Tetulang6.Value15 = Text15.text
Tetulang6.Value16 = Text16.text
Tetulang6.Value17 = Text17.text
Tetulang6.Value18 = Text18.text
Tetulang6.Value19 = Text19.text
Tetulang6.Value20 = Text20.text
Tetulang6.Value21 = Text21.text
Tetulang6.Value22 = Text22.text
Tetulang6.Value23 = Text23.text  ''''
Tetulang6.Value24 = Text24.text
Tetulang6.Value25 = Text25.text
Tetulang6.Value26 = Text26.text ''''
Tetulang6.Value27 = Text27.text
Tetulang6.Value28 = Text28.text
Tetulang6.Value29 = Text29.text
Tetulang6.Value30 = Text30.text
Tetulang6.Value31 = Text31.text
Tetulang6.Value32 = Text32.text
Tetulang6.Value33 = Text33.text
Tetulang6.Value34 = Text34.text

Tetulang6.Value51 = Text51.text   ''''''''''''
Tetulang6.Value52 = Text52.text  ''''''''''
Tetulang6.Value53 = Text53.text  '''''''''
Tetulang6.Value49 = Text49.text   ''''''''''''link
Tetulang6.Value50 = Text50.text    ''''''''''link
Tetulang6.Value54 = Text54.text   ''''''''''''
Tetulang6.Value55 = Text55.text   ''''''''''
Tetulang6.Value56 = Text56.text  '''''''''

Text1.text = Span6.LeftGrid
Text1.FontBold = True
Text2.text = Span6.LeftColumnSize
Text3.text = Span6.LeftXbeamB
Text4.text = Span6.LeftXbeamH
Text5.text = Span6.BeamLength
Text6.text = Span6.BeamBreadth
Text7.text = Span6.BeamHeight
Text8.text = Span6.BeamTopDrop
Text9.text = Span6.BeamSoffitDrop
Text10.text = Span6.RightXbeamb
Text11.text = Span6.RightXbeamh
Text12.text = Span6.RightColumnSize
Text13.text = Span6.RightGrid
Text13.FontBold = True
Text47.text = Span6.FrontSlabLv
Text48.text = Span6.BackSlabLv
''''''''''''''''''''''''''
Text14.text = Tetulang6.FirstLTBno
Text15.text = Tetulang6.FirstLTBdia
Text16.text = Tetulang6.FirstLTBcurt
Text17.text = Tetulang6.SecondLTBno
Text18.text = Tetulang6.SecondLTBdia
Text19.text = Tetulang6.SecondLTBcurt
Text20.text = Tetulang6.LinkSpacingLHS
Text21.text = Tetulang6.FirstMBBno
Text22.text = Tetulang6.FirstMBBdia
Text23.text = Tetulang6.SecondMBBcurtS ''''
Text24.text = Tetulang6.SecondMBBno
Text25.text = Tetulang6.SecondMBBdia
Text26.text = Tetulang6.SecondMBBcurtE ''''
Text27.text = Tetulang6.LinkSpacingMID
Text28.text = Tetulang6.FirstRTBno
Text29.text = Tetulang6.FirstRTBdia
Text30.text = Tetulang6.FirstRTBcurt
Text31.text = Tetulang6.SecondRTBno
Text32.text = Tetulang6.SecondRTBdia
Text33.text = Tetulang6.SecondRTBcurt
Text34.text = Tetulang6.LinkSpacingRHS

Text51.text = Tetulang6.LeftConnBarNo   ''''''''''''
Text52.text = Tetulang6.LeftConnBarDia  ''''''''''
Text53.text = Tetulang6.LeftConnBarCurtE '''''''''
Text49.text = Tetulang6.LinkCarrierNo  ''''''''''''link
Text50.text = Tetulang6.LinkCarrierDia   ''''''''''link
Text54.text = Tetulang6.RightConnBarNo ''''''''''''
Text55.text = Tetulang6.RightConnBarDia  ''''''''''
Text56.text = Tetulang6.RightConnBarCurtS '''''''''



End Sub

Private Sub Option7_Click()
Image1.Picture = LoadPicture("C:\autodraf\icon\ukad3.ico")
OpenDataFile
Command5.Visible = False
 

List1.Clear
List1.Visible = False
Dim fnum As Integer
Dim txtFile, Temp As String
fnum = FreeFile
txtFile = "C:\autodraf\rasuk\input_data\SpanSevenGET.txt"
Command3.Enabled = False
Command4.Enabled = True
Command4.Left = 3680
Command4.Caption = "Span 7"
Command6.Enabled = False
Command5.Enabled = False
NoOfSpan = 7
i = NoOfSpan

Option1.Enabled = False
Option2.Enabled = False
Option3.Enabled = False
Option4.Enabled = False
Option5.Enabled = False
Option6.Enabled = False
Option7.Enabled = True
Option8.Enabled = False
Option9.Enabled = False
Txt_StatusOne
Text7.Enabled = False
Text8.Enabled = True
Text9.Enabled = True

Transfer_DataToTxt (7)


Span7.Value1 = Text1.text
Span7.Value2 = Text2.text
Span7.Value3 = Text3.text
Span7.Value4 = Text4.text
Span7.Value5 = Text5.text
Span7.Value6 = Text6.text
Span7.Value7 = Text7.text
Span7.Value8 = Text8.text
Span7.Value9 = Text9.text
Span7.Value10 = Text10.text
Span7.Value11 = Text11.text
Span7.Value12 = Text12.text
Span7.Value13 = Text13.text
Span7.Value47 = Text47.text
Span7.Value48 = Text48.text
'''''''''''
Tetulang7.Value14 = Text14.text
Tetulang7.Value15 = Text15.text
Tetulang7.Value16 = Text16.text
Tetulang7.Value17 = Text17.text
Tetulang7.Value18 = Text18.text
Tetulang7.Value19 = Text19.text
Tetulang7.Value20 = Text20.text
Tetulang7.Value21 = Text21.text
Tetulang7.Value22 = Text22.text
Tetulang7.Value23 = Text23.text  ''''
Tetulang7.Value24 = Text24.text
Tetulang7.Value25 = Text25.text
Tetulang7.Value26 = Text26.text ''''
Tetulang7.Value27 = Text27.text
Tetulang7.Value28 = Text28.text
Tetulang7.Value29 = Text29.text
Tetulang7.Value30 = Text30.text
Tetulang7.Value31 = Text31.text
Tetulang7.Value32 = Text32.text
Tetulang7.Value33 = Text33.text
Tetulang7.Value34 = Text34.text

Tetulang7.Value51 = Text51.text   ''''''''''''
Tetulang7.Value52 = Text52.text  ''''''''''
Tetulang7.Value53 = Text53.text  '''''''''
Tetulang7.Value49 = Text49.text   ''''''''''''link
Tetulang7.Value50 = Text50.text    ''''''''''link
Tetulang7.Value54 = Text54.text   ''''''''''''
Tetulang7.Value55 = Text55.text   ''''''''''
Tetulang7.Value56 = Text56.text  '''''''''

Text1.text = Span7.LeftGrid
Text1.FontBold = True
Text2.text = Span7.LeftColumnSize
Text3.text = Span7.LeftXbeamB
Text4.text = Span7.LeftXbeamH
Text5.text = Span7.BeamLength
Text6.text = Span7.BeamBreadth
Text7.text = Span7.BeamHeight
Text8.text = Span7.BeamTopDrop
Text9.text = Span7.BeamSoffitDrop
Text10.text = Span7.RightXbeamb
Text11.text = Span7.RightXbeamh
Text12.text = Span7.RightColumnSize
Text13.text = Span7.RightGrid
Text13.FontBold = True
Text47.text = Span7.FrontSlabLv
Text48.text = Span7.BackSlabLv
''''''''''''''''''''''''''
Text14.text = Tetulang7.FirstLTBno
Text15.text = Tetulang7.FirstLTBdia
Text16.text = Tetulang7.FirstLTBcurt
Text17.text = Tetulang7.SecondLTBno
Text18.text = Tetulang7.SecondLTBdia
Text19.text = Tetulang7.SecondLTBcurt
Text20.text = Tetulang7.LinkSpacingLHS
Text21.text = Tetulang7.FirstMBBno
Text22.text = Tetulang7.FirstMBBdia
Text23.text = Tetulang7.SecondMBBcurtS ''''
Text24.text = Tetulang7.SecondMBBno
Text25.text = Tetulang7.SecondMBBdia
Text26.text = Tetulang7.SecondMBBcurtE ''''
Text27.text = Tetulang7.LinkSpacingMID
Text28.text = Tetulang7.FirstRTBno
Text29.text = Tetulang7.FirstRTBdia
Text30.text = Tetulang7.FirstRTBcurt
Text31.text = Tetulang7.SecondRTBno
Text32.text = Tetulang7.SecondRTBdia
Text33.text = Tetulang7.SecondRTBcurt
Text34.text = Tetulang7.LinkSpacingRHS

Text51.text = Tetulang7.LeftConnBarNo   ''''''''''''
Text52.text = Tetulang7.LeftConnBarDia  ''''''''''
Text53.text = Tetulang7.LeftConnBarCurtE '''''''''
Text49.text = Tetulang7.LinkCarrierNo  ''''''''''''link
Text50.text = Tetulang7.LinkCarrierDia   ''''''''''link
Text54.text = Tetulang7.RightConnBarNo ''''''''''''
Text55.text = Tetulang7.RightConnBarDia  ''''''''''
Text56.text = Tetulang7.RightConnBarCurtS '''''''''


End Sub

Private Sub Option8_Click()
Image1.Picture = LoadPicture("C:\autodraf\icon\cskp.ico")
OpenDataFile
Command5.Visible = False
 

List1.Clear
List1.Visible = False
Dim fnum As Integer
Dim txtFile, Temp As String
fnum = FreeFile
txtFile = "C:\autodraf\rasuk\input_data\SpanEightGET.txt"
Command3.Enabled = False
Command4.Enabled = True
Command4.Left = 4280
Command4.Caption = "Span 8"
Command6.Enabled = False
Command5.Enabled = False
NoOfSpan = 8
i = NoOfSpan

Option1.Enabled = False
Option2.Enabled = False
Option3.Enabled = False
Option4.Enabled = False
Option5.Enabled = False
Option6.Enabled = False
Option7.Enabled = False
Option8.Enabled = True
Option9.Enabled = False
Txt_StatusOne
Text7.Enabled = False
Text8.Enabled = True
Text9.Enabled = True

Transfer_DataToTxt (8)


Span8.Value1 = Text1.text
Span8.Value2 = Text2.text
Span8.Value3 = Text3.text
Span8.Value4 = Text4.text
Span8.Value5 = Text5.text
Span8.Value6 = Text6.text
Span8.Value7 = Text7.text
Span8.Value8 = Text8.text
Span8.Value9 = Text9.text
Span8.Value10 = Text10.text
Span8.Value11 = Text11.text
Span8.Value12 = Text12.text
Span8.Value13 = Text13.text
Span8.Value47 = Text47.text
Span8.Value48 = Text48.text
'''''''''''
Tetulang8.Value14 = Text14.text
Tetulang8.Value15 = Text15.text
Tetulang8.Value16 = Text16.text
Tetulang8.Value17 = Text17.text
Tetulang8.Value18 = Text18.text
Tetulang8.Value19 = Text19.text
Tetulang8.Value20 = Text20.text
Tetulang8.Value21 = Text21.text
Tetulang8.Value22 = Text22.text
Tetulang8.Value23 = Text23.text  ''''
Tetulang8.Value24 = Text24.text
Tetulang8.Value25 = Text25.text
Tetulang8.Value26 = Text26.text ''''
Tetulang8.Value27 = Text27.text
Tetulang8.Value28 = Text28.text
Tetulang8.Value29 = Text29.text
Tetulang8.Value30 = Text30.text
Tetulang8.Value31 = Text31.text
Tetulang8.Value32 = Text32.text
Tetulang8.Value33 = Text33.text
Tetulang8.Value34 = Text34.text

Tetulang8.Value51 = Text51.text   ''''''''''''
Tetulang8.Value52 = Text52.text  ''''''''''
Tetulang8.Value53 = Text53.text  '''''''''
Tetulang8.Value49 = Text49.text   ''''''''''''link
Tetulang8.Value50 = Text50.text    ''''''''''link
Tetulang8.Value54 = Text54.text   ''''''''''''
Tetulang8.Value55 = Text55.text   ''''''''''
Tetulang8.Value56 = Text56.text  '''''''''

Text1.text = Span8.LeftGrid
Text1.FontBold = True
Text2.text = Span8.LeftColumnSize
Text3.text = Span8.LeftXbeamB
Text4.text = Span8.LeftXbeamH
Text5.text = Span8.BeamLength
Text6.text = Span8.BeamBreadth
Text7.text = Span8.BeamHeight
Text8.text = Span8.BeamTopDrop
Text9.text = Span8.BeamSoffitDrop
Text10.text = Span8.RightXbeamb
Text11.text = Span8.RightXbeamh
Text12.text = Span8.RightColumnSize
Text13.text = Span8.RightGrid
Text13.FontBold = True
Text47.text = Span8.FrontSlabLv
Text48.text = Span8.BackSlabLv
''''''''''''''''''''''''''
Text14.text = Tetulang8.FirstLTBno
Text15.text = Tetulang8.FirstLTBdia
Text16.text = Tetulang8.FirstLTBcurt
Text17.text = Tetulang8.SecondLTBno
Text18.text = Tetulang8.SecondLTBdia
Text19.text = Tetulang8.SecondLTBcurt
Text20.text = Tetulang8.LinkSpacingLHS
Text21.text = Tetulang8.FirstMBBno
Text22.text = Tetulang8.FirstMBBdia
Text23.text = Tetulang8.SecondMBBcurtS ''''
Text24.text = Tetulang8.SecondMBBno
Text25.text = Tetulang8.SecondMBBdia
Text26.text = Tetulang8.SecondMBBcurtE ''''
Text27.text = Tetulang8.LinkSpacingMID
Text28.text = Tetulang8.FirstRTBno
Text29.text = Tetulang8.FirstRTBdia
Text30.text = Tetulang8.FirstRTBcurt
Text31.text = Tetulang8.SecondRTBno
Text32.text = Tetulang8.SecondRTBdia
Text33.text = Tetulang8.SecondRTBcurt
Text34.text = Tetulang8.LinkSpacingRHS

Text51.text = Tetulang8.LeftConnBarNo   ''''''''''''
Text52.text = Tetulang8.LeftConnBarDia  ''''''''''
Text53.text = Tetulang8.LeftConnBarCurtE '''''''''
Text49.text = Tetulang8.LinkCarrierNo  ''''''''''''link
Text50.text = Tetulang8.LinkCarrierDia   ''''''''''link
Text54.text = Tetulang8.RightConnBarNo ''''''''''''
Text55.text = Tetulang8.RightConnBarDia  ''''''''''
Text56.text = Tetulang8.RightConnBarCurtS '''''''''


End Sub

Private Sub Option9_Click()
Image1.Picture = LoadPicture("C:\autodraf\icon\ukad.ico")
OpenDataFile
Command5.Visible = False
 

List1.Clear
List1.Visible = False
Dim fnum As Integer
Dim txtFile, Temp As String
fnum = FreeFile
txtFile = "C:\autodraf\rasuk\input_data\SpanNineGET.txt"
Command3.Enabled = False
Command4.Enabled = True
Command4.Left = 4880
Command4.Caption = "Span 9"
Command6.Enabled = False
Command5.Enabled = False
NoOfSpan = 9
i = NoOfSpan

Option1.Enabled = False
Option2.Enabled = False
Option3.Enabled = False
Option4.Enabled = False
Option5.Enabled = False
Option6.Enabled = False
Option7.Enabled = False
Option8.Enabled = False
Option9.Enabled = True
Txt_StatusOne
Text7.Enabled = False
Text8.Enabled = True
Text9.Enabled = True

Transfer_DataToTxt (9)


Span9.Value1 = Text1.text
Span9.Value2 = Text2.text
Span9.Value3 = Text3.text
Span9.Value4 = Text4.text
Span9.Value5 = Text5.text
Span9.Value6 = Text6.text
Span9.Value7 = Text7.text
Span9.Value8 = Text8.text
Span9.Value9 = Text9.text
Span9.Value10 = Text10.text
Span9.Value11 = Text11.text
Span9.Value12 = Text12.text
Span9.Value13 = Text13.text
Span9.Value47 = Text47.text
Span9.Value48 = Text48.text
'''''''''''
Tetulang9.Value14 = Text14.text
Tetulang9.Value15 = Text15.text
Tetulang9.Value16 = Text16.text
Tetulang9.Value17 = Text17.text
Tetulang9.Value18 = Text18.text
Tetulang9.Value19 = Text19.text
Tetulang9.Value20 = Text20.text
Tetulang9.Value21 = Text21.text
Tetulang9.Value22 = Text22.text
Tetulang9.Value23 = Text23.text  ''''
Tetulang9.Value24 = Text24.text
Tetulang9.Value25 = Text25.text
Tetulang9.Value26 = Text26.text ''''
Tetulang9.Value27 = Text27.text
Tetulang9.Value28 = Text28.text
Tetulang9.Value29 = Text29.text
Tetulang9.Value30 = Text30.text
Tetulang9.Value31 = Text31.text
Tetulang9.Value32 = Text32.text
Tetulang9.Value33 = Text33.text
Tetulang9.Value34 = Text34.text

Tetulang9.Value51 = Text51.text   ''''''''''''
Tetulang9.Value52 = Text52.text  ''''''''''
Tetulang9.Value53 = Text53.text  '''''''''
Tetulang9.Value49 = Text49.text   ''''''''''''link
Tetulang9.Value50 = Text50.text    ''''''''''link
Tetulang9.Value54 = Text54.text   ''''''''''''
Tetulang9.Value55 = Text55.text   ''''''''''
Tetulang9.Value56 = Text56.text  '''''''''

Text1.text = Span9.LeftGrid
Text1.FontBold = True
Text2.text = Span9.LeftColumnSize
Text3.text = Span9.LeftXbeamB
Text4.text = Span9.LeftXbeamH
Text5.text = Span9.BeamLength
Text6.text = Span9.BeamBreadth
Text7.text = Span9.BeamHeight
Text8.text = Span9.BeamTopDrop
Text9.text = Span9.BeamSoffitDrop
Text10.text = Span9.RightXbeamb
Text11.text = Span9.RightXbeamh
Text12.text = Span9.RightColumnSize
Text13.text = Span9.RightGrid
Text13.FontBold = True
Text47.text = Span9.FrontSlabLv
Text48.text = Span9.BackSlabLv
''''''''''''''''''''''''''
Text14.text = Tetulang9.FirstLTBno
Text15.text = Tetulang9.FirstLTBdia
Text16.text = Tetulang9.FirstLTBcurt
Text17.text = Tetulang9.SecondLTBno
Text18.text = Tetulang9.SecondLTBdia
Text19.text = Tetulang9.SecondLTBcurt
Text20.text = Tetulang9.LinkSpacingLHS
Text21.text = Tetulang9.FirstMBBno
Text22.text = Tetulang9.FirstMBBdia
Text23.text = Tetulang9.SecondMBBcurtS ''''
Text24.text = Tetulang9.SecondMBBno
Text25.text = Tetulang9.SecondMBBdia
Text26.text = Tetulang9.SecondMBBcurtE ''''
Text27.text = Tetulang9.LinkSpacingMID
Text28.text = Tetulang9.FirstRTBno
Text29.text = Tetulang9.FirstRTBdia
Text30.text = Tetulang9.FirstRTBcurt
Text31.text = Tetulang9.SecondRTBno
Text32.text = Tetulang9.SecondRTBdia
Text33.text = Tetulang9.SecondRTBcurt
Text34.text = Tetulang9.LinkSpacingRHS

Text51.text = Tetulang9.LeftConnBarNo   ''''''''''''
Text52.text = Tetulang9.LeftConnBarDia  ''''''''''
Text53.text = Tetulang9.LeftConnBarCurtE '''''''''
Text49.text = Tetulang9.LinkCarrierNo  ''''''''''''link
Text50.text = Tetulang9.LinkCarrierDia   ''''''''''link
Text54.text = Tetulang9.RightConnBarNo ''''''''''''
Text55.text = Tetulang9.RightConnBarDia  ''''''''''
Text56.text = Tetulang9.RightConnBarCurtS '''''''''


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
  Txt = SetStandardColumnDepth(Left(Text10.text, 1))
    Text10.text = Txt
      End If
      Text10.text = Val((Text10.text))
If Val(Text10.text) = 0 Then
   Text10.text = 0
   End If
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
  Txt = SetStandardBeamDepth(Left(Text11.text, 1))
    Text11.text = Txt
      End If
      Text4.text = Val((Text4.text))
If Val(Text11.text) = 0 Then
   Text11.text = 0
   End If
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
  Txt = SetStandardColumnDepth(Left(Text12.text, 1))
    Text12.text = Txt
      End If
      Text12.text = Val((Text12.text))
If Val(Text12.text) = 0 Then
   Text12.text = 0
   End If
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
  Txt = SetStandardBarNumber(Left(Text14.text, 1))
    Text14.text = Txt
      End If
      Text14.text = Val((Text14.text))
If Val(Left(Text14.text, 2)) = 0 Then
   Text14.text = 2
   End If
End Sub

Private Sub Text14_GotFocus()
Line6.BorderWidth = 2

End Sub

Private Sub Text14_LostFocus()
Line6.BorderWidth = 1

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
  Txt = SetStandardBarSize(Left(Text15.text, 1))
    Text15.text = Txt
      End If
       
      Text15.text = Val((Text15.text))
If Val(Left(Text15.text, 2)) <= 6 Then
   Text15.text = 6
      End If
If Val(Left(Text15.text, 2)) >= 40 Then
   Text15.text = 40
       End If
   
   Text15.text = Val(Left(Text15.text, 2))
   Text16.text = Val(Text15.text) * 50
     If Val(Text16.text) > Val(Text5.text) / 4 Then
        Text16.text = Int(Val(Text5.text) / 4)
         End If
   
End Sub

Private Sub Text15_GotFocus()
Line6.BorderWidth = 2
End Sub

Private Sub Text15_LostFocus()
Line6.BorderWidth = 1
End Sub

Private Sub Text16_Change()
  
If Val(Text16.text) = 0 Then
   Text16.text = Int(Val(Text5.text) / 4)
   End If
End Sub

Private Sub Text16_Click()
Text16.text = Val(Text15.text) * 50
End Sub

Private Sub Text16_DblClick()
Text16.text = Val(Text5.text) / 4
End Sub

Private Sub Text16_GotFocus()
Line6.BorderWidth = 2
End Sub

Private Sub Text16_LostFocus()
Line6.BorderWidth = 1
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
  Txt = SetStandardBarNumber(Left(Text17.text, 1))
    Text17.text = Txt
      End If
      Text17.text = Val((Text17.text))
If Val(Left(Text17.text, 2)) = 0 Then
   Text17.text = 0
   Text18.text = 0
   Text19.text = 0
   End If
End Sub

Private Sub Text17_GotFocus()
Line7.BorderWidth = 2
End Sub

Private Sub Text17_LostFocus()
Line7.BorderWidth = 1
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
  Txt = SetStandardBarSize(Left(Text18.text, 1))
    Text18.text = Txt
      End If
      Text18.text = Val((Text18.text))
If Val(Text18.text) = 0 Then
   Text18.text = 0
   End If
If Val(Text18.text) >= 40 Then
   Text18.text = 40
   End If
   
  If Val(Text17.text) > 0 And Val(Text18.text) = 0 Then
   Text18.text = 6
   End If
  If Val(Text18.text) > 0 Then
   Text19.text = Val(Text18.text) * 50
      End If
   If Val(Text19.text) > Val(Text5.text) / 5 Then
       Text19.text = Int(Val(Text5.text) / 5)
         End If
End Sub

Private Sub Text18_GotFocus()
Line7.BorderWidth = 2
End Sub

Private Sub Text18_LostFocus()
Line7.BorderWidth = 1
End Sub

Private Sub Text19_Change()

If Val(Text19.text) = 0 Then
   Text19.text = 0
   End If
  If Val(Text19.text) = 0 And Val(Text17.text) > 0 _
     And Val(Text18.text) > 0 Then
       Text19.text = Int(Val(Text5.text) / 5)
          End If
   
   
End Sub

Private Sub Text19_Click()
Text19.text = Val(Text18.text) * 50
End Sub

Private Sub Text19_DblClick()
Text19.text = Val(Text5.text) / 5
End Sub

Private Sub Text19_GotFocus()
Line7.BorderWidth = 2
End Sub

Private Sub Text19_LostFocus()
Line7.BorderWidth = 1
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
If Val(Text2.text) = 0 Then
   Text2.text = 0
   End If

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
Left(Text20.text, 1) = "L" Or Left(Text20.text, 1) = "l" Or _
Left(Text20.text, 1) = "P" Or Left(Text20.text, 1) = "p" Or _
Left(Text20.text, 1) = "O" Or Left(Text20.text, 1) = "o" Then
  Txt = SetStandardLinkSpacing(Left(Text20.text, 1))
    Text20.text = Txt
      End If
      Text20.text = Val((Text20.text))
If Val(Text20.text) = 0 Then
   Text20.text = 100
   End If
End Sub

Private Sub Text20_GotFocus()
Line12.BorderWidth = 2
Line13.BorderWidth = 2
End Sub

Private Sub Text20_LostFocus()
Line12.BorderWidth = 1
Line13.BorderWidth = 1
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
  Txt = SetStandardBarNumber(Left(Text21.text, 1))
    Text21.text = Txt
      End If
      Text21.text = Val((Text21.text))
If Val(Text21.text) = 0 Then
   Text21.text = 2
   End If
   
End Sub

Private Sub Text21_GotFocus()
Line10.BorderWidth = 2
End Sub

Private Sub Text21_LostFocus()
Line10.BorderWidth = 1
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
  Txt = SetStandardBarSize(Left(Text22.text, 1))
    Text22.text = Txt
      End If
      Text22.text = Val((Text22.text))
If Val(Text22.text) <= 6 Then
   Text22.text = 6
   End If
   If Val(Text22.text) >= 40 Then
   Text22.text = 40
   End If
End Sub

Private Sub Text22_GotFocus()
Line10.BorderWidth = 2
End Sub

Private Sub Text22_LostFocus()
Line10.BorderWidth = 1
End Sub

Private Sub Text23_Change()

If Val(Text23.text) = 0 Then
   Text23.text = 0
   End If
End Sub

Private Sub Text23_Click()
Text23.text = Val(Text25.text) * 35
End Sub

Private Sub Text23_DblClick()
Text23.text = Val(Text5.text) / 5
End Sub

Private Sub Text23_GotFocus()
Line11.BorderWidth = 2
End Sub

Private Sub Text23_LostFocus()
Line11.BorderWidth = 1
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
  Txt = SetStandardBarNumber(Left(Text24.text, 1))
    Text24.text = Txt
      End If
      Text24.text = Val((Text24.text))
If Val(Text24.text) = 0 Then
   Text24.text = 0
   Text25.text = 0
   Text23.text = 0
   Text26.text = 0
   End If
End Sub

Private Sub Text24_GotFocus()
Line11.BorderWidth = 2
End Sub

Private Sub Text24_LostFocus()
Line11.BorderWidth = 1
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
  Txt = SetStandardBarSize(Left(Text25.text, 1))
    Text25.text = Txt
      End If
      Text25.text = Val((Text25.text))
If Val(Text25.text) = 0 Then
   Text25.text = 0
   End If
   If Val(Text25.text) >= 40 Then
   Text25.text = 40
   End If
   
 If Val(Text24.text) > 0 And Val(Text25.text) = 0 Then
   Text25.text = 6
     End If
  If Val(Text25.text) > 0 Then
    Text23.text = Val(Text25.text) * 35
    Text26.text = Val(Text25.text) * 35
      End If
   If Val(Text23.text) > Val(Text5.text) / 5 Then
      Text23.text = Val(Text5.text) / 5
         End If
   If Val(Text26.text) > Val(Text5.text) / 5 Then
      Text26.text = Val(Text5.text) / 5
         End If
End Sub

Private Sub Text25_GotFocus()
Line11.BorderWidth = 2
End Sub

Private Sub Text25_LostFocus()
Line11.BorderWidth = 1
End Sub

Private Sub Text26_Change()

If Val(Text26.text) = 0 Then
   Text26.text = 0
   End If
End Sub

Private Sub Text26_Click()
Text26.text = Val(Text25.text) * 35
End Sub

Private Sub Text26_DblClick()
Text26.text = Val(Text5.text) / 5
End Sub

Private Sub Text26_GotFocus()
Line11.BorderWidth = 2
End Sub

Private Sub Text26_LostFocus()
Line11.BorderWidth = 1
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
Left(Text27.text, 1) = "L" Or Left(Text27.text, 1) = "l" Or _
Left(Text27.text, 1) = "P" Or Left(Text27.text, 1) = "p" Or _
Left(Text27.text, 1) = "O" Or Left(Text27.text, 1) = "o" Then
  Txt = SetStandardLinkSpacing(Left(Text27.text, 1))
    Text27.text = Txt
      End If
      Text27.text = Val((Text27.text))
If Val(Text27.text) = 0 Then
   Text27.text = 100
   End If
End Sub

Private Sub Text27_GotFocus()
Line14.BorderWidth = 2
Line15.BorderWidth = 2
End Sub

Private Sub Text27_LostFocus()
Line14.BorderWidth = 1
Line15.BorderWidth = 1
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
  Txt = SetStandardBarNumber(Left(Text28.text, 1))
    Text28.text = Txt
      End If
      Text28.text = Val((Text28.text))
If Val(Text28.text) = 0 Then
   Text28.text = 2
   End If
End Sub

Private Sub Text28_GotFocus()
Line8.BorderWidth = 2
End Sub

Private Sub Text28_LostFocus()
Line8.BorderWidth = 1
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
  Txt = SetStandardBarSize(Left(Text29.text, 1))
    Text29.text = Txt
      End If
      Text29.text = Val((Text29.text))
If Val(Left(Text29.text, 2)) <= 6 Then
   Text29.text = 6
   End If
If Val(Left(Text29.text, 2)) >= 40 Then
   Text29.text = 40
   End If
   
   Text29.text = Val(Left(Text29.text, 2))
   Text30.text = Val(Text29.text) * 50
     If Val(Text30.text) > Val(Text5.text) / 4 Then
        Text30.text = Int(Val(Text5.text) / 4)
         End If
   
End Sub


Private Sub Text29_GotFocus()
Line8.BorderWidth = 2
End Sub

Private Sub Text29_LostFocus()
Line8.BorderWidth = 1
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
  Txt = SetStandardColumnDepth(Left(Text3.text, 1))
    Text3.text = Txt
      End If
      Text3.text = Val((Text3.text))
If Val(Text3.text) = 0 Then
   Text3.text = 0
   End If
End Sub

Private Sub Text30_Change()

If Val(Text30.text) = 0 Then
   Text30.text = Int(Val(Text5.text) / 4)
   End If
End Sub

Private Sub Text30_Click()
Text30.text = Val(Text29.text) * 50
End Sub

Private Sub Text30_DblClick()
Text30.text = Val(Text5.text) / 4
End Sub

Private Sub Text30_GotFocus()
Line8.BorderWidth = 2
End Sub

Private Sub Text30_LostFocus()
Line8.BorderWidth = 1
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
  Txt = SetStandardBarNumber(Left(Text31.text, 1))
    Text31.text = Txt
      End If
      Text31.text = Val((Text31.text))
If Val(Text31.text) = 0 Then
   Text31.text = 0
   Text32.text = 0
   Text33.text = 0
   End If
End Sub

Private Sub Text31_GotFocus()
Line9.BorderWidth = 2
End Sub

Private Sub Text31_LostFocus()
Line9.BorderWidth = 1
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
  Txt = SetStandardBarSize(Left(Text32.text, 1))
    Text32.text = Txt
      End If
      Text32.text = Val((Text32.text))
If Val(Text32.text) = 0 Then
   Text32.text = 0
   End If
If Val(Text32.text) >= 40 Then
   Text32.text = 40
   End If
   
  If Val(Text31.text) > 0 And Val(Text32.text) = 0 Then
   Text32.text = 6
     End If
  If Val(Text32.text) > 0 Then
    Text33.text = Val(Text32.text) * 50
      End If
   If Val(Text33.text) > Val(Text5.text) / 5 Then
      Text33.text = Int(Val(Text5.text) / 5)
         End If
End Sub

Private Sub Text32_GotFocus()
Line9.BorderWidth = 2
End Sub

Private Sub Text32_LostFocus()
Line9.BorderWidth = 1
End Sub

Private Sub Text33_Change()

If Val(Text33.text) = 0 Then
   Text33.text = 0
   End If
   If Val(Text33.text) = 0 And Val(Text31.text) > 0 _
     And Val(Text32.text) > 0 Then
       Text33.text = Int(Val(Text5.text) / 5)
          End If
End Sub

Private Sub Text33_Click()
Text33.text = Val(Text32.text) * 50
End Sub

Private Sub Text33_DblClick()
Text33.text = Val(Text5.text) / 5
End Sub

Private Sub Text33_GotFocus()
Line9.BorderWidth = 2
End Sub

Private Sub Text33_LostFocus()
Line9.BorderWidth = 1
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
Left(Text34.text, 1) = "L" Or Left(Text34.text, 1) = "l" Or _
Left(Text34.text, 1) = "P" Or Left(Text34.text, 1) = "p" Or _
Left(Text34.text, 1) = "O" Or Left(Text34.text, 1) = "o" Then
  Txt = SetStandardLinkSpacing(Left(Text34.text, 1))
    Text34.text = Txt
      End If
      Text34.text = Val((Text34.text))
If Val(Text34.text) = 0 Then
   Text34.text = 100
   End If
End Sub

Private Sub Text34_GotFocus()
Line16.BorderWidth = 2
Line17.BorderWidth = 2
End Sub

Private Sub Text34_LostFocus()
Line16.BorderWidth = 1
Line17.BorderWidth = 1
End Sub

Private Sub Text35_Change()
 
If Val(Text35.text) = 0 Then
   Text35.text = 0
   End If
End Sub

Private Sub Text36_Change()
 
If Val(Text36.text) = 0 Then
   Text36.text = 0
   End If
End Sub

Private Sub Text37_Change()
''nama rasuk
 
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
  Txt = SetStandardConcreteFcu(Left(Text38.text, 1))
    Text38.text = Txt
      End If
      Text38.text = Val((Text38.text))
If Val(Text38.text) = 0 Then
   Text38.text = 30
   End If
End Sub

Private Sub Text39_Change()
 
If Val(Text39.text) = 0 Then
   Text39.text = 460
   End If
End Sub

Private Sub Text4_Change()
Dim Txt As String
 
If Left(Text4.text, 1) = "A" Or Left(Text4.text, 1) = "a" Or _
Left(Text4.text, 1) = "S" Or Left(Text4.text, 1) = "s" Or _
Left(Text4.text, 1) = "D" Or Left(Text4.text, 1) = "d" Or _
Left(Text4.text, 1) = "F" Or Left(Text4.text, 1) = "f" Or _
Left(Text4.text, 1) = "G" Or Left(Text4.text, 1) = "g" Or _
Left(Text4.text, 1) = "H" Or Left(Text4.text, 1) = "h" Or _
Left(Text4.text, 1) = "J" Or Left(Text4.text, 1) = "j" Or _
Left(Text4.text, 1) = "K" Or Left(Text4.text, 1) = "k" Or _
Left(Text4.text, 1) = "L" Or Left(Text4.text, 1) = "l" Then
  Txt = SetStandardBeamDepth(Left(Text4.text, 1))
    Text4.text = Txt
      End If
      Text4.text = Val((Text4.text))
If Val(Text4.text) = 0 Then
   Text4.text = 0
   End If
End Sub

Private Sub Text40_Change()
 
If Val(Text40.text) = 0 Then
   Text40.text = 250
   End If
End Sub

Private Sub Text41_Change()
Dim Txt As String
 
If Left(Text41.text, 1) = "A" Or Left(Text41.text, 1) = "a" Or _
Left(Text41.text, 1) = "S" Or Left(Text41.text, 1) = "s" Or _
Left(Text41.text, 1) = "D" Or Left(Text41.text, 1) = "d" Or _
Left(Text41.text, 1) = "F" Or Left(Text41.text, 1) = "f" Then
  Txt = SetStandardShrink(Left(Text41.text, 1))
    Text41.text = Txt
      End If
      Text41.text = Val((Text41.text))
If Val(Text41.text) = 0 Then
   Text41.text = 0.0003
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
Left(Text42.text, 1) = "J" Or Left(Text42.text, 1) = "j" Then
  Txt = SetStandardCreep(Left(Text42.text, 1))
    Text42.text = Txt
      End If
      Text42.text = Val((Text42.text))
If Val(Text42.text) = 0 Then
   Text42.text = 2.5
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
Left(Text43.text, 1) = "J" Or Left(Text43.text, 1) = "j" Then
  Txt = SetStandardCover(Left(Text43.text, 1))
    Text43.text = Txt
      End If
      Text43.text = Val((Text43.text))
If Val(Text43.text) = 0 Then
   Text43.text = 20
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
  Txt = SetStandardSlabThk(Left(Text44.text, 1))
    Text44.text = Txt
      End If
      Text44.text = Val((Text44.text))
If Val(Text44.text) = 0 Then
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
  Txt = SetStandardBarSize(Left(Text45.text, 1))
    Text45.text = Txt
      End If
      Text45.text = Val((Text45.text))
If Val(Text45.text) = 0 Then
   Text45.text = 6
   End If
   If Val(Text45.text) >= 40 Then
   Text45.text = 40
   End If
End Sub

Private Sub Text46_Change()
 
If Val(Text46.text) = 0 Then
   Text46.text = 0
   End If
End Sub

Private Sub Text46_GotFocus()
Label3.Visible = True
End Sub

Private Sub Text46_LostFocus()
Label3.Visible = False
End Sub

Private Sub Text47_Change()
 
If Val(Text47.text) = 0 Then
   Text47.text = 0
   End If
End Sub

Private Sub Text48_Change()
 
If Val(Text48.text) = 0 Then
   Text48.text = 0
   End If
End Sub

Private Sub Text49_Change()
Dim Txt As String

If Left(Text49.text, 1) = "A" Or Left(Text49.text, 1) = "a" Or _
Left(Text49.text, 1) = "S" Or Left(Text49.text, 1) = "s" Or _
Left(Text49.text, 1) = "D" Or Left(Text49.text, 1) = "d" Or _
Left(Text49.text, 1) = "F" Or Left(Text49.text, 1) = "f" Or _
Left(Text49.text, 1) = "G" Or Left(Text49.text, 1) = "g" Or _
Left(Text49.text, 1) = "H" Or Left(Text49.text, 1) = "h" Or _
Left(Text49.text, 1) = "J" Or Left(Text49.text, 1) = "j" Or _
Left(Text49.text, 1) = "K" Or Left(Text49.text, 1) = "k" Or _
Left(Text49.text, 1) = "L" Or Left(Text49.text, 1) = "l" Then
  Txt = SetStandardBarNumber(Left(Text49.text, 1))
    Text49.text = Txt
      End If
      Text49.text = Val((Text49.text))
If Val(Text49.text) = 0 Then
   Text49.text = 2
   End If
End Sub

Private Sub Text49_GotFocus()
Line18.BorderWidth = 2
End Sub

Private Sub Text49_LostFocus()
Line18.BorderWidth = 1
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
  Txt = SetStandardBeamLength(Left(Text5.text, 1))
    Text5.text = Txt
      End If
      Text5.text = Val((Text5.text))
If Val(Text5.text) = 0 Then
   Text5.text = 3000
   End If
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
  Txt = SetStandardBarSize(Left(Text50.text, 1))
    Text50.text = Txt
      End If
      Text50.text = Val((Text50.text))
If Val(Text50.text) = 0 Then
   Text50.text = 6
   End If
   If Val(Text50.text) >= 32 Then
   Text50.text = 32
   End If
End Sub

Private Sub Text50_GotFocus()
Line18.BorderWidth = 2
End Sub

Private Sub Text50_LostFocus()
Line18.BorderWidth = 1
End Sub

Private Sub Text51_Change()
Dim Txt As String

If Left(Text51.text, 1) = "A" Or Left(Text51.text, 1) = "a" Or _
Left(Text51.text, 1) = "S" Or Left(Text51.text, 1) = "s" Or _
Left(Text51.text, 1) = "D" Or Left(Text51.text, 1) = "d" Or _
Left(Text51.text, 1) = "F" Or Left(Text51.text, 1) = "f" Or _
Left(Text51.text, 1) = "G" Or Left(Text51.text, 1) = "g" Or _
Left(Text51.text, 1) = "H" Or Left(Text51.text, 1) = "h" Or _
Left(Text51.text, 1) = "J" Or Left(Text51.text, 1) = "j" Or _
Left(Text51.text, 1) = "K" Or Left(Text51.text, 1) = "k" Or _
Left(Text51.text, 1) = "L" Or Left(Text51.text, 1) = "l" Then
  Txt = SetStandardBarNumber(Left(Text51.text, 1))
    Text51.text = Txt
      End If
      Text51.text = Val((Text51.text))
If Val(Left(Text51.text, 2)) = 0 Then
   Text51.text = 2
   End If
End Sub

Private Sub Text51_GotFocus()
Line3.BorderWidth = 2
End Sub

Private Sub Text51_LostFocus()
Line3.BorderWidth = 1
End Sub

Private Sub Text52_Change()
Dim Txt As String

If Left(Text52.text, 1) = "A" Or Left(Text52.text, 1) = "a" Or _
Left(Text52.text, 1) = "S" Or Left(Text52.text, 1) = "s" Or _
Left(Text52.text, 1) = "D" Or Left(Text52.text, 1) = "d" Or _
Left(Text52.text, 1) = "F" Or Left(Text52.text, 1) = "f" Or _
Left(Text52.text, 1) = "G" Or Left(Text52.text, 1) = "g" Or _
Left(Text52.text, 1) = "H" Or Left(Text52.text, 1) = "h" Or _
Left(Text52.text, 1) = "J" Or Left(Text52.text, 1) = "j" Or _
Left(Text52.text, 1) = "K" Or Left(Text52.text, 1) = "k" Or _
Left(Text52.text, 1) = "L" Or Left(Text52.text, 1) = "l" Then
  Txt = SetStandardBarSize(Left(Text52.text, 1))
    Text52.text = Txt
      End If
      Text52.text = Val((Text52.text))
If Val(Text52.text) <= 6 Then
   Text52.text = 6
   End If
If Val(Text52.text) >= 40 Then
   Text52.text = 40
   End If
   
   Text53.text = Val(Text52.text) * 50
     If Val(Text53.text) > Val(Text5.text) / 6 Then
        Text53.text = Int(Val(Text5.text) / 6)
         End If
  
End Sub

Private Sub Text52_GotFocus()
Line3.BorderWidth = 2
End Sub

Private Sub Text52_LostFocus()
Line3.BorderWidth = 1
End Sub

Private Sub Text53_Change()

If Val(Text53.text) = 0 Then
   Text53.text = Int(Val(Text5.text) / 6)
   End If
End Sub

Private Sub Text53_Click()
Text53.text = Val(Text52.text) * 50
End Sub

Private Sub Text53_DblClick()
Text53.text = Val(Text5.text) / 6
End Sub

Private Sub Text53_GotFocus()
Line3.BorderWidth = 2
End Sub

Private Sub Text53_LostFocus()
Line3.BorderWidth = 1
End Sub

Private Sub Text54_Change()
Dim Txt As String

If Left(Text54.text, 1) = "A" Or Left(Text54.text, 1) = "a" Or _
Left(Text54.text, 1) = "S" Or Left(Text54.text, 1) = "s" Or _
Left(Text54.text, 1) = "D" Or Left(Text54.text, 1) = "d" Or _
Left(Text54.text, 1) = "F" Or Left(Text54.text, 1) = "f" Or _
Left(Text54.text, 1) = "G" Or Left(Text54.text, 1) = "g" Or _
Left(Text54.text, 1) = "H" Or Left(Text54.text, 1) = "h" Or _
Left(Text54.text, 1) = "J" Or Left(Text54.text, 1) = "j" Or _
Left(Text54.text, 1) = "K" Or Left(Text54.text, 1) = "k" Or _
Left(Text54.text, 1) = "L" Or Left(Text54.text, 1) = "l" Then
  Txt = SetStandardBarNumber(Left(Text54.text, 1))
    Text54.text = Txt
      End If
      Text54.text = Val((Text54.text))
If Val(Text54.text) = 0 Then
   Text54.text = 2
   End If
End Sub

Private Sub Text54_GotFocus()
Line5.BorderWidth = 2
End Sub

Private Sub Text54_LostFocus()
Line5.BorderWidth = 1
End Sub

Private Sub Text55_Change()
Dim Txt As String

If Left(Text55.text, 1) = "A" Or Left(Text55.text, 1) = "a" Or _
Left(Text55.text, 1) = "S" Or Left(Text55.text, 1) = "s" Or _
Left(Text55.text, 1) = "D" Or Left(Text55.text, 1) = "d" Or _
Left(Text55.text, 1) = "F" Or Left(Text55.text, 1) = "f" Or _
Left(Text55.text, 1) = "G" Or Left(Text55.text, 1) = "g" Or _
Left(Text55.text, 1) = "H" Or Left(Text55.text, 1) = "h" Or _
Left(Text55.text, 1) = "J" Or Left(Text55.text, 1) = "j" Or _
Left(Text55.text, 1) = "K" Or Left(Text55.text, 1) = "k" Or _
Left(Text55.text, 1) = "L" Or Left(Text55.text, 1) = "l" Then
  Txt = SetStandardBarSize(Left(Text55.text, 1))
    Text55.text = Txt
      End If
      Text55.text = Val((Text55.text))
If Val(Text55.text) <= 6 Then
   Text55.text = 6
   End If
 If Val(Text55.text) >= 40 Then
   Text55.text = 40
   End If
   
    Text56.text = Val(Text55.text) * 50
     If Val(Text56.text) > Val(Text5.text) / 6 Then
        Text56.text = Int(Val(Text5.text) / 6)
         End If
   
End Sub

Private Sub Text55_GotFocus()
Line5.BorderWidth = 2
End Sub

Private Sub Text55_LostFocus()
Line5.BorderWidth = 1
End Sub

Private Sub Text56_Change()

If Val(Text56.text) = 0 Then
   Text56.text = Int(Val(Text5.text) / 6)
   End If
End Sub

Private Sub Text56_Click()
Text56.text = Val(Text55.text) * 50
End Sub

Private Sub Text56_DblClick()
Text56.text = Val(Text5.text) / 6
End Sub

Private Sub Text56_GotFocus()
Line5.BorderWidth = 2
End Sub

Private Sub Text56_LostFocus()
Line5.BorderWidth = 1
End Sub

Private Sub Text57_Change()

If Val(Text57.text) <= 0 Then
   Text57.text = 50
   End If
End Sub

Private Sub Text57_GotFocus()
Label3.Visible = True
End Sub

Private Sub Text57_LostFocus()
Label3.Visible = False
End Sub

Private Sub Text6_Change()
Dim Txt As String
Label3.Visible = False
If Left(Text6.text, 1) = "A" Or Left(Text6.text, 1) = "a" Or _
Left(Text6.text, 1) = "S" Or Left(Text6.text, 1) = "s" Or _
Left(Text6.text, 1) = "D" Or Left(Text6.text, 1) = "d" Or _
Left(Text6.text, 1) = "F" Or Left(Text6.text, 1) = "f" Or _
Left(Text6.text, 1) = "G" Or Left(Text6.text, 1) = "g" Or _
Left(Text6.text, 1) = "H" Or Left(Text6.text, 1) = "h" Or _
Left(Text6.text, 1) = "J" Or Left(Text6.text, 1) = "j" Or _
Left(Text6.text, 1) = "K" Or Left(Text6.text, 1) = "k" Or _
Left(Text6.text, 1) = "L" Or Left(Text6.text, 1) = "l" Then
  Txt = SetStandardBeamBreadth(Left(Text6.text, 1))
    Text6.text = Txt
      End If
      Text6.text = Val((Text6.text))
If Val(Text6.text) = 0 Then
   Text6.text = 200
   End If
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
  Txt = SetStandardBeamDepth(Left(Text7.text, 1))
    Text7.text = Txt
      End If
      Text7.text = Val((Text7.text))
If Val(Text7.text) = 0 Then
   Text7.text = 350
   End If
   

''   Dim tmpTxt8, tmpTxt9 As Double
''   tmpTxt8 = Val(Text8.text)
''   tmpTxt9 = Val(Text9.text)
   
  
''   Text8.text = beamH(1) - tmpTxt9 - Val(Text7.text)
        
''   Text9.text = beamH(1) - tmpTxt8 - Val(Text7.text)
       
End Sub

Private Sub Text8_Change()
Dim tmpTxt8, tmpTxt9 As Double
   tmpTxt8 = Val(Text8.text)
   tmpTxt9 = Val(Text9.text)
   
If Command4.Left = 80 Then
  Text8.text = 0
     Text9.text = 0
        Text7.text = beamH(1)
         End If
Text7.text = beamH(1) - Val(Text8.text) - Val(Text9.text)
        
End Sub

Private Sub Text9_Change()
Dim tmpTxt8, tmpTxt9 As Double
   tmpTxt8 = Val(Text8.text)
   tmpTxt9 = Val(Text9.text)
   
If Command4.Left = 80 Then
  Text8.text = 0
     Text9.text = 0
        Text7.text = beamH(1)
         End If
Text7.text = beamH(1) - Val(Text8.text) - Val(Text9.text)
     
End Sub
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

Private Function SetStandardColumnDepth(ByVal SetBar As String)

 If SetBar = "A" Or SetBar = "a" Then
    SetStandardColumnDepth = "50"
    End If
 If SetBar = "S" Or SetBar = "s" Then
    SetStandardColumnDepth = "75"
    End If
 If SetBar = "D" Or SetBar = "d" Then
    SetStandardColumnDepth = "100"
    End If
 If SetBar = "F" Or SetBar = "f" Then
    SetStandardColumnDepth = "125"
    End If
 If SetBar = "G" Or SetBar = "g" Then
    SetStandardColumnDepth = "150"
    End If
 If SetBar = "H" Or SetBar = "h" Then
    SetStandardColumnDepth = "175"
    End If
 If SetBar = "J" Or SetBar = "j" Then
    SetStandardColumnDepth = "200"
    End If
 If SetBar = "K" Or SetBar = "k" Then
    SetStandardColumnDepth = "225"
    End If
If SetBar = "L" Or SetBar = "l" Then
    SetStandardColumnDepth = "250"
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

 
Private Function SetStandardSlabThk(ByVal SetBar As String)

 If SetBar = "A" Or SetBar = "a" Then
    SetStandardSlabThk = "50"
    End If
 If SetBar = "S" Or SetBar = "s" Then
    SetStandardSlabThk = "75"
    End If
 If SetBar = "D" Or SetBar = "d" Then
    SetStandardSlabThk = "100"
    End If
 If SetBar = "F" Or SetBar = "f" Then
    SetStandardSlabThk = "125"
    End If
 If SetBar = "G" Or SetBar = "g" Then
    SetStandardSlabThk = "150"
    End If
 If SetBar = "H" Or SetBar = "h" Then
    SetStandardSlabThk = "175"
    End If
 If SetBar = "J" Or SetBar = "j" Then
    SetStandardSlabThk = "200"
    End If
 If SetBar = "K" Or SetBar = "k" Then
    SetStandardSlabThk = "225"
    End If
If SetBar = "L" Or SetBar = "l" Then
    SetStandardSlabThk = "250"
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
End Function


Private Function SetStandardBeamLength(ByVal SetBar As String)

 If SetBar = "A" Or SetBar = "a" Then
    SetStandardBeamLength = "1000"
    End If
 If SetBar = "S" Or SetBar = "s" Then
    SetStandardBeamLength = "2000"
    End If
 If SetBar = "D" Or SetBar = "d" Then
    SetStandardBeamLength = "3000"
    End If
 If SetBar = "F" Or SetBar = "f" Then
    SetStandardBeamLength = "4000"
    End If
 If SetBar = "G" Or SetBar = "g" Then
    SetStandardBeamLength = "5000"
    End If
 If SetBar = "H" Or SetBar = "h" Then
    SetStandardBeamLength = "6000"
    End If
 If SetBar = "J" Or SetBar = "j" Then
    SetStandardBeamLength = "7000"
    End If
 If SetBar = "K" Or SetBar = "k" Then
    SetStandardBeamLength = "8000"
    End If
If SetBar = "L" Or SetBar = "l" Then
    SetStandardBeamLength = "9000"
    End If

End Function
