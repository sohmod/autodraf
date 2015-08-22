VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Linkage to autocad14...."
   ClientHeight    =   1275
   ClientLeft      =   1065
   ClientTop       =   3345
   ClientWidth     =   9810
   Icon            =   "Linkage1.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   1275
   ScaleWidth      =   9810
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "bye !"
      Height          =   375
      Left            =   9200
      TabIndex        =   3
      Top             =   765
      Width           =   550
   End
   Begin VB.Image Image8 
      Height          =   480
      Left            =   9120
      Picture         =   "Linkage1.frx":030A
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   480
   End
   Begin VB.Image Image7 
      Height          =   480
      Left            =   8040
      Picture         =   "Linkage1.frx":074C
      Top             =   240
      Width           =   480
   End
   Begin VB.Image Image6 
      Height          =   480
      Left            =   3960
      Picture         =   "Linkage1.frx":0B8E
      Top             =   240
      Width           =   480
   End
   Begin VB.Image Image5 
      Height          =   480
      Left            =   7680
      Picture         =   "Linkage1.frx":0FD0
      Top             =   840
      Width           =   480
   End
   Begin VB.Image Image4 
      Height          =   345
      Left            =   8400
      Picture         =   "Linkage1.frx":1412
      Stretch         =   -1  'True
      Top             =   480
      Width           =   600
   End
   Begin VB.Image Image3 
      Height          =   600
      Left            =   -360
      Picture         =   "Linkage1.frx":1854
      Stretch         =   -1  'True
      Top             =   480
      Width           =   720
   End
   Begin VB.Image Image2 
      Height          =   1320
      Left            =   8160
      Picture         =   "Linkage1.frx":1C96
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1200
   End
   Begin VB.Label Label3 
      Caption         =   "TrUcTion"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   615
      Left            =   4440
      TabIndex        =   2
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "CONS"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "UnDer"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   2520
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public NamaFolder As String
Dim DeltaX, DeltaY, Counter As Integer  ' Declare variables.
Private Sub Timer1_Timer()
    
   Image3.Move Image3.Left + DeltaX, Image3.Top - DeltaY / DeltaY * Rnd(-500)
   Image6.Move Image6.Left - DeltaX / 4 * Rnd(50), Image6.Top - DeltaY / DeltaY * Sgn(DeltaY) * Rnd(100)
   Image7.Move Image7.Left - DeltaX / 2 * Rnd(100), Image7.Top - DeltaY / DeltaY * Sgn(DeltaY) * Rnd(50)
   If Sgn(DeltaX) * 1 <= 0 Then
      Image3.Picture = LoadPicture(NamaFolder & "icon\asas_jalur\asj6.ico")
        Else
           Image3.Picture = LoadPicture(NamaFolder & "icon\asas_jalur\asj5.ico")
             End If
             
   If Image3.Left < ScaleLeft Then DeltaX = 100
   If Image3.Left + Image3.Width > 0.8 * ScaleWidth + ScaleLeft Then
      DeltaX = -100
   End If
   If Image3.Top < ScaleTop Then DeltaY = 100
   If Image3.Top + Image3.Height > 0.9 * ScaleHeight + ScaleTop Then
      DeltaY = -100
   End If
End Sub

Private Sub Form_Load()
NamaFolder = "C:\autodraf\"
   Timer1.Interval = 100   ' Set Interval.
   DeltaX = 100  ' Initialize variables.
   DeltaY = 100
   Counter = 1
End Sub
Private Sub Command1_Click()
Unload Form2
End Sub


