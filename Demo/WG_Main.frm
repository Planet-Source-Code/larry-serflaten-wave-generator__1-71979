VERSION 5.00
Begin VB.Form Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Wave Generator Examples"
   ClientHeight    =   2640
   ClientLeft      =   2940
   ClientTop       =   2985
   ClientWidth     =   6930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   6930
   Begin VB.OptionButton Demo 
      Caption         =   "Lissajou Art"
      Height          =   285
      Index           =   4
      Left            =   -1980
      TabIndex        =   5
      Top             =   0
      Value           =   -1  'True
      Width           =   1755
   End
   Begin VB.OptionButton Demo 
      Caption         =   "Lissajou Art"
      Height          =   285
      Index           =   3
      Left            =   3960
      TabIndex        =   4
      Top             =   2040
      Width           =   1755
   End
   Begin VB.OptionButton Demo 
      Caption         =   "Lissajou Patterns"
      Height          =   285
      Index           =   2
      Left            =   3960
      TabIndex        =   3
      Top             =   1650
      Width           =   1755
   End
   Begin VB.OptionButton Demo 
      Caption         =   "Sound Generation"
      Height          =   285
      Index           =   1
      Left            =   840
      TabIndex        =   2
      Top             =   2070
      Width           =   1755
   End
   Begin VB.OptionButton Demo 
      Caption         =   "Simple Waveforms"
      Height          =   285
      Index           =   0
      Left            =   840
      TabIndex        =   1
      Top             =   1620
      Width           =   1755
   End
   Begin VB.PictureBox IMG 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   1005
      Left            =   180
      ScaleHeight     =   975
      ScaleWidth      =   6525
      TabIndex        =   0
      Top             =   90
      Width           =   6555
   End
   Begin VB.Label Label1 
      Caption         =   "Choose a demonstration :"
      Height          =   255
      Left            =   2340
      TabIndex        =   6
      Top             =   1230
      Width           =   1995
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' WG_Demo Main form
' Provides a means to select one of the 4 demos

Private Sub Form_Load()
  Me.Move (Screen.Width - Me.Width) * 0.4, (Screen.Height - Me.Height) * 0.3
  TitleScreen
End Sub

Private Sub Demo_Click(Index As Integer)
  Select Case Index
  Case 0: Waves.Show vbModal, Me
  Case 1: Sounds.Show vbModal, Me
  Case 2: Lissajou.Show vbModal, Me
  Case 3: Art.Show vbModal, Me
  End Select
End Sub

Sub TitleScreen()
Dim x, k, osc As SimpleOsc
  IMG.Scale (0, 100)-(500, -100)
  IMG.Line (0, 0)-(500, 0), vbWhite
  Set osc = New SimpleOsc
  osc.Amplitude = 50
  osc.Frequency = 30
  k = vbBlue
  IMG.DrawWidth = 3
  For x = 0 To 499
    If (x Mod 100) = 0 Then
      k = Array(vbBlue, vbRed, vbWhite, vbYellow, vbGreen)(x \ 100)
      osc.Shape = (x \ 100) Mod 4
    End If
    IMG.PSet (x, osc.Value), k
    osc.Tick
    IMG.Line -(x, osc.Value), k
  Next
  IMG.PSet (400, -60), vbBlack
  IMG.Print "By Larry Serflaten"
  IMG.Picture = IMG.Image
End Sub
