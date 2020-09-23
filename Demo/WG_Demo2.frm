VERSION 5.00
Begin VB.Form Lissajou 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lissajou Patterns"
   ClientHeight    =   4935
   ClientLeft      =   4770
   ClientTop       =   4245
   ClientWidth     =   3735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   3735
   Begin VB.OptionButton OPT 
      Caption         =   "Gear"
      Height          =   255
      Index           =   7
      Left            =   2220
      TabIndex        =   8
      Top             =   930
      Width           =   1455
   End
   Begin VB.OptionButton OPT 
      Caption         =   "Flower"
      Height          =   255
      Index           =   6
      Left            =   2220
      TabIndex        =   7
      Top             =   630
      Width           =   1455
   End
   Begin VB.OptionButton OPT 
      Caption         =   "Chain"
      Height          =   255
      Index           =   5
      Left            =   2220
      TabIndex        =   6
      Top             =   330
      Width           =   1455
   End
   Begin VB.OptionButton OPT 
      Caption         =   "Infinity"
      Height          =   255
      Index           =   4
      Left            =   2220
      TabIndex        =   5
      Top             =   30
      Width           =   1455
   End
   Begin VB.OptionButton OPT 
      Caption         =   "Diamond"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   4
      Top             =   930
      Width           =   1455
   End
   Begin VB.PictureBox GFX 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FF0000&
      Height          =   3615
      Left            =   60
      ScaleHeight     =   3615
      ScaleWidth      =   3615
      TabIndex        =   3
      Top             =   1260
      Width           =   3615
   End
   Begin VB.OptionButton OPT 
      Caption         =   "Square"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   630
      Width           =   1455
   End
   Begin VB.OptionButton OPT 
      Caption         =   "Oval"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   330
      Width           =   1455
   End
   Begin VB.OptionButton OPT 
      Caption         =   "Circle"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   30
      Width           =   1455
   End
End
Attribute VB_Name = "Lissajou"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Simple lissajou patterns
Private OSC As HybridOsc

Private Sub Form_Load()
  Me.Move (Screen.Width - Me.Width) * 0.5, (Screen.Height - Me.Height) * 0.4
End Sub

Private Sub OPT_Click(Index As Integer)

  ' Create oscillators
  Set OSC = New HybridOsc
  Set OSC.Amplitude = New SimpleOsc
  Set OSC.Frequency = New SimpleOsc
  Set OSC.Bias = New SimpleOsc
  
  ' Set initial values
  OSC.Amplitude.Frequency = 1
  OSC.Amplitude.Amplitude = 100
  OSC.Frequency.Frequency = 1
  OSC.Frequency.Amplitude = 100
  OSC.Bias.Bias = 1
  
  Select Case Index
  Case 0 ' Circle
    AddPhase 250
  Case 1 ' Oval
    AddPhase 150
  Case 2 'Square
    OSC.Frequency.Shape = wgSquare
    OSC.Amplitude.Shape = wgSquare
    AddPhase 250
  Case 3 ' Diamond
    OSC.Frequency.Shape = wgTriangular
    OSC.Amplitude.Shape = wgTriangular
    AddPhase 250
  Case 4 ' Infinity
    OSC.Frequency.Frequency = 2
  Case 5 ' Chain
    OSC.Frequency.Frequency = 10
    AddPhase 250
  Case 6 ' Flower
    OSC.Bias.Amplitude = 0.4
    OSC.Bias.Frequency = 20
    OSC.Bias.Bias = 0.6
    AddPhase 250
  Case 7 ' Gear
    OSC.Bias.Amplitude = 0.05
    OSC.Bias.Frequency = 40
    OSC.Bias.Bias = 0.95
    OSC.Bias.Shape = wgSquare
    AddPhase 250
  End Select
  
  DrawWave
  
End Sub

Sub DrawWave()
Dim x&
  ' Init image, draw wave
  InitGFX
  For x = 0 To 2000
    GFX.PSet (OSC.Amplitude.Value * OSC.Bias.Value, OSC.Frequency.Value * OSC.Bias.Value)
    OSC.Tick
    GFX.Line -(OSC.Amplitude.Value * OSC.Bias.Value, OSC.Frequency.Value * OSC.Bias.Value)
  Next
End Sub

Sub InitGFX()
  ' Erase image, set scale, add 0 ref. lines
  Set GFX.Picture = Nothing
  GFX.Scale (-110, 110)-(110, -110)
  GFX.Line (-110, 0)-(110, 0), &HCCCCCC
  GFX.Line (0, -110)-(0, 110), &HCCCCCC
End Sub

Sub AddPhase(Amount As Long)
Dim x&
  ' Advance wave by amount
  For x = 1 To Amount
    OSC.Frequency.Tick
  Next
End Sub
