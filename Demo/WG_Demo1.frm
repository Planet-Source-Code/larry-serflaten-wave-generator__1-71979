VERSION 5.00
Begin VB.Form Waves 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Wave forms"
   ClientHeight    =   2775
   ClientLeft      =   3345
   ClientTop       =   4200
   ClientWidth     =   5835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   185
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   389
   Begin VB.OptionButton Waves 
      Caption         =   "Pulse"
      Height          =   255
      Index           =   7
      Left            =   4440
      TabIndex        =   8
      Top             =   330
      Width           =   1350
   End
   Begin VB.OptionButton Waves 
      Caption         =   "Amp && Freq"
      Height          =   255
      Index           =   6
      Left            =   4440
      TabIndex        =   7
      Top             =   60
      Width           =   1350
   End
   Begin VB.OptionButton Waves 
      Caption         =   "Freq. Mod"
      Height          =   255
      Index           =   5
      Left            =   3000
      TabIndex        =   6
      Top             =   330
      Width           =   1350
   End
   Begin VB.OptionButton Waves 
      Caption         =   "Amp. Mod"
      Height          =   255
      Index           =   4
      Left            =   3000
      TabIndex        =   5
      Top             =   60
      Width           =   1350
   End
   Begin VB.OptionButton Waves 
      Caption         =   "Sawtooth"
      Height          =   255
      Index           =   3
      Left            =   1560
      TabIndex        =   4
      Top             =   330
      Width           =   1350
   End
   Begin VB.OptionButton Waves 
      Caption         =   "Triangular"
      Height          =   255
      Index           =   2
      Left            =   1560
      TabIndex        =   3
      Top             =   60
      Width           =   1350
   End
   Begin VB.OptionButton Waves 
      Caption         =   "Square"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   330
      Width           =   1350
   End
   Begin VB.OptionButton Waves 
      Caption         =   "Sinusoidal"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   60
      Width           =   1350
   End
   Begin VB.PictureBox GFX 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2085
      Left            =   60
      ScaleHeight     =   2085
      ScaleWidth      =   5715
      TabIndex        =   0
      Top             =   630
      Width           =   5715
   End
End
Attribute VB_Name = "Waves"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Wave generator
' Builds and shows different waveforms.

Private osc As SimpleOsc
Attribute osc.VB_VarHelpID = -1

Private Sub Form_Load()
  Me.Move (Screen.Width - Me.Width) * 0.5, (Screen.Height - Me.Height) * 0.4
End Sub

Private Sub Waves_Click(Index As Integer)
  ' Init new generator (nominal values)
  Set osc = New SimpleOsc
  osc.Frequency = 4
  osc.Amplitude = 100
  
  ' Simple wavforms
  Select Case Index
  Case 0 ' Sinusoidal
    osc.Shape = wgSinusoidal
  Case 1 ' Square
    osc.Shape = wgSquare
  Case 2 ' Triangular
    osc.Shape = wgTriangular
  Case 3 ' SawTooth
    osc.Shape = wgSawtooth
  
  ' Hybrid waveforms
  Case 4 ' Amp Mod
    ShowModulated 1
  Case 5 ' Frq Mod
    ShowModulated 2
  Case 6 ' Amp & Frq
    ShowModulated 3
  
  ' Program assisted waveforms
  Case 7 ' Pulse
    DrawPulse
  Case 8 ' Custom
  End Select
  
  ' Generic draw routine
  If Index < 4 Then DrawWave osc

End Sub

Sub DrawWave(ByVal Wav As I_Oscillator)
Dim x&
  ' Init image, draw wave
  InitGFX
  For x = 0 To 1000
    Wav.Tick
    GFX.Line -(x, Wav.Value), vbBlue
  Next
End Sub

Sub DrawPulse()
Dim x&, y&
  
  InitGFX
  
  ' Sawtooth allows for (easily) tracking wave over complete cycle
  osc.Shape = wgSawtooth
  
  For x = 0 To 1000
    osc.Tick
    ' Pulse hi
    If osc.Value < 25 Then
      y = 100
    ' Pulse lo
    Else
      y = 0
    End If
    GFX.Line -(x, y), vbBlue
  Next
End Sub


Sub ShowModulated(ByVal Index As Long)
Dim h As HybridOsc
  
  ' Create oscillator object
  Set h = New HybridOsc
  ' Create property oscillators
  Set h.Frequency = New SimpleOsc
  Set h.Amplitude = New SimpleOsc
  Set h.Bias = New SimpleOsc
  
  Select Case Index
  Case 1 ' Amp Mod
    h.Amplitude.Amplitude = 100
    h.Amplitude.Frequency = 40
    h.Frequency.Bias = 2
  Case 2 ' Freq Mod
    h.Amplitude.Bias = 100
    h.Frequency.Shape = wgSawtooth
    h.Frequency.Amplitude = 80
    h.Frequency.Frequency = 4
  Case 3  ' Amp & Freq
    h.Amplitude.Amplitude = 15
    h.Amplitude.Bias = 25
    h.Amplitude.Frequency = 4
    h.Amplitude.Shape = wgSinusoidal
    h.Bias.Amplitude = 70
    h.Bias.Frequency = 2
    h.Frequency.Amplitude = 40
    h.Frequency.Bias = 60
    h.Frequency.Frequency = 4
    h.Frequency.Shape = wgTriangular
  End Select
  
  ' Note same draw code used for simple, hybrid, and complex oscillators
  DrawWave h
End Sub

Sub InitGFX()
  ' Erase image, set scale, add 0 ref. line, set pen
  Set GFX.Picture = Nothing
  GFX.Scale (0, 110)-(1000, -110)
  GFX.Line (0, 0)-(1000, 0), &HCCCCCC
  GFX.PSet (0, 0), vbBlack
End Sub
