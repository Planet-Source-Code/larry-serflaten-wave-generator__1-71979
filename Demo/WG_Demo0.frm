VERSION 5.00
Begin VB.Form Sounds 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sound waves"
   ClientHeight    =   2655
   ClientLeft      =   4260
   ClientTop       =   4200
   ClientWidth     =   5745
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   5745
   Begin VB.OptionButton Waves 
      Caption         =   "1K Sine"
      Height          =   255
      Index           =   8
      Left            =   0
      TabIndex        =   9
      Top             =   -1000
      Value           =   -1  'True
      Width           =   1350
   End
   Begin VB.PictureBox GFX 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2085
      Left            =   0
      ScaleHeight     =   2085
      ScaleWidth      =   5715
      TabIndex        =   8
      Top             =   570
      Width           =   5715
   End
   Begin VB.OptionButton Waves 
      Caption         =   "1K Sine"
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   7
      Top             =   0
      Width           =   1350
   End
   Begin VB.OptionButton Waves 
      Caption         =   "1K Square"
      Height          =   255
      Index           =   1
      Left            =   60
      TabIndex        =   6
      Top             =   270
      Width           =   1350
   End
   Begin VB.OptionButton Waves 
      Caption         =   "1K Triangular"
      Height          =   255
      Index           =   2
      Left            =   1500
      TabIndex        =   5
      Top             =   0
      Width           =   1350
   End
   Begin VB.OptionButton Waves 
      Caption         =   "1K Sawtooth"
      Height          =   255
      Index           =   3
      Left            =   1500
      TabIndex        =   4
      Top             =   270
      Width           =   1350
   End
   Begin VB.OptionButton Waves 
      Caption         =   "Am dial tone"
      Height          =   255
      Index           =   4
      Left            =   2940
      TabIndex        =   3
      Top             =   0
      Width           =   1350
   End
   Begin VB.OptionButton Waves 
      Caption         =   "Stereo"
      Height          =   255
      Index           =   5
      Left            =   2940
      TabIndex        =   2
      Top             =   270
      Width           =   1350
   End
   Begin VB.OptionButton Waves 
      Caption         =   "Alarm"
      Height          =   255
      Index           =   6
      Left            =   4380
      TabIndex        =   1
      Top             =   0
      Width           =   1350
   End
   Begin VB.OptionButton Waves 
      Caption         =   "Computer"
      Height          =   255
      Index           =   7
      Left            =   4380
      TabIndex        =   0
      Top             =   270
      Width           =   1350
   End
End
Attribute VB_Name = "Sounds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Sound player
' Builds and plays a few selected sounds

Private PLR As Player
Private Const SampleRate = 1 / 110.25

Private Sub Form_Load()
  ' Init random generator, player object
  Randomize
  Set PLR = New Player
  Me.Move (Screen.Width - Me.Width) * 0.5, (Screen.Height - Me.Height) * 0.4
End Sub

Private Sub GFX_Click()
  PLR.PlaySound
End Sub

Private Sub Waves_Click(Index As Integer)
  ' Disable form while building wave
  Me.Enabled = False
  ' Reset player
  PLR.StopSound
  PLR.ClearData
  ' Reset image
  InitGFX
  
  ' Select option
  Select Case Index
  ' 1K Waves
  Case 0: ShapedWave wgSinusoidal
  Case 1: ShapedWave wgSquare
  Case 2: ShapedWave wgTriangular
  Case 3: ShapedWave wgSawtooth
  ' Custom waves
  Case 4: DialTone
  Case 5: Stereo
  Case 6: Alarm
  Case 7: RandomNotes
  End Select
  
  ' Add click messsage
  GFX.CurrentX = 1200
  GFX.CurrentY = 30
  GFX.Print "CLICK IMAGE TO PLAY SOUND"
  GFX.Refresh
  
  PLR.PlaySound
  ' Ignoring mouse clicks while disabled
  DoEvents
  Me.Enabled = True
End Sub

Sub InitGFX()
  ' Clear old image / add message
  Set GFX.Picture = LoadPicture()
  GFX.Cls
  GFX.Scale (0, 260)-(4000, -10)
  GFX.PSet (1500, 155), vbWhite
  GFX.Print "BUILDING WAVE"
  Set GFX.Picture = GFX.Image
  GFX.Refresh
  'DoEvents
End Sub

Sub ClearGFX()
  ' Clear message / add zero base line
  Set GFX.Picture = Nothing
  GFX.Line (0, 127)-(GFX.ScaleWidth, 127), &HC0C0C0
End Sub

Private Sub RandomNotes()
Dim idx As Long, wait As Long
Dim note As SimpleOsc

  ' Larger image scale
  GFX.Scale (0, 260)-(40000, -10)
  GFX.Line (0, 127)-(40000, 127), &HC0C0C0
  ' Init note osc
  Set note = New SimpleOsc
  note.Amplitude = 40
  note.Bias = 127
  ' Fill buffer / Draw image
  ClearGFX
  For idx = 0 To 880000 Step 2
    If idx < 40000 Then GFX.PSet (idx, note.Value), vbBlue
    note.Tick
    If idx < 40000 Then GFX.Line -(idx, note.Value), vbBlue
    ' Change note periodically
    If wait = 0 Then
      note.Frequency = (((Rnd * 120) + 15) * SampleRate) * 20
    End If
    wait = (wait + 1) Mod 8000
    ' Store L/R data
    PLR.Data(idx) = note.Value
    PLR.Data(idx + 1) = note.Value
  Next
End Sub

Private Sub DialTone()
Dim idx As Long
Dim osc As SimpleOsc, hyb As HybridOsc
  ' Tone is 600 Hz amplitude modulated at 120 Hz
  Set hyb = New HybridOsc
  ' Assign base frequency
  Set osc = New SimpleOsc
  osc.Bias = 600 * SampleRate
  Set hyb.Frequency = osc
  ' Assign modulating frequency
  Set osc = New SimpleOsc
  osc.Amplitude = 50
  osc.Frequency = 120 * SampleRate
  osc.Bias = 64
  Set hyb.Amplitude = osc
  ' Add bias
  Set osc = New SimpleOsc
  osc.Bias = 127
  Set hyb.Bias = osc
  ' Fill buffer / Draw image
  ClearGFX
  For idx = 0 To 880000 Step 2
    If idx < 4000 Then GFX.PSet (idx, hyb.Value), vbBlue
    hyb.Tick
    If idx < 4000 Then GFX.Line -(idx, hyb.Value), vbBlue
    ' Store L/R data
    PLR.Data(idx) = hyb.Value
    PLR.Data(idx + 1) = hyb.Value
  Next
End Sub

Private Sub Alarm()
Dim frq As SimpleOsc
Dim osc As SimpleOsc
Dim idx As Long
  ' Larger image scale
  GFX.Scale (0, 260)-(40000, -10)
  GFX.Line (0, 127)-(40000, 127), &HC0C0C0
    
  Set frq = New SimpleOsc
  Set osc = New SimpleOsc
  osc.Amplitude = 120
  osc.Bias = 127
  osc.Shape = wgSquare
  
  frq.Bias = 200
  frq.Amplitude = 1000
  frq.Frequency = 0.5 * SampleRate
  frq.Shape = wgSawtooth
  ' Fill buffer / Draw image
  ClearGFX
  For idx = 0 To 880000 Step 2
    If idx < 40000 Then GFX.PSet (idx, osc.Value), vbBlue
    frq.Tick
    osc.Frequency = frq.Value * SampleRate
    osc.Tick
    If idx < 40000 Then GFX.Line -(idx, osc.Value), vbBlue
    ' Store L/R data
    If frq.Value < 900 Then
      PLR.Data(idx) = osc.Value
      PLR.Data(idx + 1) = osc.Value
    End If
  Next
End Sub

Private Sub Stereo()
Dim idx As Long, r As Long
Dim osc As SimpleOsc

  Set osc = New SimpleOsc
  osc.Frequency = 1 * SampleRate
  osc.Amplitude = 0.1
  osc.Bias = 0.5
  osc.Shape = wgTriangular
  
  ' Fill buffer / Draw image
  ClearGFX
  For idx = 0 To 880000 Step 2
    If idx < 4000 Then GFX.PSet (idx, r), vbBlue
    osc.Tick
    r = Int(Rnd * 210) + 20
    ' Store L/R data
    PLR.Data(idx) = osc.Value * r
    PLR.Data(idx + 1) = (1 - osc.Value) * r
  Next
  
End Sub
Private Sub ShapedWave(Style As WaveType)
Dim idx As Long
Dim osc As SimpleOsc
  Set osc = New SimpleOsc
  ' All but Sawtooth ride above and below 0 and
  ' have to be biased to stay above 0.
  If Style = wgSawtooth Then
    osc.Bias = 2
    osc.Amplitude = 250
  Else
    osc.Bias = 127
    osc.Amplitude = 124
  End If
  osc.Frequency = 1000 * SampleRate
  osc.Shape = Style
  ' Fill buffer / Draw image
  ClearGFX
  For idx = 0 To 880000 Step 2
    If idx < 4000 Then GFX.PSet (idx, osc.Value), vbBlue
    osc.Tick
    If idx < 4000 Then GFX.Line -(idx, osc.Value), vbBlue
    ' Store L/R data
    PLR.Data(idx) = osc.Value
    PLR.Data(idx + 1) = osc.Value
  Next
End Sub
