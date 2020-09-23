VERSION 5.00
Begin VB.Form Art 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lissajou Art"
   ClientHeight    =   5820
   ClientLeft      =   3885
   ClientTop       =   3450
   ClientWidth     =   6090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   6090
   Begin VB.Timer TMR 
      Interval        =   500
      Left            =   480
      Top             =   300
   End
End
Attribute VB_Name = "Art"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Lissajou art
' Modulating values for X and values for Y
' produce some interesting patterns.

Private osc(0 To 5) As SimpleOsc
Private Ticker As ComplexOsc     ' Main oscillator
Private Run As Long              ' Number of ticks to finish art
Private Group As Long            ' Image ID
Private Delay As Date            ' Image delay time

Private Sub Form_Load()
   Set Ticker = New ComplexOsc
   ReSizeMe
End Sub

Sub Group0()
   Set osc(0) = Simple(1000, 0, 25, wgSinusoidal)
   Set osc(1) = Simple(0.3, 0.6, 276, wgSinusoidal)
   Set osc(2) = Simple(900, 0, 75, wgSinusoidal)
   Set osc(3) = Simple(0, 0, 0, wgSinusoidal)
   Set osc(4) = Simple(0, 0, 0, wgSinusoidal)
   Set osc(5) = Simple(0, 0, 0, wgSinusoidal)
   AddPhase osc(0), 0.25
   Run = 1000
End Sub

Sub Group1()
   Set osc(0) = Simple(600, 0, 0.25, wgSinusoidal)
   Set osc(1) = Simple(0, 1, 0, wgSinusoidal)
   Set osc(2) = Simple(500, 0, 0.25, wgSinusoidal)
   Set osc(3) = Simple(200, 0, 64, wgSinusoidal)
   Set osc(4) = Simple(0, 1, 0, wgSinusoidal)
   Set osc(5) = Simple(400, 0, 32, wgSinusoidal)
   AddPhase osc(0), 0.25
   Run = 4000
End Sub

Sub Group2()
   Set osc(0) = Simple(240, 0, 35, wgSinusoidal)
   Set osc(1) = Simple(0, 1, 0, wgSinusoidal)
   Set osc(2) = Simple(240, 0, 35, wgSinusoidal)
   Set osc(3) = Simple(600, 0, 0.2, wgSinusoidal)
   Set osc(4) = Simple(0.4, 0.8, 1, wgSinusoidal)
   Set osc(5) = Simple(600, 0, 0.2, wgSinusoidal)
   AddPhase osc(0), 0.25
   AddPhase osc(3), 0.25
   Run = 6000
End Sub

Sub Group3()
   Set osc(0) = Simple(1000, 0, 25, wgSinusoidal)
   Set osc(1) = Simple(0.3, 0.6, 26, wgSinusoidal)
   Set osc(2) = Simple(900, 0, 50, wgSinusoidal)
   Set osc(3) = Simple(0, 0, 0, wgSinusoidal)
   Set osc(4) = Simple(0, 0, 0, wgSinusoidal)
   Set osc(5) = Simple(0, 0, 0, wgSinusoidal)
   AddPhase osc(0), 0.25
   Run = 1000
End Sub

Sub Group4()
   Set osc(0) = Simple(300, 0, 4, wgSinusoidal)
   Set osc(1) = Simple(0.6, 0.3, 75, wgSinusoidal)
   Set osc(2) = Simple(300, 0, 4, wgSinusoidal)
   Set osc(3) = Simple(200, 0, 0.75, wgSinusoidal)
   Set osc(4) = Simple(0, 3, 0, wgSinusoidal)
   Set osc(5) = Simple(200, 0, 0.75, wgSinusoidal)
   AddPhase osc(0), 0.25
   AddPhase osc(3), 0.25
   Run = 8000
End Sub

Sub Draw()
Dim ITR As Long
Dim XX As Single, YY As Single
   ' Draws waveform
   Set Me.Picture = Nothing
   Ticker.Tick
   For ITR = 1 To Run
     PSet (osc(0).Value * osc(1).Value + osc(3).Value * osc(4).Value, osc(2).Value * osc(1).Value + osc(5).Value * osc(4).Value), vbWhite
     Ticker.Tick
     Line -(osc(0).Value * osc(1).Value + osc(3).Value * osc(4).Value, osc(2).Value * osc(1).Value + osc(5).Value * osc(4).Value), vbWhite
   Next
End Sub

Function Simple(Amp As Single, Bias As Single, Freq As Single, Shape As WaveType) As SimpleOsc
   ' Helper routine to build an oscillator
   Set Simple = New SimpleOsc
   With Simple
     .Amplitude = Amp
     .Bias = Bias
     .Frequency = Freq
     .Shape = Shape
   End With
End Function

Function Hybrid(Amp As SimpleOsc, Bias As SimpleOsc, Freq As SimpleOsc, Shape As WaveType) As HybridOsc
   ' Helper routine to build an oscillator
   Set Hybrid = New HybridOsc
   With Hybrid
     Set .Amplitude = Amp
     Set .Bias = Bias
     Set .Frequency = Freq
     .Shape = Shape
   End With
End Function

Sub Gather()
   ' Assigns oscillator array to single complex oscillator
   Set Ticker.Amplitude = Hybrid(osc(0), osc(1), osc(2), wgSinusoidal)
   Set Ticker.Frequency = Hybrid(osc(3), osc(4), osc(5), wgSinusoidal)
End Sub

Sub AddPhase(osc As SimpleOsc, ByVal phase As Single)
Dim ticks As Long
   ' Advances wave by phase %
   ticks = (1000 / osc.Frequency) * phase
   While ticks > 0
     osc.Tick
     ticks = ticks - 1
   Wend
End Sub

Sub ReSizeMe()
Dim bdr As Long, ttl As Long
Dim wid As Long, hgt As Long
   ' Makes client area square
   bdr = Width - ScaleWidth
   ttl = Height - ScaleHeight - bdr
   wid = 6000 + bdr
   hgt = 6000 + ttl + bdr
   ' Center form / Add scale
   Me.Move (Screen.Width - wid) / 2, (Screen.Height - hgt) / 2, wid, hgt
   Me.Scale (-1000, 1000)-(1000, -1000)
End Sub

Private Sub TMR_Timer()
   ' Cycles through artwork
   If Now > Delay Then
     Delay = DateAdd("s", 8, Now)
     Select Case Group
     Case 0: Group0
     Case 1: Group1
     Case 2: Group2
     Case 3: Group3
     Case 4: Group4: Group = -1
     End Select
     Gather
     Draw
     Group = Group + 1
   End If
End Sub
