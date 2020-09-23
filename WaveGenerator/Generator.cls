VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Generator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' GENERATOR CLASS
' Generates the changing value for all oscillators

' Event raised at the end of each cycle
Public Event CycleSync(ByRef Angle!)
Attribute CycleSync.VB_Description = "Fires at the end of each full cycle"

' Public properties
Public Angle!
Attribute Angle.VB_VarDescription = "Waveform angle (0 to 2pi)"
Public Frequency!
Attribute Frequency.VB_VarDescription = "Repeat rate of generator"
Public Value!
Attribute Value.VB_VarDescription = "Current waveform value per Frequency and Shape"
Public Shape As WaveType
Attribute Shape.VB_VarDescription = "Determines waveform option"

' Angle rises from 0 to MAX (2 * Pi) in sawtooth fashion
Private Const MAX! = 6.28318530717958
Attribute MAX.VB_VarDescription = "2 * Pi"
' Sample rate setting
Private Const RATE! = MAX / 1000
Attribute RATE.VB_VarDescription = "Increment factor"

Public Sub Tick()
Attribute Tick.VB_Description = "Increments Angle per Frequency"
' Waveform incrementing routine

  ' Increment angle
  Angle = Angle + (RATE * Frequency)
  
  ' Test for end of cycle (user may adjust Angle from event)
  If Angle > MAX Then
    'Angle = 0
    Angle = Angle - MAX
    RaiseEvent CycleSync(Angle)
  End If
  
  ' Determine wave Value per selected shape
  Select Case Shape
  Case WaveType.wgSinusoidal
      Value = Sin(Angle) * MAX
      
  Case WaveType.wgTriangular
      ' First quarter: value rises from 0 to MAX
      If Angle < (MAX * 0.25) Then
          Value = Angle * 4
      ' Second & third quarter: value falls from MAX to -MAX
      ElseIf Angle > (MAX * 0.75) Then
          Value = (Angle - MAX) * 4
      ' Fourth quarter: value rises from -MAX to 0
      Else
          Value = (MAX + MAX) - (Angle * 4)
      End If
  
  Case WaveType.wgSquare
      ' First half: value is at MAX
      If Angle < (MAX / 2) Then
        Value = MAX
      ' Second half: value is at -MAX
      Else
        Value = -MAX
      End If
  
  Case Else ' WaveType.SawTooth
      Value = Angle
  End Select

End Sub

Friend Property Get UnitValue() As Single
Attribute UnitValue.VB_Description = "Returns Angle in the range of -1 to 1"
  ' Returns angle in the range of -1 to 1
  UnitValue = Value / MAX
End Property



