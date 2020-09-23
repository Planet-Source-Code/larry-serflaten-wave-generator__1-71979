VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "I_Oscillator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' I_OSCILLATOR CLASS
' A simplified interface for obtaining wave output.

' Waveforms have 4 basic shapes
Public Enum WaveType
  wgSinusoidal = 0
  wgTriangular
  wgSquare
  wgSawTooth
End Enum

Public Property Get Clone() As I_Oscillator
Attribute Clone.VB_Description = "Returns a new ocillator object set to current conditions"
' Returns an identical oscillator using current values
End Property

Public Property Get Value() As Single
Attribute Value.VB_Description = "Generic Value property for all oscillators"
' Returns the current waveform value
End Property

Public Sub Initialize(ByVal Freq!, ByVal Angle!, ByVal Value!, ByVal Shape As WaveType)
Attribute Initialize.VB_Description = "Assigns internal generator values"
Attribute Initialize.VB_MemberFlags = "40"
' Sets initial values of the internal generator (used for cloning)
End Sub

Public Sub Tick()
Attribute Tick.VB_Description = "Generic clock tick routine for all osillators"
' Increments a waveform one step
End Sub
