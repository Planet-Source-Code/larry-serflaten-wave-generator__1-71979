VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ComplexOsc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' COMPLEX OSCILLATOR CLASS
' An oscillator class where wave parameters can be any oscillator type
' Meaning; all properties can be modulated in complex ways

' Generic oscillator interface
Implements I_Oscillator

' End of cycle event
Public Event CycleSync(ByRef Angle As Single)

' Public properties as oscillators
Public Bias As I_Oscillator
Public Amplitude As I_Oscillator
Public Frequency As I_Oscillator
' Internal generator
Private WithEvents mGEN As Generator
Attribute mGEN.VB_VarHelpID = -1

'===============  PUBLIC INTERFACE  ============

Public Property Get Value() As Single
Dim am!, bi!
  If Not Amplitude Is Nothing Then am = Amplitude.Value
  If Not Bias Is Nothing Then bi = Bias.Value
  Value = (mGEN.UnitValue * am) + bi
End Property

Public Property Get Shape() As WaveType
  Shape = mGEN.Shape
End Property

Public Property Let Shape(RHS As WaveType)
  mGEN.Shape = RHS
End Property

Public Sub Tick()
  If Not Amplitude Is Nothing Then Amplitude.Tick
  If Not Bias Is Nothing Then Bias.Tick
  If Not Frequency Is Nothing Then
    Frequency.Tick
    mGEN.Frequency = Frequency.Value
  End If
  mGEN.Tick
End Sub

'===============  GENERIC INTERFACE  ============

Private Property Get I_OSCILLATOR_Clone() As I_Oscillator
Dim CO As ComplexOsc

  ' Create identical object
  Set CO = New ComplexOsc
  
  ' Set properties
  With CO
    Set .Amplitude = Me.Amplitude.Clone
    Set .Bias = Me.Bias.Clone
    Set .Frequency = Me.Frequency.Clone
    .Shape = Me.Shape
  End With
  Set I_OSCILLATOR_Clone = CO
  ' Init new internal generator
  I_OSCILLATOR_Clone.Initialize Frequency.Value, mGEN.Angle, mGEN.Value, mGEN.Shape
  
End Property

Private Sub I_OSCILLATOR_Initialize(ByVal Freq!, ByVal Angle!, ByVal Value!, ByVal Shape As WaveType)
  ' Init "my" internal generator
  With mGEN
    .Frequency = Freq
    .Angle = Angle
    .Value = Value
    .Shape = Shape
  End With
End Sub

Private Sub I_OSCILLATOR_Tick()
  Me.Tick
End Sub

Private Property Get I_OSCILLATOR_Value() As Single
  I_OSCILLATOR_Value = Me.Value
End Property

'===============  PRIVATE ROUTINES  ============

Private Sub Class_Initialize()
  Set mGEN = New Generator
End Sub

Private Sub Class_Terminate()
  Set Amplitude = Nothing
  Set Bias = Nothing
  Set Frequency = Nothing
  Set mGEN = Nothing
End Sub

'===============  EVENT PASSTHROUGH ============

Private Sub mGEN_CycleSync(Angle!)
  RaiseEvent CycleSync(Angle)
End Sub
