VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SimpleOsc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' SIMPLE OSCILLATOR CLASS
' An oscillator class where wave parameters are type Single
' Meaning; all properties are a constant value

' Generic oscillator interface
Implements I_Oscillator

' End of cycle event
Public Event CycleSync(ByRef Angle!)

' Memory shadow of Public variables
Private mBias!
Private mAmplitude!
' Internal generator
Private WithEvents mGEN As Generator
Attribute mGEN.VB_VarHelpID = -1

'===============  PUBLIC INTERFACE  ============

Public Property Get Value() As Single
Attribute Value.VB_Description = "Returns current waveform value per Frequency and Shape"
  Value = (mGEN.UnitValue * mAmplitude) + mBias
End Property

Public Property Get Bias() As Single
Attribute Bias.VB_Description = "Offset amount added to oscillator Value"
  Bias = mBias
End Property

Public Property Let Bias(RHS As Single)
  mBias = RHS
End Property

Public Property Get Amplitude() As Single
Attribute Amplitude.VB_Description = "Maximum waveform value"
  Amplitude = mAmplitude
End Property

Public Property Let Amplitude(RHS As Single)
  mAmplitude = RHS
End Property

Public Property Get Shape() As WaveType
Attribute Shape.VB_Description = "Determines waveform shape"
  Shape = mGEN.Shape
End Property

Public Property Let Shape(RHS As WaveType)
  mGEN.Shape = RHS
End Property

Public Property Get Frequency() As Single
Attribute Frequency.VB_Description = "Determines waveform repeat rate"
  Frequency = mGEN.Frequency
End Property

Public Property Let Frequency(RHS As Single)
  mGEN.Frequency = RHS
End Property

Public Sub Tick()
Attribute Tick.VB_Description = "Increments Angle per Frequency"
  mGEN.Tick
End Sub

'===============  GENERIC INTERFACE  ============

Private Property Get I_OSCILLATOR_Clone() As I_Oscillator
Dim WG As SimpleOsc
  
  ' Create identical object
  Set WG = New SimpleOsc
  ' Set properties
  With WG
    .Amplitude = Me.Amplitude
    .Bias = Me.Bias
    .Frequency = Me.Frequency
    .Shape = Me.Shape
  End With
  Set I_OSCILLATOR_Clone = WG
  ' Init new internal generator
  I_OSCILLATOR_Clone.Initialize mGEN.Frequency, mGEN.Angle, mGEN.Value, mGEN.Shape
  
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
  Set mGEN = Nothing
End Sub

'===============  EVENT PASSTHROUGH ============

Private Sub mGEN_CycleSync(ByRef Angle!)
   RaiseEvent CycleSync(Angle)
End Sub
