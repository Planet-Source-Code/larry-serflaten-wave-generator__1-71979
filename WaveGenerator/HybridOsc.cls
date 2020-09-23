VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HybridOsc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' HYBRID OSCILLATOR CLASS
' An oscillator class where wave parameters are Simple Oscillators
' Meaning; all properties can be modulated in simple fashion

' Generic oscillator interface
Implements I_Oscillator

' End of cycle event
Public Event CycleSync(ByRef Angle!)

' Public properties as oscillators
Public Bias As SimpleOsc
Public Amplitude As SimpleOsc
Public Frequency As SimpleOsc
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
Dim HO As HybridOsc
Dim IW As I_Oscillator

  ' Create identical object
  Set HO = New HybridOsc
  
  ' Set properties (via I_OSCILLATOR.Clone interface)
  With HO
    Set IW = Me.Amplitude
    Set .Amplitude = IW.Clone
    Set IW = Me.Bias
    Set .Bias = IW.Clone
    Set IW = Me.Frequency
    Set .Frequency = IW.Clone
    .Shape = Me.Shape
  End With
  Set I_OSCILLATOR_Clone = HO
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

Private Sub mGEN_CycleSync(ByRef Angle!)
   RaiseEvent CycleSync(Angle)
End Sub

