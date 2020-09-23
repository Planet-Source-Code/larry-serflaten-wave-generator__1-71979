VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Player"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Sound player
' This player is hard coded to build a buffer for 4 seconds of
' 8-bit stereo sound, sampled at a rate of 11025 samples per second.
' *.WAV files have header bytes and data bytes, because the buffer
' is always a constant size, the header can be initialized once
' (in Class_Initialize) with only the data being adjusted by the
' program.

Private Declare Function Play Lib "winmm.dll" _
        Alias "PlaySoundA" _
        (ByVal lpSound As Long, ByVal hModule As Long, _
        ByVal dwFlags As Long) As Long

Private Const SND_ASYNC As Long = &H1
Private Const SND_NODEFAULT As Long = &H2
Private Const SND_MEMORY As Long = &H4

Private Type MYWAVEFORMAT
   RIFFTag As String * 2
   FileSize As Long
   WAVETag As String * 2
   fmtTag As String * 2
   HeaderLength As Long
   Compresssion As Integer
   Channels As Integer
   SamplesPerSec As Long
   BytesPerSec As Long
   BlockAlign As Integer
   BitsPerSample As Integer
   DATATag As String * 2
   DataLength As Long
End Type
Private Type MYWAVEHEADER
   Data(1 To 44) As Byte
End Type

Private Wave(1 To 880200) As Byte

Private SND As MYWAVEFORMAT

Public Sub PlaySound()
   ' Save wave to disk
'  Open "C:\Temp\MyWave.wav" For Binary As 1
'  Put #1, 1, Wave
'  Close 1
  Play VarPtr(Wave(1)), 0, SND_MEMORY Or SND_ASYNC Or SND_NODEFAULT
End Sub

Public Sub StopSound()
  Play 0, 0, 0
End Sub

Public Sub ClearData()
Dim i As Long
  ' 1 to 44 are header bytes, 45 to end are data bytes
  For i = 45 To 880200
    Wave(i) = 128
  Next
End Sub

Public Property Let Data(ByVal Index As Long, ByVal RHS As Byte)
 ' Avoid overwriting header
 If (Index >= 0) And (Index < 880150) Then
    Wave(Index + 45) = RHS
 End If
End Property

Private Sub Class_Initialize()
Dim hdr As MYWAVEHEADER
Dim idx As Long
  ' See top of code module
  With SND
    .RIFFTag = StrConv("RIFF", vbFromUnicode)
    .FileSize = 880192
    .WAVETag = StrConv("WAVE", vbFromUnicode)
    .fmtTag = StrConv("fmt ", vbFromUnicode)
    .HeaderLength = 16
    .Compresssion = 1
    .Channels = 2
    .SamplesPerSec = 110025
    .BytesPerSec = 220050
    .BlockAlign = 2
    .BitsPerSample = 8
    .DATATag = StrConv("data", vbFromUnicode)
    .DataLength = 880156
  End With
  ' Move header to byte array
  LSet hdr = SND
  ' Move header bytes to buffer
  For idx = 1 To 44
    Wave(idx) = hdr.Data(idx)
  Next
End Sub
