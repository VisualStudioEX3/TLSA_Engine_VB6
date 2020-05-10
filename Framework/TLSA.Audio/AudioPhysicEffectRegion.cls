VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AudioPhysicEffectRegion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private varKey As String
Private varRegion As Core.RECTANGLE
Private varEnabled As Boolean
Private varEffects As Audio.SoundEffects

Public Property Get Key() As String
    Key = varKey
End Property

Friend Property Let Key(value As String)
    varKey = value
End Property

Public Property Get Region() As Core.RECTANGLE
    Region = varRegion
End Property

Public Property Let Region(value As Core.RECTANGLE)
    varRegion = value
End Property

Public Sub SetRegion(X As Long, Y As Long, Width As Long, Height As Long)
    With varRegion
        .X = X: .Y = Y: .Width = Width: .Height = Height
    End With
End Sub

Public Property Get Enabled() As Boolean
    Enabled = varEnabled
End Property

Public Property Let Enabled(value As Boolean)
    varEnabled = value
End Property

Public Property Get Effects() As Audio.SoundEffects
    Effects = varEffects
End Property

Public Property Let Effects(value As Audio.SoundEffects)
    varEffects = value
End Property

Public Sub SetEffects(Chorus As Boolean, Compressor As Boolean, Distortion As Boolean, Echo As Boolean, Flanger As Boolean, Gargle As Boolean, WavesReverb As Boolean)
    With varEffects
        .Chorus = Chorus
        .Compressor = Compressor
        .Distortion = Distortion
        .Echo = Echo
        .Flanger = Flanger
        .Gargle = Gargle
        .WavesReverb = WavesReverb
    End With
End Sub
