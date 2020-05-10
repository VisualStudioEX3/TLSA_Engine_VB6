VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SoundFX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private varChorus As Boolean
Private varCompressor As Boolean
Private varDistortion As Boolean
Private varEcho As Boolean
Private varFlanger As Boolean
Private varGargle As Boolean
Private varWavesReverb As Boolean

Public Property Get Chorus() As Boolean
    Chorus = varChorus
End Property

Public Property Let Chorus(value As Boolean)
    varChorus = value
End Property

Public Property Get Compressor() As Boolean
    Compressor = varCompressor
End Property

Public Property Let Compressor(value As Boolean)
    varCompressor = value
End Property

Public Property Get Distortion() As Boolean
    Distortion = varDistortion
End Property

Public Property Let Distortion(value As Boolean)
    varDistortion = value
End Property

Public Property Get Echo() As Boolean
    Echo = varEcho
End Property

Public Property Let Echo(value As Boolean)
    varEcho = value
End Property

Public Property Get Flanger() As Boolean
    Flanger = varFlanger
End Property

Public Property Let Flanger(value As Boolean)
    varFlanger = value
End Property

Public Property Get Gargle() As Boolean
    Gargle = varGargle
End Property

Public Property Let Gargle(value As Boolean)
    varGargle = value
End Property

Public Property Get WavesReverb() As Boolean
    WavesReverb = varWavesReverb
End Property

Public Property Let WavesReverb(value As Boolean)
    varWavesReverb = value
End Property
