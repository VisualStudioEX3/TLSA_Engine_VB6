VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SoundChannel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Type SoundEffects
    Chorus As Boolean
    Compressor As Boolean
    Distortion As Boolean
    Echo As Boolean
    Flanger As Boolean
    Gargle As Boolean
    WavesReverb As Boolean
End Type

Private varIndex As Long                    ' Indica el indice del canal.
Private varSample As Audio.SoundSample      ' Referencia a la muestra de sonido a reproducir.
Private varType As Audio.Audio_Types        ' Indica el tipo de sonido que reproduce.
Private varEffects As Audio.SoundEffects    ' Efectos de sonidos aplicados al canal.

Private varSpeedValue As Long               ' Guarda el valor porcentual de la velocidad establecida.

' Indica si el canal esta activo. Si el canal no esta activo significa que la muestra de sonido ya termino de reproducirse
' y el canal esta libre:
Public Property Get Active() As Boolean
    Active = snd.SOUND_IsPlaying(varIndex)
End Property

' Categoria asignada al canal:
Public Property Get Category() As Audio.Audio_Types
    Category = varType
End Property

Friend Property Let Category(value As Audio.Audio_Types)
    varType = value
End Property

Public Property Get Effects() As Audio.SoundEffects
    Effects = varEffects
End Property

Public Property Let Effects(value As Audio.SoundEffects)
    varEffects = value
    Call SetEffects
End Property

' Aplica los efectos de sonido en el canal:
Public Sub SetEffects()
    Call snd.SOUND_FX_SetEffects(varIndex, snd.SOUND_FX_MakeArrayEffects(varEffects.Chorus, varEffects.Compressor, varEffects.Distortion, varEffects.Echo, varEffects.Flanger, varEffects.Gargle, varEffects.WavesReverb))
End Sub

' Elimina todos los efectos de sonido del canal:
Public Sub RemoveEffects()
    Call snd.SOUND_FX_DisableAllEffects(varIndex)
End Sub

Friend Property Let Index(value As Long)
    varIndex = value
End Property

' Devuelve el indice del canal que representa:
Public Property Get Index() As Long
    Index = varIndex
End Property

Friend Property Set Sample(value As Audio.SoundSample)
    Set varSample = value
End Property

' Devuelve la instancia de la muestra de sonido que reproduce:
Public Property Get Sample() As Audio.SoundSample
    Set Sample = varSample
End Property

' Pausa o retoma la reproduccion del canal:
Public Sub Pause()
    Call snd.SOUND_Pause(varIndex)
End Sub

' *** Acceso para el mezclador de sonido ***
Friend Sub SetManagedVolume(value As Long)
    Call snd.SOUND_SetVolume(varIndex, value)
End Sub

Public Property Let Volume(value As Long)
    ' Si esta configurado como manual se asigna el valor:
    If varType = Audio_Type_Manual Then Call snd.SOUND_SetVolume(varIndex, value)
End Property

Public Property Get Volume() As Long
    Call snd.SOUND_GetVolume(varIndex)
End Property

' Establece la velocidad de reproduccion usando una escala de porcentajes enteros (e.j.: 100 representa la velocidad normal de la muestra de sonido)
Public Property Let Speed(value As Long)
    ' Value * Range / 100
    Call snd.SOUND_SetFrequency(varIndex, (value / 100 * varSample.SamplesPerSecond))  ' value * varSample.SamplesPerSecond / 100)
    varSpeedValue = value
End Property

' Devuelve la velocidad de reproduccion en escala de porcentajes enteros (25%, 60%, 123%, etc...)
Public Property Get Speed() As Long
    ' Value / Range * 100
    Speed = varSpeedValue ' snd.SOUND_GetFrequency(varIndex) / varSample.SamplesPerSecond * 100
End Property

' *** Acceso para el mezclador de sonido ***
Friend Sub SetManagedBalance(value As Long)
    Call snd.SOUND_SetPan(varIndex, value)
End Sub

Public Property Let Balance(value As Long)
    ' Si esta configurado como manual se asigna el valor:
    If varType = Audio_Type_Manual Then Call snd.SOUND_SetPan(varIndex, value)
End Property

Public Property Get Balance() As Long
    Balance = snd.SOUND_GetPan(varIndex)
End Property
