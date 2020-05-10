VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MusicChannel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private varIndex As Long                    ' Indica el indice del canal.
Private varSample As Audio.MusicSample      ' Referencia a la muestra de sonido a reproducir.

' Indica si el canal esta activo. Si el canal no esta activo significa que la muestra de sonido ya termino de reproducirse
' y el canal esta libre:
Public Property Get Active() As Boolean
    Active = snd.MUSIC_IsPlaying(varIndex)
End Property

Friend Property Let Index(value As Long)
    varIndex = value
End Property

' Devuelve el indice del canal que representa:
Public Property Get Index() As Long
    Index = varIndex
End Property

Friend Property Set Sample(value As Audio.MusicSample)
    varSample = value
End Property

' Devuelve la instancia de la muestra de sonido que reproduce:
Public Property Get Sample() As Audio.MusicSample
    Set Sample = varSample
End Property

' Pausa o retoma la reproduccion del canal:
Public Sub Pause()
    Call snd.MUSIC_Pause(varIndex)
End Sub

' *** Acceso para el mezclador de sonido ***
Friend Sub SetManagedVolume(value As Long)
    Call snd.MUSIC_SetVolume(varIndex, value)
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
    Call snd.SOUND_SetFrequency(varIndex, System.Math.GetPercentFromValue(value, varSample.SamplesPerSecond))  ' value * varSample.SamplesPerSecond / 100)
End Property

' Devuelve la velocidad de reproduccion en escala de porcentajes enteros (25%, 60%, 123%, etc...)
Public Property Get Speed() As Long
    ' Value / Range * 100
    Speed = System.Math.GetValueFromPercent(snd.SOUND_GetFrequency(varIndex), varSample.SamplesPerSecond) ' snd.SOUND_GetFrequency(varIndex) / varSample.SamplesPerSecond * 100
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

