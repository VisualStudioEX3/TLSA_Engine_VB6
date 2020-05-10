VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SoundMixer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mSoundSamples As Audio.SoundSampleList
Private varPhysicEngine As Audio.AudioPhysicEngine

' Devuelve la lista de muestras de sonido cargadas:
Public Property Get Samples() As Audio.SoundSampleList
    Set Samples = mSoundSamples
End Property

' Pausa o reanuda la reproduccion de todos los canales de sonido:
Public Sub GloablPause()
    Call modAudio.snd.SOUND_Pause(-1)
End Sub

' Detiene la reproduccion de todos los canales de sonido:
Public Sub GlobalStop()
    Call modAudio.snd.SOUND_Stop(-1)
End Sub

' Obtiene o establece los efectos FX globales a los canales que lo permitan:
Public Property Let GlobalEffects(value As Audio.SoundEffects)
    fxGlobalEffects = value
    modAudio.fxUpdate = True
End Property

Public Property Get GlobalEffects() As Audio.SoundEffects
    GlobalEffects = fxGlobalEffects
End Property

' Obtiene o establece la velocidad global de reproduccion a los canales que lo permitan:
Public Property Get GlobalSpeed() As Long
    GlobalSpeed = modAudio.spdGlobalSpeed
End Property

Public Property Let GlobalSpeed(value As Long)
    Dim last As Long: last = modAudio.spdGlobalSpeed
    modAudio.spdGlobalSpeed = value
    modAudio.spdUpdate = last <> value
End Property

' Devuelve la instancia de un canal de sonido:
Public Property Get Channel(Index As Long) As Audio.SoundChannel
    Set Channel = modAudio.sndChannels(Index)
End Property

' Devuelve el numero de canales:
Public Property Get Channels() As Long
    Channels = UBound(modAudio.sndChannels) + 1
End Property

' Acceso al motor de fisicas de audio:
Public Property Get PhysicEngine() As Audio.AudioPhysicEngine
    Set PhysicEngine = varPhysicEngine
End Property

Private Sub Class_Initialize()
    Dim i As Long
    For i = 0 To UBound(modAudio.sndChannels)
        modAudio.sndChannels(i).Index = i
    Next
    
    Set mSoundSamples = New Audio.SoundSampleList
    Set varPhysicEngine = New Audio.AudioPhysicEngine
    
    modAudio.spdGlobalSpeed = 100
    modAudio.volAmbient = 100
    modAudio.volEffects = 100
    modAudio.volGUI = 100
    modAudio.volMusic = 100
    modAudio.volVoice = 100
End Sub

Private Sub Class_Terminate()
    Set mSoundSamples = Nothing
End Sub

' Actualiza los estados de los canales de sonido en reproduccion:
Friend Sub Update()
    ' Actualizamos los estados globales:
    Dim i As Long
    For i = 0 To UBound(modAudio.sndChannels)
        With modAudio.sndChannels(i)
            If Not .Category = Audio_Type_Manual And Not .Category = Audio_Type_GUI Then
                If modAudio.fxUpdate Then
                    .Effects = modAudio.fxGlobalEffects
                End If
                If modAudio.spdUpdate Then
                    .Speed = modAudio.spdGlobalSpeed
                End If
            End If
        End With
    Next
    
    modAudio.fxUpdate = False
    
    ' Actualiza el simulador de fisicas en audio si esta activado:
    If varPhysicEngine.Enabled Then Call varPhysicEngine.Update
End Sub
