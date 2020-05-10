VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MusicMixer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private varSamples As Audio.MusicSampleList
Private varCurrentSample As Audio.MusicSample
Private varCurrentIndex As Long ' Canal que esta reproduciendo.
Private varFading(1) As Long ' 1 - Fade de silencio a nivel volumen (<), 2 - Fade a silencio (>)
Private varFadeDelay As Long ' Retraso de tiempo para la progresion de los fades.
Private varTimer As Core.Timer

Public Property Get FadeDelay() As Long
    FadeDelay = varFadeDelay
End Property

Public Property Let FadeDelay(value As Long)
    varFadeDelay = value
End Property

Public Property Get Samples() As Audio.MusicSampleList
    Set Samples = varSamples
End Property

Public Sub Play_(Sample As Audio.MusicSample, Optional Looping As Boolean = False, Optional Fading As Boolean = False)
    Set varCurrentSample = Sample
    
    Dim setVol As Long
    If Fading Then
        Call Me.Stop_(True) ' Hacemos un stop con fading al canal actual (da igual si no esta en uso).
        If varCurrentIndex = 0 Then varCurrentIndex = 1 Else varCurrentIndex = 0 ' Intercambiamos el canal activo.
        varFading(varCurrentIndex) = 1 ' Activamos el fading desde silencio a volumen.
        setVol = modAudio.volMusic \ 2 ' Establecemos el volumen a 0.
    Else
        Call snd.MUSIC_Stop(0): Call snd.MUSIC_Stop(1)
        varFading(varCurrentIndex) = 0 ' Desactivamos el fading por si estuviera activo.
        setVol = modAudio.volMusic ' Establecemos el volumen al valor establecido en el control de volumen.
    End If
    
    Call snd.MUSIC_Play(varCurrentSample.Id, varCurrentIndex, Looping)
    Call snd.MUSIC_SetVolume(varCurrentIndex, setVol)
End Sub

Public Sub Pause_()
    Call snd.MUSIC_Pause(0): Call snd.MUSIC_Pause(1)
End Sub

Public Sub Stop_(Optional Fading As Boolean = False)
    If Fading Then
        varFading(varCurrentIndex) = 2
    Else
        Call snd.MUSIC_Stop(0): Call snd.MUSIC_Stop(1)
    End If
End Sub

' Indica si el canal actual esta en reproduccion:
Public Property Get IsPlaying() As Boolean
    IsPlaying = snd.MUSIC_IsPlaying(varCurrentIndex)
End Property

' Muestra de musica que se esta reproduciendo:
Public Property Get CurrentSample() As Audio.MusicSample
    Set CurrentSample = varCurrentSample
End Property

' Devuelve la posicion actual de reproduccion en milisegundos:
Public Property Get CurrentPosition() As Long
    CurrentPosition = snd.MUSIC_GetCurrentPosition(varCurrentIndex)
End Property

' Establece la posicion actual de reproduccion en milisegundos:
Public Property Let CurrentPosition(value As Long)
    Call snd.MUSIC_SetCurrentPosition(varCurrentIndex, value)
End Property

Public Property Get Balance() As Long
    Balance = snd.MUSIC_GetPan(varCurrentIndex)
End Property

Public Property Let Balance(value As Long)
    Call snd.MUSIC_SetPan(varCurrentIndex, value)
End Property

Public Property Get Speed() As Long
    Speed = snd.MUSIC_GetRate(varCurrentIndex)
End Property

Public Property Let Speed(value As Long)
    Call snd.MUSIC_SetRate(varCurrentIndex, value)
End Property

' Actualiza los estados de la reproduccion de musica:
Friend Sub Update()
    Dim i As Long
    
    If varTimer.value >= varFadeDelay Then
        For i = 0 To 1
            If varFading(i) = 1 Then
                If snd.MUSIC_GetVolume(i) < modAudio.volMusic Then
                    Call snd.MUSIC_SetVolume(i, snd.MUSIC_GetVolume(i) + 1)
                Else
                    varFading(i) = 0
                End If
                
            ElseIf varFading(i) = 2 Then
                If snd.MUSIC_GetVolume(i) > 0 Then
                    Call snd.MUSIC_SetVolume(i, snd.MUSIC_GetVolume(i) - 1)
                Else
                    Call modAudio.snd.MUSIC_Stop(i)
                    varFading(i) = 0
                End If
            End If
            
        Next
        
        Call varTimer.Reset
    End If
End Sub

Private Sub Class_Initialize()
    Set varTimer = New Core.Timer
    Set varSamples = New Audio.MusicSampleList
End Sub

Private Sub Class_Terminate()
    Set varTimer = Nothing
    Set varSamples = Nothing
End Sub
