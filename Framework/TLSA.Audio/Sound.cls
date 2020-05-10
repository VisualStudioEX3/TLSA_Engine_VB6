VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SoundSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private varKey As String
Private varId As Long
Private varFilename As String
Private varSamplesPerSecond As Long

Public Property Get Key() As String
    Key = varKey
End Property

Friend Property Let Key(value As String)
    varKey = value
End Property

Friend Property Get Id() As Long
    Id = varId
End Property

Friend Property Let Id(value As Long)
    varId = value
End Property

Public Property Get Filename() As String
    Filename = varFilename
End Property

Friend Property Let Filename(value As String)
    varFilename = value
End Property

Public Property Get SamplesPerSecond() As Long
    SamplesPerSecond = varSamplesPerSecond
End Property

Friend Sub LoadSample(Filename As String)
    varFilename = Filename
    varId = modAudio.snd.SOUND_Load(varFilename)
    If varId = dxlib32_221.SND_ErrorCodes.SND_FILENOTLOAD Then
        If Not Core.IO.FileExists(Filename) Then
            Call Err.Raise(Audio.AUDIO_EXCEPTION.RESOURCE_NOT_FOUND, , "No se encontro la muestra de sonido:" & vbNewLine & "'" & Filename & "'")
        Else
            Call Err.Raise(Audio.AUDIO_EXCEPTION.RESOURCE_NOT_LOAD, , "Error al cargar la muestra de sonido:" & vbNewLine & "'" & Filename & "'")
        End If
    End If
    varSamplesPerSecond = snd.SOUND_GetSamplesPerSecond(varId)
End Sub

Private Sub Class_Initialize()
    varId = -1
End Sub

Private Sub Class_Terminate()
    If Not modAudio.snd Is Nothing Then Call modAudio.snd.SOUND_Unload(varId)
End Sub

' Reproduce la muestra de sonido y devuelve la instancia del canal donde se reproduce:
Public Function Play(Looping As Boolean, Category As Audio.Audio_Types, Optional ChannelIndex As Long = -1) As Audio.SoundChannel
    ' ChannelIndex permite especificar un canal concreto para reproducir la muestra. Si el canal elegido esta protegido se producira una excepcion.
    ' Canales protegidos o reservados serian los que usen el flag Audio_Type_Voice por ejemplo o los que se configuren manualmente para ello.
    Dim ch As Long
    ch = snd.SOUND_Play(varId, ChannelIndex, Looping)
    
    If ch = dxlib32_221.SND_ErrorCodes.SND_CHANNELUNAVAILABLE Then
        Exit Function
    ElseIf ch = dxlib32_221.SND_ErrorCodes.SND_UNKNOWNERROR Then
        Call Err.Raise(Audio.AUDIO_EXCEPTION.UNKNOWN_EXCEPTION, , "Error al intentar reproducir muestra de sonido.")
    End If
    
    With modAudio.sndChannels(ch)
        Set .Sample = Me
        .Category = Category
        If Not .Category = Audio_Type_GUI And Not .Category = Audio_Type_Manual Then
            .Effects = modAudio.fxGlobalEffects
            .Speed = modAudio.spdGlobalSpeed
        End If
    End With
    Set Play = modAudio.sndChannels(ch)
End Function
