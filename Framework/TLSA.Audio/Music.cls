VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MusicSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private varKey As String
Private varId As Long
Private varFilename As String
Private varLength As Long

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

Public Property Get Length() As Long
    Length = varLength
End Property

Friend Sub LoadSample(Filename As String)
    varFilename = Filename
    varId = modAudio.snd.MUSIC_Load(varFilename)
    If varId = dxlib32_221.SND_ErrorCodes.SND_FILENOTLOAD Then
        If Not Core.IO.FileExists(Filename) Then
            Call Err.Raise(Audio.AUDIO_EXCEPTION.RESOURCE_NOT_FOUND, , "No se encontro la muestra de musica:" & vbNewLine & "'" & Filename & "'")
        Else
            Call Err.Raise(Audio.AUDIO_EXCEPTION.RESOURCE_NOT_LOAD, , "Error al cargar la muestra de musica:" & vbNewLine & "'" & Filename & "'")
        End If
    End If
    varLength = snd.MUSIC_GetDuration(varId)
End Sub

Private Sub Class_Initialize()
    varId = -1
End Sub

Private Sub Class_Terminate()
    If Not modAudio.snd Is Nothing Then Call modAudio.snd.MUSIC_Unload(varId)
End Sub

'' Reproduce la muestra de sonido y devuelve la instancia del canal donde se reproduce:
'Friend Function PlaySample(Looping As Boolean, Optional ChannelIndex As Long = 0) As Audio.MusicChannel
'    Dim ch As Long
'    ch = snd.MUSIC_Play(varId, ChannelIndex, Looping)
'    With modAudio.sndChannels(ch)
'        .Sample = Me
'        If Not .Category = Audio_Type_GUI And Not .Category = Audio_Type_Physic Then
'            .Effects = modAudio.fxGlobalEffects
'            .Speed = modAudio.spdGlobalSpeed
'        End If
'    End With
'End Function
'
