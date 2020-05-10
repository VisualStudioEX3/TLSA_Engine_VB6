VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VolumeControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Volumen de los efectos:
Public Property Get Effects() As Long
    Effects = modAudio.volEffects
End Property

Public Property Let Effects(value As Long)
    modAudio.volEffects = value
    modAudio.volUpdate = True
End Property

' Volumen de las voces:
Public Property Get Voice() As Long
    Voice = modAudio.volVoice
End Property

Public Property Let Voice(value As Long)
    modAudio.volVoice = value
    modAudio.volUpdate = True
End Property

' Volumen de los elementos de los menus:
Public Property Get GUI() As Long
    GUI = modAudio.volGUI
End Property

Public Property Let GUI(value As Long)
    modAudio.volGUI = value
    modAudio.volUpdate = True
End Property

' Volumen de los efectos de ambiente:
Public Property Get Ambient() As Long
    Ambient = modAudio.volAmbient
End Property

Public Property Let Ambient(value As Long)
    modAudio.volAmbient = value
    modAudio.volUpdate = True
End Property

' Volumen de los efectos de musica:
Public Property Get Music() As Long
    Music = modAudio.volMusic
End Property

Public Property Let Music(value As Long)
    modAudio.volMusic = value
    modAudio.volMusUpdate = True
End Property

' Actualiza los estados de volumen de todos los canales salvo los manuales o los que esten asociados al motor de fisica:
Friend Sub Update()
    If modAudio.volUpdate Then
        Dim i As Long
        For i = 0 To UBound(modAudio.sndChannels)
            With modAudio.sndChannels(i)
                If .Active Then
                    Select Case .Category
                        Case Audio.Audio_Types.Audio_Type_Ambient: .Volume = modAudio.volAmbient
                        Case Audio.Audio_Types.Audio_Type_Effect: .Volume = modAudio.volEffects
                        Case Audio.Audio_Types.Audio_Type_GUI: .Volume = modAudio.volGUI
                        Case Audio.Audio_Types.Audio_Type_Voice: .Volume = modAudio.volVoice
                    End Select
                End If
            End With
        Next
    End If
    
    If modAudio.volMusUpdate Then
        
    End If
End Sub
