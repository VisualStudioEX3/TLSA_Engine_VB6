VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AudioPhysicEmitter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private varKey As String
Private varEnabled As Boolean
Private varLocation As Core.POINT
Private varRadius As Single
Private varSample As Audio.SoundSample
Private varChannel As Audio.SoundChannel

Public Property Get Key() As String
    Key = varKey
End Property

Friend Property Let Key(value As String)
    varKey = value
End Property

' Activa o desactiva el emisor:
Public Property Get Enabled() As Boolean
    Enabled = varEnabled
End Property

Public Property Let Enabled(value As Boolean)
    varEnabled = value
End Property

' Posicion del emisor:
Public Property Get Location() As Core.POINT
    Location = varLocation
End Property

Public Property Let Location(value As Core.POINT)
    varLocation = value
End Property

Public Sub SetLocation(X As Long, Y As Long)
    varLocation.X = X
    varLocation.Y = Y
End Sub

' Radio de accion del emisor:
Public Property Get Radius() As Single
    Radius = varRadius
End Property

Public Property Let Radius(value As Single)
    varRadius = value
End Property

' Devuelve o establece la muestra de sonido que se reproducira:
Public Property Get Sample() As Audio.SoundSample
    Set Sample = varSample
End Property

Public Property Set Sample(value As Audio.SoundSample)
    Set varSample = value
End Property

' Devuelve el canal usado por el emisor:
Public Property Get Channel() As Audio.SoundChannel
    Set Channel = varChannel
End Property

' Activa la reproduccion de sonido del emisor:
Public Sub Play(Optional Looping As Boolean, Optional Channel As Long = -1)
    Set varChannel = varSample.Play(Looping, Audio_Type_Physic, Channel)
    Call varChannel.SetManagedVolume(0)
    varEnabled = True
End Sub
