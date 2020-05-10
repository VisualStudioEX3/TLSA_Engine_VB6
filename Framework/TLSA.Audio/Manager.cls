VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Manager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Lista de excepciones tipificadas de la libreria:
Public Enum AUDIO_EXCEPTION
    RESOURCE_NOT_LOAD = vbObjectError + 2000
    RESOURCE_NOT_FOUND = vbObjectError + 2001
    RESOURCE_REFERENCE_EXCEPTION = vbObjectError + 2003
    ITEM_KEY_DUPLICATED = 457 ' Error(457) Esta clave ya está asociada a un elemento de esta colección
    ITEM_KEY_NOT_FOUND = vbObjectError + 2004
    ARGUMENT_EXCEPTION = 5
    DEVICE_NOT_CREATE = 2005
    UNKNOWN_EXCEPTION = vbObjectError + 9999
End Enum

Public Enum Audio_Types
    Audio_Type_Manual   ' Se configura el volumen y efectos de forma manual.
    Audio_Type_Effect   ' Efectos de sonido (Disparos, chispas, etc...). Tiene su propio valor de volumen.
    Audio_Type_Voice    ' Voces de dialogos. Tiene su propio valor de volumen independientemente del resto de categorias.
    Audio_Type_GUI      ' Efectos de la interfaz de usuario. Tiene su propio valor de volumen independientemente del resto de
                        ' categorias.
    Audio_Type_Ambient  ' Sonidos de ambiente (pajaros, trafico, radio, etc...). Tiene su propio valor de volumen
                        ' independientemente del resto de categorias. La musica siempre se aplica como muestra ambiental.
    Audio_Type_Physic   ' Sonidos que se ajustaran a los estados del simulador de fisica asociado al motor de audio, aplicando
                        ' cambios de volumen y efectos segun distancia y obstaculos entre el foco de sonido y el oyente.
                        ' Si no hay asociada ninguna instancia de un simulador de fisicas el sonido se configurara automaticamente
                        ' como Audio_Type_Manual.
End Enum

Public Enum Audio_Constants
    Max_Volume = 100        ' Volumen maximo.
    Min_Volume = 0          ' Volumen minimo.
    Right_Balance = 100     ' Balance ajustado al canal derecho.
    Left_Balance = -100     ' Balance ajustado al canal izquierdo.
    Default_Balance = 0     ' Balance equilibrado en estereo.
    Max_Speed = 500         ' 500% de la velocidad real.
    Min_Speed = 25          ' 25% de la velocidad real.
    Default_Speed = 100     ' 100% de la velocidad real.
End Enum

Private varVolumeControl As Audio.VolumeControl
Private varSoundMixer As Audio.SoundMixer
Private varMusicMixer As Audio.MusicMixer

' Control de volumen:
Public Property Get VolumeControl() As Audio.VolumeControl
    Set VolumeControl = varVolumeControl
End Property

' Control del mezclador de sonido:
Public Property Get SoundMixer() As Audio.SoundMixer
    Set SoundMixer = varSoundMixer
End Property

' Control del mezclador de Musica:
Public Property Get MusicMixer() As Audio.MusicMixer
    Set MusicMixer = varMusicMixer
End Property

Private Sub Class_Initialize()
    '
End Sub

' Inicializa e instancia el motor de audio en la ventana designada:
Public Sub SetWindowHandle(hWnd As Long)
    Set snd = New dx_Sound_Class
    If snd.Init(hWnd, 64) Then
        Set varSoundMixer = New Audio.SoundMixer
        Set varMusicMixer = New Audio.MusicMixer
        Set varVolumeControl = New Audio.VolumeControl
        Exit Sub
    Else
        Call Err.Raise(Audio.AUDIO_EXCEPTION.DEVICE_NOT_CREATE, , "No se pudo crear la instancia del motor de audio.")
    End If
End Sub

Private Sub Class_Terminate()
    Set varSoundMixer = Nothing
    Set varMusicMixer = Nothing
    Set varVolumeControl = Nothing
    Call snd.Terminate
    Set snd = Nothing
End Sub

' Actualiza los estados del motor de audio:
Public Sub Update()
    Call varVolumeControl.Update
    Call varSoundMixer.Update
    Call varMusicMixer.Update
End Sub
