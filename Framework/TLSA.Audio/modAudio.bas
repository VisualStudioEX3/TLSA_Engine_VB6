Attribute VB_Name = "modAudio"
Option Explicit

Public snd As dx_Sound_Class                        ' Instancia de dx_sound_class.

Public sndChannels(63) As New Audio.SoundChannel    ' Lista de canales de muestras de sonido.

Public fxGlobalEffects As Audio.SoundEffects        ' Guarda la configuracion de efectos globales que se aplicaran a efectos de sonido y voces.
Public fxUpdate As Boolean                          ' Indica si se ha actualizar los estados de efectos en los canales de en reproduccion.

Public volEffects As Long, volVoice As Long, volGUI As Long, volAmbient As Long, volMusic As Long ' Parametros de volumenes por caterogia.
Public volUpdate As Boolean, volMusUpdate As Boolean    ' Indica si se ha de actualizar los estados de volumen en los canales en reproduccion.

Public spdGlobalSpeed As Long   ' Velocidad global de en los canales de audio.
Public spdUpdate As Boolean     ' Indica si se ha de actualizar los estados de velocidad en los canales en reproduccion.
