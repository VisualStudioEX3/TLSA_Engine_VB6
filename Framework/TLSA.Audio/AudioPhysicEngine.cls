VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AudioPhysicEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private varEmitters As Audio.AudioPhysicEmitterList                     ' Coleccion de emsiores de audio.
Private varEffectRegions As Audio.AudioPhysicEffectRegionList           ' Coleccion de zonas de efectos definidos que
                                                                        ' se aplicaran a los emisores.
Private varEnabled As Boolean
Private varListener As Core.POINT, varLastListener As Core.POINT    ' Posicion de escucha.

Private varNullRegion As Boolean                                        ' Indica al motor si debe desactivar en el proximo ciclo los efectos globales.
Private varCurrentRegion As String                                      ' Clave de la region actual.

' Lista de emisores de sonido:
Public Property Get Emitters() As Audio.AudioPhysicEmitterList
    Set Emitters = varEmitters
End Property

' Lista de regiones de efectos:
Public Property Get EffectRegions() As Audio.AudioPhysicEffectRegionList
    Set EffectRegions = varEffectRegions
End Property

' Posicion del oyente. Define la posicion con la cual se realizaran los calculos:
Public Property Get Listener() As Core.POINT
    Listener = varListener
End Property

Public Property Let Listener(value As Core.POINT)
    varLastListener = value
    varListener = value
End Property

Public Sub SetListener(X As Long, Y As Long)
    varLastListener = varListener
    varListener.X = X
    varListener.Y = Y
End Sub

' Indica si el motor de fisicas de audio esta disponible y activado:
Public Property Get Enabled() As Boolean
    Enabled = varEnabled
End Property

Public Property Let Enabled(value As Boolean)
    varEnabled = value
End Property

' Actualiza los parametros de los canales de audio configurados para aplicar fisica:
Friend Sub Update()
    If varEnabled Then
        ' Actualizamos el volumen y balance de cada emisor activo:
        Dim e As Audio.AudioPhysicEmitter
        For Each e In varEmitters
            If e.Enabled Then
                Call SetProperties(e)
                e.Enabled = e.Channel.Active
            End If
        Next
        
        ' Aplicamos los efectos del area donde se encuentre el oyente:
        Dim r As Audio.AudioPhysicEffectRegion
        For Each r In varEffectRegions
            If Core.Math.PointInRect(varListener, r.Region) Then
                ' Si no se trata de la misma region aplicamos los efectos:
                If Not varCurrentRegion = r.Key Then
                    varCurrentRegion = r.Key                ' Guardamos la clave de la region que vamos aplicar.
                    modAudio.fxGlobalEffects = r.Effects    ' Asigna los efectos de la region a la variable global.
                    modAudio.fxUpdate = True                ' Todos los canales en reproduccion actualizaran sus efectos.
                    varNullRegion = False
                End If
                Exit Sub                                    ' Se aplica los efectos y terminamos el proceso de actualizacion.
            End If
        Next
        
        If Not varNullRegion Then
            varCurrentRegion = ""
            
            ' Si no se encuentra en ninguna region se desactivan todos los efectos:
            With modAudio.fxGlobalEffects
                .Chorus = False
                .Compressor = False
                .Distortion = False
                .Echo = False
                .Flanger = False
                .Gargle = False
                .WavesReverb = False
            End With
            
            modAudio.fxUpdate = True
            
            varNullRegion = True                        ' Evita que al proximo ciclo si no se activo ningun efecto se
                                                        ' vuelva a realizar la operacion de desactivacion de efectos.
        End If
    End If
End Sub

' Metodo que calcula los parametros de volumen y balance del sonido de un emisor respecto al oyente:
Private Sub SetProperties(Emitter As Audio.AudioPhysicEmitter)
    If Not Core.Generics.POINTCompare(varListener, varLastListener) Then
        Dim varDist As Long, varValue As Long, varBalance As Long
        
        With Emitter
            ' Obtenemos la distancia absoluta al emisor de sonido:
            varDist = Core.Math.GetDistance(varListener, .Location)
            
            ' Obtenemos el factor en escala de 0 a 100 para parametrizar el volumen y balance:
            varValue = Core.Math.GetPercentFromValue(varDist, Emitter.Radius)
            
            ' Configuramos el volumen:
            Call .Channel.SetManagedVolume(Core.Math.GetPercentFromValue(Emitter.Radius - varValue, Emitter.Radius))
            
            ' Configuramos el balance:
            ' Si el emisor esta en direccion hacia la izquierda se convierte a negativo el valor:
            If .Location.X > varListener.X Then varValue = -varValue
            Call .Channel.SetManagedBalance(varValue)
        End With
    End If
End Sub

Private Sub Class_Initialize()
    Set varEmitters = New Audio.AudioPhysicEmitterList
    Set varEffectRegions = New Audio.AudioPhysicEffectRegionList
End Sub
