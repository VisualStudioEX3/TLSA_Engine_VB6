VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Animation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum AnimationPath
    Foward                                          ' Reproduce la animacion de principio a final.
    Reverse                                         ' Reproduce la animacion en orden inverso.
End Enum

Private varKey As String                            ' Clave de la animacion en la coleccion.
Private varTexture As Graphics.Texture              ' Referencia a la textura que contiene los tiles.
Private varTiles As Graphics.TileList               ' Lista de los tiles de la secuencia animada.
Private varDelay As Long                            ' Demora entre fotogramas en milisegundos.

'Private varTimerId As Long
Private varTimer As Core.Timer
Private varCurrentIndex As Long                     ' Indice del tile actual en la lista.
Private varTileCount As Long
Private varAnimPath As AnimationPath                ' Define el sentido de la animacion.

Private varPaused As Boolean                        ' Indica si la animacion esta pausada.

Private varLooping As Boolean                       ' Indica si la animacion se ejecuta indefinidamente.
Private varEnded As Boolean                         ' Indica si la se llego al final de la secuencia de animacion.

' Devuelve el indice del tile actual:
Public Property Get CurrentTileIndex() As Long
    CurrentTileIndex = varCurrentIndex
End Property

Public Property Get IsAnimationEnded() As Boolean
    IsAnimationEnded = varEnded
End Property

Friend Property Let IsAnimationEnded(value As Boolean)
    varEnded = value
End Property

Public Property Get Looping() As Boolean
    Looping = varLooping
End Property

Public Property Let Looping(value As Boolean)
    varLooping = value
End Property

Public Property Get IsPaused() As Boolean
    IsPaused = varPaused
End Property

Public Property Get Key() As String
    Key = varKey
End Property

Friend Property Let Key(value As String)
    varKey = value
End Property

Public Property Get Texture() As Graphics.Texture
    Set Texture = varTexture
End Property

Public Sub SetTexture(value As Graphics.Texture)
    Set varTexture = value
    varTileCount = varTiles.Count ' varTexture.Tiles.Count
End Sub

Public Property Get Tiles() As Graphics.TileList
    Set Tiles = varTiles
End Property

Friend Property Set Tiles(value As Graphics.TileList)
    Set varTiles = value
End Property

' Tiempo de transicion entre un fotograma y el siguiente:
Public Property Get FrameDelay() As Long
    FrameDelay = varDelay
End Property

Public Property Let FrameDelay(value As Long)
    varDelay = value
End Property

Public Property Get AnimatePath() As AnimationPath
    AnimatePath = varAnimPath
End Property

Public Property Let AnimatePath(value As AnimationPath)
    varAnimPath = value
End Property

' Devuelve la instancia del tile que se esta dibujando actualmente:
Public Property Get CurrentTile() As Graphics.Tile
    Set CurrentTile = varTiles(varCurrentIndex)
End Property

' Establece el tile de la animacion:
Public Sub SetCurrentTile(Key As Variant)
    If VarType(Key) = vbString Then
        varCurrentIndex = varTiles.IndexOf(CStr(Key))
    Else
        varCurrentIndex = Key
    End If
End Sub

' Actualiza el fotograma de la animacion:
Public Sub Update()
    If Not varEnded Then
        If Not varPaused Then
'            If sys.TIMER_GetValue(varTimerId) >= varDelay Then
            If varTimer.value >= varDelay Then
                If varAnimPath = Foward Then
                    varCurrentIndex = varCurrentIndex + 1
                    If varCurrentIndex > varTileCount Then
                        If varLooping Then
                            varCurrentIndex = 1
                        Else
                            varCurrentIndex = varTileCount
                            varEnded = True
                        End If
                    End If
                Else
                    varCurrentIndex = varCurrentIndex - 1
                    If varCurrentIndex < 1 Then
                        If varLooping Then
                            varCurrentIndex = varTileCount
                        Else
                            varCurrentIndex = 1
                            varEnded = True
                        End If
                    End If
                End If
                'Call sys.TIMER_Reset(varTimerId)
                Call varTimer.Reset
            End If
        End If
    End If
End Sub

Private Sub Class_Initialize()
    Set varTiles = New Graphics.TileList
    'varTimerId = sys.TIMER_Create()
    Set varTimer = New Core.Timer
    varCurrentIndex = 1
End Sub

Private Sub Class_Terminate()
'    Call sys.TIMER_Kill(varTimerId)
    Set varTimer = Nothing
End Sub

' Inicia o continua la animacion:
Public Sub Play()
    varPaused = False
    ' Si la animacion termino se retoma al principio de la secuencia, segun orden de reproduccion:
    If varEnded Then
        varEnded = False
        If varAnimPath = Foward Then varCurrentIndex = 1 Else varCurrentIndex = varTileCount
    End If
End Sub

' Detiene la animacion en su tile o fotograma actual:
Public Sub Pause()
    varPaused = True
End Sub

' Reinicia la secuencia de animacion al primer tile o fotograma, segun orden de reproduccion:
Public Sub Reset()
    varEnded = False
    If Me.AnimatePath = Foward Then
        Call Me.SetCurrentTile(1)
    Else
        Call Me.SetCurrentTile(Me.Tiles.Count)
    End If
End Sub
