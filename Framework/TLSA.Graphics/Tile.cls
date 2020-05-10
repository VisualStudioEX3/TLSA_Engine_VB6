VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Tile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private varTex As Graphics.Texture
Private varRect As Core.RECTANGLE
Private varCenter As Core.Point
Private varKey As String
Private varPoints As New Graphics.ControlPointMap

Public Property Get Texture() As Graphics.Texture
    Set Texture = varTex
End Property

Public Property Get ControlPoints() As Graphics.ControlPointMap
    Set ControlPoints = varPoints
End Property

' Establece los puntos a partir de la instancia de un mapa existente:
Public Sub SetControlPoints(CopyFrom As Graphics.ControlPointMap)
    Set varPoints = CopyFrom.Clone()
    Call varPoints.SetSize(varRect.Width, varRect.Height)
    Call varPoints.SetOffSet(Me.Center.X, Me.Center.Y)
End Sub

Public Sub SetTexture(value As Graphics.Texture)
    Set varTex = value
End Sub

Public Property Get Region() As Core.RECTANGLE
    Region = varRect
End Property

Public Property Let Region(value As Core.RECTANGLE)
    varRect = value
    
    ' Cambiamos el tamaño del mapa de puntos de control:
    Call varPoints.SetFixedSize(value.Width, value.Height)
End Property

Public Sub SetRegion(X As Long, Y As Long, Width As Long, Height As Long)
    With varRect
        .X = X
        .Y = Y
        .Width = Width
        .Height = Height
    End With
    
    ' Cambiamos el tamaño del mapa de puntos de control:
    Call varPoints.SetFixedSize(Width, Height)
End Sub

Public Property Get Center() As Core.Point
    Center = varCenter
End Property

Public Property Let Center(value As Core.Point)
    varCenter = value
    Call varPoints.SetPoint("Default", value.X, value.Y)
End Property

Public Sub SetCenter(X As Long, Y As Long)
    With varCenter
        .X = X
        .Y = Y
    End With
    Call varPoints.SetPoint("Default", X, Y)
End Sub

Public Property Get Key() As String
    Key = varKey
End Property

Friend Property Let Key(value As String)
    varKey = value
End Property

' Crea una superficie con el contenido del tile:
Public Function ToSurface() As Graphics.Surface
    Set ToSurface = modGraphics.varSurfaces.CreateFromTexture(varTex, varRect)
End Function

