VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Light"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Fuente de luz: Define una fuente de iluminacion radial, no direccional, que puede aplicarse para crear efectos de iluminacion
' por vertice en sprites y tiles.

Option Explicit

Private varLocation As Core.POINT
Private varRadius As Long, varColor As Long, varIntensity As Long
Private varEnabled As Boolean

Public Property Get Radius() As Long
    Radius = varRadius
End Property

Public Property Let Radius(value As Long)
    If value >= 0 Then varRadius = value Else varRadius = 0
End Property

Public Property Get Color() As Long
    Color = varColor
End Property

Public Property Let Color(value As Long)
    varColor = value
End Property

Public Property Get Intensity() As Long
    Intensity = varIntensity
End Property

Public Property Let Intensity(value As Long)
    If value > 0 Then varIntensity = value Else varIntensity = 1
End Property

Public Property Get Enabled() As Boolean
    Enabled = varEnabled
End Property

Public Property Let Enabled(value As Boolean)
    varEnabled = value
End Property

Public Property Get Location() As Core.POINT
    Location = varLocation
End Property

Public Property Let Location(value As Core.POINT)
    varLocation = value
End Property

Public Sub SetLocatation(X As Long, Y As Long, Z As Long)
    varLocation.X = X
    varLocation.Y = Y
    varLocation.Z = Z
End Sub
