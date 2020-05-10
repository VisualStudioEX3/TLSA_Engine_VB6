VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ENG_Tile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private varKey As String
Private varSprite As Graphics.Sprite
Private varRectangle As Core.Rectangle

Public Property Get Key() As String
    Key = varKey
End Property

Public Property Get Sprite() As Graphics.Sprite
    Set Sprite = varSprite
End Property

Public Property Set Sprite(value As Graphics.Sprite)
    Set varSprite = value
    varRectangle = Core.Generics.CreateRECTANGLE2(varSprite.Location, varSprite.Size)
End Property

' Area que ocupa el tile en la escena:
Public Property Get Rectangle() As Core.Rectangle
    Rectangle = varRectangle
End Property

' Devuelve el rectangulo con las coordenadas de escena a coordenadas de ventana:
Public Function ClientToScreenRectangle() As Core.Rectangle
    ClientToScreenRectangle = Core.Generics.CreateRECTANGLE2(Engine.Scene.ScreenToClient(varSprite.Location), varSprite.Size)
End Function

' Devuelve una cadena de texto con la informacion del tile:
Public Function ToString() As String
    With varSprite
        Dim animKey As String
        If .Animations.Count > 0 Then animKey = .CurrentAnimation.Key
        
        ' ToString = Clave Textura; Clave Tile; EstadoAnimacion; Clave Animacion; X; Y; Z; Angulo; Color; Efecto; Espejado
        ToString = .Texture.Key & ";" & .CurrentTile.Key & ";" & .EnabledAnimation & ";" & animKey & ";" & .Location.X & ";" & .Location.Y & ";" & .Location.Z & ";" & .Angle & ";" & .Color & ";" & .Effect & ";" & .Mirror
    End With
End Function

Private Sub Class_Initialize()
    varKey = Core.Guid.Generate ' Se autogenera la clave unica del objeto.
    Set varSprite = New Graphics.Sprite
End Sub

Private Sub Class_Terminate()
    Set varSprite = Nothing
End Sub
