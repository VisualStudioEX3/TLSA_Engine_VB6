VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FX_FireShoot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private varTextureFile As String
Private varSpriteAnim As String

Private varTexture As Graphics.Texture

Private fire As Graphics.Sprite

Private varEnabled As Boolean
Private varVisible As Boolean
Private varScale As Single

Private Sub Class_Initialize()
    varTextureFile = App.Path & ResourcePaths.Effects & "fire_shoot.png"
    varSpriteAnim = App.Path & ResourcePaths.Effects & "fire_shoot.ani"
    
    ' Cargamos la textura, si no esta cargada ya, y la asignamos:
    If Not GraphicEngine.Textures.Exists("fire_shoot") Then
        Set varTexture = GraphicEngine.Textures.LoadTexture(varTextureFile, "fire_shoot", False)
    Else
        Set varTexture = GraphicEngine.Textures("fire_shoot")
    End If
    
    ' Inicializamos los canales:
    Set fire = GraphicEngine.Helper.CreateAnimatedSprite(varTexture, varSpriteAnim) ': fire.EnabledAnimation = False: Call Head.SetCurrentTile("head1")
    'fire.Effect = Aditive
    
    varEnabled = True
    varVisible = True
End Sub

Private Sub Class_Terminate()
    Set fire = Nothing
    
    On Error Resume Next ' Evitamos la excepcion en caso de que la textura siga en uso:
    Call Engine.GraphicEngine.Textures.Unload("fire_shoot")
End Sub

Public Property Get Location() As Core.Point
    Location = fire.Location
End Property

Public Property Let Location(value As Core.Point)
    fire.Location = value
End Property

Public Property Get Angle() As Single
    Angle = fire.Angle
End Property

Public Property Let Angle(value As Single)
    fire.Angle = value
End Property

Public Property Get Mirror() As Graphics.Mirrors
    Mirror = fire.Mirror
End Property

Public Property Let Mirror(value As Graphics.Mirrors)
    fire.Mirror = value
End Property

Public Property Get Enabled() As Boolean
    Enabled = varEnabled
End Property

Public Property Let Enabled(value As Boolean)
    varEnabled = value
End Property

Public Property Get Visible() As Boolean
    Visible = varVisible
End Property

Public Property Let Visible(value As Boolean)
    varVisible = value
End Property

Public Property Get Looping() As Boolean
    Looping = fire.CurrentAnimation.Looping
End Property

Public Property Let Looping(value As Boolean)
    fire.CurrentAnimation.Looping = value
End Property

Public Property Get Color() As Long
    Color = fire.Color
End Property

Public Property Let Color(value As Long)
    fire.Color = value
End Property

Public Sub Draw()
    If varEnabled And varVisible Then Call fire.Draw
End Sub

Public Sub Update()
    If varEnabled Then Call fire.Update
End Sub

Public Sub SetScale(value As Single)
    varScale = value
    Call fire.SetScale(value)
End Sub

Public Function GetScale() As Single
    GetScale = varScale
End Function

