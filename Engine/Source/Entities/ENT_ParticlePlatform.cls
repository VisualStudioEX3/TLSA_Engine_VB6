VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ENT_ParticlePlatform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Tile As Graphics.Sprite                 ' Tile que representara la plataforma.
Private Shape As Graphics.Sprite                ' Figura que mostrara la plataforma.

Public Enabled As Boolean
Public Visible As Boolean

Private varTextureFile As String
Private varSpriteAnim As String

Private varTexture As Graphics.Texture

Private Platform As Graphics.Sprite
Private Symbol As Graphics.Sprite

Private WithEvents BBox As Physics.Body

Public ParticleType As Long

Public Property Get Location() As Core.Point
    Location = Core.Generics.VECTOR2POINT(BBox.Location)
End Property

Public Property Let Location(value As Core.Point)
    BBox.Location = Core.Generics.POINT2VECTOR(value)
End Property

Public Sub Draw()
    If Visible Then
        Call Platform.Draw
        Call Symbol.Draw
    End If
End Sub

Public Sub Update()
    
End Sub

Private Sub BBox_OnCollision(BodyCount As Long, E() As Physics.CollisionEventData)
    '
End Sub

Private Sub BBox_OnHit(Force As Single, eventData As Variant)
    '
End Sub
