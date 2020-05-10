VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ENG_TileList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private varCol As Collection

' Añade un sprite y genera un tile en la lista:
Public Function Add(Sprite As Graphics.Sprite) As TLSA.ENG_Tile
    Dim t As New TLSA.ENG_Tile
    Set t.Sprite = Engine.GraphicEngine.Helper.CreateAnimatedSprite(Sprite.Texture)
    With t.Sprite
        Call .SetCurrentTile(Sprite.CurrentTile.Key)
        Call .SetScale(Engine.Scene.LevelEditor.TileEditor.ScaleFactor)
        .Angle = Sprite.Angle
        .Color = Sprite.Color
        .Location = Sprite.Location
        .Offset = Sprite.Offset
        .Effect = Sprite.Effect
        .Mirror = Sprite.Mirror
        .EnabledAnimation = Sprite.EnabledAnimation
    End With
   
    Call varCol.Add(t, t.Key)
    Set Add = t
    Set t = Nothing
End Function

' Añade un tile mediante referencia a la coleccion:
Public Sub AddRef(Tile As TLSA.ENG_Tile)
    Call varCol.Add(Tile, Tile.Key)
End Sub

' Elimina un tile de la escena:
Public Sub Remove(Tile As TLSA.ENG_Tile)
    Call varCol.Remove(Tile.Key)
End Sub

' Accede a un tile de la lista por su clave unica:
Public Property Get Item(Key As String) As TLSA.ENG_Tile
Attribute Item.VB_UserMemId = 0
    Set Item = varCol(Key)
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'esta propiedad permite enumerar
    'esta colección con la sintaxis For...Each
    Set NewEnum = varCol.[_NewEnum]
End Property

' Devuelve el numero de tiles en la lista:
Public Property Get Count() As Long
    Count = varCol.Count
End Property

' Devuelve una coleccion con todos los tiles visibles en pantalla:
Public Function GetVisibles() As TLSA.ENG_TileList
    Dim ite As TLSA.ENG_Tile, ret As New TLSA.ENG_TileList
    Dim scrRect As Core.Rectangle
    scrRect = Core.Generics.CreateRECTANGLE(0, 0, Engine.GraphicEngine.CurrentDisplayMode.Width, Engine.GraphicEngine.CurrentDisplayMode.Height)
    For Each ite In varCol
        If Core.Math.IntersectRect(ite.ClientToScreenRectangle(), scrRect) Then Call ret.AddRef(ite)
    Next
    Set GetVisibles = ret
End Function

' Devuelve la lista en sentido inverso:
Public Function Inverse() As TLSA.ENG_TileList
    Dim ite As TLSA.ENG_Tile, arr() As TLSA.ENG_Tile, i As Long, ret As New TLSA.ENG_TileList
    
    If varCol.Count > 0 Then
        ' Rellenamos el array con los valores de la coleccion:
        ReDim arr(varCol.Count - 1)
        For Each ite In varCol
            Set arr(i) = ite
            i = i + 1
        Next
        
        ' Rellenamos la coleccion en sentido inverso:
        For i = UBound(arr) To 0 Step -1
            Call ret.AddRef(arr(i))
        Next
    End If
    
    Set Inverse = ret
End Function

Public Sub Clear()
    Set varCol = Nothing
    Set varCol = New Collection
End Sub

Private Sub Class_Initialize()
    Set varCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set varCol = Nothing
End Sub
