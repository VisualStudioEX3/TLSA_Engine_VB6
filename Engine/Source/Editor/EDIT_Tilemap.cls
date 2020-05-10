VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EDIT_Tilemap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private varTex As Graphics.Texture
Private varSpr As Graphics.Sprite

Public ScaleFactor As Single
Public Color As Long
Public Angle As Single
Public Effect As Graphics.Effects
Public Mirror As Graphics.Mirrors
Public Layer As Long                ' Valor de Z para representar la capa de dibujo.

Public Property Get Texture() As Graphics.Texture
    Set Texture = varTex
End Property

Property Get Sprite() As Graphics.Sprite
    Set Sprite = varSpr
End Property

Public Sub SetTexture(value As Graphics.Texture)
    Set varTex = value
   
    If varTex Is Nothing Then
        Set varSpr = Nothing
    Else
        ' Utilizamos el constructor de ayuda para crear el sprite y si tiene animacion asociada cargarla tambien:
        Set varSpr = Engine.GraphicEngine.Helper.CreateAnimatedSprite(varTex)
        varSpr.Fixed = True
        
        ' Si el sprite no esta animado seleccionamos el primer tile definido que no sea "Default":
        If varSpr.Animations.Count = 0 Then
            If varTex.Tiles.Count > 0 Then
                Call SetTileFromIndex(2)
            Else
                Call MsgBox("La textura no contiene tiles definidos. Se utilizara la textura entera como tile por defecto.", vbExclamation, "No hay tiles definidos")
            End If
        End If
    End If
End Sub

Public Property Get Tile() As Graphics.Tile
    Set Tile = varSpr.CurrentTile
End Property

Public Sub SetTile(Key As String)
    If varSpr.Animations.Count = 0 Then
        ' Asignamos el tile al sprite:
        Call varSpr.SetCurrentTile(Key)
        Call varSpr.SetOffSet(0, 0)     ' Forzamos el offset a 0,0.
        
        ' Generamos la previsualizacion:
        Dim varSurf As Graphics.Surface
        Set varSurf = varSpr.CurrentTile.ToSurface()
        Dim filename As String: filename = Core.IO.CreateTemporalFilename("tmp")
        Call varSurf.Save(filename)
        
        ' Cargamos el BMP en el control Image:
        Set frmTileEditor.imgPreview = LoadPicture(filename)
        
        ' Descargamos la superficie temporal de memoria:
        Call Engine.GraphicEngine.Surfaces.Unload(varSurf.Key)
        Set varSurf = Nothing
    End If
End Sub

Public Sub SetTileFromIndex(Index As Long)
    Call SetTile(varTex.Tiles(Index).Key)
End Sub

Public Sub Draw()
    If Not varTex Is Nothing Then Call varSpr.Draw
End Sub

Public Sub Update()
    If Not varTex Is Nothing Then
        Call varSpr.setLocation(Engine.Scene.LevelEditor.Brush.Location.X, Engine.Scene.LevelEditor.Brush.Location.Y, Me.Layer)
        Call varSpr.SetScale(ScaleFactor)
        varSpr.Color = Color
        varSpr.Angle = Angle
        varSpr.Effect = Effect
        varSpr.Mirror = Mirror
    End If
End Sub

' Añade un tile a la escena:
Public Sub Add()
    If Not varSpr Is Nothing Then
        ' Convertimos las coordenadas del pincel en coordenadas de escena:
        varSpr.Location = Engine.Scene.ClientToScreen(varSpr.Location)
        Call Engine.Scene.TileMap.Add(varSpr)
    End If
End Sub

' Elimina un tile de la escena:
Public Sub Remove()
    ' El algorritmo de eliminacion se basa en obtener la lista de los tiles visibles en pantalla y buscar el primer tile
    ' cuya posicion del pincel/cursor este dentro de su area.
    Dim ite As TLSA.ENG_Tile
    For Each ite In Engine.Scene.TileMap.GetVisibles().Inverse()  ' Accedemos a la lista inversa para tomar en cuenta los ultimos elementos añadidos:
        If Core.Math.PointInRect(Engine.Scene.LevelEditor.Brush.Location, ite.ClientToScreenRectangle()) Then
            Call Engine.Scene.TileMap.Remove(ite)
            Exit Sub
        End If
    Next
End Sub

Private Sub Class_Initialize()
    frmTileEditor.Left = 0 'frmMainMenu.Left + frmMainMenu.Width
    frmTileEditor.Top = frmMainMenu.Top + frmMainMenu.Height
    ScaleFactor = 1
End Sub

Private Sub Class_Terminate()
    Call Unload(frmTileEditor)
End Sub
