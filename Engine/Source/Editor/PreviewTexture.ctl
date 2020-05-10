VERSION 5.00
Begin VB.UserControl PreviewTexture 
   BackColor       =   &H00C0C000&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.PictureBox picPreview 
      BackColor       =   &H8000000C&
      Height          =   2175
      Left            =   0
      ScaleHeight     =   2115
      ScaleWidth      =   2235
      TabIndex        =   0
      Top             =   0
      Width           =   2295
      Begin VB.Timer timerRender 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   810
         Top             =   600
      End
   End
End
Attribute VB_Name = "PreviewTexture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Control para previsualizacion de texturas mediante TLSA.Graphics
' Diseñado para creacion de herramientas.
' * El tamaño de visualizacion estara fijado al valor establecido en la textura incluso a la hora de visualizar tiles.

Option Explicit

Private gfxEngine As Graphics.Manager
Private gfxTexture As Graphics.Texture
Private gfxPreview As Graphics.Sprite

Public Event Click()

Public Property Get Filename() As String
    If Not gfxTexture Is Nothing Then Filename = gfxTexture.Filename
End Property

Public Property Let Filename(value As String)
    Set gfxTexture = gfxEngine.Textures.LoadTexture(value, value, False)
    Set gfxPreview = New Graphics.Sprite
    Call gfxPreview.SetTexture(gfxTexture)
    Call gfxPreview.SetSize(256, 256)
    gfxPreview.Fixed = True
    timerRender.Enabled = True
End Property

' Devuelve una lista con todos los tiles de la textura:
Public Property Get Tiles() As Graphics.TileList
    Set Tiles = gfxPreview.Texture.Tiles
End Property

' Devuelve el tile actual:
Public Property Get CurrentTile() As Graphics.Tile
    If Not gfxPreview Is Nothing Then CurrentTile = gfxPreview.CurrentTile
End Property

' Establece el tile a mostrar:
Public Sub SetCurrentTile(Key As String)
    Call gfxPreview.SetCurrentTile(Key)
    Call gfxPreview.SetSize(256, 256)
End Sub

' Muestra toda la textura:
Public Sub SetEntireTexture()
    Call gfxPreview.SetEntireRegion
    Call gfxPreview.SetSize(256, 256)
End Sub

Public Property Get Color() As Long
    If Not gfxPreview Is Nothing Then Color = gfxPreview.Color
End Property

Public Property Let Color(value As Long)
    gfxPreview.Color = value
End Property

Private Sub picPreview_Click()
    RaiseEvent Click
End Sub

Private Sub timerRender_Timer()
    If Not gfxPreview Is Nothing Then
        Call gfxPreview.Draw
        Call gfxEngine.Render
    End If
End Sub

Private Sub UserControl_Initialize()
    Set gfxEngine = New Graphics.Manager
    Call gfxEngine.Initialize(picPreview.hWnd, picPreview.width, picPreview.height, 32, True, True)
    picPreview.Left = 0: picPreview.Top = 0 ': picPreview.width = 256: picPreview.height = 256
    'Call UserControl_Resize
End Sub

Private Sub UserControl_Resize()
    picPreview.width = width \ Screen.TwipsPerPixelX
    picPreview.height = height \ Screen.TwipsPerPixelY
    If Not gfxPreview Is Nothing Then
        Call gfxEngine.SetDisplayMode(picPreview.width, picPreview.height, 32, True)
        Call gfxPreview.SetSize(picPreview.width, picPreview.height)
    End If
End Sub

Private Sub UserControl_Terminate()
    timerRender.Enabled = False
    Set gfxEngine = Nothing
End Sub
