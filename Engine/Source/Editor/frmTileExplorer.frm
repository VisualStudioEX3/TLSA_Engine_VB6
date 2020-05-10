VERSION 5.00
Begin VB.Form frmTileExplorer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Explorador de tiles"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4755
   ControlBox      =   0   'False
   DrawMode        =   10  'Mask Pen
   Icon            =   "frmTileExplorer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   234
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   317
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmTileExplorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private varTexture As Graphics.Texture, varSurface As Graphics.Surface
Private filename As String
Dim offsetX As Long, offsetY As Long

Private varSelTile As String

' Carga la textura de la biblioteca:
Private Sub SetTexture()
    ' Obtenemos acceso a la textura:
    Set varTexture = Engine.Scene.LevelEditor.TileEditor.Texture
    
    ' Guardamos a BMP la textura:
    filename = Core.IO.CreateTemporalFilename("tmp")
    Set varSurface = varTexture.ToSurface()
    Call varSurface.Save(filename)
    Call Engine.GraphicEngine.Surfaces.Unload(varSurface.Key)
    Set varSurface = Nothing
    
    ' Redimensionamos el formulario:
    Me.Width = (varTexture.Information.Texture.Width + offsetX) * Screen.TwipsPerPixelX
    Me.Height = (varTexture.Information.Texture.Height + offsetY) * Screen.TwipsPerPixelY
    
    ' Cargamos el BMP en el formulario:
    Set Me.Picture = LoadPicture(filename)
    Call Kill(filename)
End Sub

Private Sub Form_Load()
    ' Obtenemos las medidas de diferencia entre el area de cliente y el area completa de la ventana:
    Dim winRect As RECT, cliRect As RECT, finalRect As RECT
    
    Call GetWindowRect(Me.hwnd, winRect)
    Call GetClientRect(Me.hwnd, cliRect)
    Call SubtractRect(finalRect, winRect, cliRect)
    
    offsetX = (cliRect.Left - cliRect.Right) - (winRect.Left - winRect.Right)
    offsetY = (cliRect.Top - cliRect.Bottom) - (winRect.Top - winRect.Bottom)
    
    Me.AutoRedraw = True
    
    Call SetTexture
End Sub

' Evento de seleccion:
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Engine.Scene.LevelEditor.TileEditor.SetTile(varSelTile)
    'Call Unload(Me)
End Sub

' Muestra la seleccion del tile donde se encuentre el cursor del raton:
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Recorremos la lista de tiles el tile donde se encuentre el cursor del raton:
    Dim Tile As Graphics.Tile
    For Each Tile In varTexture.Tiles
        If Not Tile.Key = "Default" Then
            If Core.Math.PointInRect(Core.Generics.CreatePOINT(CLng(X), CLng(Y)), Tile.Region) Then
                Cls
                With Tile.Region
                    ' Dibuja un rectangulo mostrando el area del tile seleccionado invirtiendo los colores para el resalte:
                    ' "DrawMode" debe estar puesto a "10 - Not Xor Pen" en tiempo de diseño en las propiedades del formulario.
                    Line (.X, .Y)-(.X + .Width, .Y + .Height), , BF
                    
                    varSelTile = Tile.Key
                End With
            End If
        End If
    Next
End Sub
