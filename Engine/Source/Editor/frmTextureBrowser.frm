VERSION 5.00
Begin VB.Form frmTextureBrowser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Biblioteca de texturas"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7980
   Icon            =   "frmTextureBrowser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   312
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   532
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSelection 
      Caption         =   "Seleccionar"
      Enabled         =   0   'False
      Height          =   435
      Left            =   3960
      TabIndex        =   3
      Top             =   4200
      Width           =   3975
   End
   Begin VB.Frame Frame3 
      Caption         =   "Previsualizacion"
      ForeColor       =   &H8000000D&
      Height          =   4155
      Left            =   3930
      TabIndex        =   2
      Top             =   0
      Width           =   4005
      Begin VB.Image imgPreview 
         BorderStyle     =   1  'Fixed Single
         Height          =   3855
         Left            =   90
         Stretch         =   -1  'True
         Top             =   210
         Width           =   3825
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Texturas"
      ForeColor       =   &H8000000D&
      Height          =   4635
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   3885
      Begin VB.ListBox lstTextures 
         Height          =   4335
         Left            =   90
         Style           =   1  'Checkbox
         TabIndex        =   1
         Top             =   210
         Width           =   3705
      End
   End
End
Attribute VB_Name = "frmTextureBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSelection_Click()
    Call Engine.Scene.LevelEditor.TileEditor.SetTexture(Engine.GraphicEngine.Textures(Me.lstTextures.Text))
    Call Unload(Me)
    Call Unload(frmTileExplorer)
    Call frmTileExplorer.Show
End Sub

Private Sub Form_Load()
    Call LoadTextures       ' Carga la lista de texturas de la biblioteca.
    Call SelectLoaded       ' Selecciona en la lista las que esten en memoria.
    Call lstTextures_Click  ' Previsualizamos el primer elemento de la lista.
End Sub

' Carga en la lista los nombres de todas las texturas de la biblioteca:
Private Sub LoadTextures()
    Dim list() As String: list = Core.IO.GetFiles(App.Path & ResourcePaths.Textures & "*.png", Archive, NotSorted)
    Call Me.lstTextures.Clear
    On Error GoTo ErrOut
    Dim tex As Variant: For Each tex In list
        ' Si el archivo PNG viene acompañado de su archivo TEX se carga en la lista:
        If Core.IO.FileExists(App.Path & ResourcePaths.Textures & VBA.Left(tex, Len(tex) - 4) & ".tex") Then _
            Call Me.lstTextures.AddItem(CStr(tex))
    Next
ErrOut:
End Sub

' Selecciona en la lista las texturas que ya estan cargadas:
Private Sub SelectLoaded()
    ' Recorremos la lista y preguntamos a la coleccion de texturas del motor si existe la textura:
    Dim i As Long
    For i = 0 To Me.lstTextures.ListCount - 1
        Me.lstTextures.Selected(i) = Engine.GraphicEngine.Textures.Exists(Me.lstTextures.list(i))
    Next
    
    Me.cmdSelection.Enabled = Me.lstTextures.SelCount > 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Engine.Scene.LevelEditor.TileEditor.Texture Is Nothing Then Call Unload(frmTileExplorer)
End Sub

' Carga o descarga una textura de la lista:
Private Sub lstTextures_Click()
    ' Si el item esta seleccionado cargamos la textura:
    If Me.lstTextures.Selected(Me.lstTextures.ListIndex) Then
        ' Cargamos la textura en la instancia del motor del juego:
        If Not Engine.GraphicEngine.Textures.Exists(Me.lstTextures.Text) Then
            ' Cargamos la textura en el motor:
            Call Engine.GraphicEngine.Textures.LoadTexture(App.Path & TLSA.ResourcePaths.Textures & Me.lstTextures.Text, Me.lstTextures.Text, False)
            ' Agregamos su clave a la lista de texturas del nivel:
            Call Engine.Scene.Textures.Add(Me.lstTextures.Text, Me.lstTextures.Text)
        End If
        Call PreviewTexture         ' Cargamos la previsualizacion.
    
    ' Si no esta seleccionado la descargamos:
    Else
        On Local Error Resume Next
        ' Intentamos eliminar la textura de la biblioteca:
        If Engine.GraphicEngine.Textures.Exists(Me.lstTextures.Text) Then _
            Call Engine.GraphicEngine.Textures.Unload(Me.lstTextures.Text)
        
        If Err.Number <> 0 Then
            Call MsgBox(Err.Description, vbCritical, "Error al eliminar textura de la biblioteca")
            
        ' Si se elimino satisfactoriamente se selecciona el primer elemento de la lista y se previsualiza:
        Else
            ' Eliminamos la referencia a la textura en la lista de texturas de la escena:
            Call Engine.Scene.Textures.Remove(Me.lstTextures.Text)
            
            ' Si no hay texturas seleccionadas o la textura eliminada es la seleccionada actualmente en el editor se
            ' elimina la instancia del editor:
            If Me.lstTextures.SelCount = 0 Or Engine.Scene.LevelEditor.TileEditor.Texture.Key = Me.lstTextures.Text Then
                Call Engine.Scene.LevelEditor.TileEditor.SetTexture(Nothing)
                'Me.lstTextures.ListIndex = 0
            End If
            
'            Me.lstTextures.ListIndex = 0
            Call PreviewTexture     ' Cargamos la previsualizacion.
        End If
    End If
    
    Me.cmdSelection.Enabled = Me.lstTextures.SelCount > 0
End Sub

'Private Sub lstTextures_Click()
'    Call PreviewTexture
'End Sub

' Carga la previsualizacion de la textura:
Private Sub PreviewTexture()
    Dim filename As String: filename = Core.IO.CreateTemporalFilename("tmp")
    Dim varSurface As Graphics.Surface
    
    ' Si esta cargada la textura se genera la superficie:
    If Engine.GraphicEngine.Textures.Exists(Me.lstTextures.Text) Then
        Set varSurface = Engine.GraphicEngine.Textures(Me.lstTextures.Text).ToSurface()
    Else
        ' Cargamos temporalmente la textura y generamos la superficie:
        Call Engine.GraphicEngine.Textures.LoadTexture(App.Path & ResourcePaths.Textures & Me.lstTextures.Text, Me.lstTextures.Text, False)
        Set varSurface = Engine.GraphicEngine.Textures(Me.lstTextures.Text).ToSurface()
        Call Engine.GraphicEngine.Textures.Unload(Me.lstTextures.Text)
    End If
        
    Call varSurface.Save(filename)                                  ' Generamos el BMP.
    Set Me.imgPreview.Picture = LoadPicture(filename)               ' Cargamos la previsualizacion.
    Call Kill(filename)                                             ' Eliminamos el archivo temporal.
    Call Engine.GraphicEngine.Surfaces.Unload(Me.lstTextures.Text)  ' Eliminamos la superficie temporal.
    Set varSurface = Nothing
End Sub

'' Seleccionar:
'Private Sub lstTextures_DblClick()
'    Call cmdSelection_Click
'End Sub
