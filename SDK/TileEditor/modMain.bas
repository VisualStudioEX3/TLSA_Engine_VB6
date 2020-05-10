Attribute VB_Name = "modMain"
Option Explicit

Public Const AppTitle As String = "TLSA SDK: Tile studio"

Public Gfx As Graphics.Manager                      ' Instancia del motor grafico.
Public Tex As Graphics.Texture                      ' Referencia a la textura cargada.
Public Info As Graphics.TEXTUREINFO                 ' Referencia a la informacion de la textura cargada.
Public Spr As Graphics.Sprite                       ' Sprite que dibujara la textura.

Public CurrentTile As Graphics.Tile                 ' Referencia al tile actual que se esta editando/visualizando.

Public Alpha As Graphics.Sprite                     ' Dibujamos el fondo para resaltar las zonas transparentes de la textura.
Public zoomScale As Single, lastScale As Single     ' Escala del zoom aplicado.

Public visorMouseCoord As Core.Point

Public selA As Core.Point, selAset As Boolean
Public selB As Core.Point
Public selC As Core.Point, selCset As Boolean
Public selP As Core.Point, selPset As Boolean
Public showControlPoint As Boolean                  ' Indica si la guia roja muestra la posicion del offset del tile o la posicion de un punto de control.

Public Declare Function GetActiveWindow Lib "user32" () As Long

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const SWP_NOMOVE = 2
Private Const SWP_NOSIZE = 1
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2

Public renderMode As Long ' Le indica al render el modo de visualizacion que tiene que mostrar.
Public showControlPointsInAnimationMode As Boolean ' Indica al render si debe dibujar los puntos de control en modo animacion.

Public Sub OnTop(f As Form)
    Call SetWindowPos(f.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
End Sub

Public Sub OffTop(f As Form)
    Call SetWindowPos(f.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
End Sub

' Configura el programa como visor de animacion:
Public Sub SetAnimationMode()
    lastScale = zoomScale
    frmMain.picMainPanel.Enabled = False
    frmMain.picToolBar.Enabled = False
    frmPointControl.Visible = False: frmPointControl.Enabled = False
    frmQuickAnimEd.Visible = False: frmQuickAnimEd.Enabled = False
    Call Gfx.SetDisplayMode(512, 512, 32, True)
    frmVisor.Left = 0: frmVisor.Top = 0
    Spr.EnabledAnimation = True
    Call Spr.SetLocation(256, 256, 0) 'Tex.Information.Texture.Width / 2, Tex.Information.Texture.Height / 2, 0)
    
    renderMode = 1
    
    ' Cargamos el panel de reproduccion:
    Call frmPrevCtrl.Show
End Sub

' Configura el programa como editor:
Public Sub SetEditorMode()
    frmMain.picMainPanel.Enabled = True
    frmMain.picToolBar.Enabled = True
    frmPointControl.Visible = True: frmPointControl.Enabled = True
    frmQuickAnimEd.Visible = True: frmQuickAnimEd.Enabled = True
    Call SetZoom(CLng(lastScale))
    Spr.EnabledAnimation = False
    Call Spr.SetLocation(0, 0, 0)
    Call Spr.SetCurrentTile("Default")
    renderMode = 0
    
    ' Seleccionamos el primer tile de la lista para resetear los valores del editor:
    Call modMain.ShowTileData(frmMain.lstTiles.List(0))
End Sub

' Dibuja la escena:
Public Sub Render()
    On Error GoTo ErrOut
    
    Call DrawAlphaBackground
    
    If renderMode = 1 Then Call Spr.Update ' Si estamos en modo animacion actualizamos el estado del sprite.
    If Not Spr Is Nothing Then Call Spr.Draw
    If renderMode = 1 And showControlPointsInAnimationMode Then Spr.CurrentTile.ControlPoints.Draw
    
    Call DrawGuides
    
    ' Si estamos en modo edicion dibujamos las guias y la seleccion actual:
    If renderMode = 0 Then Call DrawSelection
    
    Call Gfx.Render
    
    Exit Sub

ErrOut:
    If Not Err.Number = Graphics.GRAPHICS_EXCEPTION.DEVICE_LOST Then Call MsgBox(Err.Number & ": " & Err.Description, vbCritical, "Error en motor grafico")
End Sub

' Aplica la escala para definir un zoom sobre la textura para facilitar la seleccion de pixeles:
Public Sub SetZoom(value As Long)
    If Not Tex Is Nothing Then
        Dim newSize As Core.Size
        
        zoomScale = CSng(value)
        
        newSize.Width = Tex.Information.Texture.Width * value
        newSize.Height = Tex.Information.Texture.Height * value
        
        Call Gfx.SetDisplayMode(newSize.Width, newSize.Height, 32, True) ' Cambiamos el tamaño de la ventana.
        Call Spr.SetScale(zoomScale)                                     ' Aplicamos la escala al sprite de la textura.
        Call CenterVisor
    End If
End Sub

' Centra la ventana del visor en el espacio de trabajo de la ventana principal:
Public Sub CenterVisor()
    frmVisor.Left = 0 ' -((frmVisor.Width - (Gfx.CurrentDisplayMode.Width * Screen.TwipsPerPixelX)) \ 2) ' (frmMain.Width / 2) - (frmVisor.Width / 2)
    frmVisor.Top = 0 '-(frmVisor.Height - (Gfx.CurrentDisplayMode.Height * Screen.TwipsPerPixelY)) + ((frmVisor.Width - (Gfx.CurrentDisplayMode.Width * Screen.TwipsPerPixelX)) \ 2)  ' (frmMain.Height / 2) - (frmVisor.Height / 2)
End Sub

' Carga una textura:
Public Sub LoadTexture()
    Dim openFile As New Core.OpenDialog
    openFile.Title = "Seleccionar una textura"
    Call openFile.AddFilter("Formatos comunes de texturas", "*.bmp;*.tga;*.png")
        
    'If filename <> "" Then
    If openFile.Show Then
        If Gfx.Textures.Exists("tilemap") Then Call Gfx.Textures.Unload("tilemap") ' Si ya se cargo una textura previamente esta se descargara.
        Call Gfx.Textures.LoadTexture(openFile.FileName, "tilemap", False) ' Cargamos la textura.
        Set Tex = Gfx.Textures("tilemap") ' Hacemos referencia directa a la instancia de la textura en la coleccion.
        'Call Tex.ImportTiles ' Cargamos la lista de tiles si la hubiera.
        Call LoadTiles
        
        ' Creamos el sprite para dibujar la textura:
        Set Spr = New Graphics.Sprite
        Call Spr.SetTexture(Tex)
        Info = Tex.Information
        Call Gfx.SetDisplayMode(Info.Texture.Width, Info.Texture.Height, 32, False)
        
        ' Mostramos la informacion pertinente en la ventana:
        frmMain.txtFilename.Text = openFile.FileName
        frmMain.txtFilename.SelStart = Len(frmMain.txtFilename.Text)
        frmMain.lblInfo.Caption = "Tamaño: " & Info.Texture.Width & "x" & Info.Texture.Height
        Call CenterVisor
        
        Call SetZoom(CLng(zoomScale))
        
        frmMain.cmdSave.Enabled = True
        
        frmMain.Command1(0).Enabled = True
        
        ' Habilitamos las herramientas:
        frmMain.Command1(1).Enabled = True
        frmMain.Command1(2).Enabled = True
    End If
    Set openFile = Nothing
End Sub

' Guarda la lista de tiles de la textura:
Public Sub SaveTiles()
    If Not Tex Is Nothing Then Call Tex.ExportTiles
End Sub

' Dibuja un fondo para resaltar las zonas transparentes de la textura:
Private Sub DrawAlphaBackground()
    Dim X As Long, Y As Long
    Call Alpha.SetScale(CSng(zoomScale))
    For X = 0 To Gfx.CurrentDisplayMode.Width Step (16 * zoomScale)
        For Y = 0 To Gfx.CurrentDisplayMode.Height Step (16 * zoomScale)
            Call Alpha.SetLocation(X, Y, 8)
            Call Alpha.Draw
        Next
    Next
End Sub

' Dibuja las guias aplicando escala al pixel segun zoom aplicado:
Private Sub DrawGuides()
    Dim coord As Core.Point
    
    ' Guia principal:
    If selAset Then
        coord.X = selB.X * zoomScale
        coord.Y = selB.Y * zoomScale
    ElseIf Not selAset Then
        coord = visorMouseCoord
    End If
    
    If Not selCset Or Not selPset Then
        Call Gfx.Primitives.DrawBox3(0, coord.Y, Gfx.CurrentDisplayMode.Width, CLng(zoomScale), -8, &HFF000000, True, &HFF000000)
        Call Gfx.Primitives.DrawBox3(coord.X, 0, CLng(zoomScale), Gfx.CurrentDisplayMode.Height, -8, &HFF000000, True, &HFF000000)
    End If
    
    ' Centro del sprite:
    If Not selCset And Not selPset Then
        If Not showControlPoint Then
            coord.X = selC.X * zoomScale
            coord.Y = selC.Y * zoomScale
        Else
            coord.X = selP.X * zoomScale
            coord.Y = selP.Y * zoomScale
        End If
    ElseIf selCset Or selPset Then
        coord = visorMouseCoord
    End If
    
    Call Gfx.Primitives.DrawBox3(0, coord.Y, Gfx.CurrentDisplayMode.Width, CLng(zoomScale), -8, &HFFFF0000, True, &HFFFF0000)
    Call Gfx.Primitives.DrawBox3(coord.X, 0, CLng(zoomScale), Gfx.CurrentDisplayMode.Height, -8, &HFFFF0000, True, &HFFFF0000)
End Sub

' Dibuja la seleccion con un cuadro aplicando la escala al pixel segun zoom aplicado:
Public Sub DrawSelection()
    ' Aplicamos la escala a las coordenadas:
    Dim paS As Core.Point, pbS As Core.Point
    paS.X = selA.X * zoomScale: paS.Y = selA.Y * zoomScale
    pbS.X = selB.X * zoomScale: pbS.Y = selB.Y * zoomScale
    
    ' Resaltamos con una transparencia el area seleccionada:
    Call Gfx.Primitives.DrawBox2(paS, pbS, -7, Graphics.Color_Constant.Black, True, &H773300FF)
    
    ' Dibujamos el marco con escala de grosor:
    Call Gfx.Primitives.DrawBox(paS.X, paS.Y, pbS.X, (paS.Y) + CLng(zoomScale), -8, &HFF000000, True, &HFF000000)
    Call Gfx.Primitives.DrawBox(paS.X, paS.Y, (paS.X) + CLng(zoomScale), pbS.Y, -8, &HFF000000, True, &HFF000000)
    Call Gfx.Primitives.DrawBox(paS.X, pbS.Y, pbS.X, (pbS.Y) + CLng(zoomScale), -8, &HFF000000, True, &HFF000000)
    Call Gfx.Primitives.DrawBox(pbS.X, paS.Y, (pbS.X) + CLng(zoomScale), (pbS.Y) + CLng(zoomScale), -8, &HFF000000, True, &HFF000000)
End Sub

' Carga las claves de los tiles en el editor:
Public Sub LoadTiles()
    Dim t As Graphics.Tile
    
    Call frmMain.lstTiles.Clear
    
    For Each t In Tex.Tiles
        Call frmMain.lstTiles.AddItem(t.Key)
    Next
    
    If Tex.Tiles.Count > 0 Then frmMain.lstTiles.ListIndex = 0
End Sub

' Carga un mapa de puntos de control de un tile:
Public Sub LoadControlPointMap()
    Dim ite As Variant ' Iterador de la lista de claves.
    Dim keys() As String
    
    
    Call frmPointControl.lstPoints.Clear
    
    ' Esto evita que la posicion de los puntos de control se altere despues de hacer una previsualizacion de animacion al haberse
    ' aplicado transformaciones:
    Call CurrentTile.ControlPoints.ResetControlPoints
    
    If CurrentTile.ControlPoints.Count > 0 Then
        keys = CurrentTile.ControlPoints.GetKeyList()
        
        For Each ite In keys
            Call frmPointControl.lstPoints.AddItem(CStr(ite))
        Next
    Else
        frmPointControl.txtKey.Text = ""
        frmPointControl.txtX.Text = ""
        frmPointControl.txtY.Text = ""
    End If
    
    frmPointControl.lstPoints.ListIndex = 0
End Sub

' Muestra los datos del tile seleccionado:
Public Sub ShowTileData(Key As String)
    Set CurrentTile = Tex.Tiles(Key)
    
    With frmMain
        .txtKey.Text = CurrentTile.Key
        .txtKey.Tag = CurrentTile.Key ' Copia del original para la accion de cambio de clave.
        .txtX.Text = CurrentTile.Region.X
        .txtY.Text = CurrentTile.Region.Y
        .txtWidth.Text = CurrentTile.Region.Width
        .txtHeight.Text = CurrentTile.Region.Height
        .txtCenterX.Text = CurrentTile.Center.X
        .txtCenterY.Text = CurrentTile.Center.Y
    End With
    
    selA.X = CurrentTile.Region.X
    selA.Y = CurrentTile.Region.Y
    selB.X = CurrentTile.Region.Width + selA.X - 1
    selB.Y = CurrentTile.Region.Height + selA.Y - 1
    selC.X = CurrentTile.Center.X + selA.X
    selC.Y = CurrentTile.Center.Y + selA.Y
    
    Call modMain.LoadControlPointMap ' Cargamos el mapa de puntos de control.
End Sub

' Muestra los datos del tile seleccionado:
Public Sub ShowControlPointData(Key As String)
    With frmPointControl
        .txtKey.Text = Key
        .txtKey.Tag = Key ' Copia del original para la accion de cambio de clave.
        .txtX.Text = CurrentTile.ControlPoints(Key).X
        .txtY.Text = CurrentTile.ControlPoints(Key).Y
    End With
    
    selP.X = CurrentTile.ControlPoints(Key).X + selA.X
    selP.Y = CurrentTile.ControlPoints(Key).Y + selA.Y
End Sub

' Cambia la clave del tile:
Public Sub TileChangeKey(NewKey As String)
    ' Primero comprobamos que la clave no este en uso por otro tile:
    If Not Tex.Tiles.Exists(NewKey) Then
        ' Creamos el tile en base al actual pero usando la nueva clave:
        Call Tex.Tiles.Create(NewKey, CurrentTile.Texture, CurrentTile.Region.X, CurrentTile.Region.Y, CurrentTile.Region.Width, CurrentTile.Region.Height, CurrentTile.Center.X, CurrentTile.Center.Y)
        Call Tex.Tiles.Remove(CurrentTile.Key) ' Eliminamos el anterior.
        
        Call LoadTiles  ' Actualizamos la lista.
    Else
        Call MsgBox("La clave ya existe en la lista.", vbCritical, "Clave duplicada")
    End If
End Sub

' Cambia la clave del punto de control:
Public Sub ControlPointChangeKey(OldKey As String, NewKey As String)
    ' Primero comprobamos que la clave no este en uso por otro tile:
    If Not CurrentTile.ControlPoints.Exists(NewKey) Then
        ' Creamos el tile en base al actual pero usando la nueva clave:
        Call CurrentTile.ControlPoints.Add(NewKey, CurrentTile.ControlPoints(OldKey).X, CurrentTile.ControlPoints(OldKey).Y)
        Call CurrentTile.ControlPoints.Remove(OldKey) ' Eliminamos el anterior.
        
        Call LoadControlPointMap  ' Actualizamos la lista.
    Else
        Call MsgBox("La clave ya existe en la lista.", vbCritical, "Clave duplicada")
    End If
End Sub

' Copia la lista de puntos de control de un tile en otro sobreescribiendo los existentes:
Public Sub CopyFromTile(FromTile As String)
    Dim t As Graphics.Tile
    
    Set t = Tex.Tiles(FromTile)
    
    ' Comprobar si t tiene puntos de control definidos:
    If t.ControlPoints.Count > 1 Then
        If CurrentTile.ControlPoints.Count > 1 Then
            If MsgBox("El tile '" & CurrentTile.Key & "' contiene puntos de control definidos que seran sobreescritos por los del tile '" & FromTile & "'." & _
                      "¿Desea sobreescribir los puntos de control existentes?" _
                      , vbExclamation + vbYesNo, "¿Sobreescribir puntos de control existentes?") = VbMsgBoxResult.vbNo Then Exit Sub
        End If
        
        ' Copiamos los puntos de control de t en el tile actual:
        Call CurrentTile.SetControlPoints(t.ControlPoints)
    Else
        Call MsgBox("El tile '" & FromTile & "' no contiene puntos control.", vbCritical, "Error al copiar puntos de control")
    End If
End Sub
