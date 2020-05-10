VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ENG_Scene"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Escena: Representa un nivel con todos sus elementos y su logica.

Option Explicit

Public Enabled As Boolean                       ' Indica si la escena actualizara su logica.
Public Visible As Boolean                       ' Indica si se dibujara la escena.

Public Textures As Collection                   ' Coleccion de biblioteca de texturas (clave, no referencia) de la escena.

Public Entities As Collection                   ' Coleccion de entidades.
Public TileMap As TLSA.ENG_TileList             ' Coleccion de tiles que forman el escenario.
Public PhysicMap As TLSA.ENG_BBoxList           ' Mapa de fisicas del escenario.

Public PhysicSimulator As Physics.Simulator     ' Simulador de fisicas de la escena.

Private varCamera As Core.Point
Private varCameraTarget As Object               ' Entidad que seguira la camara de la escena (el objeto ha de tener propiedad Location)
Private varCameraTargetKey As String            ' Clave de la entidad que seguira la camara.
Private varCameraTargetOffset As Core.Point     ' Permite alterar la posicion de la camara definida por su entidad objetivo.

Public LevelEditor As EDIT_Manager              ' Instancia de las funciones de edicion de niveles.

Private varRenderTargetEnabled As Boolean
Private varRenderTarget As Graphics.RenderTarget
Private varRenderScene As Graphics.Sprite

Public ShowPhysicBodies As Boolean

' Indica las capas que son visibles:
Public ShowTiles As Boolean
Public ShowBackLayer As Boolean, ShowMainLayer As Boolean, ShowFrontLayer As Boolean

Private Sub Class_Initialize()
    Set Textures = New Collection
    
    Set Entities = New Collection
    Set TileMap = New TLSA.ENG_TileList
    Set PhysicMap = New TLSA.ENG_BBoxList
    
    If Engine.PHYSICS_SEPARATE_THREAD Then
        Call PhysicThread.CreateThread
        Set PhysicSimulator = PhysicThread.ThreadPhysicSimulator
    Else
        Set PhysicSimulator = New Physics.Simulator
    End If
    
    PhysicSimulator.Enabled = False
    
    Me.ShowTiles = True
    Me.ShowBackLayer = True
    Me.ShowMainLayer = True
    Me.ShowFrontLayer = True
    
    Me.ShowPhysicBodies = True
    
    If Engine.EDIT_MODE Then Set LevelEditor = New EDIT_Manager
End Sub

Private Sub Class_Terminate()
    Set Entities = Nothing
    Set TileMap = Nothing
    Set PhysicMap = Nothing
    
    If Engine.PHYSICS_SEPARATE_THREAD Then Call PhysicThread.KillThread
    Set PhysicSimulator = Nothing
    
    Call DeactivateRenderTarget
    
    Set LevelEditor = Nothing
End Sub

' Elimina todo el contenido del nivel:
Public Sub Clear()
    Call Me.TileMap.Clear   ' Tiles.
    Call Me.PhysicMap.Clear ' Fisicas.
    Set Me.Entities = New Collection            ' /// implementar codigo de gestion como en las listas anteriores ///
    
    ' Texturas:
    Dim tex As Variant
    For Each tex In Me.Textures
        Call Engine.GraphicEngine.Textures.Unload(CStr(tex))
    Next
End Sub

' --------------------------------------------------------------------------------
' Funciones del render target para aplicar efectos y transformaciones a la escena
' --------------------------------------------------------------------------------

' Activa el render target para aplicar transformaciones a la escena:
Public Sub ActivateRenderTarget()
    If Not varRenderTargetEnabled Then
        ' Creamos el render target:
        Set varRenderTarget = New Graphics.RenderTarget
        Call varRenderTarget.Initialize(1024, 1024, False)
        
        ' Configuramos el sprite que representa la escena:
        Set varRenderScene = New Graphics.Sprite
        With varRenderScene
            Call .SetTexture(varRenderTarget.GetTexture())    ' Obtenemos la textura del render target y la asignamos al sprite que mostrara la escena.
            Call .setLocation(400, 300, 0)                    ' Posicionamos la escena en las coordenadas centrales de la pantalla.
            Call .SetOffSet(400, 300)                         ' Posicionamos el centro de rotacion en el centro absoluto de la pantalla.
            Call .SetRegion(0, 0, 800, 600)                   ' Seleccionamos el area que corresponde a la resolucion de pantalla.
            .Fixed = True                                     ' Fijamos la escena para que no le afecte la camara.
        End With
        
        varRenderTargetEnabled = True
    End If
End Sub

' Desactiva el render target:
Public Sub DeactivateRenderTarget()
    If varRenderTargetEnabled Then
        Set varRenderTarget = Nothing
        Set varRenderScene = Nothing
        varRenderTargetEnabled = False
    End If
End Sub

' Indica si esta activado el render target para efectos sobre escena:
Public Property Get RenderTargetEnabled() As Boolean
    RenderTargetEnabled = varRenderTargetEnabled
End Property

' Devuelve la instancia del sprite que contiene el resultado del render target:
Public Property Get RenderTarget() As Graphics.Sprite
    If varRenderTargetEnabled Then Set RenderTarget = varRenderScene
End Property

' Angulo de la escena:
Public Property Get Angle() As Single
    If varRenderTargetEnabled Then Angle = varRenderScene.Angle
End Property

Public Property Let Angle(value As Single)
    If varRenderTargetEnabled Then varRenderScene.Angle = value
End Property

' Espejado de la escena:
Public Property Get Mirror() As Graphics.Mirrors
    If varRenderTargetEnabled Then Mirror = varRenderScene.Mirror
End Property

Public Property Let Mirror(value As Graphics.Mirrors)
    If varRenderTargetEnabled Then varRenderScene.Mirror = value
End Property

' Posicion de la escena en pantalla:
Public Property Get Location() As Core.Point
    If varRenderTargetEnabled Then Location = varRenderScene.Location
End Property

Public Property Let Location(value As Core.Point)
    If varRenderTargetEnabled Then varRenderScene.Location = value
End Property

' Tamaño de la vista de la escena (escala la representacion de la escena):
Public Property Get Size() As Core.Size
    If varRenderTargetEnabled Then Size = varRenderScene.Size
End Property

Public Property Let Size(value As Core.Size)
    If varRenderTargetEnabled Then varRenderScene.Size = value
End Property

' --------------------------------------------------------------------------------

' Convierte las coordenadas de la pantalla o vista de escena (area de la ventana) a coordenadas del espacio de la escena:
Public Function ScreenToClient(Point As Core.Point) As Core.Point
    Dim p As Core.Point
    p.X = Point.X + GraphicEngine.Offset.X
    p.Y = Point.Y + GraphicEngine.Offset.Y
    
    ScreenToClient = p
End Function

' Convierte las coordenadas de la escena a las de la pantalla o vista de escena:
Public Function ClientToScreen(Point As Core.Point) As Core.Point
    Dim p As Core.Point
    p.X = Point.X - GraphicEngine.Offset.X
    p.Y = Point.Y - GraphicEngine.Offset.Y
    p.Z = Point.Z
    
    ClientToScreen = p
End Function

' Obtiene o determina la posicion de la camara en el espacio de la escena (por ejemplo para implementacion de scrolles)
Public Property Let Camera(value As Core.Point)
    Dim Width As Long, Height As Long
    Width = GraphicEngine.CurrentDisplayMode.Width
    Height = GraphicEngine.CurrentDisplayMode.Height
    
    ' -(x + ScreenWidth \ 2) -(y + ScreenHeight \ 2)
    GraphicEngine.Offset = Core.Generics.CreatePOINT(-(value.X - Width \ 2), -(value.Y - Height \ 2))
    
    ' Actualizamos el area de accion del simulador de fisicas:
    PhysicSimulator.WorkArea = Core.Generics.CreateRECTANGLE(value.X - Width, _
                                                               value.Y - Height, _
                                                               Width * 2, Height * 2)

    varCamera = value
End Property

Public Property Get Camera() As Core.Point
    Camera = varCamera
End Property

' Asigna una entidad como objetivo de la camara de la escena:
Public Sub SetCameraTarget(Key As String) ' "" = No hay objetivo asignado. Se mantendra la posicion actual.
    Set varCameraTarget = Entities(Key)
    varCameraTargetKey = Key
End Sub

' Devuelve la clave de la entidad a la que sigue la camara:
Public Property Get CameraTarget() As String
    CameraTarget = varCameraTargetKey
End Property

Public Property Get CameraTargetOffSet() As Core.Point
    CameraTargetOffSet = varCameraTargetOffset
End Property

Public Property Let CameraTargetOffSet(value As Core.Point)
    varCameraTargetOffset = value
End Property

' Actualiza la logica de toda la escena:
Public Sub Update()
    If Engine.EDIT_MODE Then
        ' Desplazamos la camara:
        If Engine.dev_input.Press("left") Or Engine.dev_input.Hit("left_hit") Then Me.Camera = Core.Generics.CreatePOINT(Me.Camera.X - 64, Me.Camera.Y)
        If Engine.dev_input.Press("right") Or Engine.dev_input.Hit("right_hit") Then Me.Camera = Core.Generics.CreatePOINT(Me.Camera.X + 64, Me.Camera.Y)
        If Engine.dev_input.Press("up") Or Engine.dev_input.Hit("up_hit") Then Me.Camera = Core.Generics.CreatePOINT(Me.Camera.X, Me.Camera.Y - 64)
        If Engine.dev_input.Press("down") Or Engine.dev_input.Hit("down_hit") Then Me.Camera = Core.Generics.CreatePOINT(Me.Camera.X, Me.Camera.Y + 64)
        
        ' Nos permite desplazarnos en medidas de escena completa (escena completa = dimensiones de la vista de ventana):
        If Engine.dev_input.Hit("room_move_left") Then Me.Camera = Core.Generics.CreatePOINT(Me.Camera.X - 768, Me.Camera.Y)
        If Engine.dev_input.Hit("room_move_right") Then Me.Camera = Core.Generics.CreatePOINT(Me.Camera.X + 768, Me.Camera.Y)
        If Engine.dev_input.Hit("room_move_up") Then Me.Camera = Core.Generics.CreatePOINT(Me.Camera.X, Me.Camera.Y - 576)
        If Engine.dev_input.Hit("room_move_down") Then Me.Camera = Core.Generics.CreatePOINT(Me.Camera.X, Me.Camera.Y + 576)
    
        If Engine.dev_input.Hit("fx1_hit") Then
            If Me.RenderTargetEnabled Then
                Call Me.DeactivateRenderTarget
            Else
                Call Me.ActivateRenderTarget
            End If
        End If
        
        If Engine.dev_input.Hit("fps_hit") Then
            If GraphicEngine.MaxFrames = 60 Then
                GraphicEngine.MaxFrames = 30
            Else
                GraphicEngine.MaxFrames = 60
            End If
        End If
        
        If Engine.dev_input.Hit("physic_switch") Then
            PhysicSimulator.Gravity = Core.Generics.CreateVECTOR(0, 6, 0)
            Engine.PlayerInputEnabled = True
            Call Me.SetCameraTarget("deckard")
        End If
        
        If Engine.dev_input.Hit("show_bodies") Then ShowPhysicBodies = Not ShowPhysicBodies
        
        If Engine.dev_input.Hit("show_physics_guides") Then
            Engine.PHYSICS_DRAW_GUIDES = Not Engine.PHYSICS_DRAW_GUIDES
            Engine.PhysicEngine.DEBUG_DrawColliders = Engine.PHYSICS_DRAW_GUIDES
        End If
    End If
    
    If Enabled Then
        Dim Tile As TLSA.ENG_Tile
        For Each Tile In TileMap
            Call Tile.Sprite.Update
        Next
        
        ' Actualizamos el estado de todas las entidades dentro del area de accion del simulador:
        If Editor.EditMode Then
            Dim Entity As Object
            For Each Entity In Entities
                Call Entity.Update
            Next
        End If
        
        ' Actualizamos el estado del simulador de fisicas:
        If Not Engine.PHYSICS_SEPARATE_THREAD Then
            Call PhysicSimulator.Update
        End If
        
        ' Actualizamos la posicion de la camara:
        If Not varCameraTargetKey = "" Then _
            Camera = Core.Generics.CreatePOINT(varCameraTarget.Location.X + varCameraTargetOffset.X, varCameraTarget.Location.Y + varCameraTargetOffset.Y)
        
        Dim clientCursor As Core.Point: clientCursor = Engine.Scene.ClientToScreen(dev_input.ViewAxis)
        Call Engine.dev_debug.AddMessage("Cursor [Scene/Client].(x:" & clientCursor.X & " y:" & clientCursor.Y & ") [Screen/Window].(x:" & dev_input.ViewAxis.X & " y:" & dev_input.ViewAxis.Y & ")")
        
        ' Informacion de depuracion:
        Call Engine.dev_debug.AddMessage("Camera(x:" & Me.Camera.X & " y:" & Me.Camera.Y & ")")
        Call Engine.dev_debug.AddMessage("WorkArea (x:" & Me.PhysicSimulator.WorkArea.X & " y:" & Me.PhysicSimulator.WorkArea.Y & " width:" & Me.PhysicSimulator.WorkArea.Width & " height:" & Me.PhysicSimulator.WorkArea.Height & ")")
        Call Engine.dev_debug.AddMessage("Tiles in escene: " & Me.TileMap.Count)
    End If
    
    ' Si el modo de edicion esta activado se actualiza los estados del editor:
    If Engine.EDIT_MODE Then LevelEditor.Update
End Sub

' Renderiza la escena:
Public Sub Draw()
    If Visible Then
        If Me.RenderTargetEnabled Then
            ' Dibujamos una capa de color blanco para preservar el valor correcto del canal alpha en las operaciones de dibujo:
            With varRenderScene
                Call .PrecalulateVertex
                
                Dim points(3) As Core.Point
                points(0) = .GetPrecalVertex(Left_Top)
                points(0).X = points(0).X: points(0).Y = points(0).Y
                points(0).Color = Graphics.Color_Constant.White

                points(1) = .GetPrecalVertex(Right_Top)
                points(1).X = points(1).X: points(1).Y = points(1).Y
                points(1).Color = Graphics.Color_Constant.White

                points(2) = .GetPrecalVertex(Left_Bottom)
                points(2).X = points(2).X: points(2).Y = points(2).Y
                points(2).Color = Graphics.Color_Constant.White

                points(3) = .GetPrecalVertex(Right_Bottom)
                points(3).X = points(3).X: points(3).Y = points(3).Y
                points(3).Color = Graphics.Color_Constant.White
                
                Call GraphicEngine.Primitives.DrawShape(points(0), points(1), points(2), points(3), True)
            End With
        End If
        
        If Me.RenderTargetEnabled Then Call varRenderTarget.BeginTarget ' Abrimos el render target.

            ' *** Pintamos de blanco el fondo de la escena, para resaltar las operaciones de alpha ***
            If Me.RenderTargetEnabled Then
                Engine.GraphicEngine.BackColor = Graphics.Color_Constant.Black
                Call GraphicEngine.Primitives.DrawBox3(0, 0, 800, 600, 8, Graphics.Color_Constant.White, True, Graphics.Color_Constant.White, True)
            Else
                Engine.GraphicEngine.BackColor = Graphics.Color_Constant.White
            End If

            ' Dibujamos todos los tiles visibles en pantalla:
            Dim Tile As TLSA.ENG_Tile
            For Each Tile In TileMap
                If Tile.Sprite.Location.Z = 0 And ShowMainLayer Or _
                   Tile.Sprite.Location.Z = -4 And ShowBackLayer Or _
                   Tile.Sprite.Location.Z = 4 And ShowFrontLayer _
                Then _
                    Call Tile.Sprite.Draw
            Next

            ' Dibujamos todas las entidades visibles en pantalla:
            Dim Entity As Object
            For Each Entity In Entities
                Call Entity.Draw
            Next

        If Me.RenderTargetEnabled Then Call varRenderTarget.EndTarget ' Cerramos el render target.
        
            ' *** Modifica el angulo de la escena ***
        If Me.RenderTargetEnabled Then
            Me.Angle = Me.Angle + 0.05
            If Me.Angle > 360 Then Me.Angle = 0
        
            Call varRenderScene.Draw ' Dibujamos la escena.
        End If
        
        ' Dibuja el area de trabajo del simulador de fisicas:
        'Call GraphicEngine.Primitives.DrawBox3(PhysicSimulator.WorkArea.x, PhysicSimulator.WorkArea.y, PhysicSimulator.WorkArea.width, PhysicSimulator.WorkArea.height, -8, Graphics.Color_Constant.Red, , , False)
        
        ' Se dibujaran las guias de los cuerpos fisicos si la opcion esta activada o solo la caja principal si se indica:
        If Engine.PhysicEngine.DEBUG_DrawColliders Or Me.ShowPhysicBodies Then Call Engine.Scene.PhysicSimulator.Draw
        
        ' Si el modo de edicion esta activado se muestran los estados del editor:
        If Engine.EDIT_MODE Then Call LevelEditor.Draw
        
        ' Mostramos la informacion del monitor del render grafico: (se debe llamar en este punto para obtener los resultados correctos)
        Call Engine.dev_debug.AddMessage(Engine.GraphicEngine.Monitor.ToString())
    End If
End Sub

' Carga un nivel en el escenario: (este metodo sera llamado tambien por el editor)
Public Sub LoadLevel(filename As String)
    ' ***** codigo de prueba provisional *****
    With PhysicSimulator
        .Enabled = True
        .Gravity = Core.Generics.CreateVECTOR(0, 0, 0)
        .Delay = 0 '12
    End With
    
    ' Actualizamos la posicion de la camara y con ello inicializamos el area de accion del simulador:
    Me.Camera = Core.Generics.CreatePOINT(0, 0)
    
    Dim marine As New TLSA.ACT_Marine
    marine.Visible = True
    marine.Enabled = True ' Not Engine.EDIT_MODE
    marine.IsPlayer = True
    marine.Location = Core.Generics.CreatePOINT(48, 64)
    Call marine.SetScale(2)
    Call Entities.Add(marine, "deckard")
    Call marine.Update
    Set marine = Nothing

    Engine.PlayerInputEnabled = Not EDIT_MODE
    If Not EDIT_MODE Then Call Me.SetCameraTarget("deckard")
    
'    Dim enemy As New TLSA.ACT_ParticleEnemy
'    enemy.Visible = True
'    enemy.Enabled = True
'    enemy.Behavior = VerticalPath
'
'    Call Entities.Add(enemy, "particle")
'    Set enemy = Nothing
    
    'Call Me.ActivateRenderTarget
    'Me.Mirror = Horizontal
    ' ***** codigo de prueba provisional *****
End Sub

' Carga la informacion de escena con todos sus elementos y entidades:
Public Sub LoadScene(filename As String)
    Dim Buffer() As String, Count As Long
    Count = Core.IO.ReadTextFile(filename, Buffer)  ' Cargamos el archivo en memoria para proceder a interpretarlo.
    
    Call Me.Clear   ' Destruimos el escenario previo.
    
    Dim data() As String, params() As String
    Dim i As Long
    For i = 0 To Count - 1
        If Not Core.Text.StartWith(Buffer(i), "//") Then ' Or Core.Text.Find(Buffer(i), ";") Then
            data = Split(Buffer(i), "=")    ' Separamos el comando de los parametros.
            params = Split(data(1), ";")    ' Obtenemos los parametros.
            
            Select Case UCase(data(0))
                ' Gravedad de la escena:
                Case UCase("Engine.Scene.Physics.Gravity")
                    PhysicSimulator.Gravity = Core.Generics.CreateVECTOR(CSng(params(0)), CSng(params(1)), 0)
                
                ' Estado del simulador de fisicas de la escena:
                Case UCase("Engine.Scene.Physics.Enabled")
                    Me.PhysicSimulator.Enabled = CBool(params(0))
                
                ' Posicion inicial de la camara:
                Case UCase("Engine.Scene.Camera.Location")
                    Me.Camera = Core.Generics.CreatePOINT(CLng(params(0)), CLng(params(1)))
                
                ' Entidad objetivo a la que apuntara la camara:
                Case UCase("Engine.Scene.Camera.Target")
                    If Not UCase(params(0)) = "NULL" Then Call Me.SetCameraTarget(params(0))
                
                ' Agrega una textura a la escena:
                Case UCase("Engine.Scene.Textures.Add")
                    ' Cargamos la textura en el motor:
                    Call Engine.GraphicEngine.Textures.LoadTexture(App.Path & TLSA.ResourcePaths.Textures & params(0), params(0), False)
                    ' Agregamos su clave a la lista de texturas del nivel:
                    Call Engine.Scene.Textures.Add(params(0), params(0))
                
                ' Agrega un tile a la escena:
                Case UCase("Engine.Scene.Tiles.Add")
                    ' ToString = Clave Textura; Clave Tile; EstadoAnimacion; Clave Animacion; X; Y; Z; Angulo; Color; Efecto; Espejado
                    Dim Tile As Graphics.Sprite
                    Set Tile = Engine.GraphicEngine.Helper.CreateAnimatedSprite(Engine.GraphicEngine.Textures(params(0)))
                    Tile.EnabledAnimation = CBool(params(2))
                    If Tile.EnabledAnimation Then Call Tile.SetCurrentAnimation(params(3)) Else Call Tile.SetCurrentTile(params(1))
                    Call Tile.setLocation(CLng(params(4)), CLng(params(5)), CLng(params(6)))
                    Tile.Angle = CSng(params(7))
                    Tile.Color = CLng(params(8))
                    Tile.Effect = CLng(params(9))
                    Tile.Mirror = CLng(params(10))
                    
                    Call Engine.Scene.TileMap.Add(Tile)
                    
                    Set Tile = Nothing
                
                ' Agrega un cuerpo fisico a la escena:
                Case UCase("Engine.Scene.Physics.Add")
                    ' ToString = X; Y; Ancho; Alto; PerfilColisionadores; GrosorColisionadores
                    Dim b As Physics.Body
                    
                    Set b = Me.PhysicSimulator.Bodies.Add(CLng(params(0)), CLng(params(1)), 0, CLng(params(2)), CLng(params(3)), _
                                                          CLng(params(4)), 0, True, CLng(params(5)))
                    b.Color = Graphics.Color_Constant.Orange
                    b.Tag = CStr(params(4) & ";" & params(5))
                    
                    Call Engine.Scene.PhysicMap.Add(b)          ' Añadimos la referencia al motor.
                    
                    Set b = Nothing
                
                ' Comprobamos si se trata de una entidad:
                Case Else
                    If Core.Text.StartWith(UCase(data(0)), UCase("Engine.Scene.Entity.")) Then
                        Dim entityKey As String: entityKey = Core.Text.GetFromRight(UCase(data(0)), ".")
                        Select Case UCase(entityKey)
                            Case UCase("Marine")
                                Dim marine As New TLSA.ACT_Marine
                                With marine
                                    .Location = Core.Generics.CreatePOINT(CLng(params(0)), CLng(params(1)))
                                    .orientation = params(2)
                                    .IsPlayer = CBool(params(3))
                                    If Not Engine.EDIT_MODE Then Engine.PlayerInputEnabled = .IsPlayer
                                End With
                                Call Engine.Scene.Entities.Add(marine)
                                Set marine = Nothing
                                
                            Case UCase("ParticleEnemy")
                                Dim Particle As New TLSA.ACT_ParticleEnemy
                                With Particle
                                    .Location = Core.Generics.CreatePOINT(CLng(params(0)), CLng(params(1)))
                                    .Behavior = CLng(params(2))
                                    .ParticleType = CLng(params(3))
                                    .Reborn = CBool(params(4))
                                    .TriggerEventDuration = CLng(params(5))
                                End With
                                Call Engine.Scene.Entities.Add(Particle)
                                Set Particle = Nothing
                        End Select
                    End If
            End Select
        End If
    Next
End Sub
