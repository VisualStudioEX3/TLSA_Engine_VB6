VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ACT_Marine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const LASER_LENGTH As Long = 2048               ' Logitud maxima del laser del arma si no colisiona con ningun obstaculo.

Private varTextureFile As String
Private varSpriteAnim As String

' Canales de animacion del actor:
Private Head As Graphics.Sprite
Private Arms As Graphics.Sprite
Private Body As Graphics.Sprite
Private Legs As Graphics.Sprite

' Elementos adicionales del actor:
Private FireShoot As TLSA.FX_FireShoot                  ' Instancia del efecto de disparo.

Private WithEvents BBox As Physics.Body                 ' Representacion fisica del actor en el simulador.
Attribute BBox.VB_VarHelpID = -1

Private varScale As Single
Private varAction(1) As String                          ' Canales de claves de las acciones que esta ejecutando el actor.
Private clientArms As Core.Point                        ' Coordenadas en cliente de los brazos del actor.

' Variables publicas:
Public Enabled As Boolean
Public Visible As Boolean
Public Angle As Single
Public IsPlayer As Boolean                              ' Indica si es el jugador, en cuyo caso se activan los estados para lectura de entrada.
Public FixedAngle As Boolean                            ' Indica si el angulo es fijo (no se altera por la posicion del eje de vision)
Public orientation As Long                              ' Orientacion del actor: 0 - Derecha, 1 - Izquierda.
Public Life As Long                                     ' Salud del actor: 5 puntos. Cada punto se resta cuando el escudo baja a 0.
Public Shield As Long                                   ' Escudo del actor: 0 a 100.
Public Posture As Long                                  ' Postura del actor: 0 - De pie, 1 - Agachado, 2 - Tumbado.

' Variables privadas:
Private InputProfile As GameInput.Profile               ' Perfil de lectura de entrada del actor.
Private Laser As Physics.Ray, laserHit As Core.Point    ' El rayo para trazar el laser.
Private shootRecoil As Long
Private runRecoil As Long
Private fastRun As Long
Private slowRun As Long

' Cronometros para control de estados y animaciones:
Private timerRecoil As Core.Timer                       ' Cronometro para controlar la animacion de retroceso del arma al disparar.
Private timerRun As Core.Timer                          ' Cronometro para controlar la animacion de movimiento del cuerpo en carrera.

' Estados de control para las animaciones y acciones:
Private AttackMode As Boolean                           ' Indica si el actor esta en modo de ataque.
Private Running As Boolean                              ' Indica si el actor esta corriendo.
Private Jumping As Boolean                              ' Indica si el actor esta saltando.
Private Shooting As Boolean                             ' Indica si el actor esta disparando.
Private Reloading As Boolean                            ' Indica si el actor esta recargando un arma.
Private Falling As Boolean                              ' Indica si el actor esta cayendo en el vacio.
Private Crouching As Boolean                            ' Indica si el actor esta agachado o agazapado.
Private Tumble As Boolean                               ' Indica si el actor esta tumbado.

Public Property Get Location() As Core.Point
    Location = Core.Generics.CreatePOINT(CLng(BBox.Location.X), CLng(BBox.Location.Y) - 15)
    ' Corregimos la posicion vertical para que los graficos entren dentro del area de la caja de colision.
End Property

Public Property Let Location(value As Core.Point)
    BBox.Location = Core.Generics.CreateVECTOR(CSng(value.X), CSng(value.Y + 15), CSng(value.Z))
    ' Corregimos la posicion vertical para que los graficos entren dentro del area de la caja de colision.
End Property

' Selecciona un perfil de lectura de entrada para el actor:
Public Sub SetInputProfile(ProfileName As String)
    Set InputProfile = InputEngine.Profiles(ProfileName)
    
    ' Comprobamos si existen las acciones requeridas en el perfil:
    ' ...
End Sub

Public Function GetScale() As Single
    GetScale = varScale
End Function

Public Sub SetScale(value As Single)
    varScale = value
    Call Head.SetScale(value)
    Call Arms.SetScale(value)
    Call Body.SetScale(value)
    Call Legs.SetScale(value)
    Call FireShoot.SetScale(value)
End Sub

' Obtiene acceso al cuerpo fisico del actor y sus propiedades:
Public Property Get PhysicBody() As Physics.Body
    Set PhysicBody = BBox
End Property

Private Sub Class_Initialize()
    varTextureFile = App.Path & ResourcePaths.Entities & "Marine\Marine.png"
    varSpriteAnim = App.Path & ResourcePaths.Entities & "Marine\Marine.ani"
    
    ' Cargamos la textura, si no esta cargada ya, y la asignamos:
    If Not GraphicEngine.Textures.Exists("marine") Then
        Call GraphicEngine.Textures.LoadTexture(varTextureFile, "marine", False)
    End If
    
    ' Inicializamos los canales:
    Set Head = GraphicEngine.Helper.CreateAnimatedSprite(GraphicEngine.Textures("marine")): Head.EnabledAnimation = False: Call Head.SetCurrentTile("head2")
    Set Arms = GraphicEngine.Helper.CreateAnimatedSprite(GraphicEngine.Textures("marine")): Arms.EnabledAnimation = False: Call Arms.SetCurrentTile("mp7_idle")
    Set Body = GraphicEngine.Helper.CreateAnimatedSprite(GraphicEngine.Textures("marine")): Body.EnabledAnimation = False: Call Body.SetCurrentTile("body")
    Set Legs = GraphicEngine.Helper.CreateAnimatedSprite(GraphicEngine.Textures("marine")): Legs.EnabledAnimation = False: Call Legs.SetCurrentTile("legs_idle")
    
    ' *** Hide ***
'    Head.Color = 0
'    Body.Color = 0
'    Arms.Color = 0
'    Legs.Color = 0
    
    ' Constantes de espera entre fotogramas para las velocidades de la animacion de carrera segun orientacion del personaje:
    fastRun = Legs.Animations("run").FrameDelay
    slowRun = fastRun / 0.7
    
    ' Creamos una instancia del efecto de disparo:
    Set FireShoot = New TLSA.FX_FireShoot
    
    ' Creamos la instancia del actor en el simulador de fisicas:
    Set BBox = Engine.Scene.PhysicSimulator.Bodies.Add(0, 0, 0, 35, 75, AllEntireArea, 4, False, 20)
    Call BBox.SetMaxFriction(0.5, BBox.MaxFriction.Y)
    BBox.PhysicType = PhysicalEntity
    BBox.Color = Graphics.Color_Constant.Cyan
    
    ' Cargamos el perfil del jugador para la entidad marine:
    Set InputProfile = InputEngine.Profiles.Create("player", Player1, KeybAndMouse) ', Gamepad)
    Call InputProfile.Import(App.Path & ResourcePaths.InputProfiles & "player.prf")
    InputProfile.GamepadUsed = 1 '3
    Call InputProfile.SetCursorPosition(0, GraphicEngine.CurrentDisplayMode.Height \ 2)
    
    Set timerRecoil = New Core.Timer
    Set timerRun = New Core.Timer
    
    varScale = 1
    Call SetAction(0, "idle")
    Call SetAction(1, "")
    
    orientation = 0
    Life = 5
    Shield = 100
    Posture = 0
    AttackMode = True
    
    Visible = True
    Enabled = True
    
    Call Me.SetScale(2)
    
    Call Me.Update                          ' Forzamos a que se muestre correctamente al inicio.
End Sub

Private Sub Class_Terminate()
    Call InputEngine.Profiles.Remove("player")
    Set InputProfile = Nothing
    
    Set timerRecoil = Nothing
    Set timerRun = Nothing
    
    Set Head = Nothing
    Set Arms = Nothing
    Set Body = Nothing
    Set Legs = Nothing
    
    Set FireShoot = Nothing
    
    If Not Scene Is Nothing Then Call Scene.PhysicSimulator.Bodies.Remove(BBox)
    
    On Error Resume Next ' Evitamos la excepcion en caso de que la textura siga en uso:
    Call Engine.GraphicEngine.Textures.Unload("marine")
End Sub

Public Sub Draw()
    If Visible And Enabled Then
        ' Dibuja los sprites:
        Call Legs.Draw
        Call Body.Draw
        Call Head.Draw
        Call Arms.Draw
        
        ' Dibujamos la traza del puntero laser del arma:
        'If Not Editor.EditMode Then
            If AttackMode And Not Reloading Or InputProfile.Device = Gamepad Then
                Call GraphicEngine.Primitives.DrawLine2(Arms.CurrentTile.ControlPoints("laser"), laserHit, 0, &H77FF0000)
                If Shooting Then Call FireShoot.Draw ' Dibujamos el disparo.
            End If
        'End If
    End If
    
    Call Engine.dev_debug.AddMessage("Player (x:" & Me.Location.X & " y:" & Me.Location.Y & ")")
End Sub

Public Sub Update()
'    If Editor.EditMode Then
'        Angle = 0
'    End If
    
    If Enabled Then
        ' Actualizamos la lectura de entrada del jugador:
        Call UpdateInput
        
        ' Actualizamos la accion actual:
        Call UpdateAction
    
        ' Actualizamos las posiciones de los sprites:
        Legs.Location = Me.Location: Call Legs.Update
        Body.Location = Core.Generics.CreatePOINT(Legs.Location.X, Legs.Location.Y + runRecoil, 0): Call Body.Update
        Arms.Location = Body.CurrentTile.ControlPoints("arms"): Call Arms.Update
        Head.Location = Body.CurrentTile.ControlPoints("head"): Call Head.Update
        
        ' Actualizamos estados:
        Call UpdateStatus
        
        ' Si se tiene asociada la camara se aplica la correccion de camara (habra que hacer el offset en la instancia de la escena y no del jugador):
        Dim tmpOffset As Core.Point: tmpOffset = Engine.Scene.CameraTargetOffSet
        
        If orientation = 0 And tmpOffset.X < 32 Then
            tmpOffset.X = tmpOffset.X + 4
        ElseIf orientation = 1 And tmpOffset.X > -32 Then
            tmpOffset.X = tmpOffset.X - 4
        End If
        
        ' Codigo para controlar la camara y ajustar correccion de vision segun orientacion del jugador y segun altura del cursor:
        Select Case Angle
            Case 0 To 180
                If tmpOffset.Y > -32 Then tmpOffset.Y = tmpOffset.Y - 4
            Case 181 To 359
                If tmpOffset.Y < 32 Then tmpOffset.Y = tmpOffset.Y + 4
        End Select
        
        Engine.Scene.CameraTargetOffSet = tmpOffset
    End If
End Sub

' Evento de colision del actor:
Private Sub BBox_OnCollision(BodyCount As Long, E() As Physics.CollisionEventData)

End Sub

' Evento de impacto de una entidad o accion en el actor:
Private Sub BBox_OnHit(Force As Single, eventData As Variant)

End Sub

' Actualiza la lectura de entrada del usuario:
Public Sub UpdateInput()
    ' Bloque de acciones del actor (bloqueables en caso de ser necesario, como en una cinematica o evento de muerte por ejemplo):
    If Engine.PlayerInputEnabled And Me.IsPlayer Then
        'If InputProfile.Device = KeybAndMouse Then clientCursor = Engine.Scene.ClientToScreen(InputProfile.ViewAxis)
        
        If InputProfile.Press("left") Then Call SetAction(0, "run_left")
        If InputProfile.Press("right") Then Call SetAction(0, "run_right")
        If InputProfile.Press("up") Then Call BBox.SetForceY(-3)
        If InputProfile.Press("down") Then Call BBox.SetForceY(3)
        If InputProfile.Press("shoot") Then
            If Not Reloading Then
                If Not Shooting Then Call InputProfile.SetVibration(-1, 64)     ' Si hay soporte de vibracion se aplica.
                Call SetAction(1, "mp7_shoot")
            End If
        Else
            Shooting = False
            shootRecoil = 0
            Call SetAction(1, "")
        End If
        If InputProfile.Hit("reload") Then Call SetAction(1, "mp7_reload")
        If Not InputProfile.Press("right") And Not InputProfile.Press("left") Then
            Call SetAction(0, "idle")
            Running = False
        End If
        
        ' Obtenemos el angulo de vision (raton, tomando como centro la posicion del brazo, o segundo eje del gamepad):
        If Not Me.FixedAngle Then
            clientArms = Engine.Scene.ScreenToClient(Core.Generics.CreatePOINT(Arms.Location.X, Arms.Location.Y))
            Angle = InputProfile.ViewAngle(clientArms.X, clientArms.Y)
        End If
    End If
End Sub

' Actualizamos estados del actor:
Public Sub UpdateStatus()
    ' Si no hay contacto activamos estado de caida:
    'If Not BBox.Hit Then Call Me.SetAction("fall")
    
    ' Si la vida es 0 o menor activamos el estado de muerte:
    If Life = 0 Then Debug.Print "Player dead."
    
    ' Cuando la animacion de recarga de armas termina se reestablece el estado de apuntado/reposo del arma:
    If (Reloading And Arms.CurrentAnimation.IsAnimationEnded) Then
        Call Arms.CurrentAnimation.Reset
        Reloading = False
        Call SetAction(0, "idle")
        Call SetAction(1, "")
    End If
    
    ' Segun angulo aplicamos espejado o no:
    If Angle > -1 Then
        Select Case Angle
            Case 90 To 270
                orientation = 0
                Head.Mirror = WithoutMirroring
                Arms.Mirror = WithoutMirroring
                Body.Mirror = WithoutMirroring
                Legs.Mirror = WithoutMirroring
                FireShoot.Mirror = WithoutMirroring
            Case 0 To 89, 271 To 359
                orientation = 1
                Head.Mirror = Horizontal
                Arms.Mirror = Horizontal
                Body.Mirror = Horizontal
                Legs.Mirror = Horizontal
                FireShoot.Mirror = Horizontal
        End Select
    End If
    
    ' Cambia el sentido de la animacion de carrera segun orientacion:
    If (orientation = 0 And GetAction(0) = "run_left") Or (orientation = 1 And GetAction(0) = "run_right") Then
        Legs.CurrentAnimation.AnimatePath = Reverse
        Legs.CurrentAnimation.FrameDelay = slowRun
    Else
        Legs.CurrentAnimation.AnimatePath = Foward
        Legs.CurrentAnimation.FrameDelay = fastRun
    End If
    
    ' Creamos un efecto de tambaleo en la carrera:
    If (Running And Not Shooting) Or (Running And Shooting) Or (Not Running And Not Shooting) Then
        If (Running And timerRun.value >= 150) Or (Not Running And timerRun.value >= 1000) Then
            runRecoil = -(Not CBool(runRecoil))
            Call timerRun.Reset
        End If
    Else
        runRecoil = 0
    End If
    
    ' Recreamos el efecto del disparo:
    If AttackMode And Not Reloading And Shooting Then
        ' Creamos un efecto de retroceso:
        If timerRecoil.value >= 50 Then
            shootRecoil = Not CBool(shootRecoil)
            Arms.Location = Core.Math.TranslatePoint(Arms.Location, shootRecoil, Arms.Angle)
            Call timerRecoil.Reset
        End If
        
        FireShoot.Location = Arms.CurrentTile.ControlPoints("shoot")
        Call FireShoot.Update
    End If
    
    If AttackMode Then
        If Reloading Or Angle = -1 Then
            Body.Angle = 0
            Arms.Angle = 0
            Head.Angle = 0
            FireShoot.Angle = 0
            
            If InputProfile.Device = Gamepad Then
                If orientation = 0 Then
                    Angle = 180
                ElseIf orientation = 1 Then
                    Angle = 0
                End If
            End If
            
        Else
            Arms.Angle = Angle + IIf(orientation = 0, 180, 0)
            Head.Angle = Angle + IIf(orientation = 0, 180, 0)
            FireShoot.Angle = Angle + IIf(orientation = 0, 180, 0)
        End If
    End If
    
    'If AttackMode And Not Reloading And (Core.Math.PointInRect(Engine.Scene.ScreenToClient(InputProfile.ViewAxis), Scene.PhysicSimulator.WorkArea) Or InputProfile.Device = Gamepad) Then
    If AttackMode And Not Reloading Then
        ' Calculamos la trayectoria del puntero laser:
        Dim tmpHit As Core.Point, tmpTarget As Core.Point
        If Not BBox.TraceRay(Arms.CurrentTile.ControlPoints("laser"), Angle, tmpHit, tmpTarget, Physics.Helper.EmptyZDiscardArray) Is Nothing Then
            laserHit = tmpHit
        Else
            ' Calculamos la trayectoria traslando el punto en el angulo y a una distancia fija:
            laserHit = Core.Math.TranslatePoint(Arms.CurrentTile.ControlPoints("laser"), -LASER_LENGTH, Angle)
        End If
    End If
End Sub

' Actualiza las acciones del actor:
Public Sub UpdateAction()
    ' Estados de acciones primarias:
    Select Case varAction(0)
        Case "idle"

        Case "fall"
        
        Case "run_left"
            ' Aplicamos fuerza negativa en el eje X:
            If orientation = 0 Then Call BBox.SetForceX(-1) Else Call BBox.SetForceX(-3)
        Case "run_right"
            ' Aplicamos fuerza positiva en el eje X:
            If orientation = 1 Then Call BBox.SetForceX(1) Else Call BBox.SetForceX(3)
        Case Else

    End Select
    
    ' Estados de acciones secundarias:
    Select Case varAction(1)
        Case "mp7_shoot"
            
        Case "mp7_reload"
            Call InputProfile.StopVibration
        Case Else
            Call InputProfile.StopVibration
    End Select
End Sub

' Devuelve la clave del accion que se esta ejecutando en el canal indicado:
Public Function GetAction(Channel As Long) As String
    GetAction = varAction(Channel)
End Function

' Activa una accion en el canal indicado:
Public Sub SetAction(Channel As Long, Action As String)
    varAction(Channel) = LCase(Action)
    
    Select Case Channel
        Case 0 ' Accion primaria:
            Select Case Action
                Case "idle" ' Quieto:
                    ' Seleccionamos el tile de por defecto de las piernas:
                    Legs.EnabledAnimation = False
                    Call Legs.SetCurrentTile("legs_idle")
                    
                    If Not Reloading Then
                        Arms.EnabledAnimation = False
                        Call Arms.SetCurrentTile("mp7_idle")
                    End If
                Case "fall" ' Caida:
                
                Case "jump1" ' Salto vertical:
                
                Case "jump2" ' Salto largo:
                
                Case "run_left", "run_right"
                    ' Activamos la animacion de carrera:
                    Running = True
                    Legs.EnabledAnimation = True
                    Call Legs.SetCurrentAnimation("run")
                Case "dead" ' Dentro activamos aleatoriamente una animacion de muerte:
        
            End Select
        
        Case 1 ' Accion secundaria:
            Select Case Action
                Case "mp7_select", "mp7_shoot"
                    ' Seleccionamos el tile por defecto del mp7:
                    Arms.EnabledAnimation = False
                    Call Arms.SetCurrentTile("mp7_idle")
                    Shooting = (Action = "mp7_shoot")
                    
                Case "mp7_reload"
                    ' Activamos la animacion de recarga del mp7:
                    Reloading = True
                    Arms.EnabledAnimation = True
                    Call Arms.SetCurrentAnimation("mp7_reload")
            End Select
    End Select
End Sub

Public Function ToString() As String
    ' /// Codigo temporal ///
    ' <Key>:X;Y;Orientation;IsPlayer
    ToString = "Marine=" & Me.Location.X & ";" & Me.Location.Y & ";" & Me.orientation & ";" & Me.IsPlayer
End Function

Public Function Clone() As TLSA.ACT_Marine
    Dim m As New TLSA.ACT_Marine
    
    m.Angle = Me.Angle
    m.FixedAngle = Me.FixedAngle
    m.IsPlayer = Me.IsPlayer
    m.Life = Me.Life
    m.Location = Me.Location
    m.orientation = Me.orientation
    m.Posture = Me.Posture
    m.Shield = Me.Shield
    
    Set Clone = m
    Set m = Nothing
End Function
