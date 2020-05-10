VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TLSA.Engine: PhysicTest"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Form1.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim gfx As New Graphics.Manager
Dim gameInput As New dx_Input_Class
Dim sys As New dx_System_Class

Dim physic As New Physics.Manager
Dim sim As New Simulator
Dim WithEvents a As Body, WithEvents b As Body, C As Body, D As Body
Attribute a.VB_VarHelpID = -1
Attribute b.VB_VarHelpID = -1

Dim curVis As Boolean

Dim fnt As Graphics.Font
Dim lblKeyMap As New Graphics.TextString
Dim logo As Graphics.Texture
Dim spr As New Graphics.Sprite

Dim msg As String ' Anida los mensajes a imprimir en pantalla.

Dim crouch As Boolean

Private Sub A_OnCollision(BodyCount As Long, E() As Physics.CollisionEventData)
'    Stop
    Dim a As Integer, b As Integer
    msg = ""
    If BodyCount > 0 Then
        For a = 0 To BodyCount - 1
            Select Case E(a).Body.Tag ' Usamos la etiqueta TAG del objeto para obtener su tipo:
                Case "Friend": msg = msg & "Friend NPC collision!" & vbNewLine
                Case "Enemy": msg = msg & "Enemy NPC collision!" & vbNewLine
                Case "Trigger": msg = msg & "Trigger event raise!" & vbNewLine
            End Select
'            For b = 0 To E(a).ColliderCount - 1
'                If E(a).Colliders(b) = 99 Then
'                    msg = msg & "Trigger event!" & vbNewLine
'                    Debug.Print msg
'                End If
'            Next
        Next
    End If
End Sub

Private Sub Form_Load()
    Me.Show
    
    ' Inicializamos el sistema grafico:
    Call gfx.Initialize(Me.hWnd, 800, 480, 32, True, True)
    gfx.BackColor = &HFF888888
    gfx.MaxFrames = 60
    
    ' Cargamos una fuente de texto para mostrar informacion en pantalla:
    Call gfx.Fonts.LoadFont("lucida console", "lucida8", 8, False, False, False, False)
    Set fnt = gfx.Fonts("lucida8")
    
    ' Cargamos una textura como logotipo:
    Call gfx.Textures.LoadTexture(App.Path & "\tlsa.png", "logo", False)
    Set logo = gfx.Textures("logo")
    
    
    Call gameInput.Init(Me.hWnd)
        
    Call physic.SetGraphics(gfx)
    Call sim.SetWorkArea(0, 0, 800, 480)
    sim.Enabled = True
    sim.Delay = 14
    Call sim.setGravity(0, 8)
    
    physic.DEBUG_DrawColliders = False

    ' Escenario (partes fijas):
    Dim t As Body
    
    Set t = sim.Bodies.Add(326, 483 - 100, 0, 551, 12, 0, 0, True, 3)
    Set t = sim.Bodies.Add(608, 333 - 100, 0, 12, 800, 0, 0, True, 2)
    Set t = sim.Bodies.Add(57, 627 - 100, 0, 12, 300, 0, 0, True, 2)
    
    
    ' Escalera:
    Dim s As Integer
    For s = 0 To 4
        Set t = sim.Bodies.Add(500 + (24 * s) - 6, 471 - (12 * s) - 100, 0, 24, 12, 0, 0, True, 3)
    Next
    
    
    ' Entidades:
    
    ' Entidad no fisica:
    Set t = sim.Bodies.Add(250, 300, 0, 200, 200, 0, 0, False, 99)
    t.PhysicType = NoPhysicalEntity
    t.Color = &HFFFFFF00
    t.Tag = "Trigger"
    
    ' Entidad fisica:
    Set a = sim.Bodies.Add(100, 100, 0, 24, 100, 1, 2, False, 12)
    a.Color = &HFFFF7700
    a.PhysicType = PhysicalEntity
    a.Tag = "Player"
    
    ' Entidad fisica:
    Set b = sim.Bodies.Add(212, 300, 1, 24, 100, 1, 1, False, 12)
    b.Color = &HFF0000FF
    b.PhysicType = PhysicalEntity
    b.Tag = "Friend"

    ' Cuerpo normal (no entidad):
    Set C = sim.Bodies.Add(250, 100, 0, 100, 100, 1, 2, False, 12)

    ' Entidad fisica (enemigo):
    Set D = sim.Bodies.Add(312, 300, 0, 24, 100, 1, 1, False, 12)
    D.Color = &HFFFF0000
    D.PhysicType = PhysicalEntity
    D.Tag = "Enemy"
    
    Dim pt As Graphics.Point, source As Graphics.Point, target As Graphics.Point
    
    
    ' Configuramos el TextString para mostrar el mapa de teclas de la demo:
    Const helpText As String = "                    KEY MAP                    " & vbNewLine & _
                               "-----------------------------------------------" & vbNewLine & _
                               " A      - Run left                             " & vbNewLine & _
                               " D      - Run right                            " & vbNewLine & _
                               " Shift  - Walk                                 " & vbNewLine & _
                               " W      - Jump                                 " & vbNewLine & _
                               " S      - Crouch                               " & vbNewLine & _
                               " F1     - Show/Hide colliders in physic bodies " & vbNewLine & _
                               " Space  - Reset physic bodies possition        " & vbNewLine & _
                               " Escape - Exit                                 "
    
    Call lblKeyMap.Initialize(fnt, helpText, 0, 0, -8, , Right)
    
    
    
    ' Configuramos el logotipo del TLSA:
    Call spr.setTexture(logo)
    Call spr.setLocation(gfx.CurrentDisplayMode.Width - 100, gfx.CurrentDisplayMode.Height - 110, -8)
    
    Dim tim As Long
    tim = sys.TIMER_Create()
    
    Dim winCursor As Boolean
    winCursor = True
    
    Do While Not gfx Is Nothing
'If sys.TIMER_GetValue(tim) >= 25 Then
                
        ' Trazamos el rayo y obtenemos el punto de corte:
        source.X = CLng(a.Location.X): source.Y = CLng(a.Rect.Y) + 25 ' a.Rect.Height \ 4 'IIf(crouch, 0, 25)

        Dim angle As Single
        angle = sys.MATH_GetAngle(CLng(source.X), CLng(source.Y), gameInput.Mouse.X, gameInput.Mouse.Y)

        'If sim.TraceRay(source, target, pt) Then
        If a.TraceRay(source, angle, pt, target, 1) Then
            'Call gfx.DRAW_Line(source.X, source.Y, pt.X, pt.Y, -7, &HFF00FF00)
            Call gfx.Primitives.DrawLine(CLng(source.X), CLng(source.Y), CLng(pt.X), CLng(pt.Y), -7, &HFF00FF00)
            If physic.DEBUG_DrawColliders Then Call gfx.Primitives.DrawLine(CLng(pt.X), CLng(pt.Y), CLng(target.X), CLng(target.Y), -3, &HFFFF0000)
        Else
            Call gfx.Primitives.DrawLine(CLng(source.X), CLng(source.Y), CLng(target.X), CLng(target.Y), -7, &HFFFF0000)
        End If

        If physic.DEBUG_DrawColliders Then Call gfx.Primitives.DrawBox(CLng(source.X), CLng(source.Y), CLng(target.X), CLng(target.Y), -7, &HFF00FFFF)
        
        'Call sys.TIMER_Reset(tim)
'End If
        ' Prueba de estress. Probamos a calcular 20 rayos y 20 segmentos cada x tiempo.
        ' Si es posible, se optimizara algunas peticiones que use trazados de rayos o segmentos
        ' periodos o intervalos de tiempo y evitar llamadas en cada ciclo que no serian
        ' necesarias, ganando asi cierta velocidad.
'        If sys.TIMER_GetValue(tim) >= 1000 Then
'            Dim k As Long
'            For k = 0 To 20
'                Call A.TraceRay(source, angle, pt, target)
'                Call A.Tracesegment(source, target, pt)
'            Next
'            Call sys.TIMER_Reset(tim)
'        End If

        'Call gfx.DRAW_Box(source.X, source.Y, target.X, target.Y, -5, &HFF00FFFF)

        If gameInput.Key_Hit(Key_W) Then
            Call a.SetForceY(-20)
            If a.Force.X > 0 Then
                Call a.SetForceX(3.5)
            ElseIf a.Force.X < 0 Then
                Call a.SetForceX(-3.5)
            End If
        End If

        If gameInput.Key(Key_D) Then
            If gameInput.Key(Key_RShift) Or gameInput.Key(Key_LShift) Or crouch Or gameInput.Mouse.X < a.Location.X Then
                Call a.SetForceX(1)
            Else
                Call a.SetForceX(3.5)
            End If
        End If

        If gameInput.Key(Key_A) Then
            If gameInput.Key(Key_RShift) Or gameInput.Key(Key_LShift) Or crouch Or gameInput.Mouse.X > a.Location.X Then
                Call a.SetForceX(-1)
            Else
                Call a.SetForceX(-3.5)
            End If
        End If

        If gameInput.Key(Key_S) Then
            If Not crouch Then
                Call a.SetRect(a.Rect.X - 10, a.Rect.Y + 50, a.Rect.Width + 20, 50)
                Call gameInput.Mouse_SetPossition(gameInput.Mouse.X, gameInput.Mouse.Y + 50)

                crouch = True
            End If
        Else
            If crouch Then
                Call a.SetRect(a.Rect.X + 10, a.Rect.Y - 50, a.Rect.Width - 20, 100)
                Call gameInput.Mouse_SetPossition(gameInput.Mouse.X, gameInput.Mouse.Y - 50)

                crouch = False
            End If
        End If

'        Call gfx.DRAW_Line(0, CLng(a.Location.Y), 800, CLng(a.Location.Y), -8, &HFFFFFFFF)
'        Call gfx.DRAW_Line(CLng(a.Location.X), 0, CLng(a.Location.X), 480, -8, &HFFFFFFFF)

        If gameInput.Key(Key_Space) Then
            Call a.setLocation(100, 100)
            Call b.setLocation(150, 100)
            Call C.setLocation(250, 50)
        End If

        If gameInput.Key_Hit(Key_Z) Then
            Call a.SetForce(-20, -80)
        End If
        
        If gameInput.Key_Hit(Key_F1) Then physic.DEBUG_DrawColliders = Not physic.DEBUG_DrawColliders

        If gameInput.Key_Hit(Key_Escape) Then Call Unload(Me)

        a.Enabled = True

        Call gfx.Primitives.DrawAdvBox(0, gfx.CurrentDisplayMode.Height - 100, gfx.CurrentDisplayMode.Width, gfx.CurrentDisplayMode.Height, -8, &H0, &H0, &HFF000000, &HFF000000)
        Call spr.Draw
        Call gfx.Primitives.WriteText(fnt, "Physic Engine Prototype - TLSA Engine - José Miguel Sánchez Fernández, 2009", 0, gfx.CurrentDisplayMode.Height - 12, -8, , Right)

        Call gfx.Primitives.WriteText(fnt, gfx.FPS & "fps", 0, 0, -8)
        Call gfx.Primitives.WriteText(fnt, msg, 0, 20, -8)
        
        'Call gfx.Primitives.WriteText(fnt, a.Hit, 0, 60, -8)

        Call gfx.Primitives.DrawBox(gfx.CurrentDisplayMode.Width - lblKeyMap.Width, 0, gfx.CurrentDisplayMode.Width - 1, lblKeyMap.Height, -8, &HFFFFFFFF, True, &H77000000)
        Call lblKeyMap.Draw

        ' Cursor grafico:
        Call gfx.Primitives.DrawLine(gameInput.Mouse.X, gameInput.Mouse.Y - 10, gameInput.Mouse.X, gameInput.Mouse.Y + 10, -8, &HFFFFFFFF)
        Call gfx.Primitives.DrawLine(gameInput.Mouse.X - 10, gameInput.Mouse.Y, gameInput.Mouse.X + 10, gameInput.Mouse.Y, -8, &HFFFFFFFF)

        ' Actualizamos la logica:
        Call gameInput.Update
        Call sim.Update

        ' Actualizamos el render:
        Call sim.Draw
        Call gfx.Render
    Loop
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set gfx = Nothing
    End
End Sub
