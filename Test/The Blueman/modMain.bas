Attribute VB_Name = "modMain"
Option Explicit

Public gfx As Graphics.Manager
Public phx As Physics.Manager
Public sim As Physics.Simulator
Public gameInput As dx_Input_Class

Public Entities As Collection

Private boom As Physics.Explosion

Private varDrawPhx As Boolean

Public Sub Initialize()
    Set gfx = New Graphics.Manager
    Call gfx.Initialize(frmMain.hWnd, 800, 600, 32, True, True)
    gfx.BackColor = Graphics.Color_Constant.Orange
    gfx.MaxFrames = 60
    gfx.TextureFilter = None
    
    Call LoadContent
    
    Set phx = New Physics.Manager
    Call phx.SetGraphics(gfx)
    phx.DEBUG_DrawColliders = True
    
    Set sim = New Physics.Simulator
    sim.Delay = 24
    Call sim.SetGravity(0, 8)
    Call sim.SetWorkArea(-800, -600, 1600, 1200)
    sim.Enabled = True
    
    
    Set boom = sim.CreateExplosionEmitter(System.Generics.CreatePOINT(250, 250), 50, 100)
    
    Set gameInput = New dx_Input_Class
    Call gameInput.Init(frmMain.hWnd)
    
    Set Entities = New Collection
End Sub

Private Sub LoadContent()
    Call modMain.gfx.Textures.LoadTexture(App.Path & "\blueman.png", "blueman", False)
    Call modMain.gfx.Textures.LoadTexture(App.Path & "\tile100x100.png", "block", False)
End Sub

Public Sub Terminate()
    Set Entities = Nothing
    Call gameInput.Terminate
    Set gameInput = Nothing
    Set sim = Nothing
    Set phx = Nothing
    Set gfx = Nothing
End Sub

Public Sub Update()
    Dim ent As Object
    For Each ent In Entities
        Call ent.Update
    Next
    Call sim.Update
    ' -(x + ScreenWidth \ 2) -(y + ScreenHeight \ 2)
    Call gfx.SetOffSet(-(Entities("player").Location.X) + (gfx.CurrentDisplayMode.Width \ 2), -(Entities("player").Location.Y) + (gfx.CurrentDisplayMode.Height \ 2))
    
    If gameInput.Key_Hit(Key_F1) Then phx.DEBUG_DrawColliders = Not phx.DEBUG_DrawColliders ' varDrawPhx = Not varDrawPhx
    
    If gameInput.Mouse_Hit(Left_Button) Then
        'boom.Location = System.Generics.POINT2VECTOR(System.Generics.CreatePOINT(-gameInput.Mouse.X + gfx.Offset.X, -gameInput.Mouse.Y + gfx.Offset.Y))
        Call boom.Explode
    End If
    
    Call gameInput.Update
End Sub

Public Sub Draw()
    Dim ent As Object
    For Each ent In Entities
        Call ent.Draw
    Next
    'If varDrawPhx Then Call sim.Draw
    Call sim.Draw
    Call gfx.Primitives.WriteText(gfx.Fonts("SYSTEM"), gfx.FPS & "fps", 0, 0, -8, &HFFFFFFFF, Left, True)
    
    ' Cursor grafico:
    Call gfx.Primitives.DrawLine(CLng(boom.Location.X), CLng(boom.Location.Y) - 10, CLng(boom.Location.X), CLng(boom.Location.Y) + 10, -8, &HFFFFFFFF)
    Call gfx.Primitives.DrawLine(CLng(boom.Location.X) - 10, CLng(boom.Location.Y), CLng(boom.Location.X) + 10, CLng(boom.Location.Y), -8, &HFFFFFFFF)

    
    Call gfx.Render
End Sub
