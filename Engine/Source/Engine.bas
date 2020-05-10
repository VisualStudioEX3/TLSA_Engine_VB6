Attribute VB_Name = "Engine"
Option Explicit

Public Enum Engine_States
    Intro                                   ' Introduccion inicial.
    MainMenu                                ' Menu principal del juego.
    Scene_InGame                            ' Escena del juego.
    Scene_GameMenu                          ' Escena del juego con el menu de opciones intermedio.
End Enum

Public Enum MainMenu_States
    Desktop                                 ' Representa el nivel principal del menu del juego.
    Config_MainPanel                        ' Representa el panel de opciones y configuracion del juego.
    Config_InputPanel                       ' Representa el panel de configuracion de dispositivos de entrada.
    Credits                                 ' Representa el area de creditos.
    ExitConfirmation                        ' Representa el mensaje de confirmacion de salida del programa.
End Enum

' Instancias del framework de TLSA:
Public GraphicEngine As Graphics.Manager
Public AudioEngine As Audio.Manager
Public PhysicEngine As Physics.Manager
Public InputEngine As GameInput.Manager

Public Ready As Boolean                     ' Indica que el motor esta a punto y funcionando. Desactive esta variable para
                                            ' terminar la ejecucion del motor.
Public Scene As TLSA.ENG_Scene              ' Instancia de la escena.

Public PlayerInputEnabled As Boolean        ' Indica si los controles del jugador estan activados.

' Objetos de depuracion:
Public dev_input As GameInput.Profile
Public dev_debug As TLSA.DEBUG_MessagePool
'

Public EDIT_MODE As Boolean                 ' Indica que el modo de edicion esta activado.

Public PHYSICS_SEPARATE_THREAD As Boolean   ' Indica que el metodo de actualizacion del simulador de fisicas se ejecutara en un hilo independiente.
Public PHYSICS_DRAW_GUIDES As Boolean       ' Indica si el simulador de fisicas muestra las guias de los objetos y el escenario.
Public FPS_MAX As Long                      ' Indica el numero maximo de fotogramas por segundo con los que correra el motor.
Public FULL_SCREEN As Boolean               ' Indica si se ejecutara a pantalla completa.

' Punto de entrada de la aplicacion:
Public Sub Main()
    EDIT_MODE = True
    Editor.EditMode = True
    
    Call frmDebug.Show(vbModal)

    Ready = InitEngine()            ' Inicializamos la ejecucion del motor.
    
    Do While Ready
        ' Informacion de depuracion:
        Call dev_debug.AddMessage(GraphicEngine.FPS & "/" & GraphicEngine.MaxFrames & "fps")
        
        ' Actualizamos los estados del motor:
        Call AudioEngine.Update     ' Actualizamos los estados del motor de audio.
        Call InputEngine.Update     ' Actualizamos los estados del gestor de entrada.
        
        ' Actualizamos la logica de la escena:
        Call Scene.Update
        
        ' Dibujamos la escena:
        Call Scene.Draw
        Call dev_debug.Draw
        Call GraphicEngine.Render   ' Renderizamos la escena.
    Loop
    
    Call TerminateEngine            ' Terminamos la ejecucion del motor.
    
    End                             ' Forzamos la salida del programa.
End Sub

' Inicializa el motor:
Private Function InitEngine() As Boolean
    On Error GoTo ErrOut
    
    Call frmMain.Show               ' Cargamos la ventana del programa.
    
    ' Instancias del framework de TLSA:
    Set GraphicEngine = New Graphics.Manager
    Set AudioEngine = New Audio.Manager
    Set PhysicEngine = New Physics.Manager
    Set InputEngine = New GameInput.Manager
    
    ' Inicializamos el motor grafico:
    Call GraphicEngine.Initialize(frmMain.hwnd, 800, 600, 32, Not Engine.FULL_SCREEN, True)
    GraphicEngine.MaxFrames = Engine.FPS_MAX
    GraphicEngine.TextureFilter = None
    GraphicEngine.BackColor = Graphics.Color_Constant.White
    
    ' Asociamos el render grafico al motor de fisicas para poder dibujar las cajas de colision:
    Call Physics.SetGraphics(GraphicEngine)
    Physics.DEBUG_DrawColliders = Engine.PHYSICS_DRAW_GUIDES
    
    ' Inicializamos el motor de audio:
    Call AudioEngine.SetWindowHandle(frmMain.hwnd)

    ' Inicializamos el gestor de entrada de dispositivos:
    Call InputEngine.SetWindowHandle(frmMain.hwnd)
    
    ' Inicializamos el perfil de input para depurar:
    Set dev_input = InputEngine.Profiles.Create("dev", Player1, KeybAndMouse)
    Call dev_input.Import(App.Path & "\Data\InputProfiles\dev.prf")
    
    Set dev_debug = New TLSA.DEBUG_MessagePool
    
    Set Scene = New TLSA.ENG_Scene
    Scene.Visible = True
    Scene.Enabled = True
    Call Scene.LoadLevel("")
    
    ' Si el modo de edicion esta desactivado se cargara el nivel de pruebas:
    If Not EDIT_MODE Then Call Scene.LoadScene(App.Path & TLSA.ResourcePaths.Levels & "cp2k10_test.tlv")
    
    InitEngine = True
    Exit Function
    
ErrOut:
    Call MsgBox(Err.Number & ": " & Err.Description, vbCritical, "TLSA Engine: Error de inicializacion")
End Function

' Termina la ejecucion del motor:
Private Sub TerminateEngine()
    Call Unload(frmMain)            ' Destruimos la instancia de la ventana.
    
    Set Scene = Nothing
    
    ' Destruimos las instancias del TLSA.Framework:
    Set GraphicEngine = Nothing
    Set AudioEngine = Nothing
    Set PhysicEngine = Nothing
    Set InputEngine = Nothing
End Sub

' TEST: Metodo para separar en un hilo independiente la actualizacion de fisicas del motor:
Private Sub UpdatePhysics()
    If Engine.Scene.Enabled Then
        Call ThreadPhysicSimulator.Update
    End If
End Sub
