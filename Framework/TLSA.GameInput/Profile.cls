VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Profile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const PROFILE_HEADER As String = "// TLSA.Gameinput Profile"
Private Const PROFILE_MAPINFO As String = "// Action=KeybMouse;Joy"

' Indice de jugadores. Esto se utiliza bien para definir a que jugador se asociaria el perfil (un maximo de 4 jugadores):
Public Enum PlayerIndex
    Player1 = 1
    Player2 = 2
    Player3 = 3
    Player4 = 4
End Enum

Public Enum InputDevice                             ' Dispositivos de captura de entrada:
    KeybAndMouse                                    ' Teclado y raton.
    Gamepad                                         ' Gamepad o joystick.
End Enum

Public Type EffectParams
    Duration As Long                                ' Duracion del efecto (-1 para infinito)
    Force As Long                                   ' Fuerza del efecto.
End Type

Private varUser As String                           ' Usuario del perfil.
Private varDevice As GameInput.InputDevice
Private varJoyIndex As Long
Private varPlayerIndex As GameInput.PlayerIndex
Private varAction As New GameInput.ActionMap
Private varFFBack As GameInput.EffectParams
Private varXbox360Index As Long                     ' Almacena el indice del gamepad de Xbox360 asociado
                                                    ' al perfil (en caso de haber mas de uno conectado a la maquina).

Private Const VCURSOR_RANGE As Long = 255           ' Desplazamiento maximo del cursor virtual, se utilizara para hacer
                                                    ' el porcentaje de desplazamiento segun valor del eje de vision.
Private varVCursorEnabled As Boolean                ' Indica si el cursor virtual esta activado.
Private varVCursor As Core.POINT                  ' Coordenadas del cursor virtual.
Private varvCursorSen As Single                     ' Sensibilidad del cursor virtual.

Private varFilename As String                       ' Guarda el nombre y ruta del archivo del perfil si lo tuviera.

Public Property Get Filename() As String
    Filename = varFilename
End Property

' Importa un perfil desde archivo:
Public Sub Import(Filename As String)
    If Core.IO.FileExists(Filename) Then
        Dim file As Integer
        file = FreeFile
        
        Open Filename For Input As file
            Dim setUser As Boolean ' Indica si se encontro y asigno el nombre de usuario.
            Dim line As String, fields() As String, subfields() As String
            
            Do While Not EOF(file)
                Line Input #file, line ' Leemos la siguiente linea del archivo.
                
                If left(line, 2) <> "//" Then ' Evitamos las lineas de comentarios:
                    fields = Split(line, "=") ' Extraemos el campo nombre y el campo valor de la linea.
                    
                    If LCase(fields(0)) = "username" Then ' Si el campo nombre es igual a "username":
                        ' Asignamos el nombre de usuario al perfil:
                        varUser = fields(1)
                        Call varAction.Clear ' Eliminamos todas las acciones que ya estuvieran definidas.
                        setUser = True ' Se indica que se asigno el nombre de usuario.
                    Else
                        If Not setUser Then ' Si no se asigno el campo nombre de usuario el archivo esta incompleto:
                            Call Err.Raise(GameInput.GAMEINPUT_EXCEPTION.PROFILE_NOT_LOAD, , "El archivo de perfil no tiene asignado nombre de usuario: '" & Filename & "'")
                        Else
                            ' Añadimos la accion:
                            Call varAction.Add(fields(0))
                            
                            ' Obtenemos los valores de la linea leida:
                            subfields = Split(fields(1), ";")
    
                            ' Si existe el valor para el teclado se evalua y se asigna:
                            If (subfields(0) <> "") Then _
                                Call Me.SetActionButton(fields(0), modInput.dicKeybMouse.GetValue(subfields(0)), KeybAndMouse)
                                
                            ' Si existe el valor para el gamepad se evalua y se asigna:
                            If (subfields(1) <> "") Then _
                               Call Me.SetActionButton(fields(0), modInput.dicGamepad.GetValue(subfields(1)), Gamepad)
                        End If
                    End If
                End If
            Loop
        Close file
        varFilename = Filename
    Else
        Call Err.Raise(GameInput.GAMEINPUT_EXCEPTION.PROFILE_NOT_FOUND, , "No se encontro el perfil: '" & Filename & "'")
    End If
End Sub

' Exportar el perfil a un archivo:
Public Sub Export(Filename As String, Optional Overwrite As Boolean = True)
    ' Si ya existe un archivo con el mismo nombre, y la sobreescritura esta desactivada, generamos la excepcion:
    If Core.IO.FileExists(Filename) And Not Overwrite Then _
        Call Err.Raise(GameInput.GAMEINPUT_EXCEPTION.ARGUMENT_EXCEPTION, "Ya existe un perfil con el mismo nombre de archivo y en la misma ruta: '" & Filename & "'")

    Dim file As Integer
    file = FreeFile()
    
    ' Volcamos el contenido del mapa de acciones y el nombre de usuario al archivo:
    Open Filename For Output As file
        Print #file, PROFILE_HEADER
        Print #file, "username=" & varUser
        Print #file, PROFILE_MAPINFO
        Dim a As GameInput.ActionNode
        For Each a In varAction
            Print #file, a.Name & "=" & modInput.dicKeybMouse.GetKey(a.KeyboardValue) & ";" & modInput.dicGamepad.GetButton(a.GamepadValue)
        Next
    Close file
    varFilename = Filename
End Sub

Public Property Get Actions() As GameInput.ActionMap
    Set Actions = varAction
End Property

' Devuelve el nombre del usuario propietario del perfil:
Public Property Get UserName() As String
    UserName = varUser
End Property

' Establece el nombre del usuario propietario del perfil (generalmente para editar perfiles):
Public Property Let UserName(value As String)
    varUser = value
End Property

' Devuelve el indice de jugador asociado al perfil:
Public Property Get Player() As GameInput.PlayerIndex
    Player = varPlayerIndex
End Property

Friend Property Let Player(value As GameInput.PlayerIndex)
    varPlayerIndex = value
End Property

' Devuelve el dispositivo que esta usando para capturar la entrada:
Public Property Get Device() As GameInput.InputDevice
    Device = varDevice
End Property

' Indica al perfil que dispositivo se utilizara para capturar la entrada:
Public Property Let Device(value As GameInput.InputDevice)
    varDevice = value
End Property

' Indice del gamepad o joystick asignado al perfil:
Public Property Get GamepadUsed() As Long
    GamepadUsed = varJoyIndex
End Property

Public Property Let GamepadUsed(value As Long)
    varJoyIndex = value
End Property

' Establece el boton para un evento de accion:
Public Sub SetActionButton(Action As String, Button As Long, Device As GameInput.InputDevice)
    If varAction.Count > 0 Then
        ' Comprobamos que el valor del boton existe en los diccionarios:
        If modInput.dicGamepad.ExistByValue(Button) Or modInput.dicKeybMouse.ExistByValue(Button) Then
            
            ' Buscamos el evento que tenga el mismo boton asignado:
            Dim act As GameInput.ActionNode
            If Device = KeybAndMouse Then ' Raton y teclado:
                For Each act In varAction
                    If act.KeyboardValue = Button Then
                        ' Intercambiamos los botones entre las dos acciones con la otra para evitar duplicados:
                        act.KeyboardValue = varAction(Action).KeyboardValue
                        Exit For
                    End If
                Next
                
                ' Asignamos el boton a la accion indicada:
                varAction(Action).KeyboardValue = Button
                
            ElseIf Device = Gamepad Then ' Gamepad:
                For Each act In varAction
                    If act.GamepadValue = Button Then
                        ' Intercambiamos los botones entre las dos acciones con otra para evitar duplicados:
                        act.GamepadValue = varAction(Action).GamepadValue
                        Exit For
                    End If
                Next
                
                ' Asignamos el boton a la accion indicada:
                varAction(Action).GamepadValue = Button
                
            End If
            
        Else
            ' Se introdujo una constante o clave de teclado, raton o gamepad no valida:
            Call Err.Raise(GameInput.GAMEINPUT_EXCEPTION.ARGUMENT_EXCEPTION, , "La clave del boton asociado al evento no es valida.")
        End If
    Else
        Call Err.Raise(GameInput.GAMEINPUT_EXCEPTION.ACTION_LIST_EMPTY, , "La lista de acciones esta vacia.")
    End If
End Sub

' Devuelve el valor del boton asignado a un evento de accion:
Public Function GetActionButton(Action As String, Device As GameInput.InputDevice) As Long
    If varAction.Count > 0 Then
        If Device = KeybAndMouse Then ' Raton y teclado:
            GetActionButton = varAction(Action).KeyboardValue
        ElseIf Device = Gamepad Then ' Gamepad:
            GetActionButton = varAction(Action).GamepadValue
        End If
    Else
        Call Err.Raise(GameInput.GAMEINPUT_EXCEPTION.ACTION_LIST_EMPTY, , "La lista de acciones esta vacia.")
    End If
End Function

' Indica si un boton de accion esta siendo presionado:
Public Function Press(Action As String) As Boolean
    If varAction.Count > 0 Then
        If varDevice = KeybAndMouse Then
            If varAction(Action).KeyboardValue > 0 Then  ' Teclado:
                Press = gInput.Key(varAction(Action).KeyboardValue)
            ElseIf varAction(Action).KeyboardValue < 0 Then ' Raton:
                Press = (gInput.Mouse_Press() = Abs(varAction(Action).KeyboardValue))
            End If
        ElseIf (varJoyIndex > -1 And varJoyIndex < gpadEx.GamepadCount) And varDevice = Gamepad Then  ' Gamepad:
            ' Si no se asigno boton a la accion no evaluamos:
            If varAction(Action).GamepadValue <> 0 Then
                ' Si es un gamepad de XBox360 comprobamos los gatillos:
                If modInput.gpadEx.IsXBox360Gamepad(varJoyIndex) Then
                    If varAction(Action).GamepadValue = 11 Then
                        Press = modInput.gpadEx.XBox360LeftTrigger(varJoyIndex)
                    ElseIf varAction(Action).GamepadValue = 12 Then
                        Press = modInput.gpadEx.XBox360RightTrigger(varJoyIndex)
                    End If
                ' Si el valor del boton asignado es mayor que 12 y menor de 17 se trata de una direccion precompilada de dx_lib32:
                ElseIf varAction(Action).GamepadValue > 12 And varAction(Action).GamepadValue < 17 Then
                    Select Case varAction(Action).GamepadValue
                        Case dxlib32_221.Joystick_Button.Left_Arrow: Press = (modInput.gInput.Joystick(varJoyIndex).X < 2500)
                        Case dxlib32_221.Joystick_Button.Right_Arrow: Press = (modInput.gInput.Joystick(varJoyIndex).X > 7500)
                        Case dxlib32_221.Joystick_Button.Down_Arrow: Press = (modInput.gInput.Joystick(varJoyIndex).Y > 7500)
                        Case dxlib32_221.Joystick_Button.Up_Arrow: Press = (modInput.gInput.Joystick(varJoyIndex).Y < 2500)
                    End Select
                ' Si el valor del boton asignado a la accion es mayor de 16 se trata de un angulo del POV:
                ElseIf varAction(Action).GamepadValue > 16 Then
                    ' Comprobamos el angulo del POV:
                    Select Case modInput.gpadEx.POV(varJoyIndex)
                        Case 0: Press = (varAction(Action).GamepadValue = 17)
                        Case 45: Press = (varAction(Action).GamepadValue = 18)
                        Case 90: Press = (varAction(Action).GamepadValue = 19)
                        Case 135: Press = (varAction(Action).GamepadValue = 20)
                        Case 180: Press = (varAction(Action).GamepadValue = 21)
                        Case 225: Press = (varAction(Action).GamepadValue = 22)
                        Case 270: Press = (varAction(Action).GamepadValue = 23)
                        Case 315: Press = (varAction(Action).GamepadValue = 24)
                    End Select
                Else
                    ' Comprobamos si el boton pulsado es el mismo que el de la accion:
                    Press = (gInput.Joystick(varJoyIndex).Button(varAction(Action).GamepadValue))
                End If
            End If
        End If
    Else
        Call Err.Raise(GameInput.GAMEINPUT_EXCEPTION.ACTION_LIST_EMPTY, , "La lista de acciones esta vacia.")
    End If
End Function

' Indica si un boton ha sido pulsado:
Public Function Hit(Action As String) As Boolean
    If varAction.Count > 0 Then
        If varDevice = KeybAndMouse Then
            If varAction(Action).KeyboardValue > 0 Then  ' Teclado:
                Hit = gInput.Key_Hit(varAction(Action).KeyboardValue)
            Else ' Raton:
                If varAction(Action).KeyboardValue = GameInput.KeyboardMouseButtons.Mouse_Wheel_Up Then
                    Hit = gInput.Mouse.Z > 0
                ElseIf varAction(Action).KeyboardValue = GameInput.KeyboardMouseButtons.Mouse_Wheel_Down Then
                    Hit = gInput.Mouse.Z < 0
                Else
                    Hit = gInput.Mouse_Hit(Abs(varAction(Action).KeyboardValue))
                End If
            End If
        ElseIf (varJoyIndex > -1 And varJoyIndex < gpadEx.GamepadCount) And varDevice = Gamepad Then  ' Gamepad:
            ' Si no se asigno boton a la accion no evaluamos:
            If varAction(Action).GamepadValue <> 0 Then
                ' Si es un gamepad de XBox360 comprobamos los gatillos:
                If modInput.gpadEx.IsXBox360Gamepad(varJoyIndex) Then
                    If varAction(Action).GamepadValue = 11 Then
                        Hit = modInput.gpadEx.HIT_XBox360Trigger(varJoyIndex, 0)
                    ElseIf varAction(Action).GamepadValue = 12 Then
                        Hit = modInput.gpadEx.HIT_XBox360Trigger(varJoyIndex, 1)
                    End If
                ' Si el valor del boton asignado a la accion es mayor de 16 se trata de un angulo del POV:
                ElseIf varAction(Action).GamepadValue > 16 Then
                    Hit = modInput.gpadEx.HIT_AnglePOV(varJoyIndex, varAction(Action).GamepadValue)
                ' Comprobamos el resto de botones:
                Else
                    Hit = gInput.Joystick_Hit(varJoyIndex, varAction(Action).GamepadValue)
                End If
            End If
        End If
    Else
        Call Err.Raise(GameInput.GAMEINPUT_EXCEPTION.ACTION_LIST_EMPTY, , "La lista de acciones esta vacia.")
    End If
End Function

' Devuelve el codigo ASCII del ultimo caracter entrado por teclado:
Public Function GetASCII() As Integer
    GetASCII = gInput.Key_ASCII()
End Function

' Eje principal. Se asocia al eje principal del gamepad:
Public Property Get GamepadAxis() As Core.POINT
    Dim ret As Core.POINT
    If (varJoyIndex > -1 And varJoyIndex < gpadEx.GamepadCount) And varDevice = Gamepad Then
        'GamepadAxis = Core.Generics.CreatePOINT(gInput.Joystick(varJoyIndex).X, gInput.Joystick(varJoyIndex).Y)
        ret.X = gInput.Joystick(varJoyIndex).X: ret.Y = gInput.Joystick(varJoyIndex).Y
        GamepadAxis = ret
    End If
End Property

' Eje de vista. Se asocia al raton y al eje secundario del gamepad si estuviera presente:
Public Property Get ViewAxis() As Core.POINT
    If varDevice = KeybAndMouse Then ' Raton:
        'ViewAxis =  Core.Generics.CreatePOINT(gInput.Mouse.X, gInput.Mouse.Y, gInput.Mouse.Z)
        ViewAxis.X = gInput.Mouse.X: ViewAxis.Y = gInput.Mouse.Y: ViewAxis.Z = gInput.Mouse.Z
    ElseIf (varJoyIndex > -1 And varJoyIndex < gpadEx.GamepadCount) And varDevice = Gamepad Then ' Gamepad:
        If varVCursorEnabled Then ' Actualizamos la posicion del eje secundario:
            With modInput.gpadEx.SecondAxis(varJoyIndex)
                ' Calculamos el movimiento del cursor virtual:
                varVCursor.X = varVCursor.X + Core.Math.GetPercentFromValue((.X - 5000) / varvCursorSen, 10000)
                varVCursor.Y = varVCursor.Y + Core.Math.GetPercentFromValue((.Y - 5000) / varvCursorSen, 10000)
                
                ' Corregimos si es necesario la posicion de X:
                If varVCursor.X <= 0 Then
                    varVCursor.X = 0
                ElseIf varVCursor.X > (Screen.Width \ Screen.TwipsPerPixelX) Then
                    varVCursor.X = (Screen.Width \ Screen.TwipsPerPixelX)
                End If
                
                ' Corregimos si es necesario la posicion de Y:
                If varVCursor.Y <= 0 Then
                    varVCursor.Y = 0
                ElseIf varVCursor.Y > (Screen.Height \ Screen.TwipsPerPixelY) Then
                    varVCursor.Y = (Screen.Height \ Screen.TwipsPerPixelY)
                End If
            End With
            
            ' Devolvemos el valor del cursor virtual:
            ViewAxis = varVCursor
            
        Else ' Devolvemos el valor del eje secundario:
            ViewAxis = modInput.gpadEx.SecondAxis(varJoyIndex)
        End If
    End If
End Property

' Devuelve el angulo de eje de vision. En caso de ser el raton se le debera indicar las coordenadas de origen, de lo contrario se tomara la coordenada 0,0:
Public Function ViewAngle(Optional X As Long, Optional Y As Long) As Single
    Dim a As Core.POINT, b As Core.POINT
    If varDevice = KeybAndMouse Then
        a.X = X: a.Y = Y: b.X = gInput.Mouse.X: b.Y = gInput.Mouse.Y
        ViewAngle = Core.Math.GetAngle(a, b)
        'ViewAngle = Core.Math.GetAngle(Core.Generics.CreatePOINT(X, Y), Core.Generics.CreatePOINT(gInput.Mouse.X, gInput.Mouse.Y))
    ElseIf (varJoyIndex > -1 And varJoyIndex < gpadEx.GamepadCount) And varDevice = Gamepad Then
        a.X = 5000: a.Y = 5000: 'b.X = gInput.Mouse.X: b.Y = gInput.Mouse.Y
        If Core.Generics.POINTCompare(a, modInput.gpadEx.SecondAxis(varJoyIndex)) Then
        'If Core.Generics.POINTCompare(Core.Generics.CreatePOINT(5000, 5000), modInput.gpadEx.SecondAxis(varJoyIndex)) Then
            ViewAngle = -1
        Else
            ViewAngle = Core.Math.GetAngle(a, modInput.gpadEx.SecondAxis(varJoyIndex))
            'ViewAngle = Core.Math.GetAngle(Core.Generics.CreatePOINT(5000, 5000), modInput.gpadEx.SecondAxis(varJoyIndex))
        End If
    End If
End Function

' Activa el cursor virtual. El cursor virtual convierte el segundo eje del gamepad, si estuviera presente, en un cursor similar al
' del raton manteniendo su posicion y utilizando el espacio de coordenadas de la pantalla, al igual que el raton:
Public Property Let VirtualCursor(Enabled As Boolean)
    varVCursorEnabled = Enabled
    varVCursor.X = 0: varVCursor.Y = 0: varVCursor.Z = 0
    'varVCursor = Core.Generics.NullPOINT
End Property

' Indica si el cursor virtual esta activado:
Public Property Get VirtualCursor() As Boolean
    VirtualCursor = varVCursorEnabled
End Property

' Sensibilidad del cursor virtual. Define el rango pixeles que recorrera el cursor virtual segun el valor del eje secundario:
Public Property Let VirtualCursorSensitivity(value As Single)
    varvCursorSen = value
End Property

' Devuelve la sensibilidad del cursor virtual:
Public Property Get VirtualCursorSensitivity() As Single
    VirtualCursorSensitivity = varvCursorSen
End Property

' Devuelve el angulo en grados del POV:
Public Property Get POV() As Long
    If (varJoyIndex > -1 And varJoyIndex < gpadEx.GamepadCount) And varDevice = Gamepad Then POV = modInput.gpadEx.POV(varJoyIndex)
End Property

' Establece la posicion del cursor del raton (afecta al eje primario solo si esta configurado como teclado y raton):
Public Sub SetCursorPosition(X As Long, Y As Long)
    If varDevice = KeybAndMouse Then Call gInput.Mouse_SetPossition(X, Y)
End Sub

' Indica si el gamepad dispone de soporte de vibracion:
Public Property Get VibrationSupport() As Boolean
    If (varJoyIndex > -1 And varJoyIndex < gpadEx.GamepadCount) And varDevice = Gamepad Then VibrationSupport = modInput.gpadEx.HasForceFeedback(varJoyIndex)
End Property

Public Property Get Vibration() As GameInput.EffectParams
    Vibration = varFFBack
End Property

Public Property Let Vibration(value As GameInput.EffectParams)
    If (varJoyIndex > -1 And varJoyIndex < gpadEx.GamepadCount) And varDevice = Gamepad Then Call SetVibration(value.Duration, value.Force)
End Property

' Aplica la fuerza en los ejes de vibracion del gamepad si los tuviera:
Public Sub SetVibration(Duration As Long, Force As Long)
    If (varJoyIndex > -1 And varJoyIndex < gpadEx.GamepadCount) And varDevice = Gamepad Then
        varFFBack.Duration = Duration
        varFFBack.Force = Force
        Call modInput.gpadEx.SetForceFeedBack(varJoyIndex, Duration, Force)
    End If
End Sub

' Desactiva la vibracion:
Public Sub StopVibration()
    If (varJoyIndex > -1 And varJoyIndex < gpadEx.GamepadCount) And varDevice = Gamepad Then
        varFFBack.Duration = 0
        varFFBack.Force = 0
        Call modInput.gpadEx.StopForceFeedBack(varJoyIndex)
    End If
End Sub

' Indica si el gamepad asociado esta conectado:
Public Function IsConnected() As Boolean
    If (varJoyIndex > -1 And varJoyIndex < gpadEx.GamepadCount) And varDevice = Gamepad Then IsConnected = modInput.gpadEx.IsConnected(varJoyIndex)
End Function

' Obtiene el valor de la tecla o boton pulsado segun dispositivo de entrada indicado:
Public Function Capture(Device As GameInput.InputDevice) As Long
    Dim ret As Long
    If Device = KeybAndMouse Then
        ret = gInput.Key_Press ' Obtenemos la tecla presionada en el teclado.
        
        ' Si no se presiono ninguna tecla del teclado comprobamos el raton:
        If ret = 0 Then ret = Core.Math.Invert(gInput.Mouse_Press())
        
        ' Si no se presiono ningun boton del raton comprobamos la rueda:
        If ret = 0 Then
            If gInput.Mouse.Z > 0 Then
                ret = -4
            ElseIf gInput.Mouse.Z < 0 Then
                ret = -5
            End If
        End If
    ElseIf (varJoyIndex > -1 And varJoyIndex < gpadEx.GamepadCount) And Device = Gamepad Then
        ret = gInput.Joystick_Press(varJoyIndex)
        
        If ret = 0 Then
            ' Si es un gamepad de XBox360 comprobamos los gatillos:
            If modInput.gpadEx.IsXBox360Gamepad(varJoyIndex) Then
                If modInput.gpadEx.XBox360LeftTrigger(varJoyIndex) Then
                    ret = 11
                ElseIf modInput.gpadEx.XBox360RightTrigger(varJoyIndex) Then
                    ret = 12
                End If
            End If
            
            ' Si el valor sigue siendo 0 comprobamos los angulos del POV:
            If ret = 0 Then
                ' Comprobamos el angulo del POV:
                Select Case modInput.gpadEx.POV(varJoyIndex)
                    Case 0: ret = 17
                    Case 45: ret = 18
                    Case 90: ret = 19
                    Case 135: ret = 20
                    Case 180: ret = 21
                    Case 225: ret = 22
                    Case 270: ret = 23
                    Case 315: ret = 24
                End Select
            End If
        End If
    End If
    
    Capture = ret
End Function

Private Sub Class_Initialize()
    varJoyIndex = -1
    varvCursorSen = 3#
End Sub
