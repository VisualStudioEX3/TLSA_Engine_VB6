VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GamePadExt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Añade soporte para lectura del FOV y del segundo eje del gamepad y tambien soporte basico de ForceFeedBack.

Option Explicit

Private Const DI_Params = DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
Private Const BufferSize = 10

Private Type GameInputDevice
    Device As DirectInputDevice8                ' Instancia del dispositivo en DirectInput.
    Name As String                              ' Nombre del gamepad.
    State As DIJOYSTATE                         ' Devuelve los estados del gamepad.
    HasForceFeedback As Boolean                 ' Indica si el gamepad tiene funcion de Force Feedback.
    ForceFeedbackEffect As DirectInputEffect    ' Efecto de Force Feedback.
    IsXBox360Gamepad As Boolean                 ' Indica si se trata de un gamepad de Xbox 360 (se accedera via XInput).
    Xbox360Index As Long                        ' Indice del gamepad de Xbox360 (0 a 3)
    Buttons As Long                             ' Numero de botones (se toma en cuenta un maximo de 12)
    Axis As Long                                ' Numero de ejes (maximo 2).
    HasPOV As Boolean                           ' Tiene POV (solo se tendra en cuenta uno).
    Xbox360ForcefeedbackDuration As Long        ' Se utiliza para crear la duracion del efecto en el gamepad de Xbox360.
    XBox360LeftTrigger As Boolean               ' Indica si se esta pulsando el gatillo izquierdo del gamepad de Xbox360.
    XBox360RightTrigger As Boolean              ' Indica si se esta pulsando el gatillo derecho del gamepad de Xbox360.
                                                ' (para unificar con los gamepad normales, solo se tiene en cuenta si su valor es distinto de 0)
    HitPOV(17 To 24) As Byte                    ' Indica si se hizo pulsacion sobre un angulo del POV (0 - No, 1 - Si, 2 No pero se mantiene presionado ese angulo)
    HitXBox360Trigger(1) As Byte                ' Indica si se hizo pulsacion sobre los gatillos del gamepad de XBox360 (0 - No, 1 - Si, 2 o mas - No pero se mantiene presionado ese angulo)
End Type

Public Type GameInputInfo
    Name As String                              ' Nombre del gamepad.
    HasForceFeedback As Boolean                 ' Indica si el gamepad tiene funcion de Force Feedback.
    IsXBox360Gamepad As Boolean                 ' Indica si se trata de un gamepad de Xbox 360 (se accedera via XInput).
    Buttons As Long                             ' Numero de botones (se toma en cuenta un maximo de 12)
    Axis As Long                                ' Numero de ejes (maximo 2).
    HasPOV As Boolean                           ' Tiene POV (solo se tendra en cuenta uno).
End Type

Dim Dx As New DirectX8
Dim Di As DirectInput8

Dim diDevEnum As DirectInputEnumDevices8

Dim JoyCaps As DIDEVCAPS
Dim DiProp_Dead As DIPROPLONG
Dim DiProp_Range As DIPROPRANGE
Dim DiProp_Saturation As DIPROPLONG
Dim JoyCount As Long

Dim varGamepads() As GameInputDevice, varXbox360GamepadAssigned(3) As Boolean
Dim effectTimer(3) As New Core.Timer          ' Cronometro para controlar la duracion de tiempo del Force Feedback en el gamepad de Xbox360.

' Declaraciones de las funciones de VB6Xinput.dll - Wrapper de Xinput para Visual Basic 6.0 basado en el wrapper GMXinput.dll para Game Maker desarrollado por yellowsix (http://gmc.yoyogames.com/lofiversion/index.php/t360286.html)
Private Declare Function GMXinput_setRumble Lib "VB6Xinput.dll" Alias "_setRumble@24" (ByVal index As Double, ByVal left As Double, ByVal right As Double) As Double
Private Declare Function GMXinput_leftTrigger Lib "VB6Xinput.dll" Alias "_leftTrigger@8" (ByVal index As Double) As Double
Private Declare Function GMXinput_rightTrigger Lib "VB6Xinput.dll" Alias "_rightTrigger@8" (ByVal index As Double) As Double
Private Declare Function GMXinput_ctrlState Lib "VB6Xinput.dll" Alias "_getCtrlState@8" (ByVal index As Double) As Double

Friend Property Get HasForceFeedback(Gamepad As Long) As Boolean
    If JoyCount = 0 Then Exit Property
    HasForceFeedback = varGamepads(Gamepad).HasForceFeedback
End Property

Private Sub Class_Initialize()
    'On Error GoTo ErrOut
    Dim Params As CONST_DISCLFLAGS
    Dim diProp As DIPROPLONG
    
    Dim didoEnum As DirectInputEnumDeviceObjects
    Dim dido As DirectInputDeviceObjectInstance
    Dim i As Integer
    Dim sGuid As String
        
    Set Di = Dx.DirectInputCreate()

    diProp.lHow = DIPH_DEVICE
    diProp.lObj = 0
    diProp.lData = BufferSize

    'Inicializamos el control del joistick:
    Set diDevEnum = Di.GetDIDevices(DI8DEVCLASS_GAMECTRL, DIEDFL_ATTACHEDONLY And DIEDFL_FORCEFEEDBACK)

    JoyCount = diDevEnum.GetCount

    If Not JoyCount = 0 Then
        ReDim varGamepads(JoyCount - 1)

        'Set deadzone for X and Y axis to 10 percent of the range of travel:
        With DiProp_Dead
            .lData = 1000
            .lHow = DIPH_DEVICE 'DIPH_BYOFFSET
        End With
    
        'Set saturation zones for X and Y axis to 5 percent of the range:
        With DiProp_Saturation
            .lData = 9500
            .lHow = DIPH_DEVICE 'DIPH_BYOFFSET
        End With
    
        'NOTE Some devices do not let you set the range
        On Local Error Resume Next

        'Set range for all axes:
        With DiProp_Range
            .lHow = DIPH_DEVICE
            .lMin = 0
            .lMax = 10000
        End With
              
        ' Creamos el efecto de fuerza constante:
        Dim effectInfo As DIEFFECT
        
        With effectInfo
          .constantForce.lMagnitude = 10000
          .lGain = 10000
          .lDuration = &HFFFFFFFF
          .X = 0 '10000
          .Y = 0 '10000
          .lTriggerButton = DIEB_NOTRIGGER  ' No trigger button
        End With
    
        For i = 1 To JoyCount
            Set varGamepads(i - 1).Device = Di.CreateDevice(diDevEnum.GetItem(i).GetGuidInstance)
    
            With varGamepads(i - 1)
                Call .Device.GetCapabilities(JoyCaps)
                .Buttons = JoyCaps.lButtons
                .Axis = JoyCaps.lAxes
                .HasPOV = (JoyCaps.lPOVs > 0)
                .Name = .Device.GetDeviceInfo.GetProductName
                .IsXBox360Gamepad = (.Name = "Controller (XBOX 360 For Windows)") ' Es un gamepad de Xbox360, por lo tanto se utiliza XInput
                                                                                  ' para acceder a ciertos controles.
                If .IsXBox360Gamepad Then .Xbox360Index = GetAvailableXbox360Gamepad() ' Obtenemos el primer gamepad disponible.
            
                Call .Device.SetCommonDataFormat(DIFORMAT_JOYSTICK)
                
                On Error Resume Next
                DiProp_Dead.lObj = DIJOFS_X
                Call .Device.SetProperty("DIPROP_DEADZONE", DiProp_Dead)
                
                DiProp_Dead.lObj = DIJOFS_Y
                Call .Device.SetProperty("DIPROP_DEADZONE", DiProp_Dead)
                
                DiProp_Saturation.lObj = DIJOFS_X
                Call .Device.SetProperty("DIPROP_SATURATION", DiProp_Saturation)

                DiProp_Saturation.lObj = DIJOFS_Y
                Call .Device.SetProperty("DIPROP_SATURATION", DiProp_Saturation)
                
                Call .Device.SetProperty("DIPROP_RANGE", DiProp_Range)
                
                Call .Device.Acquire
                
                If .IsXBox360Gamepad Then
                    .HasForceFeedback = True
                Else
                    Call Err.Clear ' En caso de haber algun error menor lo eliminamos de memoria para poder hacer la comprobacion en la siguiente llamada
                    ' Si no se trata de un gamepad de XBox360 comprobamos si el gamepad tiene soporte de forcefeedback:
                    Set .ForceFeedbackEffect = .Device.CreateEffect("GUID_ConstantForce", effectInfo)
                    .HasForceFeedback = (Err.Number = 0)
                End If
                
                On Error GoTo ErrOut
            End With
        Next i
    End If
    Exit Sub
    
ErrOut:
    Set diDevEnum = Nothing
    Set didoEnum = Nothing
    Set dido = Nothing
    
    Call Err.Raise(GameInput.GAMEINPUT_EXCEPTION.INITIALIZATION_FAILED, , "Error al obtener informacion de los gamepad conectados al equipo. Error " & Err.Number & ": " & Err.Description)
End Sub

' Obtiene el primer gamepad Xbox360 libre que haya:
Private Function GetAvailableXbox360Gamepad() As Long
    Dim i As Long
    
    For i = 0 To 3
        ' Si el gamepad esta conectado y no esta asignado se asigna al perfil actual:
        If (GMXinput_ctrlState(i) = 0) And Not varXbox360GamepadAssigned(i) Then
            GetAvailableXbox360Gamepad = i
            varXbox360GamepadAssigned(i) = True ' Se marca como asignado.
            Exit Function
        End If
    Next
    
    GetAvailableXbox360Gamepad = -1 ' No hay ningun gamepad disponible.
End Function

' Activa el efecto de force feedback:
Friend Sub SetForceFeedBack(Gamepad As Long, Duration As Long, Force As Long)
    On Local Error GoTo ErrorHandler
    
    If JoyCount = 0 Then Exit Sub
    
    If Not varGamepads(Gamepad).HasForceFeedback Then Exit Sub
    
    If Not varGamepads(Gamepad).IsXBox360Gamepad Then
        Dim effect As DIEFFECT
        
        With varGamepads(Gamepad).ForceFeedbackEffect
            Call .Unload
            Call .GetParameters(effect)
            
            effect.lDuration = Duration * 1000
            Call .SetParameters(effect, DIEP_DURATION)
            
            effect.constantForce.lMagnitude = Core.Math.GetValueFromPercent(Force, 10000)
            Call .SetParameters(effect, DIEP_TYPESPECIFICPARAMS)
      
            Call .Start(1, 0)
        End With
    Else
        ' Codigo encargado de llamar al componente COM de XInput:
        Dim value As Double
        value = Core.Math.GetValueFromPercent(Force, Core.Math.MAX_INT_VALUE)
        Call GMXinput_setRumble(varGamepads(Gamepad).Xbox360Index, value, value)
        varGamepads(Gamepad).Xbox360ForcefeedbackDuration = Duration
        Call effectTimer(varGamepads(Gamepad).Xbox360Index).Reset   ' Reiniciamos el cronometro.
    End If
    
ErrorHandler:                                           'Handle any errors
    
    If Err.Number = DIERR_NOTEXCLUSIVEACQUIRED Then     'If the program loses exclusive use of the joystick,
        With varGamepads(Gamepad).Device
            .Unacquire                                  'Unacquire the joystick
            .SetCooperativeLevel gInput.hWnd, DISCL_BACKGROUND Or DISCL_EXCLUSIVE
                                                        'Set the cooperative level again
            .Acquire                                    'Acquire the joystick again
            Resume                                      'Resume execution on that line again, since the program
                                                        'now has control of the joystick
        End With
    End If
    
End Sub

' Detiene el efecto de force feedback:
Friend Sub StopForceFeedBack(Gamepad As Long)
    On Local Error GoTo ErrorHandler
    
    If JoyCount = 0 Then Exit Sub
    
    If Not varGamepads(Gamepad).HasForceFeedback Then Exit Sub
    
    If Not varGamepads(Gamepad).IsXBox360Gamepad Then
        Call varGamepads(Gamepad).ForceFeedbackEffect.Stop
    Else
        ' Codigo encargado de llamar al componente COM de XInput:
        Call GMXinput_setRumble(varGamepads(Gamepad).Xbox360Index, 0, 0)
        varGamepads(Gamepad).Xbox360ForcefeedbackDuration = 0
    End If
    
ErrorHandler:                                           'Handle any errors
    
    If Err.Number = DIERR_NOTEXCLUSIVEACQUIRED Then     'If the program loses exclusive use of the joystick,
        With varGamepads(Gamepad).Device
            .Unacquire                                  'Unacquire the joystick
            .SetCooperativeLevel gInput.hWnd, DISCL_BACKGROUND Or DISCL_EXCLUSIVE
                                                        'Set the cooperative level again
            .Acquire                                    'Acquire the joystick again
            Resume                                      'Resume execution on that line again, since the program
                                                        'now has control of the joystick
        End With
    End If
    
End Sub

' Lectura del segundo eje:
Friend Property Get SecondAxis(Gamepad As Long) As Core.POINT
    If JoyCount = 0 Then Exit Property
    If Not varGamepads(Gamepad).IsXBox360Gamepad Then
        SecondAxis.X = varGamepads(Gamepad).State.rz
        SecondAxis.Y = varGamepads(Gamepad).State.Z
    Else
        SecondAxis.X = varGamepads(Gamepad).State.rx
        SecondAxis.Y = varGamepads(Gamepad).State.ry
    End If
End Property

' Devuelve el angulo, en grados, del POV:
Friend Property Get POV(Gamepad As Long) As Single
    If JoyCount = 0 Then Exit Property
    Select Case varGamepads(Gamepad).State.POV(0)
        Case 0: POV = 0
        Case -1: POV = -1
        Case Else: POV = CSng(varGamepads(Gamepad).State.POV(0) / 100)
    End Select
End Property

' Actualiza los estados extra de los gamepads:
Friend Sub Update()
    If JoyCount = 0 Then Exit Sub
    Dim i As Long
    For i = 0 To UBound(varGamepads)
        With varGamepads(i)
            If IsConnected(i) Then
                Call .Device.GetDeviceStateJoystick(.State)
                
                If .IsXBox360Gamepad Then
                    ' Obtenemos los estados de los gatillos: (Evaluamos a partir de la mitad del recorrido del eje)
                    .XBox360LeftTrigger = (GMXinput_leftTrigger(.Xbox360Index) > 127)
                    .XBox360RightTrigger = (GMXinput_rightTrigger(.Xbox360Index) > 127)
                    
                    ' Eventos HIT en los gatillos:
                    If .XBox360LeftTrigger Then .HitXBox360Trigger(0) = .HitXBox360Trigger(0) + 1 Else .HitXBox360Trigger(0) = 0
                    If .XBox360RightTrigger Then .HitXBox360Trigger(1) = .HitXBox360Trigger(1) + 1 Else .HitXBox360Trigger(1) = 0
                    
                    ' Controlamos la duracion del efecto force feedback de los gamepad de Xbox360:
                    If .Xbox360ForcefeedbackDuration > 0 Then
                        If effectTimer(.Xbox360Index).value >= .Xbox360ForcefeedbackDuration Then
                            Call Me.StopForceFeedBack(i)
                            .Xbox360ForcefeedbackDuration = 0
                        End If
                    End If
                End If
            Else
                ' Devolveremos una excepcion para reaccionar al evento de desconexion del gamepad:
                Call Err.Raise(GameInput.GAMEINPUT_EXCEPTION.GAMEPAD_LOST_SIGNAL, , "Se ha perdido la señal del gamepad #" & i & ": '" & varGamepads(i).Name & "'")
            End If
        End With
    Next
End Sub

' Comprueba si un gamepad esta conectado:
Public Function IsConnected(Gamepad As Long) As Boolean
    If JoyCount = 0 Then Exit Function
    With varGamepads(Gamepad)
        If .IsXBox360Gamepad Then
            IsConnected = (GMXinput_ctrlState(.Xbox360Index) = 0)
        Else
            On Error Resume Next
            Call .Device.GetDeviceStateJoystick(varGamepads(Gamepad).State)
            IsConnected = (Err.Number = 0)
        End If
    End With
End Function

' Devuelve el numero de gamepads conectados al sistema:
Public Property Get GamepadCount() As Long
    GamepadCount = JoyCount
End Property

' Devuelve la informacion de un gamepad:
Public Sub GetGamepadInfo(Gamepad As Long, info As GameInput.GameInputInfo)
    If JoyCount = 0 Then Exit Sub
    With info
        .Name = varGamepads(Gamepad).Name
        .IsXBox360Gamepad = varGamepads(Gamepad).IsXBox360Gamepad
        .Axis = varGamepads(Gamepad).Axis
        .HasPOV = varGamepads(Gamepad).HasPOV
    End With
End Sub

Friend Function IsXBox360Gamepad(Gamepad As Long) As Boolean
    If JoyCount = 0 Then Exit Function
    IsXBox360Gamepad = varGamepads(Gamepad).IsXBox360Gamepad
End Function

Friend Function XBox360LeftTrigger(Gamepad As Long) As Boolean
    If JoyCount = 0 Then Exit Function
    XBox360LeftTrigger = varGamepads(Gamepad).XBox360LeftTrigger
End Function

Friend Function XBox360RightTrigger(Gamepad As Long) As Boolean
    If JoyCount = 0 Then Exit Function
    XBox360RightTrigger = varGamepads(Gamepad).XBox360RightTrigger
End Function

' Indica si se hizo pulsacion sobre un angulo concreto del POV
' Valor de IndexAngle: 17 = 0º, 18 = 45º, 19 = 90º, 20 = 135º, 21 = 180º, 22 = 225º, 23 = 270º, 24 = 315º
' El valor coincide con la constante definida en el diccionario del gamepad:
Friend Function HIT_AnglePOV(Gamepad As Long, IndexAngle As Long) As Boolean
    If JoyCount = 0 Then Exit Function
    HIT_AnglePOV = (varGamepads(Gamepad).HitPOV(IndexAngle) = 1)
End Function

' Indica si se hizo pulsacion sobre uno de los gatillos del gamepad de XBox360:
' Valor de Trigger: 0 = Izquierdo (LT), 1 = Derecho (RT)
Friend Function HIT_XBox360Trigger(Gamepad As Long, Trigger As Long) As Boolean
    If JoyCount = 0 Then Exit Function
    HIT_XBox360Trigger = (varGamepads(Gamepad).HitXBox360Trigger(Trigger) = 1)
End Function

Private Sub Class_Terminate()
    Set diDevEnum = Nothing
    Dim i As Long
    For i = 0 To 3
        Set effectTimer(i) = Nothing
    Next
End Sub
