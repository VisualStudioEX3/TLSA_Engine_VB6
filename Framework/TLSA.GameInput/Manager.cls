VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Manager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Lista de excepciones tipificadas de la libreria:
Public Enum GAMEINPUT_EXCEPTION
    INITIALIZATION_FAILED = vbObjectError + 3000
    PROFILE_NOT_LOAD = vbObjectError + 3001
    PROFILE_NOT_FOUND = vbObjectError + 3002
    ITEM_KEY_DUPLICATED = 457 ' Error(457) Esta clave ya está asociada a un elemento de esta colección
    ITEM_KEY_NOT_FOUND = vbObjectError + 3003
    ARGUMENT_EXCEPTION = 5
    MAX_RANGE_EXCEDED = vbObjectError + 3004
    ACTION_LIST_EMPTY = vbObjectError + 3005
    GAMEPAD_LOST_SIGNAL = vbObjectError + 3006
    UNKNOWN_EXCEPTION = vbObjectError + 9999
End Enum

Private varProfiles As GameInput.Profiles

' Asocia la ventana al gestor de entrada:
Public Sub SetWindowHandle(hWnd As Long)
    Set gInput = New dx_Input_Class
    If gInput.Init(hWnd) Then
        Set varProfiles = New GameInput.Profiles        ' Inicializamos la lista de perfiles.
        Set modInput.gpadEx = New GameInput.GamePadExt  ' Inicializamos las extensiones para los gamepads.
    Else
        Call Err.Raise(GameInput.GAMEINPUT_EXCEPTION.INITIALIZATION_FAILED, , "No se pudo inicializar los subsistemas del gestor de dispositivos de entrada.")
    End If
End Sub

' Diccionario de claves del teclado y raton:
Public Property Get KeyboadDictionary() As GameInput.KeyboardDictionary
    Set KeyboadDictionary = modInput.dicKeybMouse
End Property

' Diccionario de claves del gamepad:
Public Property Get GamepadDictionary() As GameInput.GamepadDictionary
    Set GamepadDictionary = modInput.dicGamepad
End Property

' Lista de pefiles de configuracion de teclas:
Public Property Get Profiles() As GameInput.Profiles
    If varProfiles Is Nothing Then Exit Property
    Set Profiles = varProfiles
End Property

' Devuelve el numero de gamepads conectados:
Public Property Get GamepadCount() As Long
    GamepadCount = modInput.gpadEx.GamepadCount
End Property

' Devuelve el nombre del gamepad segun su posicion en la lista:
Public Function GetGamepadName(index As Long) As String
    Dim info As GameInput.GameInputInfo
    Call modInput.gpadEx.GetGamepadInfo(index, info)
    GetGamepadName = info.Name
End Function

Public Sub Update()
    If gInput Is Nothing Then Exit Sub
    Call gInput.Update
    Call modInput.gpadEx.Update
End Sub

Private Sub Class_Terminate()
    Set varProfiles = Nothing
    Call gInput.Terminate
End Sub
