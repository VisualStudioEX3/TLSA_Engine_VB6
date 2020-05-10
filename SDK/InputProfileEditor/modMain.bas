Attribute VB_Name = "modMain"
Option Explicit

Public varInput As GameInput.Manager        ' Instancia del gestor de dispostivos de entrada.
Public varProfile As GameInput.Profile      ' Instancia del perfil a configurar.

Public Action As String                     ' Nombre de la accion a crear o editar.
Public Keyb As String                       ' Nombre de la constante de teclado o raton.
Public Joy As String                        ' Nombre de la constante de joystick o gamepad.
Public NewAction As Boolean                 ' Indica si la accion se va a crear de cero o se va a modificar una existente.

' Carga las acciones en el listBox:
Public Sub LoadActionMap(List As ListBox)
    Dim act As GameInput.ActionNode
    Dim strName As String * 28
    Dim strKeyb As String * 16
    Dim strJoy As String * 16
    Call List.Clear
    For Each act In modMain.varProfile.Actions
        strName = act.Name
        strKeyb = modMain.varInput.KeyboadDictionary.GetKey(act.KeyboardValue)
        strJoy = modMain.varInput.GamepadDictionary.GetButton(act.GamepadValue)
        Call List.AddItem(strName & strKeyb & strJoy)
    Next
End Sub

' Carga todos los gamepads que esten conectados al equipo:
Public Sub LoadGamepads(List As ComboBox)
    Dim i As Long
    For i = 0 To modMain.varInput.GamepadCount - 1
        Call List.AddItem(modMain.varInput.GetGamepadName(i))
    Next
End Sub
