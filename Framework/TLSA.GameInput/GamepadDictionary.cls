VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GamepadDictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Valores genericos del gamepad:
Public Enum GamepadButtons
    Joy_None = 0
    
    ' Botones de accion:
    Joy_Button1 = 1
    Joy_Button2 = 2
    Joy_Button3 = 3
    Joy_Button4 = 4
    
    ' Traseros:
    Joy_Button5 = 5
    Joy_Button6 = 6
    Joy_Button7 = 7
    Joy_Button8 = 8
    
    ' Start y Select:
    Joy_Button9 = 9
    Joy_Button10 = 10
    
    ' Botones extra (seguramente correspondan con los sticks):
    Joy_Button11 = 11
    Joy_Button12 = 12
    
    ' Eventos precompilados de direcciones en el eje primario:
    Joy_Left = dxlib32_221.Joystick_Button.Left_Arrow
    Joy_Up = dxlib32_221.Joystick_Button.Up_Arrow
    Joy_Right = dxlib32_221.Joystick_Button.Right_Arrow
    Joy_Down = dxlib32_221.Joystick_Button.Down_Arrow
    
    ' Eventos precompilados de direcciones del POV:
    Joy_POV_0 = 17
    Joy_POV_45 = 18
    Joy_POV_90 = 19
    Joy_POV_135 = 20
    Joy_POV_180 = 21
    Joy_POV_225 = 22
    Joy_POV_270 = 23
    Joy_POV_315 = 24
End Enum

Private varName() As String, varValue() As dxlib32_221.Joystick_Button, varReserved() As Boolean, varCount As Long

' Valores negativos corresponden al raton:
Private Sub Add(Button As String, value As dxlib32_221.Joystick_Button, Optional Reserved As Boolean)
    ReDim Preserve varName(varCount), varValue(varCount), varReserved(varCount)
    varName(varCount) = Button: varValue(varCount) = value: varReserved(varCount) = Reserved
    varCount = varCount + 1
End Sub

' Inicializa el diccionario de teclas y sus claves:
Private Sub Class_Initialize()
    Call Add("Joy_None", 0)
    
    ' Siempre corresponden a los botones de accion:
    Call Add("Joy_Button1", 1)
    Call Add("Joy_Button2", 2)
    Call Add("Joy_Button3", 3)
    Call Add("Joy_Button4", 4)
    
    ' Generalmente Gatillos y botones traseros:
    Call Add("Joy_Button5", 5)
    Call Add("Joy_Button6", 6)
    Call Add("Joy_Button7", 7)
    Call Add("Joy_Button8", 8)
    
    ' Generalmente Start y Select:
    Call Add("Joy_Button9", 9)
    Call Add("Joy_Button10", 10)
    
    ' Botones extra (seguramente correspondan con los sticks):
    Call Add("Joy_Button11", 11)
    Call Add("Joy_Button12", 12)
    
    ' Eventos precompilados de direcciones en el eje primario:
    Call Add("Joy_Left", Left_Arrow)
    Call Add("Joy_Up", Up_Arrow)
    Call Add("Joy_Right", Right_Arrow)
    Call Add("Joy_Down", Down_Arrow)
    
    ' Eventos precompilados de direcciones en el POV:
    Call Add("Joy_POV_0", 17)
    Call Add("Joy_POV_45", 18)
    Call Add("Joy_POV_90", 19)
    Call Add("Joy_POV_135", 20)
    Call Add("Joy_POV_180", 21)
    Call Add("Joy_POV_225", 22)
    Call Add("Joy_POV_270", 23)
    Call Add("Joy_POV_315", 24)
End Sub

' Obtiene el valor de un boton:
Public Function GetValue(Button As String) As Long
    Dim i As Long
    For i = 0 To varCount - 1
        If UCase(Button) = UCase(varName(i)) Then
            GetValue = varValue(i)
            Exit Function
        End If
    Next
    Call Err.Raise(GameInput.GAMEINPUT_EXCEPTION.ARGUMENT_EXCEPTION, , "La clave no corresponde con ninguna clave del gamepad.")
End Function

' Obtiene la clave de un boton:
Public Function GetButton(value As Long) As String
    Dim i As Long
    For i = 0 To varCount - 1
        If value = varValue(i) Then
            GetButton = varName(i)
            Exit Function
        End If
    Next
    Call Err.Raise(GameInput.GAMEINPUT_EXCEPTION.ARGUMENT_EXCEPTION, , "El valor no corresponde con ninguna clave del gamepad.")
End Function

' Verifica que la clave exista en el diccionario:
Public Function ExistByName(Name As String) As Boolean
    Dim i As Long
    For i = 0 To varCount - 1
        If UCase(Name) = UCase(varName(i)) Then
            ExistByName = True
            Exit Function
        End If
    Next
End Function

' Verifica que la clave exista en el diccionario:
Public Function ExistByValue(value As Long) As Boolean
    Dim i As Long
    For i = 0 To varCount - 1
        If value = varValue(i) Then
            ExistByValue = True
            Exit Function
        End If
    Next
End Function
