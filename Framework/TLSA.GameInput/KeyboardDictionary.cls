VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "KeyboardDictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum KeyboardMouseButtons
    Key_None = 0 ' Tecla no asignada.
    Key_Escape = dxlib32_221.Key_Const.Key_Escape
    Key_F1 = dxlib32_221.Key_Const.Key_F1
    Key_F2 = dxlib32_221.Key_Const.Key_F2
    Key_F3 = dxlib32_221.Key_Const.Key_F3
    Key_F4 = dxlib32_221.Key_Const.Key_F4
    Key_F5 = dxlib32_221.Key_Const.Key_F5
    Key_F6 = dxlib32_221.Key_Const.Key_F6
    Key_F7 = dxlib32_221.Key_Const.Key_F7
    Key_F8 = dxlib32_221.Key_Const.Key_F8
    Key_F9 = dxlib32_221.Key_Const.Key_F9
    Key_F10 = dxlib32_221.Key_Const.Key_F10
    Key_F11 = dxlib32_221.Key_Const.Key_F11
    Key_F12 = dxlib32_221.Key_Const.Key_F12
    Key_BackSlash = dxlib32_221.Key_Const.Key_BackSlash
    Key_1 = dxlib32_221.Key_Const.Key_1
    Key_2 = dxlib32_221.Key_Const.Key_2
    Key_3 = dxlib32_221.Key_Const.Key_3
    Key_4 = dxlib32_221.Key_Const.Key_4
    Key_5 = dxlib32_221.Key_Const.Key_5
    Key_6 = dxlib32_221.Key_Const.Key_6
    Key_7 = dxlib32_221.Key_Const.Key_7
    Key_8 = dxlib32_221.Key_Const.Key_8
    Key_9 = dxlib32_221.Key_Const.Key_9
    Key_0 = dxlib32_221.Key_Const.Key_0
    Key_Equals = dxlib32_221.Key_Const.Key_Equals
    Key_Apostrophe = dxlib32_221.Key_Const.Key_Apostrophe
    Key_Slash = dxlib32_221.Key_Const.Key_Slash
    Key_Back = dxlib32_221.Key_Const.Key_Back
    Key_Tab = dxlib32_221.Key_Const.Key_Tab
    Key_Q = dxlib32_221.Key_Const.Key_Q
    Key_W = dxlib32_221.Key_Const.Key_W
    Key_E = dxlib32_221.Key_Const.Key_E
    Key_R = dxlib32_221.Key_Const.Key_R
    Key_T = dxlib32_221.Key_Const.Key_T
    Key_Y = dxlib32_221.Key_Const.Key_Y
    Key_U = dxlib32_221.Key_Const.Key_U
    Key_I = dxlib32_221.Key_Const.Key_I
    Key_O = dxlib32_221.Key_Const.Key_O
    Key_P = dxlib32_221.Key_Const.Key_P
    Key_Enter = dxlib32_221.Key_Const.Key_Enter
    Key_LBracket = dxlib32_221.Key_Const.Key_LBracket
    Key_RBracket = dxlib32_221.Key_Const.Key_RBracket
    Key_A = dxlib32_221.Key_Const.Key_A
    Key_S = dxlib32_221.Key_Const.Key_S
    Key_D = dxlib32_221.Key_Const.Key_D
    Key_F = dxlib32_221.Key_Const.Key_F
    Key_G = dxlib32_221.Key_Const.Key_G
    Key_H = dxlib32_221.Key_Const.Key_H
    Key_J = dxlib32_221.Key_Const.Key_J
    Key_K = dxlib32_221.Key_Const.Key_K
    Key_L = dxlib32_221.Key_Const.Key_L
    Key_Grave = dxlib32_221.Key_Const.Key_Grave
    Key_SemiColon = dxlib32_221.Key_Const.Key_SemiColon
    Key_LShift = dxlib32_221.Key_Const.Key_LShift
    Key_Minus = dxlib32_221.Key_Const.Key_Minus
    Key_Z = dxlib32_221.Key_Const.Key_Z
    Key_X = dxlib32_221.Key_Const.Key_X
    Key_C = dxlib32_221.Key_Const.Key_C
    Key_V = dxlib32_221.Key_Const.Key_V
    Key_B = dxlib32_221.Key_Const.Key_B
    Key_N = dxlib32_221.Key_Const.Key_N
    Key_M = dxlib32_221.Key_Const.Key_M
    Key_Comma = dxlib32_221.Key_Const.Key_Comma
    Key_Period = dxlib32_221.Key_Const.Key_Period
    '    Key_SysRQ - Imprimir pantalla (Impr Pant PetSis
    Key_RShift = dxlib32_221.Key_Const.Key_RShift
    Key_MinorMajor = dxlib32_221.Key_Const.Key_MinorMajor
    Key_LControl = dxlib32_221.Key_Const.Key_LControl
    Key_LAlt = dxlib32_221.Key_Const.Key_LAlt
    Key_Space = dxlib32_221.Key_Const.Key_Space
    Key_RAlt = dxlib32_221.Key_Const.Key_RAlt
    Key_RControl = dxlib32_221.Key_Const.Key_RControl
    Key_Pause = dxlib32_221.Key_Const.Key_Pause
    Key_Insert = dxlib32_221.Key_Const.Key_Insert
    Key_Delete = dxlib32_221.Key_Const.Key_Delete
    Key_Home = dxlib32_221.Key_Const.Key_Home
    Key_End = dxlib32_221.Key_Const.Key_End
    Key_PagUp = dxlib32_221.Key_Const.Key_PagUp
    Key_PagDown = dxlib32_221.Key_Const.Key_PagDown
    Key_Left = dxlib32_221.Key_Const.Key_Left
    Key_Up = dxlib32_221.Key_Const.Key_Up
    Key_Down = dxlib32_221.Key_Const.Key_Down
    Key_Right = dxlib32_221.Key_Const.Key_Right
    Key_NumpadEnter = dxlib32_221.Key_Const.Key_NumpadEnter
    Key_Decimal = dxlib32_221.Key_Const.Key_Decimal
    Key_Numpad0 = dxlib32_221.Key_Const.Key_Numpad0
    Key_Numpad1 = dxlib32_221.Key_Const.Key_Numpad1
    Key_Numpad2 = dxlib32_221.Key_Const.Key_Numpad2
    Key_Numpad3 = dxlib32_221.Key_Const.Key_Numpad3
    Key_Numpad4 = dxlib32_221.Key_Const.Key_Numpad4
    Key_Numpad5 = dxlib32_221.Key_Const.Key_Numpad5
    Key_Numpad6 = dxlib32_221.Key_Const.Key_Numpad6
    Key_Numpad7 = dxlib32_221.Key_Const.Key_Numpad7
    Key_Numpad8 = dxlib32_221.Key_Const.Key_Numpad8
    Key_Numpad9 = dxlib32_221.Key_Const.Key_Numpad9
    Key_Add = dxlib32_221.Key_Const.Key_Add
    Key_SubsTract = dxlib32_221.Key_Const.Key_SubsTract
    Key_Multiply = dxlib32_221.Key_Const.Key_Multiply
    Key_Divide = dxlib32_221.Key_Const.Key_Divide
    
    ' Los valores negativos los utilizamos para identificar los botones y la rueda del raton:
    Mouse_Left = -1
    Mouse_Right = -2
    Mouse_Middle = -3
    Mouse_Wheel_Up = -4
    Mouse_Wheel_Down = -5
End Enum

Private varName() As String, varValue() As dxlib32_221.Key_Const, varReserved() As Boolean, varCount As Long

' Valores negativos corresponden al raton:
Private Sub Add(Key As String, value As dxlib32_221.Key_Const, Optional Reserved As Boolean)
    ReDim Preserve varName(varCount), varValue(varCount), varReserved(varCount)
    varName(varCount) = Key: varValue(varCount) = value: varReserved(varCount) = Reserved
    varCount = varCount + 1
End Sub

' Inicializa el diccionario de teclas y sus claves:
Private Sub Class_Initialize()
    Call Add("Key_None", 0)
    
    Call Add("Key_Escape", Key_Escape, True)
    Call Add("Key_F1", Key_F1, True)
    Call Add("Key_F2", Key_F2, True)
    Call Add("Key_F3", Key_F3, True)
    Call Add("Key_F4", Key_F4, True)
    Call Add("Key_F5", Key_F5, True)
    Call Add("Key_F6", Key_F6, True)
    Call Add("Key_F7", Key_F7, True)
    Call Add("Key_F8", Key_F8, True)
    Call Add("Key_F9", Key_F9, True)
    Call Add("Key_F10", Key_F10, True)
    Call Add("Key_F11", Key_F11, True)
    Call Add("Key_F12", Key_F12, True)
    Call Add("Key_BackSlash", Key_BackSlash)
    Call Add("Key_1", Key_1)
    Call Add("Key_2", Key_2)
    Call Add("Key_3", Key_3)
    Call Add("Key_4", Key_4)
    Call Add("Key_5", Key_5)
    Call Add("Key_6", Key_6)
    Call Add("Key_7", Key_7)
    Call Add("Key_8", Key_8)
    Call Add("Key_9", Key_9)
    Call Add("Key_0", Key_0)
    Call Add("Key_Equals", Key_Equals)
    Call Add("Key_Apostrophe", Key_Apostrophe)
    Call Add("Key_Slash", Key_Slash)
    Call Add("Key_Back", Key_Back)
    Call Add("Key_Tab", Key_Tab)
    Call Add("Key_Q", Key_Q)
    Call Add("Key_W", Key_W)
    Call Add("Key_E", Key_E)
    Call Add("Key_R", Key_R)
    Call Add("Key_T", Key_T)
    Call Add("Key_Y", Key_Y)
    Call Add("Key_U", Key_U)
    Call Add("Key_I", Key_I)
    Call Add("Key_O", Key_O)
    Call Add("Key_P", Key_P)
    Call Add("Key_Enter", Key_Enter)
    Call Add("Key_LBracket", Key_LBracket)
    Call Add("Key_RBracket", Key_RBracket)
    Call Add("Key_A", Key_A)
    Call Add("Key_S", Key_S)
    Call Add("Key_D", Key_D)
    Call Add("Key_F", Key_F)
    Call Add("Key_G", Key_G)
    Call Add("Key_H", Key_H)
    Call Add("Key_J", Key_J)
    Call Add("Key_K", Key_K)
    Call Add("Key_L", Key_L)
    Call Add("Key_Grave", Key_Grave, True)
    Call Add("Key_SemiColon", Key_SemiColon)
    Call Add("Key_LShift", Key_LShift)
    Call Add("Key_Minus", Key_Minus)
    Call Add("Key_Z", Key_Z)
    Call Add("Key_X", Key_X)
    Call Add("Key_C", Key_C)
    Call Add("Key_V", Key_V)
    Call Add("Key_B", Key_B)
    Call Add("Key_N", Key_N)
    Call Add("Key_M", Key_M)
    Call Add("Key_Comma", Key_Comma)
    Call Add("Key_Period", Key_Period)
    '    Key_SysRQ - Imprimir pantalla (Impr Pant PetSis)
    Call Add("Key_RShift", Key_RShift)
    Call Add("Key_MinorMajor", Key_MinorMajor)
    Call Add("Key_LControl", Key_LControl)
    Call Add("Key_LAlt", Key_LAlt)
    Call Add("Key_Space", Key_Space)
    Call Add("Key_RAlt", Key_RAlt)
    Call Add("Key_RControl", Key_RControl)
    Call Add("Key_Pause", Key_Pause, True)
    Call Add("Key_Insert", Key_Insert)
    Call Add("Key_Delete", Key_Delete)
    Call Add("Key_Home", Key_Home)
    Call Add("Key_End", Key_End)
    Call Add("Key_PagUp", Key_PagUp)
    Call Add("Key_PagDown", Key_PagDown)
    Call Add("Key_Left", Key_Left)
    Call Add("Key_Up", Key_Up)
    Call Add("Key_Down", Key_Down)
    Call Add("Key_Right", Key_Right)
    '    Key_NumLock - Bloque numérico.
    Call Add("Key_NumpadEnter", Key_NumpadEnter)
    Call Add("Key_Decimal", Key_Decimal)
    Call Add("Key_Numpad0", Key_Numpad0)
    Call Add("Key_Numpad1", Key_Numpad1)
    Call Add("Key_Numpad2", Key_Numpad2)
    Call Add("Key_Numpad3", Key_Numpad3)
    Call Add("Key_Numpad4", Key_Numpad4)
    Call Add("Key_Numpad5", Key_Numpad5)
    Call Add("Key_Numpad6", Key_Numpad6)
    Call Add("Key_Numpad7", Key_Numpad7)
    Call Add("Key_Numpad8", Key_Numpad8)
    Call Add("Key_Numpad9", Key_Numpad9)
    Call Add("Key_Add", Key_Add)
    Call Add("Key_SubsTract", Key_SubsTract)
    Call Add("Key_Multiply", Key_Multiply)
    Call Add("Key_Divide", Key_Divide)
    
    ' Los valores negativos los utilizamos para identificar los botones y la rueda del raton:
    Call Add("Mouse_Left", -1)
    Call Add("Mouse_Right", -2)
    Call Add("Mouse_Middle", -3)
    Call Add("Mouse_Wheel_Up", -4)
    Call Add("Mouse_Wheel_Down", -5)
End Sub

' Obtiene el valor de una tecla:
Public Function GetValue(Key As String) As Long
    Dim i As Long
    For i = 0 To varCount - 1
        If UCase(Key) = UCase(varName(i)) Then
            GetValue = varValue(i)
            Exit Function
        End If
    Next
    Call Err.Raise(GameInput.GAMEINPUT_EXCEPTION.ARGUMENT_EXCEPTION, , "La clave no corresponde con ninguna clave del teclado o el raton.")
End Function

' Obtiene la clave de una tecla:
Public Function GetKey(value As Long) As String
    Dim i As Long
    For i = 0 To varCount - 1
        If value = varValue(i) Then
            GetKey = varName(i)
            Exit Function
        End If
    Next
    Call Err.Raise(GameInput.GAMEINPUT_EXCEPTION.ARGUMENT_EXCEPTION, , "El valor no corresponde con ninguna clave del teclado o el raton.")
End Function

' Verifica que la clave exista en el diccionario:
Public Function ExistByName(Key As String) As Boolean
    Dim i As Long
    For i = 0 To varCount - 1
        If UCase(Key) = UCase(varName(i)) Then
            ExistByName = True
            Exit Function
        End If
    Next
End Function

' Verifica que la clave exista en el diccionario:
Public Function ExistByValue(value As Long) As Boolean
    Dim i As Long
    For i = 0 To varCount - 1
        If value = UCase(varValue(i)) Then
            ExistByValue = True
            Exit Function
        End If
    Next
End Function
