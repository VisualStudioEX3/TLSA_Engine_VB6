Attribute VB_Name = "modInput"
Option Explicit

Public gInput As New dx_Input_Class

Public dicKeybMouse As New GameInput.KeyboardDictionary
Public dicGamepad As New GameInput.GamepadDictionary

Public gpadEx As GameInput.GamePadExt

' Convierte la estructura Vertex de dx_lib32 a POINT:
Public Function Vertex2POINT(v As dxlib32_221.Vertex) As Core.POINT
    Dim pt As Core.POINT
    
    pt.X = v.X
    pt.Y = v.Y
    pt.Z = v.Z
    pt.Color = v.Color
    
    Vertex2POINT = pt
End Function
