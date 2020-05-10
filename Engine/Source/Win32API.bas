Attribute VB_Name = "Win32API"
Option Explicit

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, ByRef lpRect As RECT) As Long
Public Declare Function GetClientRect Lib "user32.dll" (ByVal hwnd As Long, ByRef lpRect As RECT) As Long
Public Declare Function SubtractRect Lib "user32.dll" (lprcDst As RECT, lprcSrc1 As RECT, lprcSrc2 As RECT) As Long
Public Declare Function GetDesktopWindow Lib "user32.dll" () As Long

