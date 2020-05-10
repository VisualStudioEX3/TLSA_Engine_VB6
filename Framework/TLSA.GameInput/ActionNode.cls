VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ActionNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private varName As String
Private varKeybMouse As Long
Private varJoyButton As Long

Public Property Get Name() As String
    Name = varName
End Property

Friend Property Let Name(value As String)
    varName = value
End Property

Public Property Get KeyboardValue() As Long
    KeyboardValue = varKeybMouse
End Property

Friend Property Let KeyboardValue(value As Long)
    varKeybMouse = value
End Property

Public Property Get GamepadValue() As Long
    GamepadValue = varJoyButton
End Property

Friend Property Let GamepadValue(value As Long)
    varJoyButton = value
End Property
