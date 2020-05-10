VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Random"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Establece una semilla en el generador aleatorio:
Public Sub SetSeed(Value As Single)
    Call sys.MATH_RandSeed(Value)
End Sub

' Devuelve un valor aleatorio:
Public Function GetRandomNumber() As Single
    GetRandomNumber = Rnd()
End Function

' Devuelve un valor aleatorio comprendido en un rango definido:
Public Function GetRandomNumber2(MaxValue As Single, MinValue As Single) As Single
    GetRandomNumber2 = sys.MATH_Rand(MaxValue, MinValue)
End Function
