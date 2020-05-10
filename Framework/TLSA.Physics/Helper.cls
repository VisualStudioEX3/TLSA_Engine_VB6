VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Helper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Define una lista vacia de valores Z de descarte para ser utilizado en las funciones de trazado de rayos:
Public Property Get EmptyZDiscardArray() As Long(): End Function

' Crea un array de valores para descarte por coordenada Z utilizado en las funciones de trazado de rayos:
Public Function CreateZDiscardArray(ParamArray ZValues()) As Long()
    On Error GoTo ErrOut
    Dim ret() As Long: ReDim ret(UBound(ZValues))
    Dim i As Variant
    For i = 0 To UBound(ZValues)
        ret(i) = ZValues(i)
    Next
    CreateZDiscardArray = ret
    Exit Function
ErrOut:
    Call Err.Raise(9, , "La lista de valores esta vacia.")
End Function
