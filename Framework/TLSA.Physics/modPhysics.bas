Attribute VB_Name = "modPhysics"
Option Explicit

Global indexKey As Long

Public Type PHXVECTOR
    X As Double
    Y As Double
    D As Long ' Distancia.
    b As Physics.Body ' Referencia del cuerpo del simulador. Se utilizara para devolver la referencia al cuerpo en las funciones de trazado de rayos.
End Type

Public gfx As Graphics.Manager
Public debugDrawColliders As Boolean
Public varHelper As New Physics.Helper

Public Function GetIndexKey() As Long
    indexKey = indexKey + 1
    GetIndexKey = indexKey
End Function

' Convierte la estructura PHXVECTOR a POINT de System:
Public Function PHXVECTOR2POINT(v As Physics.PHXVECTOR) As Core.POINT
    Dim ver As Core.POINT

    ver.X = CLng(v.X): ver.Y = CLng(v.Y): ver.Z = 0

    PHXVECTOR2POINT = ver
End Function

' Convierte la estructura POINT de System a PHXVECTOR:
Public Function POINT2PHXVECTOR(v As Core.POINT) As Physics.PHXVECTOR
    Dim vec As Physics.PHXVECTOR

    vec.X = CSng(v.X): vec.Y = CSng(v.Y)

    POINT2PHXVECTOR = vec
End Function

' Convierte la estructura PHXVECTOR a VECTOR de System:
Public Function PHXVECTOR2VECTOR(v As Physics.PHXVECTOR) As Core.VECTOR
    Dim ver As Core.VECTOR

    ver.X = CSng(v.X): ver.Y = CSng(v.Y): ver.Z = 0#

    PHXVECTOR2VECTOR = ver
End Function

' Convierte la estructura VECTOR de System a PHXVECTOR:
Public Function VECTOR2PHXVECTOR(v As Core.VECTOR) As Physics.PHXVECTOR
    Dim vec As Physics.PHXVECTOR

    vec.X = v.X: vec.Y = v.Y

    VECTOR2PHXVECTOR = vec
End Function
