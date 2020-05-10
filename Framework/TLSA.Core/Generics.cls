VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Generics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Type POINT
    X As Long
    Y As Long
    Z As Long
    Color As Long
End Type

Public Type RECTANGLE
    X As Long
    Y As Long
    Width As Long
    Height As Long
End Type

Public Type VECTOR
    X As Single
    Y As Single
    Z As Single
End Type

Public Type Size
    Width As Long
    Height As Long
End Type

' *** Los Null<types> se utilizan para rellenar los parametros de salida de funciones que utilicen el tipo de dato y no se requiera
' devolver dato alguno. ***

' Devuelve una instancia vacia de POINT:
Public Property Get NullPOINT() As Core.POINT: End Property

' Devuelve una instancia vacia de RECTANGLE:
Public Property Get NullRECTANGLE() As Core.RECTANGLE: End Property

' Devuelve una instancia vacia de VECTOR:
Public Property Get NullVECTOR() As Core.VECTOR: End Property

' Devuelve una instancia vacia de SIZE:
Public Property Get NullSIZE() As Core.Size: End Property

' Convierte la estructura VECTOR a POINT:
Public Function VECTOR2POINT(v As Core.VECTOR) As Core.POINT
    Dim ver As Core.POINT
    
    ver.X = CLng(v.X): ver.Y = CLng(v.Y): ver.Z = CLng(v.Z)
    
    VECTOR2POINT = ver
End Function

' Convierte la estructura POINT a VECTOR:
Public Function POINT2VECTOR(v As Core.POINT) As Core.VECTOR
    Dim vec As Core.VECTOR
    
    vec.X = CSng(v.X): vec.Y = CSng(v.Y): vec.Z = CSng(v.Z)
    
    POINT2VECTOR = vec
End Function

' Compara dos estructuras POINT y devuelve verdadero si son iguales:
Public Function POINTCompare(A As Core.POINT, B As Core.POINT, Optional ZCompare As Boolean = False, Optional ColorCompare As Boolean = False) As Boolean
    POINTCompare = (A.X = B.X) And (A.Y = B.Y) And _
                    IIf(ZCompare, (A.Z = B.Z), True) And _
                    IIf(ColorCompare, (A.Color = B.Color), True)
End Function

' Compara dos estructuras VECTOR y devuelve verdadero si son iguales:
Public Function VECTORCompare(A As Core.VECTOR, B As Core.VECTOR, Optional ZCompare As Boolean = False) As Boolean
    VECTORCompare = (A.X = B.X) And (A.Y = B.Y) And IIf(ZCompare, (A.Z = B.Z), True)
End Function

' Compara dos estructuras RECTANGLE y devuelve verdadero si son iguales:
Public Function RECTANGLECompare(A As Core.RECTANGLE, B As Core.RECTANGLE) As Boolean
    RECTANGLECompare = (A.X = B.X) And (A.Y = B.Y) And (A.Width = B.Width) And (A.Height = B.Height)
End Function

' Compara dos tamaños:
Public Function SIZECompare(A As Core.Size, B As Core.Size) As Boolean
    SIZECompare = (A.Width = B.Width) And (A.Height = B.Height)
End Function

' Crea un punto:
Public Function CreatePOINT(X As Long, Y As Long, Optional Z As Long, Optional Color As Long) As Core.POINT
    With CreatePOINT
        .X = X
        .Y = Y
        .Z = Z
        .Color = Color
    End With
End Function

' Crea un vector:
Public Function CreateVECTOR(X As Single, Y As Single, Z As Single) As Core.VECTOR
    With CreateVECTOR
        .X = X
        .Y = Y
        .Z = Z
    End With
End Function

' Crea un rectangulo:
Public Function CreateRECTANGLE(X As Long, Y As Long, Width As Long, Height As Long) As Core.RECTANGLE
    With CreateRECTANGLE
        .X = X
        .Y = Y
        .Width = Width
        .Height = Height
    End With
End Function

Public Function CreateRECTANGLE2(Location As Core.POINT, Size As Core.Size) As Core.RECTANGLE
    With CreateRECTANGLE2
        .X = Location.X
        .Y = Location.Y
        .Width = Size.Width
        .Height = Size.Height
    End With
End Function

' Crea un tamaño:
Public Function CreateSIZE(Width As Long, Height As Long) As Core.Size
    With CreateSIZE
        .Width = Width
        .Height = Height
    End With
End Function

