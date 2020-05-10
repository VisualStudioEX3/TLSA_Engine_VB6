VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Math"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private varRandom As New Core.Random

Public Property Get PI() As Single: PI = 3.14159265358979: End Property

Public Property Get MAX_BYTE_VALUE() As Byte: MAX_BYTE_VALUE = 255: End Property
Public Property Get MIN_BYTE_VALUE() As Byte: MIN_BYTE_VALUE = 0: End Property

Public Property Get MAX_INT_VALUE() As Integer: MAX_INT_VALUE = 32767: End Property
Public Property Get MIN_INT_VALUE() As Integer: MIN_INT_VALUE = -32768: End Property

Public Property Get MAX_LONG_VALUE() As Long: MAX_LONG_VALUE = 2147483647: End Property
Public Property Get MIN_LONG_VALUE() As Long: MIN_LONG_VALUE = -2147483648#: End Property

Public Property Get MAX_SINGLE_VALUE() As Single: MAX_SINGLE_VALUE = 3.4028235E+38: End Property ' 1.797693E+38: End Property
Public Property Get MIN_SINGLE_VALUE() As Single: MIN_SINGLE_VALUE = -3.4028235E+38: End Property ' -1.797693E+38: End Property

Public Property Get MAX_DOUBLE_VALUE() As Double: MAX_DOUBLE_VALUE = 1.79769313486231E+308: End Property
Public Property Get MIN_DOUBLE_VALUE() As Double: MIN_DOUBLE_VALUE = -1.79769313486231E+308: End Property

' Instancia del generador de numeros aleatorios:
Public Property Get Random() As Core.Random
    Set Random = varRandom
End Property

' Determina si dos rectangulos se intersectan entre si:
Public Function IntersectRect(A As Core.RECTANGLE, B As Core.RECTANGLE) As Boolean
    IntersectRect = sys.MATH_IntersectRect(modSystem.RECTANGLE2GFX_RECT(A), modSystem.RECTANGLE2GFX_RECT(B))
End Function

' Determina si un punto esta dentro del area de un rectangulo:
Public Function PointInRect(P As Core.POINT, r As Core.RECTANGLE) As Boolean
    PointInRect = sys.MATH_PointInRect(P.X, P.Y, modSystem.RECTANGLE2GFX_RECT(r))
End Function

' Determina si un punto esta en el trazado de una linea:
Public Function PointInLine(P As Core.POINT, A As Core.POINT, B As Core.POINT) As Boolean
    PointInLine = sys.MATH_PointInLine(A.X, A.Y, B.X, B.Y, P.X, P.Y)
End Function

' Calcula la distancia en pixeles entre dos puntos:
Public Function GetDistance(A As Core.POINT, B As Core.POINT) As Long
    GetDistance = sys.MATH_GetDist(A.X, A.Y, B.X, B.Y)
End Function

' Calcula el angulo entre dos puntos:
Public Function GetAngle(A As Core.POINT, B As Core.POINT) As Single
    GetAngle = sys.MATH_GetAngle(A.X, A.Y, B.X, B.Y)
End Function

' Funcion matematica para calcular el arcotangente de un punto:
Public Function Atan2(X As Double, Y As Double) As Double
    Atan2 = sys.MATH_Atan2(X, Y)
End Function

' Traslada un punto en un radio y angulo:
Public Function TranslatePoint(SourcePoint As Core.POINT, Radius As Long, Angle As Single) As Core.POINT
    Dim pt As Core.POINT
    pt = SourcePoint
    
    ' Rotamos el punto en el radio y angulo indicado:
    Call sys.MATH_RotatePoint(pt.X, pt.Y, Radius, Angle)
    
    TranslatePoint = pt
End Function

' Invierte el valor de un numero:
Public Function Invert(Value As Variant) As Variant
    If IsNumeric(Value) Then Invert = Value - (Value * 2)
End Function

' Obtiene el porcentaje de un valor respecto a un rango de valor:
Public Function GetPercentFromValue(Value As Variant, Total As Variant) As Variant
    If IsNumeric(Value) And IsNumeric(Total) Then GetPercentFromValue = (Value / Total * 100) Else GetPercentFromValue = 0
End Function

' Obtiene el valor de un porcentaje respecto a un rango de valor:
Public Function GetValueFromPercent(Percent As Variant, Total As Variant) As Variant
    If IsNumeric(Percent) And IsNumeric(Total) Then GetValueFromPercent = (Percent * Total / 100) Else GetValueFromPercent = 0
End Function

' Calcula el punto de corte entre dos segmentos:
Public Function IntersectLine(A As Core.VECTOR, B As Core.VECTOR, C As Core.VECTOR, D As Core.VECTOR, r As Core.VECTOR) As Boolean
    Dim xD1 As Double, yD1 As Double, xD2 As Double, yD2 As Double, xD3 As Double, yD3 As Double
    Dim dot As Double, deg As Double, len1 As Double, len2 As Double
    Dim segmentLen1 As Double, segmentLen2 As Double
    Dim ua As Double, ub As Double, div As Double
    
    ' *** Optimizacion por José Miguel Sánchez Fernández ***
    ' Primero comprobamos si son perpendiculares y en caso de serlo aplicar igualacion para obtener el punto
    ' de corte y asi evitar los calculos complejos evitando de esa forma carga de procesamiento:
    If ((A.X = B.X) And (C.Y = D.Y)) Or ((A.Y = B.Y) And (C.X = D.X)) Then
        If (A.X = B.X) Then
            r.X = A.X
            r.Y = C.Y
        Else
            r.X = C.X
            r.Y = A.Y
        End If
                
        ' Comprobamos que el punto se encuentre en ambas lineas:
        Dim ptR As Core.POINT: ptR.X = CLng(r.X): ptR.Y = CLng(r.Y)
        Dim ptA As Core.POINT: ptA.X = CLng(A.X): ptA.Y = CLng(A.Y)
        Dim ptB As Core.POINT: ptB.X = CLng(B.X): ptB.Y = CLng(B.Y)
        Dim ptC As Core.POINT: ptC.X = CLng(C.X): ptC.Y = CLng(C.Y)
        Dim ptD As Core.POINT: ptD.X = CLng(D.X): ptD.Y = CLng(D.Y)
        IntersectLine = PointInLine(ptR, ptA, ptB) And PointInLine(ptR, ptC, ptD)
    ' *** ---------------------------------------------- ***
    Else
        
        ' calculate differences
        xD1 = B.X - A.X
        xD2 = D.X - C.X
        yD1 = B.Y - A.Y
        yD2 = D.Y - C.Y
        xD3 = A.X - C.X
        yD3 = A.Y - C.Y
        
        ' calculate the lengths of the two lines
        len1 = Sqr(xD1 * xD1 + yD1 * yD1)
        len2 = Sqr(xD2 * xD2 + yD2 * yD2)
    
        ' calculate angle between the two lines.
        dot = (xD1 * xD2 + yD1 * yD2) ' dot product
        deg = dot / (len1 * len2)
    
        ' if abs(angle)==1 then the lines are parallell,
        ' so no intersection is possible
        If (Abs(deg) = 1) Then Exit Function
        
        ' find intersection Pt between two lines
        Dim pt As Core.VECTOR
        
        div = yD2 * xD1 - xD2 * yD1
        ua = (xD2 * yD3 - yD2 * xD3) / div
        ub = (xD1 * yD3 - yD1 * xD3) / div
        pt.X = A.X + ua * xD1
        pt.Y = A.Y + ua * yD1
        
        ' calculate the combined length of the two segments
        ' between Pt-p1 and Pt-p2
        xD1 = pt.X - A.X
        xD2 = pt.X - B.X
        yD1 = pt.Y - A.Y
        yD2 = pt.Y - B.Y
        segmentLen1 = Sqr(xD1 * xD1 + yD1 * yD1) + Sqr(xD2 * xD2 + yD2 * yD2)
        
        ' calculate the combined length of the two segments
        ' between Pt-p3 and Pt-p4
        xD1 = pt.X - C.X
        xD2 = pt.X - D.X
        yD1 = pt.Y - C.Y
        yD2 = pt.Y - D.Y
        segmentLen2 = Sqr(xD1 * xD1 + yD1 * yD1) + Sqr(xD2 * xD2 + yD2 * yD2)
        
        ' if the lengths of both sets of segments are the same as
        ' the lenghts of the two lines the VECTOR is actually
        ' on the line segment.
    
        ' if the VECTOR isn't on the line, return null
        If (Abs(len1 - segmentLen1) > 0.01 Or Abs(len2 - segmentLen2) > 0.01) Then Exit Function
    
        ' return the valid intersection
        r = pt
        IntersectLine = True
        
    End If
End Function

' Calcula la interseccion entre un circulo y un rectangulo:
' Adaptado del ejemplo de http://stackoverflow.com/questions/401847/circle-rectangle-collision-detection-intersection/402010#402010
Public Function IntersectCircleWithRectangle(Center As Core.POINT, Radius As Long, RECTANGLE As Core.RECTANGLE) As Boolean
    Dim circleDistance As Core.POINT
    circleDistance.X = Abs(Center.X - RECTANGLE.X - RECTANGLE.Width / 2)
    circleDistance.Y = Abs(Center.Y - RECTANGLE.Y - RECTANGLE.Height / 2)
    
    If (circleDistance.X > (RECTANGLE.Width / 2 + Radius)) Or (circleDistance.Y > (RECTANGLE.Height / 2 + Radius)) Then _
        Exit Function

    If (circleDistance.X <= (RECTANGLE.Width / 2)) Or (circleDistance.Y <= (RECTANGLE.Height / 2)) Then
        IntersectCircleWithRectangle = True
        Exit Function
    End If

    Dim cornerDistance_sq As Double
    cornerDistance_sq = (circleDistance.X - RECTANGLE.Width / 2) ^ 2 + (circleDistance.Y - RECTANGLE.Height / 2) ^ 2

    IntersectCircleWithRectangle = (cornerDistance_sq <= (Radius ^ 2))
End Function
