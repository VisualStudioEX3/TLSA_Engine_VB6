Attribute VB_Name = "modGraphics"
Option Explicit

Public gfx As New dx_GFX_Class

Public varOffset As Core.Point, varLastOffSet As Core.Point

Public varHelper As New Graphics.Helper

Public varFilter As Graphics.TextureFilter
Public varSystemFont As Graphics.Font

' Accesibles desde toda la libreria:
Public varTextures As New Graphics.TextureList
Public varSurfaces As New Graphics.SurfaceList
Public varFonts As New Graphics.FontList
Public varPrimitives As New Graphics.Primitives

Public varScreenRect As Core.RECTANGLE

Public varRenderMonitor As Graphics.RenderMonitor

' Convierte la estructura POINT a Vertex de dx_lib32:
Public Function POINT2Vertex(pt As Core.Point) As dxlib32_221.Vertex
    Dim v As dxlib32_221.Vertex
    
    v.X = pt.X
    v.Y = pt.Y
    v.Z = pt.Z
    v.Color = pt.Color
    
    POINT2Vertex = v
End Function

' Convierte la estructura Vertex de dx_lib32 a POINT:
Public Function Vertex2POINT(v As dxlib32_221.Vertex) As Core.Point
    Dim pt As Core.Point
    
    pt.X = v.X
    pt.Y = v.Y
    pt.Z = v.Z
    pt.Color = v.Color
    
    Vertex2POINT = pt
End Function

' Convierte la estructura RECTANGLE a GFX_RECT de dx_lib32:
Public Function RECTANGLE2GFX_RECT(r As Core.RECTANGLE) As dxlib32_221.GFX_Rect
    Dim gR As dxlib32_221.GFX_Rect
    
    gR.X = r.X
    gR.Y = r.Y
    gR.Width = r.Width
    gR.Height = r.Height
    
    RECTANGLE2GFX_RECT = gR
End Function

' Calcula el punto de corte entre dos segmentos:
Public Function IntersectLine(A As dxlib32_221.Vertex, B As dxlib32_221.Vertex, C As dxlib32_221.Vertex, D As dxlib32_221.Vertex, r As dxlib32_221.Vertex) As Boolean
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
'        IntersectLine = sys.MATH_PointInLine(CLng(A.X), CLng(A.Y), CLng(B.X), CLng(B.Y), CLng(r.X), CLng(r.Y)) _
'                     And sys.MATH_PointInLine(CLng(C.X), CLng(C.Y), CLng(D.X), CLng(D.Y), CLng(r.X), CLng(r.Y))
        IntersectLine = Core.Math.PointInLine(modGraphics.Vertex2POINT(r), modGraphics.Vertex2POINT(A), modGraphics.Vertex2POINT(B)) _
                        And Core.Math.PointInLine(modGraphics.Vertex2POINT(r), modGraphics.Vertex2POINT(C), modGraphics.Vertex2POINT(D))
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
        Dim pt As Vertex
        
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
        ' the lenghts of the two lines the vertex is actually
        ' on the line segment.
    
        ' if the vertex isn't on the line, return null
        If (Abs(len1 - segmentLen1) > 0.01 Or Abs(len2 - segmentLen2) > 0.01) Then Exit Function
    
        ' return the valid intersection
        r = pt
        IntersectLine = True
        
    End If
End Function
