VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Primitives"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Objeto de dibujo de primitivas:

Option Explicit

Private varTmpOffSet As Core.Point

' Dibujar linea:
Public Sub DrawLine(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, Z As Long, Color As Long, Optional Fixed As Boolean)
    If Not Fixed Then
        varTmpOffSet = modGraphics.varOffset
    Else
        varTmpOffSet.X = 0: varTmpOffSet.Y = 0
    End If
    
    If LineInScreen(Core.Generics.CreatePOINT(X1 + varTmpOffSet.X, Y1 + varTmpOffSet.Y), Core.Generics.CreatePOINT(X2 + varTmpOffSet.X, Y2 + varTmpOffSet.Y)) Then
        Call gfx.DRAW_Line(X1 + varTmpOffSet.X, Y1 + varTmpOffSet.Y, X2 + varTmpOffSet.X, Y2 + varTmpOffSet.Y, Z, Color)
        Call modGraphics.varRenderMonitor.Add("LINE")
    End If
End Sub

Public Sub DrawLine2(A As Core.Point, B As Core.Point, Z As Long, Color As Long, Optional Fixed As Boolean)
    If Not Fixed Then
        varTmpOffSet = modGraphics.varOffset
    Else
        varTmpOffSet.X = 0: varTmpOffSet.Y = 0
    End If
    
    If LineInScreen(Core.Generics.CreatePOINT(A.X + varTmpOffSet.X, A.Y + varTmpOffSet.Y), Core.Generics.CreatePOINT(B.X + varTmpOffSet.X, B.Y + varTmpOffSet.Y)) Then
        Call gfx.DRAW_Line(A.X + varTmpOffSet.X, A.Y + varTmpOffSet.Y, B.X + varTmpOffSet.X, B.Y + varTmpOffSet.Y, Z, Color)
        Call modGraphics.varRenderMonitor.Add("LINE")
    End If
End Sub

' Comprueba si una linea esta dentro del area de vision:
Private Function LineInScreen(A As Core.Point, B As Core.Point) As Boolean
    If Core.Math.PointInRect(A, modGraphics.varScreenRect) Or Core.Math.PointInRect(B, modGraphics.varScreenRect) Then
        LineInScreen = True
    Else
        Dim screen As Core.VECTOR
        screen.X = CSng(modGraphics.varScreenRect.Width)
        screen.Y = CSng(modGraphics.varScreenRect.Height)
        
        If Core.Math.IntersectLine(Core.Generics.POINT2VECTOR(A), Core.Generics.POINT2VECTOR(B), _
           Core.Generics.CreateVECTOR(0, 0, 0), Core.Generics.CreateVECTOR(screen.X, 0, 0), _
           Core.Generics.NullVECTOR) Then
            LineInScreen = True
            Exit Function
        End If
        
        If Core.Math.IntersectLine(Core.Generics.POINT2VECTOR(A), Core.Generics.POINT2VECTOR(B), _
           Core.Generics.CreateVECTOR(0, 0, 0), Core.Generics.CreateVECTOR(0, screen.Y, 0), _
           Core.Generics.NullVECTOR) Then
            LineInScreen = True
            Exit Function
        End If
    
        If Core.Math.IntersectLine(Core.Generics.POINT2VECTOR(A), Core.Generics.POINT2VECTOR(B), _
           Core.Generics.CreateVECTOR(0, screen.Y, 0), Core.Generics.CreateVECTOR(screen.X, screen.Y, 0), _
           Core.Generics.NullVECTOR) Then
            LineInScreen = True
            Exit Function
        End If
    
        If Core.Math.IntersectLine(Core.Generics.POINT2VECTOR(A), Core.Generics.POINT2VECTOR(B), _
           Core.Generics.CreateVECTOR(screen.X, 0, 0), Core.Generics.CreateVECTOR(screen.X, screen.Y, 0), _
           Core.Generics.NullVECTOR) Then
            LineInScreen = True
            Exit Function
        End If
    End If
End Function


' Dibujar caja simple:
Public Sub DrawBox(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, Z As Long, BorderColor As Long, Optional Fill As Boolean, Optional FillColor As Long, Optional Fixed As Boolean)
    If Not Fixed Then
        varTmpOffSet = modGraphics.varOffset
    Else
        varTmpOffSet.X = 0: varTmpOffSet.Y = 0
    End If
    
    If BoxInScreen(Core.Generics.CreatePOINT(X1 + varTmpOffSet.X, Y1 + varTmpOffSet.Y), Core.Generics.CreatePOINT(X2 + varTmpOffSet.X, Y2 + varTmpOffSet.Y)) Then
        Call gfx.DRAW_Box(X1 + varTmpOffSet.X, Y1 + varTmpOffSet.Y, X2 + varTmpOffSet.X, Y2 + varTmpOffSet.Y, Z, BorderColor, Fill, FillColor)
        Call modGraphics.varRenderMonitor.Add("BOX")
    End If
End Sub

Public Sub DrawBox2(A As Core.Point, B As Core.Point, Z As Long, BorderColor As Long, Optional Fill As Boolean, Optional FillColor As Long, Optional Fixed As Boolean)
    If Not Fixed Then
        varTmpOffSet = modGraphics.varOffset
    Else
        varTmpOffSet.X = 0: varTmpOffSet.Y = 0
    End If
    
    If BoxInScreen(Core.Generics.CreatePOINT(A.X + varTmpOffSet.X, A.Y + varTmpOffSet.Y), Core.Generics.CreatePOINT(B.X + varTmpOffSet.X, B.Y + varTmpOffSet.Y)) Then
        Call gfx.DRAW_Box(A.X + varTmpOffSet.X, A.Y + varTmpOffSet.Y, B.X + varTmpOffSet.X, B.Y + varTmpOffSet.Y, Z, BorderColor, Fill, FillColor)
        Call modGraphics.varRenderMonitor.Add("BOX")
    End If
End Sub

Public Sub DrawBox3(X As Long, Y As Long, Width As Long, Height As Long, Z As Long, BorderColor As Long, Optional Fill As Boolean, Optional FillColor As Long, Optional Fixed As Boolean)
    If Not Fixed Then
        varTmpOffSet = modGraphics.varOffset
    Else
        varTmpOffSet.X = 0: varTmpOffSet.Y = 0
    End If
    
    If BoxInScreen2(X + varTmpOffSet.X, Y + varTmpOffSet.Y, Width + varTmpOffSet.X, Height + varTmpOffSet.Y) Then
        Call gfx.DRAW_Box(X + varTmpOffSet.X, Y + varTmpOffSet.Y, X + Width + varTmpOffSet.X, Y + Height + varTmpOffSet.Y, Z, BorderColor, Fill, FillColor)
        Call modGraphics.varRenderMonitor.Add("BOX")
    End If
End Sub

Public Sub DrawBox4(Rect As Core.RECTANGLE, Z As Long, BorderColor As Long, Optional Fill As Boolean, Optional FillColor As Long, Optional Fixed As Boolean)
    If Not Fixed Then
        varTmpOffSet = modGraphics.varOffset
    Else
        varTmpOffSet.X = 0: varTmpOffSet.Y = 0
    End If
    
    If BoxInScreen2(Rect.X + varTmpOffSet.X, Rect.Y + varTmpOffSet.Y, Rect.Width + varTmpOffSet.X, Rect.Height + varTmpOffSet.Y) Then
        Call gfx.DRAW_Box(Rect.X + varTmpOffSet.X, Rect.Y + varTmpOffSet.Y, Rect.X + Rect.Width + varTmpOffSet.X, Rect.Y + Rect.Height + varTmpOffSet.Y, Z, BorderColor, Fill, FillColor)
        Call modGraphics.varRenderMonitor.Add("BOX")
    End If
End Sub


' Dibujar caja avanzada:
Public Sub DrawAdvBox(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, Z As Long, Color1 As Long, Color2 As Long, Color3 As Long, Color4 As Long, Optional Fixed As Boolean)
    If Not Fixed Then
        varTmpOffSet = modGraphics.varOffset
    Else
        varTmpOffSet.X = 0: varTmpOffSet.Y = 0
    End If
    
    If BoxInScreen(Core.Generics.CreatePOINT(X1 + varTmpOffSet.X, Y1 + varTmpOffSet.Y), Core.Generics.CreatePOINT(X2 + varTmpOffSet.X, Y2 + varTmpOffSet.Y)) Then
        Call gfx.DRAW_AdvBox(X1 + varTmpOffSet.X, Y1 + varTmpOffSet.Y, X2 + varTmpOffSet.X, Y2 + varTmpOffSet.Y, Z, Color1, Color2, Color3, Color4)
        Call modGraphics.varRenderMonitor.Add("ADVBOX")
    End If
End Sub

Public Sub DrawAdvBox2(A As Core.Point, B As Core.Point, Z As Long, Color1 As Long, Color2 As Long, Color3 As Long, Color4 As Long, Optional Fixed As Boolean)
    If Not Fixed Then
        varTmpOffSet = modGraphics.varOffset
    Else
        varTmpOffSet.X = 0: varTmpOffSet.Y = 0
    End If
    
    If BoxInScreen(Core.Generics.CreatePOINT(A.X + varTmpOffSet.X, A.Y + varTmpOffSet.Y), Core.Generics.CreatePOINT(B.X + varTmpOffSet.X, B.Y + varTmpOffSet.Y)) Then
        Call gfx.DRAW_AdvBox(A.X + varTmpOffSet.X, A.Y + varTmpOffSet.Y, B.X + varTmpOffSet.X, B.Y + varTmpOffSet.Y, Z, Color1, Color2, Color3, Color4)
        Call modGraphics.varRenderMonitor.Add("ADVBOX")
    End If
End Sub

Public Sub DrawAdvBox3(X As Long, Y As Long, Width As Long, Height As Long, Z As Long, Color1 As Long, Color2 As Long, Color3 As Long, Color4 As Long, Optional Fixed As Boolean)
    If Not Fixed Then
        varTmpOffSet = modGraphics.varOffset
    Else
        varTmpOffSet.X = 0: varTmpOffSet.Y = 0
    End If
    
    If BoxInScreen2(X + varTmpOffSet.X, Y + varTmpOffSet.Y, Width + varTmpOffSet.X, Height + varTmpOffSet.Y) Then
        Call gfx.DRAW_AdvBox(X + varTmpOffSet.X, Y + varTmpOffSet.Y, X + Width + varTmpOffSet.X, Y + Height + varTmpOffSet.Y, Z, Color1, Color2, Color3, Color4)
        Call modGraphics.varRenderMonitor.Add("ADVBOX")
    End If
End Sub

Public Sub DrawAdvBox4(Rect As Core.RECTANGLE, Z As Long, Color1 As Long, Color2 As Long, Color3 As Long, Color4 As Long, Optional Fixed As Boolean)
    If Not Fixed Then
        varTmpOffSet = modGraphics.varOffset
    Else
        varTmpOffSet.X = 0: varTmpOffSet.Y = 0
    End If
    
    If BoxInScreen2(Rect.X + varTmpOffSet.X, Rect.Y + varTmpOffSet.Y, Rect.X + Rect.Width + varTmpOffSet.X, Rect.Y + Rect.Height + varTmpOffSet.Y) Then
        Call gfx.DRAW_AdvBox(Rect.X + varTmpOffSet.X, Rect.Y + varTmpOffSet.Y, Rect.X + Rect.Width + varTmpOffSet.X, Rect.Y + Rect.Height + varTmpOffSet.Y, Z, Color1, Color2, Color3, Color4)
        Call modGraphics.varRenderMonitor.Add("ADVBOX")
    End If
End Sub

' Indica si la caja esta en el area de la pantalla (los parametros se deben pasan con el offset aplicado):
Private Function BoxInScreen(A As Core.Point, B As Core.Point) As Boolean
    BoxInScreen = BoxInScreen2(A.X, A.Y, B.X - A.X, B.Y - A.Y)
End Function

Private Function BoxInScreen2(X As Long, Y As Long, Width As Long, Height As Long) As Boolean
    ' Definimos el rectangulo a intersectar:
    Dim boxRect As Core.RECTANGLE: boxRect = Core.Generics.CreateRECTANGLE(X, Y, Width, Height)
    
    ' Comprobamos si intersectan ambos rectangulos:
    BoxInScreen2 = Core.Math.IntersectRect(modGraphics.varScreenRect, boxRect)
End Function


' Dibujar trapezoide:
Public Sub DrawShape(LeftTop As Core.Point, RightTop As Core.Point, LeftBottom As Core.Point, RightBottom As Core.Point, Optional Fixed As Boolean)
    Dim v(3) As Vertex
    
    If Not Fixed Then
        varTmpOffSet = modGraphics.varOffset
    Else
        varTmpOffSet.X = 0: varTmpOffSet.Y = 0
    End If
        
    LeftTop.X = LeftTop.X + varTmpOffSet.X: LeftTop.Y = LeftTop.Y + varTmpOffSet.Y
    v(0) = modGraphics.POINT2Vertex(LeftTop)
    
    RightTop.X = RightTop.X + varTmpOffSet.X: RightTop.Y = RightTop.Y + varTmpOffSet.Y
    v(1) = modGraphics.POINT2Vertex(RightTop)
    
    LeftBottom.X = LeftBottom.X + varTmpOffSet.X: LeftBottom.Y = LeftBottom.Y + varTmpOffSet.Y
    v(2) = modGraphics.POINT2Vertex(LeftBottom)
    
    RightBottom.X = RightBottom.X + varTmpOffSet.X: RightBottom.Y = RightBottom.Y + varTmpOffSet.Y
    v(3) = modGraphics.POINT2Vertex(RightBottom)
    
    Call gfx.DRAW_Trapezoid(v())
    Call modGraphics.varRenderMonitor.Add("SHAPE")
End Sub



' Dibuja una cadena de texto usando una fuente predefinida:
Public Sub WriteText(Font As Graphics.Font, Text As String, X As Long, Y As Long, Optional Z As Long, Optional Color As Long = &HFFFFFFFF, Optional Align As TextAlign, Optional Fixed As Boolean)
    If Not Fixed Then
        varTmpOffSet = modGraphics.varOffset
    Else
        varTmpOffSet.X = 0: varTmpOffSet.Y = 0
    End If
    
    Call gfx.DRAW_Text(Font.GetID(), Text, X + varTmpOffSet.X, Y + varTmpOffSet.Y, Z, Color, Align)
    Call modGraphics.varRenderMonitor.Add("TEXT")
End Sub
