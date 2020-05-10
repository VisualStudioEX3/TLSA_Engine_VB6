VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PixelArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Objeto para dibujar listas de pixeles en una sola llamada:
Option Explicit

Private pxArray() As dxlib32_221.Vertex, pxCount As Long

Private varFixed As Boolean                           ' No aplica el desplazamiento de las coordenadas OffSet.
Private varTmpOffSet As Core.Point                    ' Variable temporal donde se asigna el OffSet de la escena si el objeto no esta fijo.

Public Property Get Fixed() As Boolean
    Fixed = varFixed
End Property

Public Property Let Fixed(value As Boolean)
    varFixed = value
End Property

' Agrega un pixel al final de la lista:
Public Sub Add(X As Long, Y As Long, Z As Long, Color As Long)
    ReDim Preserve pxArray(pxCount)
    
    ' Si el objeto no esta configurado como fijo se le aplica desplazamiento de posicion segun variable OffSet global:
    If Not varFixed Then
        varTmpOffSet = modGraphics.varOffset
    Else
        varTmpOffSet.X = 0: varTmpOffSet.Y = 0
    End If
    
    With pxArray(pxCount)
        .X = X + varTmpOffSet.X: .Y = Y + varTmpOffSet.Y: .Z = Z: .Color = Color
    End With
    
    pxCount = pxCount + 1
End Sub

' Elimna un pixel de la lista:
Public Sub Remove(Index As Long)
    Dim i As Long
    
    For i = Index To UBound(pxArray)
        pxArray(i) = pxArray(i + 1)
    Next
    
    pxCount = pxCount - 1
    
    If pxCount = 0 Then
        Erase pxArray
    Else
        ReDim Preserve pxArray(UBound(pxArray) - 1)
    End If
End Sub

' Vacia la lista:
Public Sub Clear()
    Erase pxArray
    pxCount = 0
End Sub

' Devuelve le numero de elementos de la lista:
Public Property Get Count() As Long
    Count = pxCount
End Property

Public Property Get Item(Index As Long) As Core.Point
Attribute Item.VB_UserMemId = 0
    ' Si el objeto no esta configurado como fijo se le aplica desplazamiento de posicion segun variable OffSet global:
    If Not varFixed Then
        varTmpOffSet = modGraphics.varOffset
    Else
        varTmpOffSet.X = 0: varTmpOffSet.Y = 0
    End If
    
    Item.X = pxArray(Index).X - varTmpOffSet.X
    Item.Y = pxArray(Index).Y - varTmpOffSet.Y
    Item.Z = pxArray(Index).Z
    Item.Color = pxArray(Index).Color
End Property

Public Property Let Item(Index As Long, value As Core.Point)
    ' Si el objeto no esta configurado como fijo se le aplica desplazamiento de posicion segun variable OffSet global:
    If Not varFixed Then
        varTmpOffSet = modGraphics.varOffset
    Else
        varTmpOffSet.X = 0: varTmpOffSet.Y = 0
    End If
    
    pxArray(Index).X = value.X + varTmpOffSet.X
    pxArray(Index).Y = value.Y + varTmpOffSet.Y
    pxArray(Index).Z = value.Z
    pxArray(Index).Color = value.Color
End Property

Public Sub SetPoint(Index As Long, X As Long, Y As Long, Z As Long, Color As Long)
    ' Si el objeto no esta configurado como fijo se le aplica desplazamiento de posicion segun variable OffSet global:
    If Not varFixed Then
        varTmpOffSet = modGraphics.varOffset
    Else
        varTmpOffSet.X = 0: varTmpOffSet.Y = 0
    End If
    
    With pxArray(Index)
        .X = X + varTmpOffSet.X: .Y = Y + varTmpOffSet.Y: .Z = Z: .Color = Color
    End With
End Sub

' Procesa y dibuja la lista de pixeles:
Public Sub Draw()
    Call gfx.DRAW_Pixels(pxArray())
    Call modGraphics.varRenderMonitor.Add("PIXELARRAY")
End Sub

' Llamar solo si se cambia el offset:
Public Sub Update()
    If Not varFixed Then
        If pxCount > 0 Then
            Dim i As Long
            For i = 0 To pxCount - 1
                With pxArray(i)
                    .X = .X - modGraphics.varLastOffSet.X: .Y = .Y - modGraphics.varLastOffSet.Y ' Corregimos el offset anterior.
                    .X = .X + modGraphics.varOffset.X: .Y = .Y + modGraphics.varOffset.Y ' Actualizamos el offset con los nuevos valores.
                End With
            Next i
        End If
    End If
End Sub

