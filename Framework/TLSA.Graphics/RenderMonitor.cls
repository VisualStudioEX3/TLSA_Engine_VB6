VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RenderMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private varSprites As Long
Private varTextStrings As Long
Private varLines As Long
Private varBoxes As Long
Private varAdvBoxes As Long
Private varShapes As Long
Private varPixelArrays As Long

Public Property Get Sprites() As Long
    Sprites = varSprites
End Property

Public Property Get TextStrings() As Long
    TextStrings = varTextStrings
End Property

Public Property Get Lines() As Long
    Lines = varLines
End Property

Public Property Get Boxes() As Long
    Boxes = varBoxes
End Property

Public Property Get AdvanceBoxes() As Long
    AdvanceBoxes = varAdvBoxes
End Property

Public Property Get Shapes() As Long
    Shapes = varShapes
End Property

Public Property Get PixelArrays() As Long
    PixelArrays = varPixelArrays
End Property

Public Property Get Total() As Long
    Total = (varSprites + varTextStrings + varLines + varBoxes + varAdvBoxes + varShapes + varPixelArrays)
End Property

Friend Sub Reset()
    varSprites = 0
    varTextStrings = 0
    varLines = 0
    varBoxes = 0
    varAdvBoxes = 0
    varShapes = 0
    varPixelArrays = 0
End Sub

Friend Sub Add(Element As String)
    Select Case UCase(Element)
        Case "SPRITE": varSprites = varSprites + 1
        Case "TEXT": varTextStrings = varTextStrings + 1
        Case "LINE": varLines = varLines + 1
        Case "BOX": varBoxes = varBoxes + 1
        Case "ADVBOX": varAdvBoxes = varAdvBoxes + 1
        Case "SHAPE": varShapes = varShapes + 1
        Case "PIXELARRAY": varPixelArrays = varPixelArrays + 1
    End Select
End Sub

' Devuelve un resumen con los valores de los elementos dibujados:
Public Function ToString() As String
    ToString = "Elementos dibujados en pantalla:" & vbNewLine & _
               " + Sprites: " & varSprites & vbNewLine & _
               " + TextStrings: " & varTextStrings & vbNewLine & _
               " + Lineas: " & varLines & vbNewLine & _
               " + Cajas: " & varBoxes & vbNewLine & _
               " + Cajas avanzadas: " & varAdvBoxes & vbNewLine & _
               " + Figuras (cajas de vertices libres): " & varShapes & vbNewLine & _
               " + Arrays de pixeles: " & varPixelArrays & vbNewLine & _
               "TOTAL = " & (varSprites + varTextStrings + varLines + varBoxes + varAdvBoxes + varShapes + varPixelArrays)
End Function
