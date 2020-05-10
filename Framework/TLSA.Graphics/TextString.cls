VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TextString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum TextAlign
    Left = Text_Align.Align_Left
    Right = Text_Align.Align_Right
    Center = Text_Align.Align_Center
End Enum

Private mfont As Graphics.Font
Private strText As String
Private vecLocation As Core.POINT
Private lngColor As Long
Private enuAlign As Graphics.TextAlign

Private varFixed As Boolean                             ' No aplica el desplazamiento de las coordenadas OffSet.

Public Property Get Fixed() As Boolean
    Fixed = varFixed
End Property

Public Property Let Fixed(value As Boolean)
    varFixed = value
End Property

' Devuelve la fuente que se usara para representar los caracteres de la cadena de texto:
Public Property Get Font() As Font
    Set Font = mfont
End Property

' Establece la fuente que se usara para representar los caracteres de la cadena de texto:
Public Property Set Font(value As Font)
    Set mfont = value
End Property

' Devuelve la cadena de texto que se dibujara:
Public Property Get Text() As String
    Text = strText
End Property

' Establece la cadena de texto que se dibujara:
Public Property Let Text(value As String)
    strText = value
End Property

' Devuelve las coordenadas donde se dibujara la cadena de texto:
Public Property Get Location() As Core.POINT
    Location = vecLocation
End Property

' Establece las coordenadas donde se dibujara la cadena de texto:
Public Property Let Location(value As Core.POINT)
    vecLocation = value
End Property

' Devuelve el color que usara el texto:
Public Property Get Color() As Long
    Color = lngColor
End Property

' Establece el color que usara el texto:
Public Property Let Color(value As Long)
    lngColor = value
End Property

' Devuelve la alineacion en pantalla del texto:
Public Property Get Align() As Graphics.TextAlign
    Align = enuAlign
End Property

' Establece la alineacion en pantalla del texto:
Public Property Let Align(value As Graphics.TextAlign)
    enuAlign = value
End Property

Public Property Get Width() As Long
    Width = gfx.FONT_SystemGetTextWidth(mfont.GetID(), strText)
End Property

Public Property Get Height() As Long
    Height = gfx.FONT_SystemGetTextHeight(mfont.GetID(), strText)
End Property

' Constructor de la clase:
Public Sub Initialize(Font As Graphics.Font, Optional Text As String, Optional X As Long, Optional Y As Long, Optional Z As Long, Optional Color As Long = Graphics.Color_Constant.White, Optional Align As Graphics.TextAlign = Left)
    Set mfont = Font
    strText = Text
    vecLocation.X = X: vecLocation.Y = Y: vecLocation.Z = Z
    lngColor = Color
    enuAlign = Align
End Sub

' Dibuja la cadena de texto con la informacion de los parametros de la clase:
Public Sub Draw()
    ' Si el objeto no esta configurado como fijo se le aplica desplazamiento de posicion segun variable OffSet global:
    Dim varTmpOffSet As Core.POINT
    If Not varFixed Then varTmpOffSet = modGraphics.varOffset
    If InScreen() Then
        Call gfx.DRAW_Text(mfont.GetID(), strText, vecLocation.X + varTmpOffSet.X, vecLocation.Y + varTmpOffSet.Y, vecLocation.Z, lngColor, enuAlign)
        Call modGraphics.varRenderMonitor.Add("TEXT")
    End If
End Sub

Private Function InScreen() As Boolean
    Dim labelRect As Core.RECTANGLE
    Dim varTmpOffSet As Core.POINT: If Not varFixed Then varTmpOffSet = modGraphics.varOffset
    
    labelRect = Core.Generics.CreateRECTANGLE(vecLocation.X + varTmpOffSet.X, vecLocation.Y + varTmpOffSet.Y, Me.Width, Me.Height)
    InScreen = Core.Math.IntersectRect(labelRect, modGraphics.varScreenRect)
End Function
