VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EDIT_Grid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Genera una rejilla para la edicion y la vuelca a una textura.
' Es una via de optimizacion para evitar el redibujado de todas las primitivas que generan la rejilla en tiempo real.

Option Explicit

Private varTarget As Graphics.RenderTarget
Private varSprite As Graphics.Sprite

Public Property Get Visible() As Boolean
    Visible = varSprite.Visible
End Property

Public Property Let Visible(value As Boolean)
    varSprite.Visible = value
End Property

Private Sub Class_Initialize()
    Set varTarget = New Graphics.RenderTarget
    Call varTarget.Initialize(1024, 1024, False)
    Call varTarget.BeginTarget
    Call DrawGrid
    Call varTarget.EndTarget
    Set varSprite = New Graphics.Sprite
    Call varSprite.SetTexture(varTarget.GetTexture())
    Call varSprite.SetRegion(0, 0, 800, 600)
    varSprite.Effect = XOR_Exclusion
    varSprite.Fixed = True
End Sub

Private Sub Class_Terminate()
    Set varTarget = Nothing
    Set varSprite = Nothing
End Sub

Private Sub DrawGrid()
    Dim s As Long, c As Long, X As Long, Y As Long
    s = 8
    Do
        Select Case s
            Case 64: c = Graphics.Color_Constant.Gray
            Case 32: c = Graphics.Color_Constant.DarkCyan
            Case 16: c = Graphics.Color_Constant.DarkYellow
            Case 8: c = Graphics.Color_Constant.DarkRed
        End Select
        For X = 0 To Engine.GraphicEngine.CurrentDisplayMode.Width Step s
            Call Engine.GraphicEngine.Primitives.DrawLine(X, 0, X, Engine.GraphicEngine.CurrentDisplayMode.Height, -6, c, True)
        Next
        For Y = 0 To Engine.GraphicEngine.CurrentDisplayMode.Height Step s
            Call Engine.GraphicEngine.Primitives.DrawLine(0, Y, Engine.GraphicEngine.CurrentDisplayMode.Width, Y, -6, c, True)
        Next
        s = s * 2
    Loop Until s = 128
End Sub

Public Sub Draw()
    Call varSprite.Draw
End Sub
