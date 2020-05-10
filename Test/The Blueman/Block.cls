VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Block"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private varSpr As Graphics.Sprite
Private varBody As Physics.Body
Private varLocation As System.Point

Public Property Get Location() As System.Point
    Location = varLocation
End Property

Public Property Let Location(value As System.Point)
    varLocation = value
    Call varBody.SetLocation(value.X, value.Y)
End Property

Public Sub SetLocation(X As Long, Y As Long)
    varLocation.X = X
    varLocation.Y = Y
    Call varBody.SetLocation(X, Y)
End Sub

Public Sub Update()
    Call varSpr.SetLocation(CLng(varBody.Location.X), CLng(varBody.Location.Y), 0)
End Sub

Public Sub Draw()
    Call varSpr.Draw
End Sub

Private Sub Class_Initialize()
    Set varSpr = New Graphics.Sprite
    Call varSpr.SetTexture(modMain.gfx.Textures("block"))
    varSpr.Centered = True
    Set varBody = sim.Bodies.Add(0, 0, 0, 100, 100, 1, 2, True, 24)
End Sub

Private Sub Class_Terminate()
    Call modMain.sim.Bodies.Remove(varBody)
    Set varBody = Nothing
    Set varSpr = Nothing
End Sub
