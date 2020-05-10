VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Blueman"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private varSpr As Graphics.Sprite
Private varBody As physics.Body

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

Private Sub Class_Initialize()
    Set varSpr = modMain.gfx.Helper.CreateAnimatedSprite(modMain.gfx.Textures("blueman"))
    Call varSpr.SetCurrentAnimation("standby")
    
    Set varBody = sim.Bodies.Add(0, 0, 0, 50, 100, 1, 3.5, False, 12)
    varBody.Color = Graphics.Color_Constant.Orange
    varBody.PhysicType = PhysicalEntity
    varBody.Tag = "Player"
    Call varBody.SetMaxFriction(2, 2)
    Call varBody.SetMinFriction(4, 4)
End Sub

Private Sub Class_Terminate()
    Call modMain.sim.Bodies.Remove(varBody)
    Set varBody = Nothing
    Set varSpr = Nothing
End Sub

Public Sub Update()
    If modMain.gameInput.Key(Key_A) Then
        varSpr.Mirror = Horizontal
        Call varBody.SetForceX(-10)
        Call varSpr.SetCurrentAnimation("run")
    ElseIf modMain.gameInput.Key(Key_D) Then
        varSpr.Mirror = WithoutMirroring
        Call varBody.SetForceX(10)
        Call varSpr.SetCurrentAnimation("run")
    ElseIf modMain.gameInput.Key(Key_Space) Then
        Call varSpr.SetCurrentAnimation("shoot")
        'Call varBody.SetForceX(0)
    Else
        Call varSpr.SetCurrentAnimation("standby")
        'Call varBody.SetForceX(0)
    End If
    
    If modMain.gameInput.Key_Hit(Key_W) Then Call varBody.SetForceY(-45)
    
    Call varSpr.SetLocation(CLng(varBody.Location.X), CLng(varBody.Location.Y), 0)
    varLocation = varSpr.Location
    
    Call varSpr.Update
End Sub

Public Sub Draw()
    Call varSpr.Draw
    Call varSpr.CurrentTile.ControlPoints.Draw
End Sub
