VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ENG_BBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private varKey As String
Private varBody As Physics.Body
Private varColliderProfile As Physics.ColliderProfiles
Private varParam As Long

Public Property Get Key() As String
    Key = varKey
End Property

Public Property Get Body() As Physics.Body
    Set Body = varBody
End Property

Public Property Set Body(value As Physics.Body)
    Set varBody = value
    If Not IsEmpty(varBody.Tag) Then
        Dim arr() As String: arr = Split(varBody.Tag, ";")  ' Extraemos la informacion de los colisionadores del Tag del cuerpo.
        varColliderProfile = CLng(arr(0))
        varParam = CLng(arr(1))
    End If
End Property

Public Property Get Rectangle() As Core.Rectangle
    Rectangle = varBody.RECT
End Property

Public Property Get ColliderProfile() As Physics.ColliderProfiles
    ColliderProfile = varColliderProfile
End Property

Public Property Get Param() As Long
    Param = varParam
End Property

Public Function ClientToScreenRectangle() As Core.Rectangle
    ClientToScreenRectangle = Core.Generics.CreateRECTANGLE2(Engine.Scene.ScreenToClient(Core.Generics.CreatePOINT(varBody.RECT.X, varBody.RECT.Y)), _
                                                             Core.Generics.CreateSIZE(varBody.RECT.Width, varBody.RECT.Height))
End Function

Public Function ToString() As String
    With varBody
        ' ToString = X; Y; Ancho; Alto; PerfilColisionadores; GrosorColisionadores
        ToString = .Location.X & ";" & .Location.Y & ";" & .RECT.Width & ";" & .RECT.Height & ";" & varColliderProfile & ";" & Param
    End With
End Function

Private Sub Class_Initialize()
    varKey = Core.Guid.Generate ' Se autogenera la clave unica del objeto.
End Sub

Private Sub Class_Terminate()
    If Not Engine.Scene Is Nothing Then Call Engine.Scene.PhysicSimulator.Bodies.Remove(varBody)
    Set varBody = Nothing
End Sub

