VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RenderTarget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private id As Long              ' Id del RenderTarget en dx_lib32.
Private tex As Graphics.Texture ' Referencia a la textura asociada.

' Devuelve la textura del Render Target:
Public Function GetTexture() As Graphics.Texture
    Set GetTexture = tex
End Function

' Inicializa el Render Target:
Public Sub Initialize(Width As Long, Height As Long, GrayScale As Boolean)
    Set tex = New Graphics.Texture
    Call tex.Create(Width, Height, GrayScale, True)
    id = gfx.TARGET_Create(tex.GetID)
End Sub

' Abre el render target:
Public Sub BeginTarget()
    Call gfx.TARGET_Open(id)
End Sub

' Cierra el render target:
Public Sub EndTarget()
    Call gfx.TARGET_Close
End Sub

Private Sub Class_Initialize()
    id = -1
End Sub

Private Sub Class_Terminate()
    Call gfx.TARGET_Destroy(id)
End Sub
