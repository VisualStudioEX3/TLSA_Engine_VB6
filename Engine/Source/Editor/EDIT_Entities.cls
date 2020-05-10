VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EDIT_Entities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public EntityForm As Form           ' Representa el formulario de parametros de la entidad seleccionada.
Public Entity As Object             ' Representa la instancia de la previsualizacion de la entidad.

Public Sub Draw()
    If Not Entity Is Nothing Then Call Entity.Draw
End Sub

Public Sub Update()
    If Not Entity Is Nothing Then Entity.Location = Engine.Scene.ClientToScreen(Engine.Scene.LevelEditor.Brush.Location)
End Sub

Public Sub SetEntity(Key As String)
    If Not EntityForm Is Nothing Then Call Unload(EntityForm)
    Select Case Key
        Case "Jugador"
        
        Case "Particula"
            Set EntityForm = frm_ACT_ParticleEnemy
            Set Entity = New TLSA.ACT_ParticleEnemy
        Case "Plataforma"
        
        Case "Salida"
        
    End Select
    If Not EntityForm Is Nothing Then Call EntityForm.Show
End Sub

Public Sub ShowFormParams()
    If Not EntityForm Is Nothing Then Call EntityForm.Show
End Sub

Public Sub HideFormParams()
    If Not EntityForm Is Nothing Then Call EntityForm.Hide
End Sub

Public Sub Add()
    'Entity.Location = Engine.Scene.ClientToScreen(Core.Generics.CreatePOINT(Entity.Location.X, Entity.Location.Y))
    Call Engine.Scene.Entities.Add(Entity.Clone())
End Sub

Public Sub Remove()

End Sub

Private Sub Class_Terminate()
    Set Entity = Nothing
End Sub
