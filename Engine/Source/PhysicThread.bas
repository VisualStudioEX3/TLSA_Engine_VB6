Attribute VB_Name = "PhysicThread"
' Este modulo permite separar la llamada de actualizacion de estados del simulador de fisicas en un evento externo al hilo principal de ejecucion
' simulando una suerte de multihilo.

Option Explicit

Private thread As New Core.TimerEvent
Public ThreadPhysicSimulator As Physics.Simulator

Public Sub CreateThread()
    Set ThreadPhysicSimulator = New Physics.Simulator
    Call thread.SetEvent(0, AddressOf UpdatePhysics)
End Sub

Public Sub KillThread()
    Set thread = Nothing
End Sub

Private Sub UpdatePhysics()
    If Not ThreadPhysicSimulator Is Nothing Then Call ThreadPhysicSimulator.Update
End Sub
