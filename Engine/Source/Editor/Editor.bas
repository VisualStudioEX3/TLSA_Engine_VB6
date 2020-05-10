Attribute VB_Name = "Editor"
Option Explicit

Private tempFile As String

Public EditMode As Boolean

' Ejecuta la escena actual del editor:
Public Sub Run()
    ' Creamos una copia temporal en disco de la escena:
    tempFile = Core.IO.CreateTemporalFilename("tmp")
    Call Engine.Scene.LevelEditor.Export(tempFile)
    
    ' Habilitamos la logica de la escena:
    Call EnabledUpdate
End Sub

' Detiene momentaneamnete la ejecucion la escena:
Public Sub Pause()
    Call SetEnabled(Not Engine.Scene.Enabled)
End Sub

' Reinicia la escena desde el estado previo:
Public Sub Restart()
    ' Restauramos el estado anterior:
    Call RestorePrevius
    
    ' Habilitamos la logica de la escena:
    Call EnabledUpdate
End Sub

' Termina la ejecucion de la escena y la restaura a su estado previo:
Public Sub Terminate()
    ' Reiniciamos escena:
    Call Restart
    Call Engine.Scene.Update
    
    ' Deshabilitamos las logica de la escena:
    Call SetEnabled(False)
End Sub

Private Sub EnabledUpdate()
    Engine.Scene.PhysicSimulator.Gravity = Core.Generics.CreateVECTOR(0, 6, 0)
    Call SetEnabled(True)
End Sub

Private Sub SetEnabled(Status As Boolean)
    Engine.Scene.PhysicSimulator.Enabled = Status
    Engine.PlayerInputEnabled = Status
    Engine.Scene.Enabled = Status
    EditMode = Status
End Sub

' Restaura el nivel al estado anterior despues de la ejecucion:
Private Sub RestorePrevius()
    ' Borramos el contenido de la escena:
    Call Engine.Scene.Clear
    
    ' Cargamos la escena temporal:
    Call Engine.Scene.LoadScene(tempFile)
End Sub
