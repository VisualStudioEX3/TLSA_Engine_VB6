VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EDIT_Manager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' El gestor del editor es la interfaz principal para acceder y controlar las acciones de edicion del editor de niveles
' del motor:

Option Explicit

Public Enum EDIT_Modes
    TileEdition
    PhysicEdition
    EntityEdition
End Enum

Public EditMode As TLSA.EDIT_Modes

' Objeto pincel. Sera el encargado de indicar la posicion de insercion de objetos en la vista de edicion:
Public WithEvents Brush As TLSA.EDIT_Brush
Attribute Brush.VB_VarHelpID = -1
Public Grid As TLSA.EDIT_Grid

Public PhysicEditor As TLSA.EDIT_Physics    ' Instancia del editor de fisicas.
Public TileEditor As TLSA.EDIT_Tilemap      ' Instancia del editor de mapas de tiles.
Public EntityEditor As TLSA.EDIT_Entities   ' Instancia del editor de entidades.
    
Public ShowGrid As Boolean

Public Sub Draw()
    Call Grid.Draw              ' Dibuja la rejilla de edicion.
    Call Brush.Draw             ' Dibuja el pincel.
    Select Case EditMode
        Case EDIT_Modes.TileEdition: Call TileEditor.Draw
        Case EDIT_Modes.PhysicEdition: Call PhysicEditor.Draw
        Case EDIT_Modes.EntityEdition: Call EntityEditor.Draw
    End Select
End Sub

Public Sub Update()
    'Engine.Scene.Enabled = False
    
    ' Implementar controles para desplazar la posicion de la camara en el modo de edicion:
        
    Call Brush.Update           ' Actualiza el estado del pincel.
    Select Case EditMode
        Case EDIT_Modes.TileEdition: Call TileEditor.Update
        Case EDIT_Modes.PhysicEdition: Call PhysicEditor.Update
        Case EDIT_Modes.EntityEdition: Call EntityEditor.Update
    End Select
End Sub

' Evento donde controlaremos la insercion de objetos:
Private Sub Brush_Hit(X As Long, Y As Long)
    Select Case EditMode
        Case TLSA.EDIT_Modes.TileEdition: Call TileEditor.Add
        Case TLSA.EDIT_Modes.PhysicEdition: Call PhysicEditor.Add
        Case TLSA.EDIT_Modes.EntityEdition: Call EntityEditor.Add
    End Select
End Sub

' Evento donde controlaremos la eliminacion de objetos:
Private Sub Brush_AltHit(X As Long, Y As Long)
    Select Case EditMode
        Case TLSA.EDIT_Modes.TileEdition: Call TileEditor.Remove
        Case TLSA.EDIT_Modes.PhysicEdition: Call PhysicEditor.Remove
        Case TLSA.EDIT_Modes.EntityEdition: Call EntityEditor.Remove
    End Select
End Sub

' Activa el editor seleccionado:
Public Sub SetEditor(Mode As TLSA.EDIT_Modes)
    Select Case EditMode
        Case EDIT_Modes.TileEdition: Call frmTileEditor.Hide
        Case EDIT_Modes.PhysicEdition: Call frmPhysParams.Hide: Call Engine.Scene.LevelEditor.PhysicEditor.HidePreview
        Case EDIT_Modes.EntityEdition: Call frmEntityList.Hide:: Call Engine.Scene.LevelEditor.EntityEditor.HideFormParams
    End Select
    
    Select Case Mode
        Case EDIT_Modes.TileEdition: Call frmTileEditor.Show
        Case EDIT_Modes.PhysicEdition: Call frmPhysParams.Show: Call Engine.Scene.LevelEditor.PhysicEditor.ShowPreview
        Case EDIT_Modes.EntityEdition: Call frmEntityList.Show: Call Engine.Scene.LevelEditor.EntityEditor.ShowFormParams
    End Select
    
    EditMode = Mode
End Sub

' Exporta la informacion del nivel a disco:
Public Sub Export(filename As String)
    Dim file As Integer: file = FreeFile()
    Open filename For Output As file
        ' Parametros del simulador de fisicas:
        Print #file, "// Physic Simulator"
        Print #file, "Engine.Scene.Physics.Gravity=6.0;0.0" ' Gravedad en X e Y.
        Print #file, "Engine.Scene.Physics.Enabled=1"
        
        ' Parametros iniciales de la camara:
        Print #file, "// Camera_Initialization"
        Print #file, "Engine.Scene.Camera.Location=0;0"
        Print #file, "Engine.scene.Camera.Target=Null"
        
        ' Texturas
        Print #file, "// Textures"
        Dim tex As Variant
        For Each tex In Engine.Scene.Textures
            Print #file, "Engine.Scene.Textures.Add=" & tex
        Next
        
        ' Tiles
        Print #file, "// Tiles"
        Dim Tile As TLSA.ENG_Tile
        For Each Tile In Engine.Scene.TileMap
            Print #file, "Engine.Scene.Tiles.Add=" & Tile.ToString()
        Next
        
        ' Fisica
        Print #file, "// Level Physics"
        Dim BBox As TLSA.ENG_BBox
        For Each BBox In Engine.Scene.PhysicMap
            Print #file, "Engine.Scene.Physics.Add=" & BBox.ToString()
        Next
        
        ' Entidades
        Print #file, "// Level Entities"
        Dim Entity As Object
        For Each Entity In Engine.Scene.Entities
            ' Scene.Entity. + (<objetc>.Key + <object>.Data)
            Print #file, "Engine.Scene.Entity." & Entity.ToString()
        Next
    Close file
End Sub

Private Sub Class_Initialize()
    Set Grid = New TLSA.EDIT_Grid
    Set Brush = New TLSA.EDIT_Brush
    
    Brush.Size = Core.Generics.CreateSIZE(64, 64)
    
    Set TileEditor = New TLSA.EDIT_Tilemap
    Set PhysicEditor = New TLSA.EDIT_Physics
    Set EntityEditor = New TLSA.EDIT_Entities
    
    Call SetEditor(TileEdition)
    
    Call frmMainMenu.Show
End Sub

Private Sub Class_Terminate()
    Set PhysicEditor = Nothing
    Set TileEditor = Nothing
    Set EntityEditor = Nothing
    
    Set Brush = Nothing
    Set Grid = Nothing
    
    Call Unload(frmMainMenu)
End Sub
