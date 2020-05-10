VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EDIT_Physics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Preview As Physics.Body
Public SelectColliderSpecification As Physics.ColliderProfiles
Public Param As Variant
'Public BlockEnable As Boolean

Public Sub HidePreview()
    Preview.Visible = False
    Preview.Enabled = False
End Sub

Public Sub ShowPreview()
    ' Si no se creo el cuerpo de modelo de previsualizacion se crea en este punto:
    If Preview Is Nothing Then Call CreatePreview
    
    Preview.Visible = True
    Preview.Enabled = False
End Sub

Public Sub Draw()
    ' Toda la gestion de dibujo de cuerpos fisicos se encarga el propio simulador de la escena.
End Sub

Public Sub Update()
    ' Actualiza la posicion del cuerpo previsualizado:
    With Engine.Scene.LevelEditor.Brush
        Preview.Location = Core.Generics.POINT2VECTOR(Engine.Scene.ClientToScreen(.Location))
        Call Preview.SetColliders(Me.SelectColliderSpecification, Param)
        Call Preview.SetRect(Preview.Location.X, Preview.Location.Y, .Size.Width, .Size.Height)
    End With
    
    ' Informacion de depuracion:
    Call Engine.dev_debug.AddMessage(Engine.Scene.PhysicSimulator.Monitor.ToString())
End Sub

' Añade un cuerpo fisico a la escena:
Public Sub Add()
    ' Agrega un nuevo cuerpo fisico a la escena utilizando una copia de la previsualizacion
    ' a la que se editaran un par de parametros:
    Dim b As Physics.Body: Set b = Preview.Clone()
    
    b.Color = Graphics.Color_Constant.Orange
    b.Z = 0

    ' Parametros del cuerpo fisico en la escena para generar el archivo en disco:
    b.Tag = CStr(Me.SelectColliderSpecification & ";" & Param)  ' Guardamos la seleccion de colisionadores y el parametro de envergadura.
    b.Enabled = True ' BlockEnable
    Call Engine.Scene.PhysicSimulator.Bodies.AddObject(b)       ' Añadimos el cuerpo al simulador.
    Call Engine.Scene.PhysicMap.Add(b)                      ' Añadimos la referencia al motor (esto es por temas de edicion).
    Set b = Nothing
End Sub

' Elimina un cuerpo fisico de la escena:
Public Sub Remove()
    ' El algorritmo de eliminacion se basa en obtener la lista de los cuerpos visibles en pantalla y buscar el primero
    ' cuya posicion del pincel/cursor este dentro de su area.
    Dim ite As TLSA.ENG_BBox
    For Each ite In Engine.Scene.PhysicMap.GetVisibles().Inverse() ' Accedemos a la lista inversa para tomar en cuenta los ultimos elementos añadidos:
        If Core.Math.PointInRect(Engine.Scene.LevelEditor.Brush.Location, ite.ClientToScreenRectangle()) Then
            Call Engine.Scene.PhysicMap.Remove(ite)
            Exit Sub
        End If
    Next
End Sub

Private Sub Class_Initialize()
    SelectColliderSpecification = Floor
    Param = 32
    'BlockEnable = True
    frmPhysParams.Left = 0 'frmMainMenu.Left + frmMainMenu.Width
    frmPhysParams.Top = frmMainMenu.Top + frmMainMenu.Height
End Sub

' Crea un cuerpo para mostrar la previsualizacion de lo que se va a insertar en el nivel:
Private Sub CreatePreview()
    Set Preview = Engine.Scene.PhysicSimulator.Bodies.Add(0, 0, -8, 32, 32, Me.SelectColliderSpecification, 0, True, Param)
    Preview.Color = Graphics.Color_Constant.Red
    Preview.Tag = "Preview"
    Preview.Enabled = True
End Sub

Private Sub Class_Terminate()
    Call Unload(frmPhysParams)
End Sub
