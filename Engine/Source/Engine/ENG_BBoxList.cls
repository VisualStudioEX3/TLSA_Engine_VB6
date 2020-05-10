VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ENG_BBoxList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private varCol As Collection

' Añade un cuerpo fisico y genera un BBox en la lista:
Public Function Add(Body As Physics.Body) As TLSA.ENG_BBox
    Dim b As New TLSA.ENG_BBox
    Set b.Body = Body
    Call varCol.Add(b, b.Key)
    Set Add = b
    Set b = Nothing
End Function

' Añade un BBox mediante referencia a la coleccion:
Public Sub AddRef(BBox As TLSA.ENG_BBox)
    Call varCol.Add(BBox, BBox.Key)
End Sub

' Elimina una BBox de la escena:
Public Sub Remove(BBox As TLSA.ENG_BBox)
    Call varCol.Remove(BBox.Key)
End Sub

' Accede un BBox de la lista por su clave unica:
Public Property Get Item(Key As String) As TLSA.ENG_BBox
Attribute Item.VB_UserMemId = 0
    Set Item = varCol(Key)
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'esta propiedad permite enumerar
    'esta colección con la sintaxis For...Each
    Set NewEnum = varCol.[_NewEnum]
End Property

' Devuelve el numero de BBoxes en la lista:
Public Property Get Count() As Long
    Count = varCol.Count
End Property

' Devuelve una coleccion con todos los BBoxes visibles en pantalla:
Public Function GetVisibles() As TLSA.ENG_BBoxList
    Dim ite As TLSA.ENG_BBox, ret As New TLSA.ENG_BBoxList
    Dim scrRect As Core.Rectangle
    scrRect = Core.Generics.CreateRECTANGLE(0, 0, Engine.GraphicEngine.CurrentDisplayMode.Width, Engine.GraphicEngine.CurrentDisplayMode.Height)
    For Each ite In varCol
        If Core.Math.IntersectRect(ite.ClientToScreenRectangle(), scrRect) Then Call ret.AddRef(ite)
    Next
    Set GetVisibles = ret
End Function

' Devuelve la lista en sentido inverso:
Public Function Inverse() As TLSA.ENG_BBoxList
    Dim ite As TLSA.ENG_BBox, arr() As TLSA.ENG_BBox, i As Long, ret As New TLSA.ENG_BBoxList
    
    If varCol.Count > 0 Then
        ' Rellenamos el array con los valores de la coleccion:
        ReDim arr(varCol.Count - 1)
        For Each ite In varCol
            Set arr(i) = ite
            i = i + 1
        Next
        
        ' Rellenamos la coleccion en sentido inverso:
        For i = UBound(arr) To 0 Step -1
            Call ret.AddRef(arr(i))
        Next
    End If
    
    Set Inverse = ret
End Function

Public Sub Clear()
    Set varCol = Nothing
    Set varCol = New Collection
End Sub

Private Sub Class_Initialize()
    Set varCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set varCol = Nothing
End Sub

