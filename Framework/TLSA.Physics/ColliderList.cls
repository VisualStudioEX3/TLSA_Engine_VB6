VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ColliderList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"Box"
Attribute VB_Ext_KEY = "Member0" ,"Box"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

'variable local para contener colección
Private mCol As Collection

Public Function Add(X As Long, Y As Long, Z As Long, Width As Long, Height As Long, TypeBox As Long, BodyFather As Body) As Physics.Collider
    'crear un nuevo objeto
    Dim objNewMember As Physics.Collider
    Set objNewMember = New Physics.Collider

    'establecer las propiedades que se transfieren al método
    With objNewMember
        .TypeBox = TypeBox
        .Enabled = True
        Set .BodyFather = BodyFather
        Call .SetRect(X + BodyFather.Rect.X, Y + BodyFather.Rect.Y, Width, Height) ' Inicializa el rectangulo que define el area.
    End With
        
    Call mCol.Add(objNewMember, objNewMember.Key)

    'devolver el objeto creado
    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Friend Sub AddObject(b As Physics.Collider)
    Call mCol.Add(b, b.Key)
End Sub


Public Property Get Item(Index As Long) As Physics.Collider
Attribute Item.VB_UserMemId = 0
    'se usa al hacer referencia a un elemento de la colección
    'vntIndexKey contiene el índice o la clave de la colección,
    'por lo que se declara como un Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
    Set Item = mCol(Index)
End Property


Public Property Get Count() As Long
    'se usa al obtener el número de elementos de la
    'colección. Sintaxis: Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(b As Physics.Collider)
    'se usa al quitar un elemento de la colección
    'vntIndexKey contiene el índice o la clave, por lo que se
    'declara como un Variant
    'Sintaxis: x.Remove(xyz)
    mCol.Remove b.Key
End Sub


Public Sub Clear()
    Dim C As Physics.Collider
    For Each C In mCol
        Call mCol.Remove(C.Key)
    Next
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'esta propiedad permite enumerar
    'esta colección con la sintaxis For...Each
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'crea la colección cuando se crea la clase
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    'destruye la colección cuando se termina la clase
    Set mCol = Nothing
End Sub
