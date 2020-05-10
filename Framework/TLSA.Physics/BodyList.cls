VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BodyList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'variable local para contener colección
Private mCol As Collection
Private varFatherSim As Simulator

Friend Property Get FatherSim() As Simulator
    Set FatherSim = varFatherSim
End Property

Friend Property Set FatherSim(sim As Simulator)
    Set varFatherSim = sim
End Property


Public Function Add(X As Long, Y As Long, Z As Long, Width As Long, Height As Long, ColliderBaseType As Physics.ColliderProfiles, Weight As Long, Fixed As Boolean, Param As Variant) As Body
    'crear un nuevo objeto
    Dim objNewMember As Physics.Body
    Set objNewMember = New Physics.Body

    'establecer las propiedades que se transfieren al método
    With objNewMember
        .Weight = Weight
        .Enabled = True
        .Fixed = Fixed
        .PhysicType = FixedBody
        Call .SetLocation(X, Y)
        .Z = Z ' La Z se usa para determinar las capas de colision donde se ubica el objeto, para discriminar colisiones entre objetos.
        Call .SetRect(X - (Width \ 2), Y - (Height \ 2), Width, Height) ' Inicializa el rectangulo que define el area.
        Set .FatherSim = Me.FatherSim
    End With
    
    ' Generamos los colisionadores base acorde con ColliderBaseType y Param si fuera necesario:
    Call objNewMember.SetColliders(ColliderBaseType, Param)
        
    Call mCol.Add(objNewMember, objNewMember.Key)

    'devolver el objeto creado
    Set Add = objNewMember
    Set objNewMember = Nothing

End Function


' Agrega una referencia o instancia de un objeto tipo Entidad:
Public Sub AddObject(E As Body)
    Call mCol.Add(E, E.Key)
End Sub


Public Property Get Item(Index As Long) As Body
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


Public Sub Remove(E As Body)
    'se usa al quitar un elemento de la colección
    'vntIndexKey contiene el índice o la clave, por lo que se
    'declara como un Variant
    'Sintaxis: x.Remove(xyz)
    mCol.Remove E.Key
End Sub


Public Sub Clear()
    Dim b As Body
    For Each b In mCol
        Call mCol.Remove(b.Key)
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

