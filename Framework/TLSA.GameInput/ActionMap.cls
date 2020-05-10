VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ActionMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private varCol As Collection

Private Sub Class_Initialize()
    Set varCol = New Collection
End Sub

' Añade una nueva accion al mapa:
Public Sub Add(Action As String)
    If varCol.Count < 20 Then
        Dim act As New ActionNode
        act.Name = LCase(Action)
        On Error Resume Next
        Call varCol.Add(act, act.Name)
        If Err.Number = GameInput.GAMEINPUT_EXCEPTION.ITEM_KEY_DUPLICATED Then _
            Call Err.Raise(GameInput.GAMEINPUT_EXCEPTION.ITEM_KEY_DUPLICATED, , "Ya existe una accion con el mismo nombre: '" & Action & "'")
        On Error GoTo 0
        Set act = Nothing
    Else
        Call Err.Raise(GameInput.GAMEINPUT_EXCEPTION.MAX_RANGE_EXCEDED, , "El mapa de accion no puede contener mas de 20 acciones definidas.")
    End If
End Sub

' Elimina una accion del mapa:
Public Sub Remove(Action As String)
    On Error GoTo ErrOut
    Call varCol.Remove(LCase(Action))
    Exit Sub
ErrOut:
    Call Err.Raise(GameInput.GAMEINPUT_EXCEPTION.ITEM_KEY_NOT_FOUND, , "No se encontro ninguna accion con ese nombre: '" & LCase(Action) & "'")
End Sub

' Cambia el nombre de una accion:
Public Sub Rename(Action As String, NewAction As String)
    If LCase(Action) = LCase(NewAction) Then _
        Call Err.Raise(GameInput.GAMEINPUT_EXCEPTION.ITEM_KEY_DUPLICATED, , "Se esta intentando renombrar una accion con la misma clave: '" & LCase(NewAction) & "'")
        
    Dim act As GameInput.ActionNode
    
    On Error Resume Next
    
    Set act = varCol(LCase(Action))
    If Err.Number = 5 Then _
        Call Err.Raise(GameInput.GAMEINPUT_EXCEPTION.ITEM_KEY_NOT_FOUND, , "No se encontro ninguna accion con ese nombre: '" & LCase(Action) & "'")
    
    act.Name = LCase(NewAction)
    
    Call varCol.Add(act, act.Name)
    If Err.Number = GameInput.GAMEINPUT_EXCEPTION.ITEM_KEY_DUPLICATED Then _
        Call Err.Raise(GameInput.GAMEINPUT_EXCEPTION.ITEM_KEY_DUPLICATED, , "Ya existe una accion con el mismo nombre: '" & LCase(NewAction) & "'")
    
    On Error GoTo 0
    
    Call varCol.Remove(Action)
    Set act = Nothing
End Sub

Public Property Get Item(Action As String) As GameInput.ActionNode
Attribute Item.VB_UserMemId = 0
    On Error GoTo ErrOut
    Set Item = varCol(LCase(Action))
    Exit Property
ErrOut:
    Call Err.Raise(GameInput.GAMEINPUT_EXCEPTION.ITEM_KEY_NOT_FOUND, , "No se encontro ninguna accion con ese nombre: '" & LCase(Action) & "'")
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'esta propiedad permite enumerar
    'esta colección con la sintaxis For...Each
    Set NewEnum = varCol.[_NewEnum]
End Property

Public Property Get Count() As Long
    Count = varCol.Count
End Property

Public Sub Clear()
    Set varCol = New Collection
End Sub

' Devuelve un elemento de la coleccion por su indice:
Public Function GetValue(index As Long) As GameInput.ActionNode
    If (index > 0) And (index < (varCol.Count + 1)) Then Set GetValue = varCol(index)
End Function

Private Sub Class_Terminate()
    Set varCol = Nothing
End Sub
