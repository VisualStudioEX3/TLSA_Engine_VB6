VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AnimationList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private varCol As New Collection

Public Sub Add(Secuence As Graphics.Animation, Key As String)
    On Error GoTo ErrOut
    
    Dim A As New Graphics.Animation
    
    A.AnimatePath = Secuence.AnimatePath
    A.FrameDelay = Secuence.FrameDelay
    A.Looping = Secuence.Looping
    Set A.Tiles = Secuence.Tiles
    A.Key = LCase(Key)
    Call A.SetTexture(Secuence.Texture)
    
    Call varCol.Add(A, A.Key)
    Set A = Nothing
    
    Exit Sub
    
ErrOut:
    Call Err.Raise(Graphics.GRAPHICS_EXCEPTION.ITEM_KEY_DUPLICATED, , "Ya existe una animacion con esta clave: '" & LCase(Key) & "'")
End Sub

Public Sub Remove(Key As String)
    On Error GoTo ErrOut
    Call varCol.Remove(LCase(Key))
    Exit Sub
ErrOut:
    Call Err.Raise(Graphics.GRAPHICS_EXCEPTION.ITEM_KEY_NOT_FOUND, , "No existe ninguna animacion con esta clave: '" & LCase(Key) & "'")
End Sub

' Nombre o indice del tile en la coleccion:
Public Property Get Item(Key As Variant) As Graphics.Animation
Attribute Item.VB_UserMemId = 0
    On Error GoTo ErrOut
    Set Item = varCol(Key)
    Exit Sub
ErrOut:
    Call Err.Raise(Graphics.GRAPHICS_EXCEPTION.ITEM_KEY_NOT_FOUND, , "No existe ninguna animacion con esta clave: '" & LCase(Key) & "'")
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'esta propiedad permite enumerar
    'esta colección con la sintaxis For...Each
    Set NewEnum = varCol.[_NewEnum]
End Property

Public Sub Clear()
    If varCol.Count > 0 Then
        Dim A As Graphics.Animation
        For Each A In varCol
            Call varCol.Remove(A.Key)
            Set A = Nothing
        Next
        Set varCol = New Collection
    End If
End Sub

' Devuelve verdadero si la fuente existe:
Public Function Exists(Key As String) As Boolean
    On Error GoTo ErrOut
    Dim A As Graphics.Animation
    Set A = varCol(LCase(Key))
    Exists = True
    Exit Function
ErrOut:
End Function

' Devuelve el indice de la clave en la lista:
Public Function IndexOf(Key As String) As Long
    Dim i As Long
    i = 1
    Dim A As Graphics.Animation
    For Each A In varCol
        If LCase(Key) = A.Key Then
            IndexOf = i
            Exit Function
        End If
        i = i + 1
    Next
End Function

Public Property Get Count() As Integer
    Count = varCol.Count
End Property
