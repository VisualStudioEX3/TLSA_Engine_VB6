VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TileList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private varCol As New Collection

Public Sub Add(Tile As Graphics.Tile, Key As String)
    On Error GoTo ErrOut
    
    Dim t As New Graphics.Tile
    
    t.Center = Tile.Center
    t.Key = Key
    t.Region = Tile.Region
    Call t.SetControlPoints(Tile.ControlPoints.Clone)
    Call t.SetTexture(Tile.Texture)
    
    Call varCol.Add(t, Key)
    Set t = Nothing
    
    Exit Sub
    
ErrOut:
    Call Err.Raise(Graphics.GRAPHICS_EXCEPTION.ITEM_KEY_DUPLICATED, , "Ya existe un tile con esta clave: '" & Key & "'")
End Sub

' Añade el tile como referencia y no como copia a la lista:
Public Sub AddRef(Tile As Graphics.Tile, Key As String)
    On Error GoTo ErrOut
        
    Call varCol.Add(Tile, Key)
    Exit Sub
    
ErrOut:
    Call Err.Raise(Graphics.GRAPHICS_EXCEPTION.ITEM_KEY_DUPLICATED, , "Ya existe un tile con esta clave: '" & Key & "'")
End Sub

Public Function Create(Key As String, Texture As Graphics.Texture, X As Long, Y As Long, Width As Long, Height As Long, CenterX As Long, CenterY As Long) As Graphics.Tile
    On Error GoTo ErrOut
    
    Dim t As New Graphics.Tile
    
    t.Key = Key
    Call t.SetRegion(X, Y, Width, Height)
    Call t.SetCenter(CenterX, CenterY)
    Call t.SetTexture(Texture)
    
    Call varCol.Add(t, Key)
    Set Create = t
    Set t = Nothing
    
    Exit Function
    
ErrOut:
    Call Err.Raise(Graphics.GRAPHICS_EXCEPTION.ITEM_KEY_DUPLICATED, , "Ya existe un tile con esta clave: '" & Key & "'")
End Function

Public Sub Remove(Key As String)
    Call varCol.Remove(Key)
End Sub

' Nombre o indice del tile en la coleccion:
Public Property Get Item(Key As Variant) As Tile
Attribute Item.VB_UserMemId = 0
    'If IsNumeric(Key) Then Key = Key + 1 ' Nosotros trabajamos en base 0 los indices pero las colecciones en VB6.0 trabajan en base 1.
    Set Item = varCol(Key)
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
        Dim t As Graphics.Tile
        For Each t In varCol
            Call varCol.Remove(t.Key)
            Set t = Nothing
        Next
        Set varCol = New Collection
    End If
End Sub

' Devuelve verdadero si la fuente existe:
Public Function Exists(Key As String) As Boolean
    On Error GoTo ErrOut
    Dim t As Graphics.Tile
    Set t = varCol(Key)
    Exists = True
    Exit Function
ErrOut:
End Function

' Devuelve el indice de la clave en la lista:
Public Function IndexOf(Key As String) As Long
    Dim i As Long
    For i = 1 To varCol.Count
        If Key = varCol(i) Then
            IndexOf = i
            Exit Function
        End If
    Next
End Function

Public Property Get Count() As Integer
    Count = varCol.Count
End Property
