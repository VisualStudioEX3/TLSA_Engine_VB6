VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AudioPhysicEffectRegionList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private varCol As Collection

Public Function Create(Key As String, X As Long, Y As Long, Width As Long, Height As Long) As Audio.AudioPhysicEffectRegion
    Dim e As New Audio.AudioPhysicEffectRegion
    
    e.Key = Key
    e.Enabled = True
    Call e.SetRegion(X, Y, Width, Height)
        
    On Error Resume Next
    Call varCol.Add(e, Key)
    If Err.Number = Audio.AUDIO_EXCEPTION.ITEM_KEY_DUPLICATED Then _
        Call Err.Raise(Audio.AUDIO_EXCEPTION.ITEM_KEY_DUPLICATED, , "Ya existe una region de efectos con esta clave: '" & Key & "'")
    On Error GoTo 0
    
    Set Create = e
    Set e = Nothing
End Function

Public Sub Remove(Key As String)
    On Error GoTo ErrOut
    Call varCol.Remove(Key)
    Exit Sub
ErrOut:
    Call Err.Raise(Audio.AUDIO_EXCEPTION.ITEM_KEY_NOT_FOUND, , "No se encontro ninguna region de efectos con la clave '" & Key & "'")
End Sub

Public Property Get Item(Key As String) As Audio.AudioPhysicEffectRegion
Attribute Item.VB_UserMemId = 0
    On Error GoTo ErrOut
    Set Item = varCol(Key)
    Exit Property
ErrOut:
    Call Err.Raise(Audio.AUDIO_EXCEPTION.ITEM_KEY_NOT_FOUND, , "No se encontro ninguna region de efectos con la clave '" & Key & "'")
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'esta propiedad permite enumerar
    'esta colección con la sintaxis For...Each
    Set NewEnum = varCol.[_NewEnum]
End Property

Public Sub Clear()
    Set varCol = Nothing
End Sub

Public Property Get Count() As Long
    Count = varCol.Count
End Property

Private Sub Class_Initialize()
    Set varCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set varCol = Nothing
End Sub
