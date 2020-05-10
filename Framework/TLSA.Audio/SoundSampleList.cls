VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SoundSampleList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private varCol As New Collection

' Carga una muestra de sonido en la biblioteca:
Public Function LoadSample(Filename As String, Key As String) As Audio.SoundSample
    Dim s As New Audio.SoundSample
    
    ' Comprobamos que no se cargo ya el archivo anteriormente:
    If ExistsByFilename(Filename) Then Call Err.Raise(Audio.AUDIO_EXCEPTION.RESOURCE_NOT_LOAD, , "La muestra de sonido ya se cargo anteriormente en memoria: '" & Filename & "'")
        
    Call s.LoadSample(Filename)
    s.Key = Key
    
    On Error Resume Next
    Call varCol.Add(s, Key)
    If Err.Number = Audio.AUDIO_EXCEPTION.ITEM_KEY_DUPLICATED Then _
        Call Err.Raise(Audio.AUDIO_EXCEPTION.ITEM_KEY_DUPLICATED, , "Ya existe una muestra de sonido con esta clave: '" & Key & "'")
    On Error GoTo 0
    
    Set LoadSample = s
    Set s = Nothing
End Function

'Public Sub Add(Sample As Audio.SoundSample, Key As String)
'    On Error GoTo ErrOut
'
'    Dim s As New Audio.SoundSample
'
'    s.Key = Key
'    s.Id = Sample.Id
'    s.Filename = Sample.Filename
'
'    Call varCol.Add(s, Key)
'    Set s = Nothing
'
'    Exit Sub
'
'ErrOut:
'    Call Err.Raise(Audio.AUDIO_EXCEPTION.ITEM_KEY_DUPLICATED, , "Ya existe una muestra de sonido con esta clave: '" & Key & "'")
'End Sub

Public Sub Remove(Key As String)
    On Error GoTo ErrOut
    Call varCol.Remove(Key)
    Exit Sub
ErrOut:
    Call Err.Raise(Audio.AUDIO_EXCEPTION.ITEM_KEY_NOT_FOUND, , "No se encontro ninguna muestra de sonido con la clave '" & Key & "'")
End Sub

' Nombre o indice del tile en la coleccion:
Public Property Get Item(Key As Variant) As Audio.SoundSample
Attribute Item.VB_UserMemId = 0
    On Error GoTo ErrOut
    Set Item = varCol(Key)
    Exit Property
ErrOut:
    Call Err.Raise(Audio.AUDIO_EXCEPTION.ITEM_KEY_NOT_FOUND, , "No se encontro ninguna muestra de sonido con la clave '" & Key & "'")
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
        Dim s As Audio.SoundSample
        For Each s In varCol
            Call varCol.Remove(s.Key)
            Set s = Nothing
        Next
        Set varCol = New Collection
    End If
End Sub

' Devuelve verdadero si la muestra existe:
Public Function Exists(Key As String) As Boolean
    On Error GoTo ErrOut
    Dim s As Audio.SoundSample
    Set s = varCol(Key)
    Exists = True
    Exit Function
ErrOut:
End Function

' Devuelve verdadero si ya se cargo el archivo:
Public Function ExistsByFilename(Filename As String) As Boolean
    On Error GoTo ErrOut
    Dim s As Audio.SoundSample
    For Each s In varCol
        If LCase(s.Filename) = LCase(Filename) Then
            ExistsByFilename = True
            Exit Function
        End If
    Next
ErrOut:
End Function

Public Property Get Count() As Integer
    Count = varCol.Count
End Property

