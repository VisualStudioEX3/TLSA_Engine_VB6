VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TextureList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mCol As New Collection

' Numero de texturas en la lista:
Public Property Get Count() As Integer
    Count = mCol.Count
End Property

' Devuelve la instancia de una textura en la lista:
Public Property Get Item(Key As String) As Graphics.Texture
Attribute Item.VB_UserMemId = 0
    Set Item = mCol(Key)
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'esta propiedad permite enumerar
    'esta colección con la sintaxis For...Each
    Set NewEnum = mCol.[_NewEnum]
End Property

' Crea una textura vacia:
Public Function CreateTexture(Key As String, Width As Long, Height As Long, GrayScale As Boolean, RenderTarget As Boolean) As Texture
    Dim tex As New Graphics.Texture
    
    tex.Key = Key
    Call tex.Create(Width, Height, GrayScale, RenderTarget)
        
    On Error Resume Next
    Call mCol.Add(tex, Key)
    If Err.Number = Graphics.GRAPHICS_EXCEPTION.ITEM_KEY_DUPLICATED Then
        Call Err.Raise(Graphics.GRAPHICS_EXCEPTION.ITEM_KEY_DUPLICATED, , "Ya existe una textura con la misma clave: '" & Key & "'")
    End If
    On Error GoTo 0
    
    Set CreateTexture = tex
    Set tex = Nothing
End Function

' Carga una textura desde disco:
Public Function LoadTexture(Filename As String, Key As String, GrayScale As Boolean) As Texture
    Dim tex As New Graphics.Texture
    
    ' Comprobamos que no se cargo ya el archivo anteriormente:
    If ExistsByFilename(Filename) Then Call Err.Raise(Graphics.GRAPHICS_EXCEPTION.RESOURCE_NOT_LOAD, , "La textura ya se cargo anteriormente en memoria: '" & Filename & "'")
    
    tex.Key = Key
    Call tex.LoadTexture(Filename, GrayScale)
        
    On Error Resume Next
    Call mCol.Add(tex, Key)
    If Err.Number = Graphics.GRAPHICS_EXCEPTION.ITEM_KEY_DUPLICATED Then
        Call Err.Raise(Graphics.GRAPHICS_EXCEPTION.ITEM_KEY_DUPLICATED, , "Ya existe una textura con la misma clave: '" & Key & "'")
    End If
    On Error GoTo 0
    
    Set LoadTexture = tex
    Set tex = Nothing
End Function

' Carga una textura desde memoria:
Public Function LoadTextureFromMemory(buffer() As Byte, Key As String, GrayScale As Boolean) As Texture
    Dim tex As New Graphics.Texture
    
    Call tex.LoadTextureFromMemory(buffer(), GrayScale)
    tex.Key = Key
    
    On Error Resume Next
    Call mCol.Add(tex, Key)
    If Err.Number = Graphics.GRAPHICS_EXCEPTION.ITEM_KEY_DUPLICATED Then
        Call Err.Raise(Graphics.GRAPHICS_EXCEPTION.ITEM_KEY_DUPLICATED, , "Ya existe una textura con la misma clave: '" & Key & "'")
    End If
    On Error GoTo 0
    
    Set LoadTextureFromMemory = tex
    Set tex = Nothing
End Function

' Descarga una textura:
Public Sub Unload(Key As String)
    If Me.Item(Key).IsFreeRef() Then Call mCol.Remove(Key)
End Sub

' Descarga todas las texturas de la lista:
Public Sub UnloadAll()
    If mCol.Count > 0 Then
        Dim t As Graphics.Texture
        For Each t In mCol
            Call t.ClearRef         ' Anulamos las referencias que tuviera la textura para poder ser eliminada.
            Call mCol.Remove(t.Key)
        Next
    End If
End Sub

' Devuelve verdadero si la textura existe:
Public Function Exists(Key As String) As Boolean
    On Error GoTo ErrOut
    Dim t As Graphics.Texture
    Set t = mCol(Key)
    Exists = True
    Exit Function
ErrOut:
End Function

' Devuelve verdadero si ya se cargo el archivo:
Public Function ExistsByFilename(Filename As String) As Boolean
    On Error GoTo ErrOut
    Dim t As Graphics.Texture
    For Each t In mCol
        If LCase(t.Filename) = LCase(Filename) Then
            ExistsByFilename = True
            Exit Function
        End If
    Next
ErrOut:
End Function

Private Sub Class_Terminate()
    Call UnloadAll
    Set mCol = Nothing
End Sub
