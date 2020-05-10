VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SurfaceList"
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
Public Property Get Item(Key As String) As Graphics.Surface
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

' Crea una superficie vacia:
Public Function Create(Key As String, Width As Long, Height As Long) As Graphics.Surface
    Dim surf As New Graphics.Surface
    
    surf.Key = Key
    If Me.Exists(Key) Then
        Call Err.Raise(Graphics.GRAPHICS_EXCEPTION.ITEM_KEY_DUPLICATED, , "Ya existe una superficie con la misma clave: '" & Key & "'")
    Else
        Call surf.Create(Width, Height)
        Call mCol.Add(surf, Key)
    End If
    
    Set Create = surf
    Set surf = Nothing
End Function

' Carga una superficie desde disco:
Public Function LoadSurface(Filename As String, Key As String) As Graphics.Surface
    Dim surf As New Graphics.Surface
    
    ' Comprobamos que no se cargo ya el archivo anteriormente:
    If ExistsByFilename(Filename) Then Call Err.Raise(Graphics.GRAPHICS_EXCEPTION.RESOURCE_NOT_LOAD, , "La superficie ya se cargo anteriormente en memoria: '" & Filename & "'")
    
    surf.Key = Key
    If Me.Exists(Key) Then
        Call Err.Raise(Graphics.GRAPHICS_EXCEPTION.ITEM_KEY_DUPLICATED, , "Ya existe una superficie con la misma clave: '" & Key & "'")
    Else
        Call surf.LoadSurface(Filename)
        Call mCol.Add(surf, Key)
    End If
    
    Set LoadSurface = surf
    Set surf = Nothing
End Function

' Carga una superficie desde memoria:
Public Function LoadSurfaceFromMemory(buffer() As Byte, Key As String) As Graphics.Surface
    Dim surf As New Graphics.Surface
    
    surf.Key = Key
    If Me.Exists(Key) Then
        Call Err.Raise(Graphics.GRAPHICS_EXCEPTION.ITEM_KEY_DUPLICATED, , "Ya existe una superficie con la misma clave: '" & Key & "'")
    Else
        Call surf.LoadSurfaceFromMemory(buffer())
        Call mCol.Add(surf, Key)
    End If
    
    Set LoadSurfaceFromMemory = surf
    Set surf = Nothing
End Function

' Crea una superficie con el contenido de la escena:
Public Function CreateScreenCapture(Key As String) As Graphics.Surface
    Dim surf As New Graphics.Surface
    
    surf.Key = Key
    If Me.Exists(Key) Then
        Call Err.Raise(Graphics.GRAPHICS_EXCEPTION.ITEM_KEY_DUPLICATED, , "Ya existe una superficie con la misma clave: '" & Key & "'")
    Else
        Call surf.CreateScreenCapture
        Call mCol.Add(surf, Key)
    End If
    
    Set CreateScreenCapture = surf
    Set surf = Nothing
End Function

' Llamada desde Graphics.Texture.ToSurface() y Graphics.Tile.ToSurface():
Friend Function CreateFromTexture(Texture As Graphics.Texture, Region As Core.RECTANGLE) As Graphics.Surface
    If Me.Exists(Texture.Key) Then
        Call Err.Raise(Graphics.GRAPHICS_EXCEPTION.ITEM_KEY_DUPLICATED, , "La clave de la textura o tile ya esta siendo usada por una superficie: '" & Texture.Key & "'")
    Else
        Dim varSurf As Graphics.Surface
        Set varSurf = Me.Create(Texture.Key, Texture.Information.Texture.Width, Texture.Information.Texture.Height)
        Call varSurf.CopyFromTexture(Texture, Region)
        Set CreateFromTexture = varSurf
    End If
End Function

' Descarga una superficie:
Public Sub Unload(Key As String)
    Call mCol.Remove(Key)
End Sub

' Descarga todas las superficies de la lista:
Public Sub UnloadAll()
    If mCol.Count > 0 Then
        Dim s As Graphics.Surface
        For Each s In mCol
            Call mCol.Remove(s.Key)
        Next
    End If
End Sub

' Devuelve verdadero si la fuente existe:
Public Function Exists(Key As String) As Boolean
    On Error GoTo ErrOut
    Dim s As Graphics.Surface
    Set s = mCol(Key)
    Exists = True
    Exit Function
ErrOut:
End Function

' Devuelve verdadero si ya se cargo el archivo:
Public Function ExistsByFilename(Filename As String) As Boolean
    On Error GoTo ErrOut
    Dim s As Graphics.Surface
    For Each s In mCol
        If LCase(s.Filename) = LCase(Filename) Then
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

