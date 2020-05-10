VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FontList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mCol As New Collection

' Numero de fuentes en la lista:
Public Property Get Count() As Integer
    Count = mCol.Count
End Property

' Devuelve la instancia de una fuente en la lista:
Public Property Get Item(Key As String) As Font
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

' Carga una fuente de Windows:
Public Sub LoadFont(Font As String, Key As String, Size As Long, Bold As Boolean, Italic As Boolean, Underline As Boolean, Strikethrough As Boolean)
    
    Dim fon As New Graphics.Font
    
    Call fon.LoadFont(Font, Size, Bold, Italic, Underline, Strikethrough)
    fon.Key = Key
    
    On Error Resume Next
    Call mCol.Add(fon, Key)
    If Err.Number = Graphics.GRAPHICS_EXCEPTION.ITEM_KEY_DUPLICATED Then
        Call Err.Raise(Graphics.GRAPHICS_EXCEPTION.ITEM_KEY_DUPLICATED, , "Ya existe una fuente cargada con esa clave: '" & Key & "'")
    End If
    On Error GoTo 0
    
    Set fon = Nothing
End Sub

' Carga una fuente desde archivo:
Public Sub LoadFontFromFile(Filename As String, Key As String, Size As Long, Bold As Boolean, Italic As Boolean, Underline As Boolean, Strikethrough As Boolean)
    Dim fon As New Graphics.Font
    
    Call fon.LoadFontFromFile(Filename, Size, Bold, Italic, Underline, Strikethrough)
    fon.Key = Key
    
    On Error Resume Next
    Call mCol.Add(fon, Key)
    If Err.Number = Graphics.GRAPHICS_EXCEPTION.ITEM_KEY_DUPLICATED Then
        Call Err.Raise(Graphics.GRAPHICS_EXCEPTION.ITEM_KEY_DUPLICATED, , "Ya existe una fuente cargada con esa clave.")
    End If
    On Error GoTo 0
    
    Set fon = Nothing
End Sub

' Carga una fuente desde un archivo de definicion de fuentes:
Public Sub LoadFontFileDefinition(Filename As String, Key As String)
    Dim fon As New Graphics.Font
    
    Call fon.Import(Filename)
    fon.Key = Key
    
    On Error Resume Next
    Call mCol.Add(fon, Key)
    If Err.Number = Graphics.GRAPHICS_EXCEPTION.ITEM_KEY_DUPLICATED Then
        Call Err.Raise(Graphics.GRAPHICS_EXCEPTION.ITEM_KEY_DUPLICATED, , "Ya existe una fuente cargada con esa clave.")
    End If
    On Error GoTo 0
    
    Set fon = Nothing
End Sub

' Descarga una fuente:
Public Sub UnloadFont(Key As String)
    'Set mCol(Key) = Nothing
    Call mCol.Remove(Key)
End Sub

' Descarga todas las texturas de la lista:
Public Sub UnloadAll()
    If mCol.Count > 0 Then
        Dim f As Graphics.Font
        For Each f In mCol
    '        Set f = Nothing
            Call mCol.Remove(f.Key)
        Next
    End If
End Sub

' Devuelve verdadero si la fuente existe:
Public Function Exists(Key As String) As Boolean
    On Error GoTo ErrOut
    Dim f As Graphics.Font
    Set f = mCol(Key)
    Exists = True
    Exit Function
ErrOut:
End Function

Private Sub Class_Terminate()
    Call UnloadAll
    Set mCol = Nothing
End Sub

