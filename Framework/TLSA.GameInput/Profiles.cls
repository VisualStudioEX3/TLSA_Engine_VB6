VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Profiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private varCol As Collection

Private Sub Class_Initialize()
    Set varCol = New Collection
End Sub

' Crea un nuevo perfil:
Public Function Create(UserName As String, PlayerIndex As GameInput.PlayerIndex, Device As GameInput.InputDevice) As GameInput.Profile
    If varCol.Count < 4 Then
        Dim prf As New GameInput.Profile
        
        prf.UserName = LCase(UserName)
        prf.Player = PlayerIndex
        prf.Device = Device
        
        On Error Resume Next
        
        Call varCol.Add(prf, UserName)
        
        If Err.Number = GameInput.GAMEINPUT_EXCEPTION.ITEM_KEY_DUPLICATED Then _
            Call Err.Raise(GameInput.GAMEINPUT_EXCEPTION.ITEM_KEY_DUPLICATED, , "Ya existe un perfil con esa clave: '" & UserName & "'")
        
        On Error GoTo 0
        
        Set Create = prf
        Set prf = Nothing
    Else
        Call Err.Raise(GameInput.GAMEINPUT_EXCEPTION.MAX_RANGE_EXCEDED, , "No se pueden crear mas de 4 perfiles.")
    End If
End Function

' Elimina un perfil de la lista con la posibilidad de eliminar tambien el archivo del perfil permanentemente:
Public Sub Remove(UserName As String, Optional DeleteProfile As Boolean = False)
    On Error GoTo ErrOut
    Call varCol.Remove(LCase(UserName))
    If DeleteProfile Then
        If Not varCol(LCase(UserName)).Filename = "" Then Call Kill(varCol(LCase(UserName)).Filename)
    End If
    Exit Sub
ErrOut:
    If Err.Number = GameInput.GAMEINPUT_EXCEPTION.ITEM_KEY_NOT_FOUND Then
        Call Err.Raise(GameInput.GAMEINPUT_EXCEPTION.ITEM_KEY_NOT_FOUND, , "No se encontro ningun perfil con esa clave.")
     Else
        Call Err.Raise(Err.Number, , Err.Description)
    End If
End Sub

' Permite cambiar el nombre de usuario de un perfil existente siempre y cuando no coincida con otro:
Public Sub Rename(UserName As String, NewUserName As String)
    If LCase(UserName) = LCase(NewUserName) Then _
        Call Err.Raise(GameInput.GAMEINPUT_EXCEPTION.ITEM_KEY_DUPLICATED, , "Se esta intentando renombrar un perfil con la misma clave: '" & LCase(NewUserName) & "'")
    
    Dim prf As New GameInput.Profile
    
    On Error Resume Next
    
    Set prf = varCol(LCase(UserName))
    If Err.Number = 5 Then _
        Call Err.Raise(GameInput.GAMEINPUT_EXCEPTION.ITEM_KEY_NOT_FOUND, , "No se encontro ningun perfil con esa clave.")
    
    prf.UserName = LCase(NewUserName)
    
    Call varCol.Add(prf, prf.UserName)
    If Err.Number = GameInput.GAMEINPUT_EXCEPTION.ITEM_KEY_DUPLICATED Then _
        Call Err.Raise(GameInput.GAMEINPUT_EXCEPTION.ITEM_KEY_DUPLICATED, , "Ya existe un perfil con esa clave: '" & LCase(NewUserName) & "'")
    
    On Error GoTo 0
    
    Call varCol.Remove(LCase(UserName))
    Set prf = Nothing
End Sub

Public Property Get Item(UserName As String) As GameInput.Profile
Attribute Item.VB_UserMemId = 0
    On Error GoTo ErrOut
    Set Item = varCol(LCase(UserName))
ErrOut:
    Call Err.Raise(GameInput.GAMEINPUT_EXCEPTION.ITEM_KEY_NOT_FOUND, , "No se encontro ningun perfil con esa clave.")
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'esta propiedad permite enumerar
    'esta colección con la sintaxis For...Each
    Set NewEnum = varCol.[_NewEnum]
End Property

' Devuelve el numero de perfiles disponibles:
Public Property Get Count() As Long
    Count = varCol.Count
End Property

' Elimina todos los perfiles de la lista:
Public Sub Clear()
    Set varCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set varCol = Nothing
End Sub
