VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "INIFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private varFilename As String

' Lista las secciones de un archivo INI:
Private Declare Function GetPrivateProfileSectionNames Lib "kernel32.dll" Alias "GetPrivateProfileSectionNamesA" (ByVal lpszReturnBuffer As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public Property Get Filename() As String
    Filename = varFilename
End Property

Friend Sub OpenFile(Filename As String)
    If sys.FILE_Exists(Filename) Then
        varFilename = Filename
    Else
        Call Err.Raise(Core.CORE_EXCEPTIONS.FILE_NOT_FOUND, , "El archivo INI no existe [" & Filename & "]")
    End If
End Sub

Friend Sub Create(Filename As String)
    If sys.FILE_Exists(Filename) Then
        Call Err.Raise(Core.CORE_EXCEPTIONS.FILE_NOT_FOUND, , "El archivo INI ya existe [" & Filename & "]")
    Else
        varFilename = Filename
    End If
End Sub

' Lee un valor del archivo:
Public Function GetValue(Section As String, Key As String, Optional DefaultValue As String) As String
    If varFilename <> "" Then
        GetValue = sys.INI_Read(varFilename, Section, Key, DefaultValue)
    Else
        Call Err.Raise(Core.CORE_EXCEPTIONS.FILE_NOT_LOAD, , "No se cargo ningun archivo INI.")
    End If
End Function

' Escribe un valor en el archivo:
Public Sub SetValue(Section As String, Key As String, Value As String)
    If varFilename <> "" Then
        Call sys.INI_Write(varFilename, Section, Key, Value)
    Else
        Call Err.Raise(Core.CORE_EXCEPTIONS.FILE_NOT_LOAD, , "No se cargo ningun archivo INI.")
    End If
End Sub

' Devuelve una lista con las secciones que componen el archivo INI:
Public Function GetSectionNames() As String()
    If varFilename <> "" Then
        Dim szBuf As String, Length As Integer
        szBuf = String$(255, 0)
        Length = GetPrivateProfileSectionNames(szBuf, 255, vbNullChar)
        szBuf = Left$(szBuf, Length)
        GetSectionNames = Split(szBuf, vbNullChar)
    Else
        Call Err.Raise(Core.CORE_EXCEPTIONS.FILE_NOT_LOAD, , "No se cargo ningun archivo INI.")
    End If
End Function
