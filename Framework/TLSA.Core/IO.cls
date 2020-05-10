VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum FileAttribute
    Archive = File_Attribute.Archive
    Compressed = File_Attribute.Compressed
    Directory = File_Attribute.Directory
    Hidden = File_Attribute.Hidden
    Normal = File_Attribute.Normal
    ReadOnly = File_Attribute.Read_Only
    System = File_Attribute.System
    Temporary = File_Attribute.Temporary
End Enum

Public Enum SortMode
    NotSorted = -1
    Decrement = Sort_Mode.Decrement
    Increment = Sort_Mode.Increment
End Enum

Private varDesktopFolder As String, varMyDocsFolder As String, varTempFolder As String

Private Sub Class_Initialize()
    ' Obtenemos las rutas del Escritorio, del directorio Mis Documentos del usuario y del directorio temporal:
    varDesktopFolder = modSystem.GetSpecialfolder(CSIDL_DESKTOP) & "\"
    varMyDocsFolder = modSystem.GetSpecialfolder(CSIDL_PERSONAL) & "\"
    varTempFolder = sys.SYS_GetPath(TEMP_DIR) & "\"
End Sub

' Devuelve la ruta al Escritorio del usuario: (en Windows 98 y Me devuelve el Escritorio del sistema)
Public Property Get DesktopFolder() As String
    DesktopFolder = varDesktopFolder
End Property

' Devuelve la ruta al directorio Mis Documentos del usuario: (en Windows 98 y Me devuelve el directorio Mis Documentos del sistema)
Public Property Get MyDocumentsFolder() As String
    MyDocumentsFolder = varMyDocsFolder
End Property

' Devuelve la ruta del directorio temporal del usuario: (en Windows 98 y Me devuelve el directorio temporal del sistema)
Public Property Get TemporalFolder() As String
    TemporalFolder = varTempFolder
End Property

' Genera un nombre y ruta de un archivo temporal:
Public Function CreateTemporalFilename(Prefix As String) As String
    Dim sTemp As String
    
    sTemp = VBA.String(260, 0)
    Call modSystem.GetTempFileName(Me.TemporalFolder, Prefix, 0, sTemp)
    sTemp = VBA.Left(sTemp, VBA.InStr(1, sTemp, VBA.Chr(0)) - 1)
    Call modSystem.SetFileAttributes(sTemp, FILE_ATTRIBUTE_TEMPORARY)
    
    CreateTemporalFilename = sTemp
End Function

' Indica si un archivo existe:
Public Function FileExists(Filename As String) As Boolean
    FileExists = sys.FILE_Exists(Filename)
End Function

Public Function PathExists(Path As String) As Boolean
    PathExists = sys.DIR_Find(Path)
End Function

Public Function IsDirectoryEmpty(Path As String) As Boolean
    IsDirectoryEmpty = sys.DIR_IsEmpty(Path)
End Function

' Devuelve un listado con los archivos de un directorio (no incluye subdirectorios):
Public Function GetFiles(Path As String, Attrib As FileAttribute, Sort As Core.SortMode) As String()
    Dim list() As String, ret As SYS_ErrorCodes
    
    ret = sys.FILE_List(list, Path, CLng(Attrib), Sort <> NotSorted, CLng(Sort)) ' IIf(Sort = SortMode.Decrement, Sort_Mode.Decrement, Sort_Mode.Increment))
    Select Case ret
        Case SYS_ErrorCodes.SYS_EMPTYLIST
            Call Err.Raise(CORE_EXCEPTIONS.EMPTY_LIST, , "El directorio no contiene archivos que coincidan con los parametros de busqueda.")
        Case SYS_ErrorCodes.SYS_UNKNOWNERROR
            Call Err.Raise(CORE_EXCEPTIONS.UNKNOWN_EXCEPTION, , "Error desconocido al intentar crear el listado de archivos.")
        Case Else
            GetFiles = list
    End Select
End Function

' Devuelve el nombre de un archivo sin la ruta de acceso:
Public Function GetFileNameWithoutPath(Filename As String) As String
     GetFileNameWithoutPath = VBA.Right(Filename, VBA.Len(Filename) - VBA.InStrRev(Filename, "\"))
End Function

' Devuelve la ruta de un archivo sin el nombre del archivo:
Public Function GetPathWithoutFilename(Filename As String) As String
     GetPathWithoutFilename = VBA.Left(Filename, VBA.InStrRev(Filename, "\"))
End Function

' Abre un archivo INI y crea una instancia del objeto:
Public Function OpenINIFile(Filename As String) As INIFile
    Dim obj As New INIFile
    Call obj.OpenFile(Filename)
    Set OpenINIFile = obj
    Set obj = Nothing
End Function

' Crea un nuevo archivo INI y crea una instancia del objeto:
Public Function CreateINIFile(Filename As String) As INIFile
    Dim obj As New INIFile
    Call obj.Create(Filename)
    Set CreateINIFile = obj
    Set obj = Nothing
End Function

' Abre un archivo de texto y vuelca su contenido en un array. Devuelve tambien el numero de lineas leidas:
Public Function ReadTextFile(Filename As String, Lines() As String) As Long
    Dim buffer() As String, count As Integer
    Dim file As Integer: file = FreeFile()
    Dim line As String
    
    ' Cargamos el archivo en memoria:
    Open Filename For Input As file
        Do While Not EOF(file)
            Line Input #file, line
            ReDim Preserve buffer(count)
            buffer(count) = line
            count = count + 1
        Loop
    Close file
    
    Lines = buffer
    ReadTextFile = count
End Function
