Attribute VB_Name = "modSystem"
Option Explicit

Public sys As New dx_System_Class

Public lastTimerValueForRandomKey As Long ' Almacena el ultimo valor del cronometro normal en caso de usarse para generar la clave unica.

Public Enum WIN32_SPECIALFOLDERS
    CSIDL_DESKTOP = &H0
    CSIDL_PROGRAMS = &H2
    CSIDL_CONTROLS = &H3
    CSIDL_PRINTERS = &H4
    CSIDL_PERSONAL = &H5
    CSIDL_FAVORITES = &H6
    CSIDL_STARTUP = &H7
    CSIDL_RECENT = &H8
    CSIDL_SENDTO = &H9
    CSIDL_BITBUCKET = &HA
    CSIDL_STARTMENU = &HB
    CSIDL_DESKTOPDIRECTORY = &H10
    CSIDL_DRIVES = &H11
    CSIDL_NETWORK = &H12
    CSIDL_NETHOOD = &H13
    CSIDL_FONTS = &H14
    CSIDL_TEMPLATES = &H15
End Enum

Private Const MAX_PATH = 260

Private Type SHITEMID
    cb As Long
    abID As Byte
End Type
Private Type ITEMIDLIST
    mkid As SHITEMID
End Type

Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long

Public Const FILE_ATTRIBUTE_TEMPORARY = &H100

Public Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Public Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long

' Devuelve la ruta especial indicada:
Public Function GetSpecialfolder(CSIDL As WIN32_SPECIALFOLDERS) As String
    Dim r As Long
    Dim IDL As ITEMIDLIST
    'Get the special folder
    r = SHGetSpecialFolderLocation(100, CSIDL, IDL)
    If r = 0 Then
        'Create a buffer
        Dim Path As String
        Path = Space(512)
        'Get the path from the IDList
        r = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal Path)
        'Remove the unnecessary chr(0)'s
        GetSpecialfolder = Left(Path, InStr(Path, Chr(0)) - 1)
        Exit Function
    End If
    GetSpecialfolder = ""
End Function


' Convierte la estructura POINT a Vertex de dx_lib32:
Public Function POINT2Vertex(pt As Core.POINT) As dxlib32_221.Vertex
    Dim v As dxlib32_221.Vertex
    
    v.X = pt.X
    v.Y = pt.Y
    v.Z = pt.Z
    v.Color = pt.Color
    
    POINT2Vertex = v
End Function

' Convierte la estructura Vertex de dx_lib32 a POINT:
Public Function Vertex2POINT(v As dxlib32_221.Vertex) As Core.POINT
    Dim pt As Core.POINT
    
    pt.X = v.X
    pt.Y = v.Y
    pt.Z = v.Z
    pt.Color = v.Color
    
    Vertex2POINT = pt
End Function

' Convierte la estructura RECTANGLE a GFX_RECT de dx_lib32:
Public Function RECTANGLE2GFX_RECT(r As Core.RECTANGLE) As dxlib32_221.GFX_Rect
    Dim gR As dxlib32_221.GFX_Rect
    
    gR.X = r.X
    gR.Y = r.Y
    gR.Width = r.Width
    gR.Height = r.Height
    
    RECTANGLE2GFX_RECT = gR
End Function
