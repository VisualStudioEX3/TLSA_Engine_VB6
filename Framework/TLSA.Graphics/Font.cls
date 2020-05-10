VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Font"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const FONT_DEF_KEY As String = "TLSA_SPRITE_FONT"

Private id As Long
Private fontName As String
Private fontFilename As String
Private fontSize As Long
Private fontBold As Boolean
Private fontItalic As Boolean
Private fontUnderline As Boolean
Private fontStrikethrough As Boolean
Private fromFile As Boolean

Private varKey As String

Friend Property Get Key() As String
    Key = varKey
End Property

Friend Property Let Key(value As String)
    varKey = value
End Property


Public Property Get LoadFromFile() As Boolean
     LoadFromFile = fromFile
End Property

Friend Property Let LoadFromFile(value As Boolean)
    fromFile = value
End Property


' Nombre de la fuente:
Public Property Get Name() As String
    Name = fontName
End Property

' Nombre y ruta del archivo si se cargo desde archivo:
Friend Property Let Filename(value As String)
    fontFilename = value
End Property

Public Property Get Filename() As String
    Filename = fontFilename
End Property

' Tamaño de la fuente establecido:
Public Property Get SIZE() As Long
    SIZE = fontSize
End Property

' Indica si el estilo negrita esta establecido:
Public Property Get Bold() As Boolean
    Bold = fontBold
End Property

' Indica si el estilo cursiva esta establecido:
Public Property Get Italic() As Boolean
    Italic = fontItalic
End Property

' Indica si el estilo subrayado esta establecido:
Public Property Get Underline() As Boolean
    Underline = fontUnderline
End Property

' Indica si el estilo tachado esta establecido:
Public Property Get Strikethrough() As Boolean
    Strikethrough = fontStrikethrough
End Property

' Carga una fuente de Windows:
Friend Sub LoadFont(Font As String, SIZE As Long, Bold As Boolean, Italic As Boolean, Underline As Boolean, Strikethrough As Boolean)
    If id > -1 And Not gfx Is Nothing Then Call gfx.FONT_UnloadSystemFont(id)
    
    id = gfx.FONT_LoadSystemFont(Font, SIZE, Bold, Italic, Underline, Strikethrough)
    
    If id = GFX_ErrorCodes.GFX_FILENOTLOAD Then
        Call Err.Raise(GRAPHICS_EXCEPTION.RESOURCE_NOT_LOAD, , "Error al cargar la fuente de texto.")
    Else
        fromFile = False
        fontFilename = ""
        fontName = Name
        fontSize = SIZE
        fontBold = Bold
        fontItalic = Italic
        fontUnderline = Underline
        fontStrikethrough = Strikethrough
    End If
End Sub

' Carga una fuente desde archivo:
Friend Sub LoadFontFromFile(Filename As String, SIZE As Long, Bold As Boolean, Italic As Boolean, Underline As Boolean, Strikethrough As Boolean)
    If id > -1 And Not gfx Is Nothing Then Call gfx.FONT_UnloadSystemFont(id)
    
    id = gfx.FONT_LoadSystemFontFromFile(Filename, fontName, SIZE, Bold, Italic, Underline, Strikethrough)
    
    If id = GFX_ErrorCodes.GFX_FILENOTLOAD Then
        Call Err.Raise(GRAPHICS_EXCEPTION.RESOURCE_NOT_LOAD, , "Error al cargar la fuente de texto.")
    Else
        fromFile = True
        fontFilename = Filename
        fontSize = SIZE
        fontBold = Bold
        fontItalic = Italic
        fontUnderline = Underline
        fontStrikethrough = Strikethrough
    End If
End Sub

' Importa un archivo de definiciones de fuentes:
Public Sub Import(Filename As String)
    Dim ini As Core.INIFile
    
    'If sys.FILE_Exists(Filename) Then
    If Core.IO.FileExists(Filename) Then
        Set ini = Core.IO.OpenINIFile(Filename)
        
        ' Si se trata de una fuente del sistema:
        If Not CBool(ini.GetValue(FONT_DEF_KEY, "IsFromFile")) Then
            Call Me.LoadFont(ini.GetValue(FONT_DEF_KEY, "Name"), _
                             CInt(ini.GetValue(FONT_DEF_KEY, "Size")), _
                             CBool(ini.GetValue(FONT_DEF_KEY, "Bold")), _
                             CBool(ini.GetValue(FONT_DEF_KEY, "Italic")), _
                             CBool(ini.GetValue(FONT_DEF_KEY, "Underline")), _
                             CBool(ini.GetValue(FONT_DEF_KEY, "Strikethrough")))
        
        ' En caso contrario se trata de una fuente externa:
        Else
            ' Generamos el nombre del archivo de la fuente tomando su nombre original pero asignandole la ruta del archivo de definiciones:
            ' *** El archivo de fuente debe estar con su nombre original en la misma ubicacion que el archivo de definicion ***
            Dim ttfFile As String: ttfFile = Core.IO.GetPathWithoutFilename(Filename) & ini.GetValue(FONT_DEF_KEY, "Filename")
            If Core.IO.FileExists(ttfFile) Then
                Call Me.LoadFontFromFile(ttfFile, _
                                 CInt(ini.GetValue(FONT_DEF_KEY, "Size")), _
                                 CBool(ini.GetValue(FONT_DEF_KEY, "Bold")), _
                                 CBool(ini.GetValue(FONT_DEF_KEY, "Italic")), _
                                 CBool(ini.GetValue(FONT_DEF_KEY, "Underline")), _
                                 CBool(ini.GetValue(FONT_DEF_KEY, "Strikethrough")))
            Else
                Call Err.Raise(53, , "No se ha encontrado el archivo de la fuente de texto:" & vbNewLine & "'" & ttfFile & "'")
            End If
        End If
        
    Else
        Call Err.Raise(53, , "No se ha encontrado el archivo de especificacion de fuentes:" & vbNewLine & "'" & Filename & "'")
    End If
    
    Set ini = Nothing
End Sub

' Exporta la fuente como un archivo de definicion:
Public Sub Export(Filename As String)
    Dim ini As Core.INIFile
    Set ini = Core.IO.CreateINIFile(Filename)
    
    Call ini.SetValue(FONT_DEF_KEY, "IsFromFile", CStr(fromFile))
    Call ini.SetValue(FONT_DEF_KEY, "Filename", Core.IO.GetFileNameWithoutPath(fontFilename)) ' Se guarda sin ruta de acceso
    ' ya que la ruta de acceso debe ser la misma que la del archivo de definicion.
    Call ini.SetValue(FONT_DEF_KEY, "Size", CStr(fontSize))
    Call ini.SetValue(FONT_DEF_KEY, "Bold", CStr(fontBold))
    Call ini.SetValue(FONT_DEF_KEY, "Italic", CStr(fontItalic))
    Call ini.SetValue(FONT_DEF_KEY, "Underline", CStr(fontUnderline))
    Call ini.SetValue(FONT_DEF_KEY, "Strikethrough", CStr(fontStrikethrough))
    
    Set ini = Nothing
End Sub

Private Sub Class_Initialize()
    id = -1
End Sub

' Descarga la fuente de memoria:
Private Sub Class_Terminate()
    If Not gfx Is Nothing Then Call gfx.FONT_UnloadSystemFont(id)
End Sub

' Devuelve el codigo interno del recurso:
Friend Function GetID() As Long
    GetID = id
End Function
