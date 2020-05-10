VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Texture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Type TEXTUREINFO
    Lenght As Long
    Image As Core.SIZE
    Texture As Core.SIZE
End Type

Private id As Long
Private info As dxlib32_221.GFX_Info, varInfo As Graphics.TEXTUREINFO
Private isGray As Boolean
Private varFilename As String
Private varTiles As New Graphics.TileList

Private varRenderTarget As Boolean

Private varKey As String


' --- Control de referencias a la textura ---
Private varRefs As Long
Friend Property Get References() As Long
    References = varRefs
End Property

Friend Sub AddRef()
    varRefs = varRefs + 1
End Sub

Friend Sub RemoveRef()
    varRefs = varRefs - 1
End Sub

Friend Sub ClearRef()
    varRefs = 0
End Sub

Friend Function IsFreeRef() As Boolean
    IsFreeRef = (varRefs = 0)
End Function
' -------------------------------------------


Public Property Get Key() As String
    Key = varKey
End Property

Friend Property Let Key(value As String)
    varKey = value
End Property

' *** Usada desde Clone() ***
Friend Sub SetID(value As Long)
    id = value
End Sub

' Indica si la textura es un render target:
Public Property Get IsRenderTarget() As Boolean
    IsRenderTarget = varRenderTarget
End Property

Public Property Get Filename() As String
    Filename = varFilename
End Property

' Devuelve informacion sobre la textura:
Public Property Get Information() As Graphics.TEXTUREINFO
    Information = varInfo
End Property

' *** Usada desde Clone() ***
Friend Property Let Information(value As Graphics.TEXTUREINFO)
    varInfo = value
End Property

' Indica si la textura se transformo a escala de grises:
Public Property Get IsGrayScale() As Boolean
    IsGrayScale = isGray
End Property

' *** Usada desde Clone() ***
Friend Property Let IsGrayScale(value As Boolean)
    isGray = value
End Property

' Lista de tiles:
Public Property Get Tiles() As Graphics.TileList
    Set Tiles = varTiles
End Property

' Crea una textura vacia:
Friend Sub Create(Width As Long, Height As Long, GrayScale As Boolean, RenderTarget As Boolean)
    id = gfx.MAP_Create(Width, Height, RenderTarget, GrayScale)
    If id = GFX_ErrorCodes.GFX_NOTCREATED Then
        Call Err.Raise(GRAPHICS_EXCEPTION.RESOURCE_NOT_CREATED, , "Error al crear la textura.")
    Else
        varRenderTarget = RenderTarget
        isGray = GrayScale
        Call gfx.MAP_GetInfo(id, info)
        varInfo.Lenght = info.SIZE
        varInfo.Texture.Width = info.Width
        varInfo.Texture.Height = info.Height
        varInfo.Image.Width = info.Image_Width
        varInfo.Image.Height = info.Image_Height
        
        Call varTiles.Clear ' Limpiamos la lista.
            
        ' Creamos un tile por defecto que define toda la textura:
        Call varTiles.Create("Default", Me, 0, 0, varInfo.Image.Width, varInfo.Image.Height, 0, 0)
    End If
End Sub

' Carga una textura desde disco:
Friend Sub LoadTexture(Filename As String, GrayScale As Boolean)
    id = gfx.MAP_Load(Filename, 0, False, GrayScale)
    If id = GFX_ErrorCodes.GFX_FILENOTLOAD Then
        Call Err.Raise(GRAPHICS_EXCEPTION.RESOURCE_NOT_LOAD, , "Error al cargar la textura desde disco:" & vbNewLine & "'" & Filename & "'")
    Else
        varFilename = Filename
        isGray = GrayScale
        Call gfx.MAP_GetInfo(id, info)
        varInfo.Lenght = info.SIZE
        varInfo.Texture.Width = info.Width
        varInfo.Texture.Height = info.Height
        varInfo.Image.Width = info.Image_Width
        varInfo.Image.Height = info.Image_Height
        
        Call Me.ImportTiles
    End If
End Sub

' Carga una textura desde memoria:
Friend Sub LoadTextureFromMemory(buffer() As Byte, GrayScale As Boolean)
    id = gfx.MAP_LoadFromMemory(buffer(), 0, False, GrayScale)
    If id = GFX_ErrorCodes.GFX_FILENOTLOAD Then
        Call Err.Raise(GRAPHICS_EXCEPTION.RESOURCE_NOT_LOAD, , "Error al cargar la textura desde memoria.")
    Else
        varFilename = Filename
        isGray = GrayScale
        Call gfx.MAP_GetInfo(id, info)
        varInfo.Lenght = info.SIZE
        varInfo.Texture.Width = info.Width
        varInfo.Texture.Height = info.Height
        varInfo.Image.Width = info.Image_Width
        varInfo.Image.Height = info.Image_Height
        
        Call Me.ImportTiles
    End If
End Sub

' Crea una copia de la clase:
Public Function Clone() As Graphics.Texture
    Dim tex As New Graphics.Texture
    
    Call tex.SetID(id)
    tex.Information = varInfo
    tex.IsGrayScale = isGray
    Call tex.ImportTiles
    
    Set Clone = tex
    Set tex = Nothing
End Function

' Descarga la textura de memoria:
Private Sub Class_Terminate()
    If varRefs = 0 Then
        If Not gfx Is Nothing Then Call gfx.MAP_Unload(id)
        If Not varTiles Is Nothing Then
            Call varTiles.Clear
            Set varTiles = Nothing
        End If
    Else
        Call Err.Raise(Graphics.GRAPHICS_EXCEPTION.RESOURCE_REFERENCE_EXCEPTION, , _
                       "Existen referencias a la textura '" & Key & _
                       "'. Elimine cualquier instancia del objeto Sprite que haga referencia a esta textura.")
    End If
End Sub

' Devuelve el codigo interno del recurso:
Friend Function GetID() As Long
    GetID = id
End Function

' Exporta a un archivo los parametros de los tiles de la textura:
Public Sub ExportTiles()
    Dim fName As String
    Dim file As Integer
    
    Dim datTile As String, datControlPointMap As String, datFinal As String
    
    file = FreeFile
    ' Se utiliza el mismo nombre de archivo que la textura solo que modificando su extension a ".tex":
    fName = VBA.Left(varFilename, Len(varFilename) - 3) & "tex"
    
    Dim i As Long, keys() As String
    Open fName For Output As file
        Dim t As Tile
        For Each t In varTiles
            If Not t.Key = "Default" Then ' El Tile por defecto de la textura nunca se guarda.
                ' Informacion del tile:
                datTile = t.Key & ";" & t.Region.X & ";" & t.Region.Y & ";" & t.Region.Width & ";" & t.Region.Height & ";" & t.Center.X & ";" & t.Center.Y
                
                ' Informacion del mapa de puntos de control:
                If t.ControlPoints.Count > 1 Then
                    keys = t.ControlPoints.GetKeyList()
                    datControlPointMap = ""
                    
                    For i = 1 To UBound(keys)
                        datControlPointMap = datControlPointMap & keys(i) & ";" & t.ControlPoints(keys(i)).X & ";" & t.ControlPoints(keys(i)).Y & ":"
                    Next
                    ' Eliminamos el ultimo punto y coma:
                    datControlPointMap = VBA.Left(datControlPointMap, Len(datControlPointMap) - 1)
                    
                    datTile = datTile & "|" & datControlPointMap
                End If
                
                ' Grabamos en el archivo los parametros del tile separados por punto y coma:
                Print #file, datTile
            End If
        Next
    Close file
End Sub

' Importa desde archivo los parametros de los tiles de la textura si se definieron:
Public Sub ImportTiles()
    Dim fName As String
    Dim file As Integer
    Dim readLine As String
    
    Dim datTile() As String ' {Tile data|ControlPoint:ControlPoint:(...)} - Contenido del registro leido.
    
    Dim tileFields() As String ' Separaciones por ; - Listas de valores del tile.
    Dim ControlPoints() As String ' Separaciones por : - Lista de puntos de control.
    Dim pointFields() As String ' Separaciones ; - Lista de valores del punto de control.
    
    Dim t As Graphics.Tile
    
    file = FreeFile
    ' Buscamos el mismo nombre de archivo que la textura solo que modificando su extension a ".tex":
    fName = VBA.Left(varFilename, Len(varFilename) - 3) & "tex"
    
    ' Si el archivo existe se carga en memoria:
    If IO.FileExists(fName) Then
        Open fName For Input As file
            Call varTiles.Clear ' Limpiamos la lista.
            
            ' Creamos un tile por defecto que define toda la textura:
            Call varTiles.Create("Default", Me, 0, 0, varInfo.Image.Width, varInfo.Image.Height, 0, 0)
            
            Do While Not EOF(file)
                Line Input #file, readLine ' Leemos las linea.
                
                ' Obtenemos las dos partes de informacion del registro: el tile y su mapa de puntos de control si lo tuviera:
                datTile = Split(readLine, "|")
                
                tileFields = Split(datTile(0), ";") ' Obtenemos los valores del tile.
                               
                ' Creamos un tile con los parametros leidos:
                Set t = varTiles.Create(tileFields(0), Me, CLng(tileFields(1)), CLng(tileFields(2)), CLng(tileFields(3)), CLng(tileFields(4)), CLng(tileFields(5)), CLng(tileFields(6)))
                
                ' Informacion de depuracion:
                t.ControlPoints.TextureKey = varKey
                t.ControlPoints.TileKey = t.Key
                
                Call t.ControlPoints.Clear
                
                t.ControlPoints.OffSet = t.Center ' El offset se usara para corregir la posicion de los puntos de control respecto al centro del sprite.
                
                ' Si hay puntos de control se cargan:
                If UBound(datTile) > 0 Then
                    ControlPoints = Split(datTile(1), ":") ' Obtenemos los puntos de control si los hubiera.
                    
                    Dim i As Long
                    For i = 0 To UBound(ControlPoints)
                        pointFields = Split(ControlPoints(i), ";") ' Extraemos los valores del punto de control.
                        Call t.ControlPoints.Add(pointFields(0), CLng(pointFields(1)), CLng(pointFields(2)))
                    Next
                End If
            Loop
        
        Close file
    
    Else
        ' Creamos un tile por defecto que define toda la textura:
        Call varTiles.Create("Default", Me, 0, 0, varInfo.Image.Width, varInfo.Image.Height, 0, 0)
    End If
End Sub

' Crea una superficie con el contenido de la textura:
Public Function ToSurface() As Graphics.Surface
    Set ToSurface = modGraphics.varSurfaces.CreateFromTexture(Me, Core.Generics.CreateRECTANGLE(0, 0, varInfo.Texture.Width, varInfo.Texture.Height))
End Function
