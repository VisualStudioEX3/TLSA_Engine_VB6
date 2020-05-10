VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Surface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Type SURFACEINFO
    Lenght As Long
    Surface As SIZE
End Type

Private id As Long, pxId As Long
Private info As GFX_Info, varInfo As SURFACEINFO
Private varFilename As String

Private varKey As String

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

' *** Usada desde Clone() ***
Friend Sub SetPixelID(value As Long)
    pxId = value
End Sub

Public Property Get Filename() As String
    Filename = varFilename
End Property

' *** Usada desde Clone() ***
Friend Property Let Filename(value As String)
    varFilename = value
End Property

' Devuelve informacion sobre la textura:
Public Property Get Information() As SURFACEINFO
    Information = varInfo
End Property

' *** Usada desde Clone() ***
Friend Property Let Information(value As SURFACEINFO)
    varInfo = value
End Property

' Crea una superficie vacia:
Friend Sub Create(Width As Long, Height As Long)
    id = gfx.SURF_Create(Width, Height)
    If id = GFX_ErrorCodes.GFX_NOTCREATED Then
        Call Err.Raise(GRAPHICS_EXCEPTION.RESOURCE_NOT_CREATED, , "Error al crear la superficie.")
    Else
        varFilename = ""
        Call gfx.SURF_GetInfo(id, info)
        varInfo.Lenght = info.SIZE
        varInfo.Surface.Width = info.Width
        varInfo.Surface.Height = info.Height
        
        pxId = gfx.SURF_GetPixelData(id)
    End If
End Sub

' Crea una superficie con una copia del backbuffer:
Friend Sub CreateScreenCapture()
    id = gfx.SURF_ScreenCapture()
    If id = GFX_ErrorCodes.GFX_NOTCREATED Then
        Call Err.Raise(GRAPHICS_EXCEPTION.RESOURCE_NOT_CREATED, , "Error al crear la superficie.")
    Else
        varFilename = ""
        Call gfx.SURF_GetInfo(id, info)
        varInfo.Lenght = info.SIZE
        varInfo.Surface.Width = info.Width
        varInfo.Surface.Height = info.Height
        
        pxId = gfx.SURF_GetPixelData(id)
    End If
End Sub


' Carga una superficie desde disco:
Friend Sub LoadSurface(Filename As String)
    id = gfx.SURF_Load(Filename, False)
    If id = GFX_ErrorCodes.GFX_FILENOTLOAD Then
        Call Err.Raise(GRAPHICS_EXCEPTION.RESOURCE_NOT_LOAD, , "Error al cargar la superficie desde disco:" & vbNewLine & "'" & Filename & "'")
    Else
        varFilename = Filename
        Call gfx.SURF_GetInfo(id, info)
        varInfo.Lenght = info.SIZE
        varInfo.Surface.Width = info.Width
        varInfo.Surface.Height = info.Height
        
        pxId = gfx.SURF_GetPixelData(id)
    End If
End Sub

' Carga una superficie desde memoria:
Friend Sub LoadSurfaceFromMemory(buffer() As Byte)
    id = gfx.SURF_LoadFromMemory(buffer(), False)
    If id = GFX_ErrorCodes.GFX_FILENOTLOAD Then
        Call Err.Raise(GRAPHICS_EXCEPTION.RESOURCE_NOT_LOAD, , "Error al cargar la superficie desde memoria.")
    Else
        varFilename = Filename
        Call gfx.SURF_GetInfo(id, info)
        varInfo.Lenght = info.SIZE
        varInfo.Surface.Width = info.Width
        varInfo.Surface.Height = info.Height
        
        pxId = gfx.SURF_GetPixelData(id)
    End If
End Sub

' Copia una textura o parte de ella en la superficie:
Friend Sub CopyFromTexture(Texture As Graphics.Texture, Region As Core.RECTANGLE)
    Dim destRect As dxlib32_221.GFX_Rect: destRect.Width = varInfo.Surface.Width: destRect.Height = varInfo.Surface.Height
    Call gfx.MAP_AdvCopyRectsToSurf(Texture.GetID(), id, modGraphics.RECTANGLE2GFX_RECT(Region), destRect)
End Sub

' Crea una copia de la clase:
Public Function Clone() As Graphics.Surface
    Dim surf As New Graphics.Surface
    
    Call surf.SetID(id)
    Call surf.SetPixelID(pxId)
    surf.Information = varInfo
    surf.Filename = varFilename
    
    Set Clone = surf
    Set surf = Nothing
End Function

' Descarga la superficie de memoria:
Private Sub Class_Terminate()
    If Not gfx Is Nothing Then Call gfx.SURF_Unload(id)
End Sub

' Devuelve el codigo interno del recurso:
Friend Function GetID() As Long
    GetID = id
End Function

' Devuelve el codigo interno de la memoria de pixeles de la superficie:
Friend Function GetPixelID() As Long
    GetPixelID = pxId
End Function

' Desbloquea el buffer de pixeles y habilita la lectura y escritura de pixeles mediante GetPixel() y SetPixel():
Public Sub UnlockPixels()
    If pxId > -1 Then
        Call Err.Raise(Graphics.GRAPHICS_EXCEPTION.SURFACE_ALREADY_UNLOCKED, , "La superficie ya esta desbloqueada para lectura y escritura.")
    Else
        pxId = gfx.SURF_GetPixelData(id)
    End If
End Sub

' Aplica los cambios realizados, bloquea el buffer de pixeles y deshabilita la lectura y escritura de pixeles mediante GetPixel() y SetPixel():
Public Sub LockPixels()
    Call gfx.SURF_SetPixelData(pxId)
    pxId = -1
End Sub


' Devuelve el valor del pixel indicado:
Public Function GetPixel(X As Long, Y As Long) As Long
    If pxId > -1 Then
        GetPixel = gfx.SURF_GetPixel(pxId, X, Y)
    Else
        Call Err.Raise(Graphics.GRAPHICS_EXCEPTION.SURFACE_LOCKED, , "La superficie esta bloqueada para lectura y escritura.")
    End If
End Function

Public Function GetPixel2(Pixel As Core.Point) As Long
    If pxId > -1 Then
        GetPixel2 = Me.GetPixel(Pixel.X, Pixel.Y)
    Else
        Call Err.Raise(Graphics.GRAPHICS_EXCEPTION.SURFACE_LOCKED, , "La superficie esta bloqueada para lectura y escritura.")
    End If
End Function

' Establece el color del pixel indicado:
Public Sub SetPixel(X As Long, Y As Long, Color As Long)
    If pxId > -1 Then
        Call gfx.SURF_PutPixel(pxId, X, Y, Color)
    Else
        Call Err.Raise(Graphics.GRAPHICS_EXCEPTION.SURFACE_LOCKED, , "La superficie esta bloqueada para lectura y escritura.")
    End If
End Sub

Public Sub SetPixel2(Pixel As Core.Point)
    If pxId > -1 Then
        Call SetPixel(Pixel.X, Pixel.Y, Pixel.Color)
    Else
        Call Err.Raise(Graphics.GRAPHICS_EXCEPTION.SURFACE_LOCKED, , "La superficie esta bloqueada para lectura y escritura.")
    End If
End Sub

' Devuelve todo el buffer de pixeles en un array unidimensional con los valores de color:
Public Function GetPixelArray() As Long()
    Dim buffer() As Long
    Call gfx.SURF_ReadPixelArray(pxId, buffer)
    GetPixelArray = buffer
End Function

' Escribe todo el buffer de pixeles usando un array unidimensional, generalmente el que devuelve GetPixelArray(), con los valores de cada pixel:
Public Sub SetPixelArray(buffer() As Long)
    If Not gfx.SURF_WritePixelArray(pxId, buffer) Then
        Call Err.Raise(GRAPHICS_EXCEPTION.ARGUMENT_EXCEPTION, , "Error al escribir informacion en el buffer. Es posible que el array no tenga la longitud esperada.")
    End If
End Sub

' Guarda el contenido de la superficie en un archivo BMP:
Public Sub Save(Filename As String)
    Call gfx.SURF_Save(id, Filename)
End Sub
