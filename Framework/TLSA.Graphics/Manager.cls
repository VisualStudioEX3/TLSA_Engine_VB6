VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Manager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Lista de excepciones tipificadas de la libreria:
Public Enum GRAPHICS_EXCEPTION
    DEVICE_LOST = vbObjectError + 1000
    DEVICE_NOT_RESET = vbObjectError + 1001
    DISPLAY_MODE_EXCEPTION = vbObjectError + 1002
    RESOURCE_NOT_LOAD = vbObjectError + 1003
    RESOURCE_NOT_FOUND = vbObjectError + 1004
    RESOURCE_REFERENCE_EXCEPTION = vbObjectError + 1005
    RESOURCE_NOT_CREATED = vbObjectError + 1006
    ITEM_KEY_DUPLICATED = 457 ' Error(457) Esta clave ya está asociada a un elemento de esta colección
    ITEM_KEY_NOT_FOUND = vbObjectError + 1007
    ARGUMENT_EXCEPTION = 5
    SPRITE_DEFINITION_EXCEPTION = vbObjectError + 1008
    SURFACE_ALREADY_UNLOCKED = vbObjectError + 1009 ' La superficie ya esta desbloqueada.
    SURFACE_LOCKED = vbObjectError + 1010 ' La superficie esta bloqueda.
    UNKNOWN_EXCEPTION = vbObjectError + 9999
End Enum

Public Enum TextureFilter
    None = Blit_Filter.Filter_None
    Bilinear = Blit_Filter.Filter_Bilinear
    Trilinear = Blit_Filter.Filter_Trilinear
End Enum

Public Type DisplayMode
    Width As Long
    Height As Long
    Bpp As Long
End Type

Private varBackColor As Long
Private varMaxFrames As Integer

Public Property Get Monitor() As Graphics.RenderMonitor
    Set Monitor = modGraphics.varRenderMonitor
End Property

Public Property Get TextureFilter() As Graphics.TextureFilter
    TextureFilter = modGraphics.varFilter
End Property

Public Property Let TextureFilter(value As Graphics.TextureFilter)
    modGraphics.varFilter = value
End Property

Public Property Get Gamma() As Single
    Gamma = gfx.Gamma()
End Property

Public Property Let Gamma(value As Single)
    Call gfx.DEVICE_SetGamma(value)
End Property

' Devuelve si el modo de video esta establecido en ventana o pantalla completa:
Public Property Get Windowed() As Boolean
    Windowed = gfx.Windowed
End Property

' Establece el modo de video a ventana o pantalla completa:
Public Property Let Windowed(value As Boolean)
    If Not gfx.DEVICE_SetDisplayMode(gfx.screen.Width, gfx.screen.Height, gfx.screen.Bpp, value, True, gfx.VSync) Then
        If value Then
            Call Err.Raise(GRAPHICS_EXCEPTION.DISPLAY_MODE_EXCEPTION, , "No se pudo cambiar el modo a pantalla completa.")
        Else
            Call Err.Raise(GRAPHICS_EXCEPTION.DISPLAY_MODE_EXCEPTION, , "No se pudo cambiar el modo a ventana.")
        End If
    End If
End Property

' Devuelve el controlador de la ventana:
Public Property Get Hwnd() As Long
    Hwnd = gfx.Hwnd
End Property

' Devuelve la coleccion de texturas en memoria:
Public Property Get Textures() As Graphics.TextureList
    Set Textures = varTextures
End Property

' Devuelve la coleccion de superficies en memoria:
Public Property Get Surfaces() As Graphics.SurfaceList
    Set Surfaces = varSurfaces
End Property

' Devuelve la coleccion de fuentes en memoria:
Public Property Get Fonts() As Graphics.FontList
    Set Fonts = varFonts
End Property

' Inicializa el modulo grafico:
Public Sub Initialize(Hwnd As Long, Width As Long, Height As Long, Bpp As Long, Windowed As Boolean, VerticalSync As Boolean)
    If gfx.Init(Hwnd, Width, Height, Bpp, Windowed, True, VerticalSync) Then
        ' Definimos el rectangulo del area de la pantalla:
        varScreenRect = Core.Generics.CreateRECTANGLE(0, 0, modGraphics.gfx.screen.Width, modGraphics.gfx.screen.Height)
        
        ' Cargamos la fuente "Lucida Console" como fuente por defecto:
        Call Me.Fonts.LoadFont("Lucida Console", "SYSTEM", 8, False, False, False, False)
    Else
        Call Err.Raise(GRAPHICS_EXCEPTION.DISPLAY_MODE_EXCEPTION, , "No se pudo aplicar el modo de video ( " & Width & "x" & Height & "x" & Bpp & ").")
    End If
End Sub

' Cambia el modo de video:
Public Sub SetDisplayMode(Width As Long, Height As Long, Bpp As Long, VerticalSync As Boolean)
    If gfx.DEVICE_SetDisplayMode(Width, Height, Bpp, gfx.Windowed, True, VerticalSync) Then
        ' Definimos el rectangulo del area de la pantalla para las operaciones de descarte:
        modGraphics.varScreenRect = Core.Generics.CreateRECTANGLE(0, 0, modGraphics.gfx.screen.Width, modGraphics.gfx.screen.Height)
    Else
        Call Err.Raise(GRAPHICS_EXCEPTION.DISPLAY_MODE_EXCEPTION, , "No se pudo aplicar el modo de video ( " & Width & "x" & Height & "x" & Bpp & ").")
    End If
End Sub

Public Property Get CurrentDisplayMode() As Graphics.DisplayMode
    Dim mode As Graphics.DisplayMode
    
    mode.Width = gfx.screen.Width
    mode.Height = gfx.screen.Height
    mode.Bpp = gfx.screen.Bpp
    
    CurrentDisplayMode = mode
End Property

Public Property Get FPS() As Long
    FPS = gfx.FPS
End Property

Public Property Get BackColor() As Long
    BackColor = varBackColor
End Property

Public Property Let BackColor(value As Long)
    varBackColor = value
End Property

Public Property Get MaxFrames() As Integer
    MaxFrames = varMaxFrames
End Property

Public Property Let MaxFrames(value As Integer)
    varMaxFrames = value
End Property

Public Property Get OffSet() As Core.Point
    OffSet = varOffset
End Property

Public Property Let OffSet(value As Core.Point)
    varLastOffSet = varOffset
    varOffset = value
End Property

Public Sub SetOffSet(X As Long, Y As Long)
    varLastOffSet = varOffset
    varOffset.X = X: varOffset.Y = Y
End Sub

Public Property Get Primitives() As Graphics.Primitives
    Set Primitives = varPrimitives
End Property

Public Property Get Helper() As Graphics.Helper
    Set Helper = varHelper
End Property

Public Property Get VerticalSync() As Boolean
    VerticalSync = gfx.VSync
End Property

' Renderiza la escena:
Public Sub Render()
    
    
    Select Case gfx.Frame(varBackColor, varMaxFrames)
        Case GFX_ErrorCodes.GFX_DEVICELOST
            Call Err.Raise(GRAPHICS_EXCEPTION.DEVICE_LOST, , "Se perdio el dispositivo del contexto grafico.")
            
        Case GFX_ErrorCodes.GFX_DEVICENOTRESET
            Call Err.Raise(GRAPHICS_EXCEPTION.DEVICE_NOT_RESET, , "No se pudo reiniciar el dispositivo del contexto grafico.")
            
        Case GFX_ErrorCodes.GFX_UNKNOWNERROR
            Call Err.Raise(GRAPHICS_EXCEPTION.UNKNOWN_EXCEPTION, , "Error desconocido en el modulo grafico.")
            
    End Select
    
    Call varRenderMonitor.Reset
End Sub

Private Sub Class_Initialize()
    modGraphics.varFilter = None
    Set varRenderMonitor = New Graphics.RenderMonitor
End Sub

Private Sub Class_Terminate()
    Set varTextures = Nothing
    Set varSurfaces = Nothing
    Set varFonts = Nothing
    
    Call gfx.Terminate
    
    Set gfx = Nothing
End Sub
