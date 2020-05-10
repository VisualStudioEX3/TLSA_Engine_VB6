VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Sprite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Objeto para dibujar sprites:
Option Explicit

Public Enum Effects
    Default = Blit_Alpha.Blendop_Color
    Aditive = Blit_Alpha.Blendop_Aditive
    Sustrative = Blit_Alpha.Blendop_Sustrative
    Negative = Blit_Alpha.Blendop_Inverse
    XOR_Exclusion = Blit_Alpha.Blendop_XOR
    Greyscale = Blit_Alpha.Blendop_GreyScale
    Crystaline = Blit_Alpha.Blendop_Crystaline
End Enum

Public Enum Mirrors
    WithoutMirroring = Blit_Mirror.Mirror_None
    Horizontal = Blit_Mirror.Mirror_Horizontal
    Vertical = Blit_Mirror.Mirror_Vertical
    Both = Blit_Mirror.Mirror_both
End Enum

Public Enum Perspectives
    WithoutPerspective = 7
    Caballera_Height = Blit_Perspective.Caballera_Height
    Caballera_Height_Negative = Blit_Perspective.Caballera_Height_Negative
    Caballera_Width = Blit_Perspective.Caballera_Width
    Caballera_Width_Negative = Blit_Perspective.Caballera_Width_Negative
    Isometric_Base = Blit_Perspective.Isometric_Base
    Isometric_Height = Blit_Perspective.Isometric_Height
    Isometric_Height_Negative = Blit_Perspective.Isometric_Height_Negative
End Enum

Public Enum VertexPosition
    Left_Top
    Right_Top
    Left_Bottom
    Right_Bottom
End Enum

' Define un valor de luminosidad utilizando 32 bits de color y 32 bits del canal specular:
Public Type LIGHTVALUE
    Color As Long
    Specular As Long
End Type

' Define los valores de luminosidad de los 4 vertices de un objeto grafico:
Public Type VERTEXLIGHTS
    LeftTop As Graphics.LIGHTVALUE
    RightTop As Graphics.LIGHTVALUE
    LeftBottom As Graphics.LIGHTVALUE
    RightBottom As Graphics.LIGHTVALUE
End Type

' Vertices precalculados:
Dim lt As Vertex, rt As Vertex, lb As Vertex, rb As Vertex

Private varTex As Graphics.Texture
Private varLocation As Core.Point
Private varSize As Core.SIZE
Private varEffect As Graphics.Effects
Private varMirror As Graphics.Mirrors
Private varAngle As Single
Private varCenter As Boolean
Private varPerspective As Graphics.Perspectives
Private varPerspFlag As Long
Private varColor As Long
Private varOffset As Core.Point
Private varLight As Graphics.VERTEXLIGHTS
Private varLighting As Boolean
Private varRegion As Core.RECTANGLE, varSetRegion As Boolean
Private varVisible As Boolean
Private varAnimations As New Graphics.AnimationList     ' Lista de secuencias de animacion del sprite.
Private varCurrentAnim As Long                          ' Animacion actual.
Private varEnabledAnim As Boolean                       ' Indica si se ejecutan las animaciones del sprite.
Private varScale As Single                              ' Almacena el valor de la escala aplicada. Se resetea a 1 cuando se aplican dimensiones manualmente.

Private varCurrentTile As Graphics.Tile                 ' Tile seleccionado. Valdra Nothing cuando se aplique manualmente la region a pintar.

Private varFixed As Boolean                             ' No aplica el desplazamiento de las coordenadas OffSet.
Private varTmpOffSet As Core.Point                    ' Variable temporal donde se asigna el OffSet de la escena si el objeto no esta fijo.

'Private varTag As Variant                               ' Variable contenedor.
'
'' Permite almacenar informacion util para el objeto, se del tipo que sea:
'Public Property Get Tag() As Variant
'    Tag = varTag
'End Property
'
'Public Property Let Tag(value As Variant)
'    varTag = value
'End Property

Public Property Get Fixed() As Boolean
    Fixed = varFixed
End Property

Public Property Let Fixed(value As Boolean)
    varFixed = value
End Property

Public Property Get EnabledAnimation() As Boolean
    EnabledAnimation = varEnabledAnim
End Property

Public Property Let EnabledAnimation(value As Boolean)
    varEnabledAnim = value
End Property

' Define le tile actual del sprite:
Public Property Get CurrentTile() As Graphics.Tile
    Set CurrentTile = varCurrentTile
End Property

' Establece el tile actual del sprite de la lista de tiles de la textura asociada:
Public Sub SetCurrentTile(Key As String)
    Set varCurrentTile = varTex.Tiles(Key)
    If varCurrentTile.ControlPoints.Count > 0 Then Call varCurrentTile.ControlPoints.ResetControlPoints '
    
    Me.Region = varCurrentTile.Region
    Me.OffSet = varCurrentTile.Center
End Sub

Public Property Get Texture() As Graphics.Texture
    Set Texture = varTex
    'Call Me.SetCurrentTile("Default") ' ?
End Property

Public Sub SetTexture(tex As Graphics.Texture)
    If Not varTex Is Nothing Then Call varTex.RemoveRef
    
    Set varTex = tex
    
    Call tex.AddRef
    
    varSize.Width = varTex.Information.Image.Width
    varSize.Height = varTex.Information.Image.Height
    Call Me.SetCurrentTile("Default") ' Establecemos el tile por defecto.
    Call Me.SetEntireRegion
End Sub

Public Property Get Animations() As Graphics.AnimationList
    Set Animations = varAnimations
End Property

' Devuelve la referencia a la secuencia de animacion actual:
Public Property Get CurrentAnimation() As Graphics.Animation
    Set CurrentAnimation = varAnimations(varCurrentAnim)
End Property

' Establece el indice o clave de la secuencia de animacion a reproducir:
Public Sub SetCurrentAnimation(Key As Variant)
    If VarType(Key) = vbString Then
        varCurrentAnim = varAnimations.IndexOf(CStr(Key))
    Else
        varCurrentAnim = varAnimations.IndexOf(varAnimations(Key).Key)
    End If
End Sub

Public Property Get Location() As Core.Point
    Location = varLocation
End Property

Public Property Let Location(value As Core.Point)
    varLocation = value
    
    ' *** Se aplican los valores de transformacion en el momento de asignar el tile en el metodo Update() del sprite ***
    ' ControlPointMap
End Property

Public Sub SetLocation(X As Long, Y As Long, Z As Long)
    With varLocation
        .X = X
        .Y = Y
        .Z = Z
    End With
    
    ' *** Se aplican los valores de transformacion en el momento de asignar el tile en el metodo Update() del sprite ***
    ' ControlPointMap
End Sub

Public Property Get Color() As Long
    Color = varColor
End Property

Public Property Let Color(value As Long)
    varColor = value
End Property

Public Property Get SIZE() As Core.SIZE
    SIZE = varSize
End Property

Public Property Let SIZE(value As Core.SIZE)
    varSize = value
    varScale = 1
End Property

Public Sub SetSize(Width As Long, Height As Long)
    varSize.Width = Width
    varSize.Height = Height

' *** Se aplican los valores de transformacion en el momento de asignar el tile en el metodo Update() del sprite ***
'    ' Aplicamos el nuevo tamaño al mapa de puntos de control:
'    varCurrentTile.ControlPoints.Size = varSize
    
    varScale = 1
End Sub

Public Sub SetScale(value As Single)
    If varSetRegion Then
        varSize.Width = varRegion.Width * value
        varSize.Height = varRegion.Height * value
    Else
        varSize.Width = varTex.Information.Image.Width * value
        varSize.Height = varTex.Information.Image.Height * value
    End If

' *** Se aplican los valores de transformacion en el momento de asignar el tile en el metodo Update() del sprite ***
    ' Aplicamos el nuevo tamaño al mapa de puntos de control:
    'varCurrentTile.ControlPoints.SIZE = varSize
    
    varScale = value
End Sub

Public Property Get Angle() As Single
    Angle = varAngle
End Property

Public Property Let Angle(value As Single)
    varAngle = value
    
' *** Se aplican los valores de transformacion en el momento de asignar el tile en el metodo Update() del sprite ***
'    ' Aplicamos el angulo al mapa de puntos de control:
'    varCurrentTile.ControlPoints.Angle = value
End Property

' Coordenada centro u origen de dibujo del sprite. Permite definir la coordenada de dibujo del sprite:
Public Property Get OffSet() As Core.Point
    OffSet = varOffset
End Property

Public Property Let OffSet(value As Core.Point)
    varOffset = value
End Property

Public Sub SetOffSet(X As Long, Y As Long)
    varOffset.X = X
    varOffset.Y = Y
End Sub

Public Property Get Effect() As Graphics.Effects
    Effect = varEffect
End Property

Public Property Let Effect(value As Graphics.Effects)
    varEffect = value
End Property

Public Property Get Mirror() As Graphics.Mirrors
    Mirror = varMirror
End Property

Public Property Let Mirror(value As Graphics.Mirrors)
    varMirror = value
    
' *** Se aplican los valores de transformacion en el momento de asignar el tile en el metodo Update() del sprite ***
'    ' Aplicamos el espejado al mapa de puntos de control:
'    varCurrentTile.ControlPoints.Mirror = value
End Property

Public Property Get Perspective() As Graphics.Perspectives
    Perspective = varPerspective
End Property

Public Property Let Perspective(value As Graphics.Perspectives)
    varPerspective = value
End Property

Public Property Get PerspectiveCorrection() As Long
    PerspectiveCorrection = varPerspFlag
End Property

Public Property Let PerspectiveCorrection(value As Long)
    varPerspFlag = value
End Property

Public Property Get Centered() As Boolean
    Centered = varCenter
End Property

' *** No se aplica al mapa de puntos de control ***
Public Property Let Centered(value As Boolean)
    varCenter = value
End Property

Public Property Get Light() As Graphics.VERTEXLIGHTS
    Light = varLight
End Property

Public Property Let Light(value As Graphics.VERTEXLIGHTS)
    varLight = value
End Property

' Aplica un valor absoluto de iluminacion a los vertices del sprite:
Public Sub SetLight(LeftTop As Graphics.LIGHTVALUE, RightTop As Graphics.LIGHTVALUE, LeftBottom As Graphics.LIGHTVALUE, RightBottom As Graphics.LIGHTVALUE)
    If varLighting Then
        With varLight
            .LeftTop = LeftTop
            .RightTop = RightTop
            .LeftBottom = LeftBottom
            .RightBottom = RightBottom
        End With
    End If
End Sub

' Aplica los parametros de iluminacion de una fuente de luz:
Public Sub ApplyLight(LightSource As Graphics.Light)
    If LightSource.Enabled And varLighting Then
    '    With varLight
    '        .LeftTop = LeftTop
    '        .RightTop = RightTop
    '        .LeftBottom = LeftBottom
    '        .RightBottom = RightBottom
    '    End With
    End If
End Sub

' Precalcula la posicion de los vertices al dibujar el sprite:
Public Sub PrecalulateVertex()
    ' Precalculamos la posicion de los vertices:
    Call gfx.PRECAL_WriteSpriteTransformVertex
    
    ' Indicamos el centro de dibujo y rotacion del sprite:
    Call gfx.DEVICE_SetDrawCenter(varOffset.X, varOffset.Y)
    
    If varFixed Then
        If varPerspective = WithoutPerspective Then
            Call gfx.DRAW_MapEx(varTex.GetID(), varLocation.X, varLocation.Y, varLocation.Z, varSize.Width, varSize.Height, varAngle, varEffect, varColor, varMirror, modGraphics.varFilter, varCenter)
        Else
            Call gfx.DRAW_AdvMap(varTex.GetID(), varLocation.X, varLocation.Y, varLocation.Z, varSize.Width, varSize.Height, varEffect, varColor, varMirror, modGraphics.varFilter, varPerspective, varPerspFlag)
        End If
    Else
        If varPerspective = WithoutPerspective Then
            Call gfx.DRAW_MapEx(varTex.GetID(), varLocation.X + modGraphics.varOffset.X, varLocation.Y + modGraphics.varOffset.Y, varLocation.Z, varSize.Width, varSize.Height, varAngle, varEffect, varColor, varMirror, modGraphics.varFilter, varCenter)
        Else
            Call gfx.DRAW_AdvMap(varTex.GetID(), varLocation.X + modGraphics.varOffset.X, varLocation.Y + modGraphics.varOffset.Y, varLocation.Z, varSize.Width, varSize.Height, varEffect, varColor, varMirror, modGraphics.varFilter, varPerspective, varPerspFlag)
        End If
    End If
    
    ' Obtenemos los vertices precalculados:
    Call gfx.PRECAL_ReadSpriteTransformVertex(lt, rt, lb, rb)
End Sub

' Obtiene la posicion y color del vertice precalculado:
Public Function GetPrecalVertex(Vertex As Graphics.VertexPosition) As Core.Point
    Dim pt As Core.Point
    
    Select Case Vertex
        Case VertexPosition.Left_Top: pt = modGraphics.Vertex2POINT(lt)
        Case VertexPosition.Left_Bottom: pt = modGraphics.Vertex2POINT(lb)
        Case VertexPosition.Right_Top: pt = modGraphics.Vertex2POINT(rt)
        Case VertexPosition.Right_Bottom: pt = modGraphics.Vertex2POINT(rb)
    End Select
    
    GetPrecalVertex = pt
End Function

' Indica si se aplican parametros de iluminacion al sprite:
Public Property Get Lighting() As Boolean
    Lighting = varLighting
End Property

Public Property Let Lighting(value As Boolean)
    varLighting = value
End Property

Public Property Get Region() As Core.RECTANGLE
    Region = varRegion
End Property

Public Property Let Region(value As Core.RECTANGLE)
    varRegion = value
    varSize.Width = varRegion.Width
    varSize.Height = varRegion.Height
    varSetRegion = True
    Call Me.SetScale(varScale)
End Property

Public Property Get EntireRegion() As Core.RECTANGLE
    Dim r As RECTANGLE
    
    r.X = 0
    r.Y = 0
    r.Width = varTex.Information.Image.Width
    r.Height = varTex.Information.Image.Height
    
    varSize.Width = r.Width
    varSize.Height = r.Height
    
    EntireRegion = r
    
    varSetRegion = False
End Property

Public Sub SetRegion(X As Long, Y As Long, Width As Long, Height As Long)
    Dim r As Core.RECTANGLE
    
    r.X = X
    r.Y = Y
    r.Width = Width
    r.Height = Height
    
    varSize.Width = r.Width
    varSize.Height = r.Height
    
    varRegion = r
    
    varSetRegion = True
    
    Call Me.SetScale(varScale)
End Sub

Public Sub SetEntireRegion()
    varRegion = Me.EntireRegion
    varSize.Width = varRegion.Width
    varSize.Height = varRegion.Height
    varSetRegion = False
    Call Me.SetScale(varScale)
End Sub

Public Property Get Visible() As Boolean
    Visible = varVisible
End Property

Public Property Let Visible(value As Boolean)
    varVisible = value
End Property

' Actualiza las animaciones y demas datos dinamicos:
Public Sub Update()
    If Me.Animations.Count > 0 Then
        With Me.CurrentAnimation
            If varEnabledAnim Then
                If .Tiles.Count > 0 Then
                    .Update ' Actualizamos la animacion.
                    
                    With .CurrentTile
                        Me.Region = .Region ' Actualizamos el tile a dibujar.
                        Me.OffSet = .Center ' Actualizamos el centro de dibujo.
                    End With
                    
                    Set varCurrentTile = .CurrentTile ' Obtenemos la referencia al tile actual de la animacion al sprite.
                    
                End If
            End If
            
            ' Aplicamos las transformaciones del sprite al mapa de puntos de control:
            With varCurrentTile.ControlPoints
                If .Count > 0 Then
                    ' *** Aplicar en la clase ControlPointMap la verificacion de cambio de valores para evitar actualizaciones innecesarias ***
                    .Angle = varAngle
                    .SIZE = varSize
                    .Location = varLocation
                    .Mirror = varMirror
                    
                    Call .Update ' Actualizamos el estado del mapa de puntos de control.
                End If
            End With
        End With
    End If
End Sub

' Dibuja el sprite si esta dentro del area de la pantalla y aplica todos los atributos que mostrara el sprite:
Public Sub Draw()
    If varVisible Then
        If varTex Is Nothing Then
            Call Err.Raise(GRAPHICS_EXCEPTION.RESOURCE_REFERENCE_EXCEPTION, , "Se descargo la textura asociada o no se asocio textura al sprite.")
        Else
            ' Indicamos el centro de dibujo y rotacion del sprite:
            Call gfx.DEVICE_SetDrawCenter(varOffset.X, varOffset.Y)
            
            ' Si la iluminacion esta habilitada se aplican los parametros:
            If varLighting Then
                With varLight
                    Call gfx.DEVICE_SetVertexColor(.LeftTop.Color, .RightTop.Color, .LeftBottom.Color, .RightBottom.Color)
                    Call gfx.DEVICE_SetSpecularChannel(.LeftTop.Specular, .RightTop.Specular, .LeftBottom.Specular, .RightBottom.Specular)
                End With
            End If
            
            ' Seleccionamos la region de la textura a dibujar:
            Call gfx.MAP_SetRegion(varTex.GetID(), modGraphics.RECTANGLE2GFX_RECT(varRegion))
            
            ' Si el objeto no esta configurado como fijo se le aplica desplazamiento de posicion segun variable OffSet global:
            If Not varFixed Then
                varTmpOffSet = modGraphics.varOffset
            Else
                varTmpOffSet.X = 0: varTmpOffSet.Y = 0
            End If
            
            ' Comprobamos que el sprite este dentro del area de la pantalla:
            If InScreen() Then
                ' Dibujamos segun perspectiva aplicada:
                If varPerspective = WithoutPerspective Then
                    Call gfx.DRAW_MapEx(varTex.GetID(), varLocation.X + varTmpOffSet.X, varLocation.Y + varTmpOffSet.Y, varLocation.Z, varSize.Width, varSize.Height, varAngle, varEffect, varColor, varMirror, modGraphics.varFilter, varCenter)
                Else
                    Call gfx.DRAW_AdvMap(varTex.GetID(), varLocation.X + varTmpOffSet.X, varLocation.Y + varTmpOffSet.Y, varLocation.Z, varSize.Width, varSize.Height, varEffect, varColor, varMirror, modGraphics.varFilter, varPerspective, varPerspFlag)
                End If
                Call modGraphics.varRenderMonitor.Add("SPRITE")
            End If
            
            ' Limpiamos los valores de iluminacion para la siguiente pasada:
            With varLight
                With .LeftBottom: .Color = 0: .Specular = 0: End With
                With .LeftTop: .Color = 0: .Specular = 0: End With
                With .RightBottom: .Color = 0: .Specular = 0: End With
                With .RightTop: .Color = 0: .Specular = 0: End With
            End With
        End If
    End If
End Sub

' Indica si el sprite esta en el area de la pantalla:
Private Function InScreen() As Boolean
    Dim spriteRect As Core.RECTANGLE
    Dim Radius As Long
    
    ' Obtenemos el centro en el espacio de dibujo (no el centro interno del sprite):
    If varCenter Then ' Tomamos el centro del sprite como centro del rectangulo:
        spriteRect.X = (varLocation.X + IIf(varFixed, 0, varTmpOffSet.X)) \ 2
        spriteRect.Y = (varLocation.Y + IIf(varFixed, 0, varTmpOffSet.Y)) \ 2
    Else ' Agregamos el offset de su eje de rotacion definido:
        spriteRect.X = (varLocation.X + IIf(varFixed, 0, varTmpOffSet.X)) + varOffset.X
        spriteRect.Y = (varLocation.Y + IIf(varFixed, 0, varTmpOffSet.Y)) + varOffset.Y
    End If
    
    ' Ancho y alto del rectangulo del sprite:
    spriteRect.Width = varSize.Width
    spriteRect.Height = varSize.Height
    
    If varAngle = 0 Then ' Si no hay rotacion...
        ' Comprobamos si intersectan ambos rectangulos:
        InScreen = Core.Math.IntersectRect(modGraphics.varScreenRect, spriteRect)
    Else ' Si el sprite esta rotado...
        ' Definimos una circunferencia que encierre el area del sprite y calculamos su interseccion con el area de la pantalla:
        Radius = IIf(varSize.Width > varSize.Height, varSize.Width \ 2, varSize.Height \ 2)
        InScreen = Core.Math.IntersectCircleWithRectangle(Core.Generics.CreatePOINT(spriteRect.X, spriteRect.Y), Radius, modGraphics.varScreenRect)
    End If
End Function

' Devuelve el centro real del sprite. Cuando la propiedad Centered esta activa devuelve la posicion del sprite:
Public Function GetCenter() As Core.Point
    If Me.Centered Then
        GetCenter = Me.Location
    Else
        Dim C As Core.Point
        
        ' Precalculamos la posicion de los vertices:
        Call Me.PrecalulateVertex
                
        ' Obtenemos los vertices precalculados:
        Call gfx.PRECAL_ReadSpriteTransformVertex(lt, rt, lb, rb)
        
        ' Obtenemos el centro del sprite calculando la interseccion de sus 4 vertices:
        If Me.Angle = 0 Then
            C.X = lt.X + ((rt.X - lt.X) \ 2)
            C.Y = lt.Y + ((lb.Y - lt.Y) \ 2)
        
        ElseIf Me.Angle = 90 Then
            C.X = rt.X + ((rb.X - rt.X) \ 2)
            C.Y = rt.Y + ((lt.Y - rt.Y) \ 2)
        
        ElseIf Me.Angle = 180 Then
            C.X = rb.X + ((lb.X - rb.X) \ 2)
            C.Y = rb.Y + ((rt.Y - rb.Y) \ 2)
        
        ElseIf Me.Angle = 270 Then
            C.X = lb.X + ((lt.X - lb.X) \ 2)
            C.Y = lb.Y + ((rb.Y - lb.Y) \ 2)
            
        Else ' Para el resto de angulos:
            Dim r As Vertex
            Call modGraphics.IntersectLine(lt, rb, lb, rt, r)
            C.X = r.X: C.Y = r.Y
        End If
        
        GetCenter = C
    End If
End Function

Private Sub Class_Initialize()
    varColor = &HFFFFFFFF
    varPerspective = Perspectives.WithoutPerspective
    varVisible = True
    varCurrentAnim = 1
    varScale = 1
End Sub

Private Sub Class_Terminate()
    If Not varTex Is Nothing Then Call varTex.RemoveRef
End Sub
