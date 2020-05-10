VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ControlPointMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Puntos de control de la entidad:
Private varFixedPoints() As Core.POINT ' Puntos de control fijos de la entidad. Mantienen las coordenadas de de la textura/tile.
Private varPoints() As Core.POINT, varCount As Long ' Puntos de control para transformaciones. Copia de los puntos de control fijos a los que se les aplicara las transformaciones necesarias.
Private varKeyPoints() As String ' Lista de claves de los puntos.

Private varLocation As Core.POINT ' Posicion de la entidad.
Private varAngle As Single ' Angulo de la entidad.
Private varMirror As Graphics.Mirrors ' Espejado de la entidad.
Private varSize As Core.SIZE, varFixedSize As Core.SIZE ' Tamaños para calcular la escala.
Private varScaleX As Single, varScaleY As Single ' Escala que se aplicara.

Private varTransform As Boolean ' Indica si se ha modificado algun parametro del objeto y que afecte a los puntos de control.

Private varOffset As Core.POINT

Private varTextureKey As String     ' Se utiliza para indicar a que textura pertenece el mapa de puntos de control.
Private varTileKey As String        ' Se utiliza para indicar a que tile pertenece el mapa de puntos de control.

Friend Property Get TextureKey() As String
    TextureKey = varTextureKey
End Property

Friend Property Let TextureKey(value As String)
    varTextureKey = value
End Property

Friend Property Get TileKey() As String
    TileKey = varTileKey
End Property

Friend Property Let TileKey(value As String)
    varTileKey = value
End Property

Friend Property Get OffSet() As Core.POINT
    OffSet = varOffset
End Property

Friend Property Let OffSet(value As Core.POINT)
    varOffset = value
End Property

Friend Sub SetOffSet(X As Long, Y As Long)
    varOffset.X = X
    varOffset.Y = Y
End Sub

' Indica si existe un elemento con la clave indicada:
Public Function Exists(Key As String) As Boolean
    If varCount > 0 Then
        Dim i As Long
        For i = 0 To UBound(varKeyPoints)
            If varKeyPoints(i) = Key Then
                Exists = True
                Exit Function
            End If
        Next
    End If
End Function

' Devuelve una copia de la lista de claves de los puntos de control:
Public Function GetKeyList() As String()
    GetKeyList = varKeyPoints
End Function

' Obtiene o establece la posicion del punto primario de la lista trasladando el resto de puntos en relacion al punto primario:
Public Property Get Location() As Core.POINT
    Location = varLocation
End Property

Public Property Let Location(value As Core.POINT)
    varLocation = value
    varTransform = True
End Property

Public Sub SetLocation(X As Long, Y As Long)
    varLocation.X = X + varOffset.X
    varLocation.Y = Y + varOffset.Y
    varTransform = True
End Sub


Public Property Get SIZE() As Core.SIZE
    SIZE = varSize
End Property

Public Property Let SIZE(value As Core.SIZE)
    varSize = value
    
    ' ************************
    If varFixedSize.Width = 0 Then varFixedSize.Width = varSize.Width
    If varFixedSize.Height = 0 Then varFixedSize.Height = varSize.Height
    ' ************************
    
    varScaleX = varSize.Width / varFixedSize.Width
    varScaleY = varSize.Height / varFixedSize.Height
    
    varTransform = True
End Property

Public Sub SetSize(Width As Long, Height As Long)
    varSize.Width = Width
    varSize.Height = Height
    
    ' ************************
    If varFixedSize.Width = 0 Then varFixedSize.Width = varSize.Width
    If varFixedSize.Height = 0 Then varFixedSize.Height = varSize.Height
    ' ************************
    
    varScaleX = varSize.Width / varFixedSize.Width
    varScaleY = varSize.Height / varFixedSize.Height
    
    varTransform = True
End Sub

Public Sub SetFixedSize(Width As Long, Height As Long)
    varFixedSize.Width = Width
    varFixedSize.Height = Height
    
    Me.SIZE = varFixedSize
    
    varTransform = True
End Sub


' Obtiene o establece el angulo del mapa de puntos de control:
Public Property Get Angle() As Single
    Angle = varAngle
End Property

Public Property Let Angle(value As Single)
    varAngle = value
    varTransform = True
End Property


' Obtiene o establece el espejado que se aplicara al mapa de puntos de control:
Public Property Get Mirror() As Graphics.Mirrors
    Mirror = varMirror
End Property

Public Property Let Mirror(value As Graphics.Mirrors)
    varMirror = value
    varTransform = True
End Property

' Aplica todas las transformaciones al mapa de puntos de control si las hubiera:
Public Sub Update()
    If varTransform Then ' Actualizamos las coordenadas de los puntos de control:
        varPoints = varFixedPoints
        
        ' Corregimos las posiciones respecto al centro del sprite:
'        If varCount > 0 Then
'            Dim i As Integer
'            For i = 1 To UBound(varPoints)
'                varPoints(i).X = varPoints(i).X + varOffset.X: varPoints(i).Y = varPoints(i).Y + varOffset.Y
'            Next
'        End If
        
        Call varHelper.ScalePointList(varPoints(), varScaleX, varScaleY) ' Escalamos el mapa.
        Call varHelper.TransltatePointList(varPoints(), varLocation.X, varLocation.Y)  ' Trasladamos el mapa.
        
        Call varHelper.MirrorPointList(varPoints(), varMirror) ' Aplicamos espejado al mapa.
        Call varHelper.RotatePointList(varPoints(), varAngle) ' Rotamos el mapa.
        varTransform = False
    End If
End Sub

' Dibuja la posicion de los puntos de control mediante huesos tomando de origen el primer punto de la lista:
Public Sub Draw()
    If varCount > 0 Then
        Dim i As Long, Color As Long
        For i = 1 To UBound(varPoints)
            Select Case i
                Case 1: Color = Graphics.Color_Constant.Red
                Case 2: Color = Graphics.Color_Constant.Orange
                Case 3: Color = Graphics.Color_Constant.Yellow
                Case 4: Color = Graphics.Color_Constant.Green
                Case 5: Color = Graphics.Color_Constant.Cyan
                Case 6: Color = Graphics.Color_Constant.Blue
                Case 7: Color = Graphics.Color_Constant.Magenta
                Case Else: Color = Graphics.Color_Constant.White
            End Select
            
            Call gfx.DRAW_Line(varPoints(0).X + modGraphics.varOffset.X, varPoints(0).Y + modGraphics.varOffset.Y, varPoints(i).X + modGraphics.varOffset.X, varPoints(i).Y + modGraphics.varOffset.Y, -7, Color)
        Next
    End If
End Sub

' *** Solo en tiempo de diseño ***
' Añade un nuevo elemento a la lista:
Public Sub Add(Key As String, X As Long, Y As Long)
    If Not Me.Exists(Key) Then
        ReDim Preserve varPoints(varCount)
        ReDim Preserve varFixedPoints(varCount)
        ReDim Preserve varKeyPoints(varCount)
        varCount = varCount + 1
        
        With varPoints(varCount - 1)
            .X = X
            .Y = Y
        End With
        varFixedPoints(varCount - 1) = varPoints(varCount - 1)
        
        varKeyPoints(varCount - 1) = Key
    Else
        Call Err.Raise(Graphics.GRAPHICS_EXCEPTION.ITEM_KEY_DUPLICATED, , "Ya existe un punto de control con la misma clave: '" & Key & "' en Textura: '" & varTextureKey & "', Tile: '" & varTileKey & "'")
    End If
End Sub

' *** Solo en tiempo de diseño ***
' Elimina un elemento de la lista:
Public Sub Remove(Key As String)
    If UCase(Key) = UCase("DEFAULT") Then
        Call Err.Raise(Graphics.GRAPHICS_EXCEPTION.ARGUMENT_EXCEPTION, , "No se puede eliminar el punto principal 'Default' en Textura: '" & varTextureKey & "', Tile: '" & varTileKey & "'")
    End If
    
    Dim i As Long, j As Long
    For i = 0 To varCount - 1
        If varKeyPoints(i) = Key Then ' Si la clave coincide:
            
            varCount = varCount - 1 ' Descontamos un elemento en el contador.
            
            ' Recorremos la lista desde el elemento a borrar y empezamos a escribir el siguiente elemento a partir del actual:
            For j = i To varCount - 1
                varPoints(j) = varPoints(j + 1)
                varKeyPoints(j) = varKeyPoints(j + 1)
            Next
            
            ' Redimensionamos los arrays:
            If varCount = 0 Then ' Si el contador vale 0:
                Erase varPoints, varFixedPoints, varKeyPoints ' Borramos la instancia de los arrays.
            Else ' De lo contrario descontamos un elemento en el tamaño del array:
                ReDim Preserve varPoints(varCount - 1)
                ReDim Preserve varKeyPoints(varCount - 1)
            End If
            
            varFixedPoints = varPoints ' Guardamos la copia de puntos fijos para los calculos y transformaciones.
            
            Exit Sub
        End If
    Next

    ' Si no se encontro el elemento devolvemos una excepcion:
    Call Err.Raise(Graphics.GRAPHICS_EXCEPTION.ITEM_KEY_NOT_FOUND, , "No se encuentra ningun punto de control con la clave indicada: '" & Key & "' en Textura: '" & varTextureKey & "', Tile: '" & varTileKey & "'")
End Sub

' Borra todos los puntos de control menos el punto primario o por defecto:
Public Sub Clear()
    ' Borramos todo menos el primer elemento:
    ReDim Preserve varPoints(0)
    ReDim Preserve varFixedPoints(0)
    ReDim Preserve varKeyPoints(0)
    varCount = 1
End Sub

' Obtiene los valores de elemento:
Public Function GetPoint(Key As String) As Core.POINT
Attribute GetPoint.VB_UserMemId = 0
    If varCount > 0 Then
        Dim i As Long
        For i = 0 To varCount - 1
            If varKeyPoints(i) = Key Then
                GetPoint = varPoints(i)
                Exit Function
            End If
        Next
        
        ' Si no se encontro el elemento devolvemos una excepcion:
        Call Err.Raise(Graphics.GRAPHICS_EXCEPTION.ITEM_KEY_NOT_FOUND, , "No se encuentra ningun punto de control con la clave indicada: '" & Key & "' en Textura: '" & varTextureKey & "', Tile: '" & varTileKey & "'")
    End If
End Function

' *** Solo en tiempo de diseño ***
' Establece los valores del elemento:
Public Sub SetPoint(Key As String, X As Long, Y As Long)
    If varCount > 0 Then
        Dim i As Long
        For i = 0 To varCount - 1
            If varKeyPoints(i) = Key Then
                varPoints(i).X = X
                varPoints(i).Y = Y
                
                varFixedPoints(i) = varPoints(i) ' *** Actualiza el valor en la lista fija ***
                
                Exit Sub
            End If
        Next
        
        ' Si no se encontro el elemento devolvemos una excepcion:
        Call Err.Raise(Graphics.GRAPHICS_EXCEPTION.ITEM_KEY_NOT_FOUND, , "No se encuentra ningun punto de control con la clave indicada: '" & Key & "' en Textura: '" & varTextureKey & "', Tile: '" & varTileKey & "'")
    End If
End Sub

' Devuelve el numero de elementos en la lista:
Public Property Get Count() As Long
    Count = varCount
End Property

' Reestablece el mapa de puntos de control a sus valores iniciales, al igual que el resto de sus propiedades:
Public Sub ResetControlPoints()
    varPoints = varFixedPoints
    varLocation = varPoints(0)
    varAngle = 0
    varMirror = WithoutMirroring
End Sub


' *** Solo en tiempo de diseño ***
' Escribe una copia de los arrays con los puntos y claves de la lista:
Friend Sub SetArrayList(TransformPoints() As Core.POINT, FixedPoints() As Core.POINT, keys() As String)
    varPoints = TransformPoints
    varFixedPoints = FixedPoints
    varKeyPoints = keys
End Sub

' Crea una copia de la instancia:
Public Function Clone() As Graphics.ControlPointMap
    Dim ret As New Graphics.ControlPointMap
    Call ret.SetArrayList(varPoints, varFixedPoints, varKeyPoints)
    Call ret.UpdateCount
    Set Clone = ret
    Set ret = Nothing
End Function

' *** Solo en tiempo de diseño ***
' Actualiza el contador en la operacion de copia/clonacion:
Friend Sub UpdateCount()
    varCount = UBound(varFixedPoints) + 1
End Sub

Private Sub Class_Initialize()
    Call Me.Add("Default", 0, 0) ' Añadimos el punto base de la lista (el offset)
End Sub

Private Sub Class_Terminate()
    Erase varPoints, varFixedPoints, varKeyPoints
End Sub
