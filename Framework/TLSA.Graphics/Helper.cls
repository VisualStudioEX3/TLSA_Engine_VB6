VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Helper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Clase de funciones adicionales para complementar a otras funciones u objetos de la libreria:
Option Explicit

Private Const varPI As Single = 3.14159265358979

Public Type ARGBColor
    Alpha As Integer
    Red As Integer
    Green As Integer
    Blue As Integer
End Type

' Enumeracion de colores base:
Public Enum Color_Constant
    Black = &HFF000000
    DarkRed = &HFF800000
    DarkGreen = &HFF008000
    DarkYellow = &HFF808000
    DarkBlue = &HFF000080
    DarkPurple = &HFF800080
    DarkCyan = &HFF008080
    Gray = &HFFC0C0C0
    Silver = &HFF808080
    Red = &HFFFF0000
    Green = &HFF00FF00
    Yellow = &HFFFFFF00
    Blue = &HFF0000FF
    Magenta = &HFFFF00FF
    Cyan = &HFF00FFFF
    Brown = &HFF804000
    Orange = &HFFFF8000
    Lima = &HFF80FF00
    Teal = &HFF008080
    Azure = &HFF0080FF
    Aqua = &HFF408080
    White = &HFFFFFFFF
End Enum

Public Property Get PI() As Single
    PI = varPI
End Property

' Traslada la posicion de los puntos tomando como referencia la posicion del primer punto de la lista:
Public Sub TransltatePointList(Points() As Core.Point, X As Long, Y As Long)
    Dim i As Long, dif As Core.Point
    dif.X = X - Points(0).X: Points(0).X = X
    dif.Y = Y - Points(0).Y: Points(0).Y = Y
    For i = 1 To UBound(Points)
        Points(i).X = Points(i).X + dif.X
        Points(i).Y = Points(i).Y + dif.Y
    Next
End Sub

Public Sub ScalePointList(Points() As Core.Point, X As Single, Y As Single)
    Dim rad As Single ' Angulo en radianes
    Dim matResult As D3DMATRIX, matScale As D3DMATRIX, matTras As D3DMATRIX, matRot As D3DMATRIX
    Dim Vec As D3DVECTOR2
    Dim i As Long
    
    'rad = (varPI / 180) * Angle
    
    ' Tomamos con centro o eje de rotacion el primer punto de la lista:
    For i = 1 To UBound(Points)
        Vec.X = Points(i).X
        Vec.Y = Points(i).Y

        Call D3DXMatrixIdentity(matResult)

        '1 Traslacion para ubicar el centro de rotacion:
        Call D3DXMatrixTranslation(matTras, CSng(-Points(0).X), CSng(-Points(0).Y), 0)
        Call D3DXMatrixMultiply(matResult, matResult, matTras)

        '2 Escala
        Call D3DXMatrixScaling(matScale, X, Y, 0)
        Call D3DXMatrixMultiply(matResult, matResult, matScale)
    
'        '3 Rotacion
'        Call D3DXMatrixRotationZ(matRot, rad)
'        Call D3DXMatrixMultiply(matResult, matResult, matRot)
'
        '4 Traslacion para ubicar la posicion real:
        Call D3DXMatrixTranslation(matTras, CSng(Points(0).X), CSng(Points(0).Y), 0)
        Call D3DXMatrixMultiply(matResult, matResult, matTras)
        
        '5 Obtener vector
        Call D3DXVec2TransformCoord(Vec, Vec, matResult)

        Points(i).X = Vec.X
        Points(i).Y = Vec.Y
    Next
End Sub

' Aplica la rotacion a las coordenadas de la lista tomando como centro o eje de rotacion el primer punto de la lista: (requiere referencia a Direct3D)
Public Sub RotatePointList(Points() As Core.Point, Angle As Single)
    Dim rad As Single ' Angulo en radianes
    Dim matResult As D3DMATRIX, matScale As D3DMATRIX, matTras As D3DMATRIX, matRot As D3DMATRIX
    Dim Vec As D3DVECTOR2
    Dim i As Long
    
    rad = (varPI / 180) * Angle
    
    ' Tomamos con centro o eje de rotacion el primer punto de la lista:
    For i = 1 To UBound(Points)
        Vec.X = Points(i).X
        Vec.Y = Points(i).Y

        Call D3DXMatrixIdentity(matResult)

        '1 Traslacion para ubicar el centro de rotacion:
        Call D3DXMatrixTranslation(matTras, CSng(-Points(0).X), CSng(-Points(0).Y), 0)
        Call D3DXMatrixMultiply(matResult, matResult, matTras)

'        '2 Escala
'        Call D3DXMatrixScaling(matScale, s.Width / Texture(Map).Texture_Width, s.Height / Texture(Map).Texture_Height, 0)
'        Call D3DXMatrixMultiply(matResult, matResult, matScale)
    
        '3 Rotacion
        Call D3DXMatrixRotationZ(matRot, rad)
        Call D3DXMatrixMultiply(matResult, matResult, matRot)
        
        '4 Traslacion para ubicar la posicion real:
        Call D3DXMatrixTranslation(matTras, CSng(Points(0).X), CSng(Points(0).Y), 0)
        Call D3DXMatrixMultiply(matResult, matResult, matTras)
        
        '5 Obtener vector
        Call D3DXVec2TransformCoord(Vec, Vec, matResult)

        Points(i).X = Vec.X
        Points(i).Y = Vec.Y
    Next
End Sub

' Invierte la posicion de los puntos tomando como eje o referencia el primer punto de la lista:
Public Sub MirrorPointList(Points() As Core.Point, Mirror As Graphics.Mirrors)
    Dim i As Long, dif As Core.Point
    For i = 1 To UBound(Points)
        dif.X = Points(i).X - Points(0).X
        dif.Y = Points(i).Y - Points(0).Y
        Select Case Mirror
            Case Graphics.Mirrors.Horizontal
                Points(i).X = Points(i).X - (dif.X * 2)
            Case Graphics.Mirrors.Vertical
                Points(i).Y = Points(i).Y - (dif.Y * 2)
            Case Graphics.Mirrors.Both
                Points(i).X = Points(i).X - (dif.X * 2)
                Points(i).Y = Points(i).Y - (dif.Y * 2)
        End Select
    Next
End Sub

' Obtiene un color a partir de las componentes ARGB:
Public Function SetARGB(Alpha As Integer, Red As Integer, Green As Integer, Blue As Integer) As Long
    SetARGB = gfx.ARGB_Set(Alpha, Red, Green, Blue)
End Function

' Obtiene un color a partir de las componentes ARGB:
Public Function SetARGB2(Color As Graphics.ARGBColor) As Long
    SetARGB2 = gfx.ARGB_Set(Color.Alpha, Color.Red, Color.Green, Color.Blue)
End Function

' Obtiene las componenes ARGB de un color:
Public Function GetARGB(Color As Long) As Graphics.ARGBColor
    Dim D As ARGB, g As Graphics.ARGBColor
    Call gfx.ARGB_Get(Color, D)
    With g
        .Alpha = D.Alpha
        .Red = D.Red
        .Green = D.Green
        .Blue = D.Blue
    End With
    GetARGB = g
End Function

' Crea un objeto Sprite e inicializa sus propiedades de textura segun la definicion del archivo de animacion
' asi como las secuencias de animacion establecidas acordes a los tiles definidos en la textura asociada.
' Para poder crear satisfactoriamente el sprite y su instancia la textura asociada debera estar cargada en memoria
' contener los tiles definidos en el archivo de animacion, de lo contrario se producira un error:
' Por defecto el nombre del archivo con las secuencias de animacion sera el mismo que el de la textura con extension *.ani.
' Se puede especificar un nombre concreto a traves del parametro AnimationFile.
Public Function CreateAnimatedSprite(ByVal Texture As Graphics.Texture, Optional AnimationFile As String) As Graphics.Sprite
    On Error GoTo ErrOut
    
    Dim spr As New Graphics.Sprite
    Dim ani As Graphics.Animation
    
    Dim Filename As String
    Dim file As Integer
    Dim line As String
    Dim fields() As String
    Dim subFields() As String
        
    Dim i As Long
    
    Call spr.SetTexture(Texture) ' Asociamos la textura al sprite y cargamos los tiles y puntos de control.
    
    ' Cargamos el archivo en memoria:
    file = FreeFile()
    
    If AnimationFile = "" Then
        Filename = VBA.Left(Texture.Filename, Len(Texture.Filename) - 3) & "ani"
    Else
        Filename = AnimationFile
    End If
    
    'If sys.FILE_Exists(Filename) Then
    If IO.FileExists(Filename) Then
        Open Filename For Input As file
            Do While Not EOF(file)
                Line Input #file, line
                fields = Split(line, "|")
                
                Set ani = New Animation
                
                ' Establecemos los parametros de la secuencia:
                subFields = Split(fields(0), ";")
                ani.Key = subFields(0)
                ani.FrameDelay = subFields(1)
                If UBound(subFields) = 2 Then ani.Looping = subFields(2)
                
                ' Añadimos la referencia al tile de la secuencia:
                subFields = Split(fields(1), ";")
                For i = 0 To UBound(subFields)
                    Call ani.Tiles.AddRef(Texture.Tiles(subFields(i)), subFields(i))
                Next
                
                ' Añadimos la secuencia a la lista:
                Call spr.Animations.Add(ani, ani.Key)
            Loop
        Close file
        
        ' Activamos por defecto el estado de animacion del sprite:
        spr.EnabledAnimation = True
        
        ' *** Si no existiera el archivo de animaciones se crea igualmente el sprite con la informacion disponible de la textura ***
'    Else
'        Call Err.Raise(53, , "No se ha encontrado el archivo de especificacion de sprites:" & vbNewLine & "'" & Filename & "'")
    End If
    
    Set CreateAnimatedSprite = spr
    Set spr = Nothing
    
    Exit Function
    
ErrOut:
    Close file
    Set spr = Nothing
    
    Call Err.Raise(Graphics.GRAPHICS_EXCEPTION.SPRITE_DEFINITION_EXCEPTION, , "Error al cargar parametros del sprite. Revise definicion de tiles en la textura asociada y compare con las secuencias de animacion definidas en el archivo de animacion." & vbNewLine & "'" & Filename & "'")
End Function

' Genera un archivo que contiene la textura asociada y las secuencias de animacion definidas de un sprite:
' Por defecto el nombre del archivo sera el mismo que el de la textura con extension *.ani. Se puede especificar un nombre concreto
' a traves del parametro AnimationFile.
Public Sub SaveAnimatedSprite(Sprite As Graphics.Sprite, Optional AnimationFile As String)
    Dim file As Integer
    Dim Filename As String
    Dim line As String
    Dim ani As Graphics.Animation
    Dim Tile As Graphics.Tile
    
    file = FreeFile()
    
    If AnimationFile = "" Then
        Filename = VBA.Left(Sprite.Texture.Filename, Len(Sprite.Texture.Filename) - 3) & "ani"
    Else
        Filename = AnimationFile
    End If
    
    Open Filename For Output As file
        For Each ani In Sprite.Animations
            line = ani.Key & ";" & ani.FrameDelay & ";" & ani.Looping & "|"
            For Each Tile In ani.Tiles
                line = line & Tile.Key & ";"
            Next
            line = VBA.Left(line, Len(line) - 1)
            Print #file, line
        Next
    Close file
    
    Exit Sub
    
ErrOut:
    Close file
    Call Err.Raise(Err.Number, , Err.Description)
End Sub
