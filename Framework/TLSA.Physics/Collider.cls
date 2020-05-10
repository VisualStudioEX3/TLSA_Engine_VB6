VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Collider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

'variables locales para almacenar los valores de las propiedades
Private varRect As Core.RECTANGLE 'copia local
Private varEnabled As Boolean 'copia local
Private varLocation As VECTOR 'copia local
Private varKey As String 'copia local
Private varColor As Long 'copia local
Private varTypeBox As Long 'copia local
Private varBodyFather As Body

Private halfColor As Long

' Referencia a la entidad a la que esta asociada la caja de colision:
Friend Property Set BodyFather(value As Physics.Body)
    Set varBodyFather = value
End Property

Public Property Get BodyFather() As Physics.Body
    Set BodyFather = varBodyFather
End Property

' Define el tipo de colision:
'   0 - Suelo
'   1 - Pared derecha
'   2 - Pared izquierda
'   3 - Techo
'   5 o > - Tipos especiales (Ej. 5 - Trigger para evento de escalada del jugador, interactua con tipo 0)
Friend Property Let TypeBox(ByVal vData As Long)
'se usa al asignar un objeto a la propiedad, en la parte izquierda de una instrucción Set.
'Syntax: Set x.TypeId = Form1
    varTypeBox = vData
    
    ' Segun el tipo, si es predefinido, se asigna un color:
    Select Case vData
        Case 0: varColor = &HFF00FF00: halfColor = &H7700FF00
        Case 1: varColor = &HFFFFFF00: halfColor = &H77FFFF00
        Case 2: varColor = &HFF0000FF: halfColor = &H770000FF
        Case 3: varColor = &HFFFF0000: halfColor = &H77FF0000
        Case Else: varColor = &HFFFFFFFF: halfColor = &H77FFFFFF
    End Select
End Property

Public Property Get TypeBox() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.TypeId
    TypeBox = varTypeBox
End Property


Public Property Let Color(ByVal vData As Long)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Color = 5
    varColor = vData
End Property

Public Property Get Color() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Color
    Color = varColor
End Property

Friend Property Let Key(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Key = 5
    varKey = vData
End Property

Friend Property Get Key() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Key
    Key = varKey
End Property


' Cambia la posicion de la caja. Z indica la profundidad de colision:
Public Sub SetLocation(ByVal X As Long, ByVal Y As Long, Optional ByVal Z As Long)
'    With varLocation
'        .X = X
'        .Y = Y
'        .Z = Z
'    End With
    
    varRect.X = X '- (varRect.Width \ 2)
    varRect.Y = Y '- (varRect.Height \ 2)
End Sub

Friend Property Let Location(vData As Core.VECTOR)
'se usa al asignar un objeto a la propiedad, en la parte izquierda de una instrucción Set.
'Syntax: Set x.Location = Form1
    
    'varLocation = vData
    
    varRect.X = vData.X '- (varRect.Width \ 2)
    varRect.Y = vData.Y '- (varRect.Height \ 2)
End Property

Public Property Get Location() As Core.VECTOR
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Location
    varLocation.X = varRect.X
    varLocation.Y = varRect.Y
    Location = varLocation
End Property


Public Property Let Enabled(ByVal vData As Boolean)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Enabled = 5
    varEnabled = vData
End Property

Public Property Get Enabled() As Boolean
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Enabled
    Enabled = varEnabled
End Property


Friend Property Let Rect(vData As Core.RECTANGLE)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Rect = 5
    varRect = vData
End Property

Public Property Get Rect() As Core.RECTANGLE
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Rect
    Rect = varRect
End Property

Friend Sub SetRect(ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long)
    With varRect
        .X = X
        .Y = Y
        .Width = Width
        .Height = Height
    End With
End Sub

' Dibuja un cuadrado que representa la caja:
Public Sub Draw() ' Z no se aplica al dibujar la caja, solo a nivel de colsiones. Segun Z se eligira un color:
    ' Dibujamos dos lineas diagonales para definir el centro de la caja:
    'Call gfx.DRAW_Line(Me.Rect.X, Me.Rect.Y, Me.Rect.X + Me.Rect.Width, Me.Rect.Y + Me.Rect.Height, -7, Color)
    'Call gfx.DRAW_Line(Me.Rect.X, Me.Rect.Y + Me.Rect.Height, Me.Rect.X + Me.Rect.Width, Me.Rect.Y, -7, Color)
    Call gfx.Primitives.DrawLine(Me.Rect.X, Me.Rect.Y, Me.Rect.X + Me.Rect.Width, Me.Rect.Y + Me.Rect.Height, -7, Color)
    Call gfx.Primitives.DrawLine(Me.Rect.X, Me.Rect.Y + Me.Rect.Height, Me.Rect.X + Me.Rect.Width, Me.Rect.Y, -7, Color)
    
    ' Dibujamos la caja:
    'Call gfx.DRAW_Box(Me.Rect.X, Me.Rect.Y, Me.Rect.X + Me.Rect.Width, Me.Rect.Y + Me.Rect.Height, -7, Color)
    
    
    Call gfx.Primitives.DrawBox(Me.Rect.X, Me.Rect.Y, Me.Rect.X + Me.Rect.Width, Me.Rect.Y + Me.Rect.Height, -7, Color, True, halfColor)
End Sub

Public Function Clone() As Physics.Collider
    Dim b As New Collider
    
    b.Color = Me.Color
    b.Enabled = Me.Enabled
    Set b.BodyFather = Me.BodyFather
    b.Location = Me.Location
    b.Rect = Me.Rect
    b.TypeBox = Me.TypeBox
    
    Set Clone = b
    Set b = Nothing
End Function

'' Metodo para invocar el evento OnCollision desde el gestor principal:
'Friend Sub RaiseOnCollision(b As Collider, C As VECTOR)
'
'End Sub

' Constructor de la clase:
Private Sub Class_Initialize()
    ' Autogenera una clave unica al objeto:
    varKey = Core.Guid.Generate()
End Sub
