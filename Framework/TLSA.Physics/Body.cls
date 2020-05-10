VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Body"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Enum BodyType ' Todos los tipos devuelven estado de colision:
    FixedBody           ' Responde fisicamente con cualquier cuerpo del simulador pero no se ve afectado por la gravedad.
                        ' Se usa para definir cuerpos solidos como partes del escenario fijas.
                        
    PhysicalEntity      ' Responde fisicamente con cualquier cuerpo del simulador y se ve afectado por la gravedad.
                        
    NotPhysicalEntity    ' El cuerpo esta asociado a una entidad pero no se vera afectado por la gravedad ni por las colisiones.
End Enum

Public Enum ColliderProfiles
    Floor                   ' El colisionador representa suelo ocupando todo el area del cuerpo fisico.
    Roof                    ' El colisionador representa techo ocupando todo el area del cuerpo fisico.
    LeftWall                ' El colisionador representa pared izquierda ocupando todo el area del cuerpo fisico.
    Rightwall               ' El colisionador representa pared derecha ocupando todo el area del cuerpo fisico.
    FloorAndRoofEntireArea  ' 2 colisionadores que representan suelo y techo ocupando todo el area del cuerpo fisico.
    WallsEntireArea         ' 2 colisionadores que representan ambas paredes ocupando todo el area del cuerpo fisico.
    FloorAndRoofHalfArea    ' 2 colisionadores que representan suelo y techo ocupando cada colisionador la mitad del area del cuerpo fisico.
    WallsHalfArea           ' 2 colisionadores que representan ambas paredes ocupando cada colisionador la mitad del area del cuerpo fisico.
    AllEntireArea           ' 4 colisionadores que representan suelo, techo y ambas paredes, con un grosor definido por el programador.
    AllPriorityFloorRoof    ' Igual que All, solo que las esquinas del cuerpo dan prioridad a los colisionadores de suelo y techo.
    AllWihtoutCorners       ' Igual que All, solo que las esquinas del cuerpo estan libres de colisionadores.
    Manual                  ' El programador indica el tipo de colisionador y su tamaño manualmente.
End Enum

Public Type CollisionEventData
    Body As Physics.Body        ' Entidad con la que hubo colision.
    Colliders() As Long         ' Lista de Colliders (su TypeId) que han colisionado en la entidad.
    ColliderCount As Long       ' Numero de Colliders que han colisionado (Ubound + 1).
End Type

'variables locales para almacenar los valores de las propiedades
Private varLocation As Core.VECTOR, varLastLocation As Core.VECTOR
Private varRect As Core.RECTANGLE
Private varColliders As Physics.ColliderList
Private varEnabled As Boolean
Private varFixed As Boolean
Private varKey As String
Private varWeight As Long
Private varForce As Core.VECTOR
Private varColor As Long, halfColor As Long

'Para activar este evento, use RaiseEvent con la siguiente sintaxis:
'RaiseEvent OnCollision[(arg1, arg2, ... , argn)]
Public Event OnCollision(BodyCount As Long, E() As Physics.CollisionEventData)

' Evento que sucede cuando se recibe fuerza o un golpe de otro objeto o evento externo (impacto de un rayo trazado por ejemplo):
Public Event OnHit(Force As Single, eventData As Variant)

Private varPhysicType As Physics.BodyType
Private varHit As Boolean
Private varFatherSim As Physics.Simulator
Private varTag As Variant

Private varBaseType As Long, varSize As Long

Private varForceAngle As Single ' El angulo trazado entre la posicion anterior y la actual. Se usa en el simulador para aplicar la direccion
                                ' a la fuerza aplicada al choque con otro cuerpo.

' Constantes de para simular friccion:
Private varMinFriction As Core.VECTOR ' Indica el factor de friccion en X e Y independientemente de la fuerza cuando se aplica.
' Esto hace por ejemplo que el cuerpo al aplicarsele fuerza negativa en Y tarde menos o tarde mas en bajar dada la fuerza de gravedad.

Private varMaxFriction As Core.VECTOR ' La minima friccion se aplica siempre cuando no haya colision alguna. De esta forma se
' conserva mas tiempo la fuerza de empuje. La maxima se aplica cuando el cuerpo colisiona con otro simulando la absorcion de parte la
' fuerza por parte del otro cuerpo.

Private varVisible As Boolean

Public Property Get Visible() As Boolean
    Visible = varVisible
End Property

Public Property Let Visible(value As Boolean)
    varVisible = value
End Property

' Devuelve el angulo que define la direccion que aplica la fuerza en el objeto:
Public Property Get Direction() As Single
    Direction = varForceAngle
End Property

Friend Property Let Direction(value As Single)
    varForceAngle = value
End Property

Public Property Get MinFriction() As Core.VECTOR
    MinFriction = varMinFriction
End Property

Public Property Let MinFriction(value As Core.VECTOR)
    varMinFriction = value
End Property

Public Sub SetMinFriction(X As Single, Y As Single)
    varMinFriction.X = X
    varMinFriction.Y = Y
End Sub

Public Property Get MaxFriction() As Core.VECTOR
    MaxFriction = varMaxFriction
End Property

Public Property Let MaxFriction(value As Core.VECTOR)
    varMaxFriction = value
End Property

Public Sub SetMaxFriction(X As Single, Y As Single)
    varMaxFriction.X = X
    varMaxFriction.Y = Y
End Sub

Public Property Get PhysicType() As Physics.BodyType
    PhysicType = varPhysicType
End Property

Public Property Let PhysicType(value As Physics.BodyType)
    varPhysicType = value
End Property


' Propiedad que permite guardar un valor variant, inclusive un array o un object:
Public Property Get Tag() As Variant
    Tag = varTag
End Property

Public Property Let Tag(value As Variant)
    varTag = value
End Property


Friend Property Get FatherSim() As Physics.Simulator
    Set FatherSim = varFatherSim
End Property

Friend Property Set FatherSim(sim As Physics.Simulator)
    Set varFatherSim = sim
End Property

' Indica que el cuerpo esta en colision:
Public Property Get Hit() As Boolean
    Hit = varHit
End Property

Friend Property Let Hit(value As Boolean)
    varHit = value
End Property

' Color que se usara para dibujar el area del cuerpo independientemente de sus colisionadores si el dibujado de depuracion esta activado:
Public Property Get Color() As Long
    Color = varColor
End Property

Public Property Let Color(value As Long)
    varColor = value
    
    Dim argb As Graphics.ARGBColor: argb = gfx.Helper.GetARGB(value)
    argb.Alpha = 127
    halfColor = gfx.Helper.SetARGB2(argb)
End Property

' Clave unica del cuerpo:
Friend Property Let Key(value As String)
    varKey = value
End Property

Public Property Get Key() As String
    Key = varKey
End Property

' Fuerza independiente en ejes X e Y:
Friend Property Let Force(value As Core.VECTOR)
    varForce = value
End Property

Public Property Get Force() As Core.VECTOR
    Force = varForce
End Property

' Aplica fuerza en una direccion determinada por un angulo:
Public Sub SetForce(Force As Single, Direction As Single, Optional Add As Boolean) ' Parametro Add = Se incrementa la fuerza actual con la fuerza indicada.
    ' Calculamos el punto destino al rotar el punto origen en el angulo y radio/distancia indicado:
    Dim p As Core.Point: p = Core.Math.TranslatePoint(Core.Generics.VECTOR2POINT(varLocation), CLng(Force), Direction)
    
    ' Aplicamos la fuerza en X e Y:
    Call Me.SetForceXY(CLng(varLocation.X) - p.X, CLng(varLocation.Y) - p.Y, Add)
End Sub

' Devuelve un valor entero que representa la fuerza aplicada en una direccion (un angulo que aplica la fuerza en los ejes X e Y)
Public Function GetForce() As Single
    ' Obtenemos la distancia entre la posicion anterior y la actual para representar un valor aproximado de la fuerza ejercida sobre
    ' el objeto en el angulo actual:
    Dim A As Core.Point, b As Core.Point
    A = Core.Generics.VECTOR2POINT(varLocation)
    b = Core.Generics.VECTOR2POINT(varLastLocation)
    
    GetForce = CSng(Core.Math.GetDistance(A, b))
End Function

' Aplica valores de fuerza independientes a X e Y:
Public Sub SetForceXY(X As Single, Y As Single, Optional Add As Boolean) ' Parametro Add = Se incrementa la fuerza actual con la fuerza indicada.
    If Add Then
        varForce.X = varForce.X + X
        varForce.Y = varForce.Y + Y
    Else
        varForce.X = X
        varForce.Y = Y
    End If
End Sub

' Aplica fuerza en el eje X:
Public Sub SetForceX(value As Long, Optional Add As Boolean) ' Parametro Add = Se incrementa la fuerza actual con la fuerza indicada.
    If Add Then varForce.X = varForce.X + value Else varForce.X = value
End Sub

' Aplica fuerza en el eje Y:
Public Sub SetForceY(value As Long, Optional Add As Boolean) ' Parametro Add = Se incrementa la fuerza actual con la fuerza indicada.
    If Add Then varForce.Y = varForce.Y + value Else varForce.Y = value
End Sub

' Devuelve verdadero si el cuerpo esta fijo o es de tipo FixedBodie:
Public Property Get Fixed() As Boolean
    Fixed = varFixed Or varPhysicType = FixedBody
End Property

Public Property Let Fixed(value As Boolean)
    varFixed = value
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


Public Property Get Colliders() As Physics.ColliderList
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Colliders
    Set Colliders = varColliders
End Property


Public Property Let Rect(vData As Core.RECTANGLE)
'se usa al asignar un objeto a la propiedad, en la parte izquierda de una instrucción Set.
'Syntax: Set x.Rect = Form1
    varRect = vData
End Property

Public Property Get Rect() As Core.RECTANGLE
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Rect
    Rect = varRect
End Property

Public Sub SetRect(ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long)
    With varRect
        .X = X
        .Y = Y
        .Width = Width
        .Height = Height
    End With
    
    varLastLocation = varLocation
    
    With varLocation
        .X = varRect.X + (varRect.Width / 2)
        .Y = varRect.Y + (varRect.Height / 2)
    End With
    
    Call Me.SetColliders(varBaseType, varSize)
End Sub


Public Property Get Weight() As Long
    Weight = varWeight
End Property

Public Property Let Weight(value As Long)
    varWeight = value
End Property


Public Property Let Location(vData As Core.VECTOR)
'se usa al asignar un objeto a la propiedad, en la parte izquierda de una instrucción Set.
'Syntax: Set x.Location = Form1
    Call Me.SetLocation(vData.X, vData.Y)
End Property

Public Property Get Location() As Core.VECTOR
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Location
    Location = varLocation
End Property

Public Property Let Z(value As Long)
    varLocation.Z = value
End Property

Public Property Get Z() As Long
    Z = varLocation.Z
End Property

Public Sub SetLocation(ByVal X As Long, ByVal Y As Long)
    varLastLocation = varLocation
    
    With varLocation
        .X = X
        .Y = Y
    End With
    
    varRect.X = X - (varRect.Width \ 2)
    varRect.Y = Y - (varRect.Height \ 2)
    
    Call UpdateCollideresLocations
End Sub

' Actualiza la posicion de todas las cajas de colision de la entidad en base a la posicion de la misma:
Private Sub UpdateCollideresLocations()
    Dim diff As Core.VECTOR ' Diferencia de posicion en X e Y respecto a la anterior.
    Dim IsChangeX As Boolean, IsChangeY As Boolean
    
    IsChangeX = varLastLocation.X <> varLocation.X
    IsChangeY = varLastLocation.Y <> varLocation.Y
        
    ' Si la nueva localizacion es distinta de la anterior realizamos los cambios:
    If IsChangeX Or IsChangeY Then
        ' Sacamos la diferencia con la posicion anterior:
        diff.X = varLocation.X - varLastLocation.X
        diff.Y = varLocation.Y - varLastLocation.Y
        
        Dim b As Collider
        For Each b In Me.Colliders
            Call b.SetLocation(diff.X + b.Location.X, diff.Y + b.Location.Y)
        Next
        
        ' Obtenemos el angulo que define la direccion hacia donde se aplica la fuerza:
        Dim locA As Core.Point, locB As Core.Point
        locA = Core.Generics.VECTOR2POINT(varLocation)
        locB = Core.Generics.VECTOR2POINT(varLastLocation)
        varForceAngle = Core.Math.GetAngle(locA, locB)
    End If
End Sub

' Si PrefType es 1 Param define la anchura de los colisionadores de la caja:
Public Sub SetColliders(PrefType As Physics.ColliderProfiles, Optional Param As Variant)
    Call varColliders.Clear ' Eliminamos los colisionadores actuales.
    
    varBaseType = PrefType
    varSize = CLng(Param)
    
    Select Case PrefType
        Case Physics.ColliderProfiles.Manual
            If IsEmpty(Param) Then Call Err.Raise(vbObjectError, , "Physics.Body.SetColliders() - Debe indicar un valor al parametro Param para definir el tipo de colisionador.")
            Call varColliders.Add(0, 0, 0, varRect.Width, varRect.Height, CLng(Param), Me)
            
        Case Physics.ColliderProfiles.AllEntireArea
            If IsEmpty(Param) Or Param = 0 Then Call Err.Raise(vbObjectError, , "Debe indicar un valor al parametro Param para definir la anchura de los colisionadores.")
            
            ' Suelo:
            Call varColliders.Add(0, varRect.Height - CLng(Param), 0, varRect.Width, CLng(Param), 0, Me)
            ' Techo:
            Call varColliders.Add(0, 0, 0, varRect.Width, CLng(Param), 3, Me)
            
            ' Pared derecha:
            Call varColliders.Add(varRect.Width - CLng(Param), 0, 0, CLng(Param), varRect.Height, 1, Me)
            ' Pared izquierda:
            Call varColliders.Add(0, 0, 0, CLng(Param), varRect.Height, 2, Me)

        Case Physics.ColliderProfiles.AllPriorityFloorRoof
            If IsEmpty(Param) Or Param = 0 Then Call Err.Raise(vbObjectError, , "Debe indicar un valor al parametro Param para definir la anchura de los colisionadores.")
            
            ' Suelo:
            Call varColliders.Add(0, varRect.Height - CLng(Param), 0, varRect.Width, CLng(Param), 0, Me)
            ' Techo:
            Call varColliders.Add(0, 0, 0, varRect.Width, CLng(Param), 3, Me)
            
            ' Pared derecha:
            Call varColliders.Add(varRect.Width - CLng(Param), CLng(Param), 0, CLng(Param), varRect.Height - CLng(Param * 2), 1, Me)
            ' Pared izquierda:
            Call varColliders.Add(0, CLng(Param), 0, CLng(Param), varRect.Height - CLng(Param * 2), 2, Me)
        
        Case Physics.ColliderProfiles.AllWihtoutCorners  ' Experimental, las esquinas del cuerpo no estan acotadas por ninguno de los colisionadores:
            If IsEmpty(Param) Or Param = 0 Then Call Err.Raise(vbObjectError, , "Debe indicar un valor al parametro Param para definir la anchura de los colisionadores.")
            
            ' Suelo:
            Call varColliders.Add(CLng(Param), varRect.Height - CLng(Param), 0, varRect.Width - CLng(Param * 2), CLng(Param), 0, Me)
            ' Techo:
            Call varColliders.Add(CLng(Param), 0, 0, varRect.Width - CLng(Param * 2), CLng(Param), 3, Me)
            
            ' Pared derecha:
            Call varColliders.Add(varRect.Width - CLng(Param), CLng(Param), 0, CLng(Param), varRect.Height - CLng(Param * 2), 1, Me)
            ' Pared izquierda:
            Call varColliders.Add(0, CLng(Param), 0, CLng(Param), varRect.Height - CLng(Param * 2), 2, Me)
            
        
        Case Physics.ColliderProfiles.FloorAndRoofEntireArea   ' Suelo/Techo ocupando el area completa del cuerpo definido:
            ' Suelo:
            Call varColliders.Add(0, 0, 0, varRect.Width, varRect.Height, 3, Me)
            ' Techo:
            Call varColliders.Add(0, 0, 0, varRect.Width, varRect.Height, 0, Me)
            
        
        Case Physics.ColliderProfiles.WallsEntireArea          ' Paredes derecha e izquierda ocupando el area completa del cuerpo definido:
            ' Derecha:
            Call varColliders.Add(0, 0, 0, varRect.Width, varRect.Height, 2, Me)
            ' Izquiera:
            Call varColliders.Add(0, 0, 0, varRect.Width, varRect.Height, 1, Me)
            
        
        Case Physics.ColliderProfiles.FloorAndRoofHalfArea    ' Suelo/Techo ocupando cada colisionador el la mitad del area del cuerpo definido:
            ' Suelo:
            Call varColliders.Add(0, 0, 0, varRect.Width, varRect.Height \ 2, 3, Me)
            ' Techo:
            Call varColliders.Add(0, varRect.Height \ 2, 0, varRect.Width, varRect.Height \ 2, 0, Me)
            
        
        Case Physics.ColliderProfiles.WallsHalfArea           ' Paredes derecha e izquierda ocupando cada colisionador el la mitad del area del cuerpo definido:
            ' Derecha:
            Call varColliders.Add(0, 0, 0, varRect.Width \ 2, varRect.Height, 2, Me)
            ' Izquiera:
            Call varColliders.Add(varRect.Width \ 2, 0, 0, varRect.Width \ 2, varRect.Height, 1, Me)
            
        
        Case Physics.ColliderProfiles.Floor                   ' Suelo:
            Call varColliders.Add(0, 0, 0, varRect.Width, varRect.Height, 3, Me)
        
        
        Case Physics.ColliderProfiles.Roof                    ' Techo:
            Call varColliders.Add(0, 0, 0, varRect.Width, varRect.Height, 0, Me)


        Case Physics.ColliderProfiles.LeftWall                ' Pared izquierda:
            Call varColliders.Add(0, 0, 0, varRect.Width, varRect.Height, 2, Me)
            
        
        Case Physics.ColliderProfiles.Rightwall               ' Pared derecha:
            Call varColliders.Add(0, 0, 0, varRect.Width, varRect.Height, 1, Me)
            
            
        Case Else ' Cargamos la informacion de un archivo externo: (no implementado)
        
    End Select
End Sub


Public Sub Draw()
    If varVisible Then
        If modPhysics.debugDrawColliders Then
            ' Dibujamos las guias de los colisionadores:
            Dim b As Collider
            For Each b In Me.Colliders
                b.Draw
            Next
        Else
            ' Dibujamos la caja:
            Call gfx.Primitives.DrawBox(Me.Rect.X, Me.Rect.Y, Me.Rect.X + Me.Rect.Width, Me.Rect.Y + Me.Rect.Height, -7, Me.Color, _
            True, halfColor)
        End If
    End If
End Sub

' Actualiza la aceleracion de la entidad:
Public Sub Update()
    ' Fuerza:
    If varForce.X <> 0 Then
        If varForce.X > 0 Then
            If varHit Then
                varForce.X = varForce.X - varMinFriction.X ' 0.25
            Else
                varForce.X = varForce.X - varMaxFriction.X '0.05
            End If
            If varForce.X < 0 Then varForce.X = 0
            
        ElseIf varForce.X < 0 Then
            If varHit Then
                varForce.X = varForce.X + varMinFriction.X '0.25
            Else
                varForce.X = varForce.X + varMaxFriction.X '0.05
            End If
            
            If varForce.X > 0 Then varForce.X = 0
        End If
    End If
    
    If varForce.Y <> 0 Then
        If varForce.Y > 0 Then
            varForce.Y = varForce.Y - varMaxFriction.Y  '0.5
        ElseIf varForce.Y < 0 Then
            varForce.Y = varForce.Y + varMaxFriction.Y  '0.5
        End If
    End If
End Sub

' Crea una copia de la entidad:
Public Function Clone() As Physics.Body
    Dim E As New Physics.Body
    
    E.Enabled = Me.Enabled
    E.Fixed = Me.Fixed
    E.Location = Me.Location
    E.Rect = Me.Rect
    E.Force = Me.Force
    E.MaxFriction = Me.MaxFriction
    E.MinFriction = Me.MinFriction
    E.Weight = Me.Weight
    E.Color = Me.Color
    E.PhysicType = Me.PhysicType
    E.Tag = Me.Tag
    E.Direction = Me.Direction
    
    Dim b As Physics.Collider
    For Each b In Me.Colliders
        Call E.Colliders.AddObject(b.Clone())
    Next
    
    Set Clone = E
    Set E = Nothing
End Function

' *** Este evento solo se puede invocar desde el simulador ***:
Friend Sub RaiseOnCollision(Count As Long, E() As Physics.CollisionEventData)
    RaiseEvent OnCollision(Count, E())
End Sub

' Permite invocar publicamente el evento OnHit():
Public Sub RaiseOnHitEvent(Force As Single, eventData As Variant)
    RaiseEvent OnHit(Force, eventData)
End Sub

Public Function TraceRay(Source As Core.Point, Direction As Single, Hit As Core.Point, Target As Core.Point, ZDiscard() As Long) As Physics.Body
    Dim r As Physics.Ray
    Set r = varFatherSim.CreateRayTracer(Source)
    r.Direction = Direction
    Set TraceRay = r.Trace(Me, ZDiscard): Hit = r.Hit: Target = r.Target
End Function

' Traza un segmento entre un punto origen y un punto destino y devuelve el punto de corte mas cercano al punto origen:
Public Function TraceSegment(Source As Core.Point, Target As Core.Point, Hit As Core.Point, ZDiscard() As Long) As Physics.Body
    Dim r As Physics.Ray
    Set r = varFatherSim.CreateRayTracer(Core.Generics.NullPOINT)
    Set TraceSegment = r.TraceSegment(Source, Target, Me, ZDiscard): Hit = r.Hit
End Function

' Constructor de la clase:
Private Sub Class_Initialize()
    ' Autogenera una clave unica al objeto:
    varKey = Core.Guid.Generate()
    
    Set varColliders = New ColliderList
    varColor = &HFFFFFFFF
    varPhysicType = FixedBody
    varVisible = True
    
    ' Valores por defecto de friccion del cuerpo:
    Call Me.SetMaxFriction(0.18, 0.25)
    Call Me.SetMinFriction(0.5, 0.5)
End Sub
