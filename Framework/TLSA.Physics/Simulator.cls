VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Simulator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private varTimer As Core.Timer

Private varEnabled As Boolean
Private varGravity As Core.VECTOR
Private varList As New Physics.BodyList
Private varDelay As Long
Private varWorkArea As Core.RECTANGLE

Private varMonitor As Physics.SimulatorMonitor

Public Property Get Monitor() As Physics.SimulatorMonitor
    Set Monitor = varMonitor
End Property

' Area que trabajara el simulador de fisica. Fuera de ese area no se contempla ningun calculo:
Public Sub SetWorkArea(X As Long, Y As Long, Width As Long, Height As Long)
    With varWorkArea
        .X = X
        .Y = Y
        .Width = Width
        .Height = Height
    End With
End Sub

Public Property Let WorkArea(vData As Core.RECTANGLE)
    varWorkArea = vData
End Property

Public Property Get WorkArea() As Core.RECTANGLE
    WorkArea = varWorkArea
End Property


Public Property Let Delay(ByVal vData As Long)
    varDelay = vData
End Property

Public Property Get Delay() As Long
    Delay = varDelay
End Property


Public Property Get Bodies() As BodyList
    Set Bodies = varList
End Property


Public Sub SetGravity(X As Long, Y As Long)
    varGravity.X = X
    varGravity.Y = Y
End Sub

Public Property Let Gravity(ByRef vData As Core.VECTOR)
    varGravity = vData
End Property

Public Property Get Gravity() As Core.VECTOR
    Gravity = varGravity
End Property


Public Property Let Enabled(ByVal vData As Boolean)
    varEnabled = vData
End Property

Public Property Get Enabled() As Boolean
    Enabled = varEnabled
End Property

' Genera las listas con todos los cuerpos que esten dentro del area de trabajo y clasificando los cuerpos sin son fijos o entidades:
Private Sub GetActiveBodies(ActiveBodies As Physics.BodyList, EntityBodies As Physics.BodyList, FixedBodies As Physics.BodyList, NotPhysicBodies As Physics.BodyList)
    Dim b As Physics.Body
    
    Set ActiveBodies = New Physics.BodyList
    Set FixedBodies = New Physics.BodyList
    Set EntityBodies = New Physics.BodyList
    Set NotPhysicBodies = New Physics.BodyList
    
    ' Obtenemos todos los cuerpos que esten habilitados y dentro del area de trabajo del simulador:
    For Each b In Me.Bodies
        If b.Enabled Then
            If Core.Math.IntersectRect(b.Rect, Me.WorkArea) Then
                Call ActiveBodies.AddObject(b)          ' A�adimos el cuerpo a la lista de cuerpos activos.
                
                ' Clasificamos el cuerpo y lo agregamos a la lista que corresponda:
                If b.PhysicType = PhysicalEntity Then
                    Call EntityBodies.AddObject(b)      ' Entidad fisica.
                ElseIf b.PhysicType = NotPhysicalEntity Then
                    Call NotPhysicBodies.AddObject(b)   ' Entidad no fisica.
                ElseIf b.PhysicType = FixedBody Then
                    Call FixedBodies.AddObject(b)       ' Escenario y demas objetos fijos.
                End If
            End If
        End If
    Next
End Sub

' Actualiza los estados del simulador:
Public Sub Update()
    Dim Correction As Core.VECTOR, tmp As Long
    Dim arrayE() As CollisionEventData, aECount As Long
    Dim ColliderID As Long
    
    Dim ActiveBodies As Physics.BodyList
    Dim FixedBodies As Physics.BodyList
    Dim EntityBodies As Physics.BodyList
    Dim NotPhysicBodies As Physics.BodyList
    
    If Me.Enabled Then
        If varTimer.value > Me.Delay Then
            Dim eA As Body, eB As Body
            Dim bA As Collider, bB As Collider
            
            ' Obtenemos las listas con los cuerpos activos dentro del area de simulacion, la lista de entidades activas
            ' y la lista de cuerpos fijos:
            Call GetActiveBodies(ActiveBodies, EntityBodies, FixedBodies, NotPhysicBodies)
            
            ' Completamos los valores que mostrara el monitor de actividad del simulador:
            With varMonitor
                .Actives = ActiveBodies.Count
                .Fixed = FixedBodies.Count
                .Entities = EntityBodies.Count
                .NotPhysical = NotPhysicBodies.Count
                .Total = Bodies.Count
            End With
            
            For Each eA In EntityBodies ' Recorremos la lista de cuerpos que sean entidades y no esten fijos.
                eA.Hit = False ' Reiniciamos la variable de la propiedad que indica colision.
                    
                ' Actualizamos las posiciones de todas las entidades que no sean fijas y que esten habilitadas:
                Call eA.SetLocation(eA.Location.X + (Me.Gravity.X * eA.Weight) + eA.Force.X, eA.Location.Y + (Me.Gravity.Y * eA.Weight) + eA.Force.Y)
                
                ' Calculamos las colisiones con el resto de cuerpos activos:
                For Each eB In ActiveBodies
                    ' Si la entidad A no es igual a la entidad B:
                    If Not eA.Key = eB.Key Then
                        
                        ' Si A colisiona con B:
                        If Core.Math.IntersectRect(eA.Rect, eB.Rect) Then
                            ' Corrije la posicion de A respecto a la de B si A no es una entidad fija:
                            If Not eA.PhysicType = NotPhysicalEntity Then
                                For Each bA In eA.Colliders
                                    For Each bB In eB.Colliders
                                        '   0 - Suelo
                                        '   1 - Pared derecha
                                        '   2 - Pared izquierda
                                        '   3 - Techo
                                        ' Agregamos la referencia de la entidad a la lista para devolverla en el evento:
                                        ReDim Preserve arrayE(aECount)
                                        Set arrayE(aECount).Body = eB  ' Entidad
                                        
                                        If Not (eA.PhysicType = PhysicalEntity And eB.PhysicType = PhysicalEntity) Then
                                            If Core.Math.IntersectRect(bA.Rect, bB.Rect) Then
                                                ColliderID = -1
                                                ' Estados predefinidos para reubicacion de objetos por colision:
                                                
                                                ' *** Aqui se deberia implementar la fuerza de respuesta con otro objetos ***
                                                ' *** La fuerza de respuesta se aplica solo cuando el objeto no este colisionando ***
                                                ' La formula aplicada es: If Not eB.Hit And eB.Weight <= eA.Weight Then eb.SetForce()
                                                
                                                If bA.TypeBox = 0 And bB.TypeBox = 3 Then
                                                    Call eA.SetLocation(eA.Location.X, bB.Rect.Y - (eA.Rect.Height / 2) - 1)
                                                    ColliderID = 0
                                                    eA.Hit = True
                                                    ' Rebajamos la fuerza para evitar atravesar superficies finas:
                                                    If Not eB.Hit And eB.Weight <= eA.Weight Then Call eB.SetForce(eA.GetForce() / 3, eA.Direction)
                                            
                                                ElseIf bA.TypeBox = 1 And bB.TypeBox = 2 Then
                                                    Call eA.SetLocation(bB.Rect.X - (eA.Rect.Width / 2) - 1, eA.Location.Y)
                                                    ColliderID = 2
                                                    Call eA.SetForceX(eA.Force.X * -1 / 1.5)
                                                    eA.Hit = True
                                                    If Not eB.Hit And eB.Weight <= eA.Weight Then Call eB.SetForce(eA.GetForce() / 1.5, eA.Direction)
                            
                                                ElseIf bA.TypeBox = 2 And bB.TypeBox = 1 Then
                                                    Call eA.SetLocation((bB.Rect.X + bB.Rect.Width) + (eA.Rect.Width / 2) + 1, eA.Location.Y)
                                                    ColliderID = 1
                                                    Call eA.SetForceX(eA.Force.X * -1 / 1.5)
                                                    eA.Hit = True
                                                    If Not eB.Hit And eB.Weight <= eA.Weight Then Call eB.SetForce(eA.GetForce() / 1.5, eA.Direction)
                                                
                                                ElseIf bA.TypeBox = 3 And bB.TypeBox = 0 Then
                                                    Call eA.SetLocation(eA.Location.X, (bB.Rect.Y + bB.Rect.Height) + (eA.Rect.Height / 2) + 1)
                                                    ColliderID = 3
                                                    eA.Hit = True
                                                    If Not eB.Hit And eB.Weight <= eA.Weight Then Call eB.SetForce(eA.GetForce() / 1.5, eA.Direction)
                                                
                                                ElseIf bB.TypeBox > 3 Then ' Si no coincide ningun estado simplemente devolvemos el codigo del colisionador:
                                                    ColliderID = bB.TypeBox
                                                    
                                                End If
                                                
                                                If ColliderID > -1 Then
                                                    ' Agregamos el codigo de la caja de colision de la entidad a la lista para devolverla en el evento:
                                                    ReDim arrayE(aECount).Colliders(arrayE(aECount).ColliderCount)
                                                    arrayE(aECount).Colliders(arrayE(aECount).ColliderCount) = ColliderID   ' Codigo del tipo de colisionador de B.
                                                    arrayE(aECount).ColliderCount = arrayE(aECount).ColliderCount + 1       ' Incrementamos el contador.
                                                End If
                                            End If
                                        End If
                                    Next
                                Next
                                aECount = aECount + 1
                            End If
                        End If
                    End If
                Next
                ' Si hubo colisiones se invoca el evento de la entidad y se le pasa la lista de entidades que han colisionado y las coordenadas del desplazamiento de correccion:
                If aECount > 0 Then
                    Call eA.RaiseOnCollision(aECount, arrayE)
                    ' Reiniciamos los valores de las variables temporales:
                    Erase arrayE()
                    aECount = 0
                    Correction.X = 0: Correction.Y = 0
                End If
                
                Call eA.Update ' Actualiza los datos internos de la entidad.
                'Debug.Print " "
'            Else
'                'eA.Enabled = False
'                'eA.OutOfWorkArea = False
'                'Debug.Print "Fuera de area!"
'            End If
            Next
            Call varTimer.Reset
        End If
    End If
End Sub

' DEBUG: Representa en pantalla todas las cajas de la lista:
Public Sub Draw()
    Dim ite As Physics.Body
    For Each ite In Bodies
        Call ite.Draw
    Next
End Sub

' Crea un trazador de rayos asociado a este simulador:
Public Function CreateRayTracer(Source As Core.POINT) As Physics.Ray
    Dim r As New Physics.Ray
    Set r.FatherSim = Me
    r.Source = Source
    Set CreateRayTracer = r
    Set r = Nothing
End Function

' Crea un emisor de explosiones asociado a este simulador:
Public Function CreateExplosionEmitter(Location As Core.POINT, Force As Long, Radius As Long) As Physics.Explosion
    Dim exp As New Physics.Explosion
    Set exp.FatherSim = Me
    exp.Location = Core.Generics.POINT2VECTOR(Location)
    exp.Force = Force
    exp.Radius = Radius
    
    Set CreateExplosionEmitter = exp
    Set exp = Nothing
End Function

Private Sub Class_Initialize()
    Set varTimer = New Core.Timer
    Set Me.Bodies.FatherSim = Me
    Set varMonitor = New Physics.SimulatorMonitor
End Sub

Private Sub Class_Terminate()
    Set varTimer = Nothing
End Sub
