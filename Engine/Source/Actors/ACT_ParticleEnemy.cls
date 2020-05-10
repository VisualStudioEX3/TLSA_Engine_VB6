VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ACT_ParticleEnemy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Campus Party Valencia 2010, Competicion rapida de juegos

Option Explicit

Private Const PHYSIC_ENTITY_TYPE As Long = 69
Private Const STEP As Long = 3
Private Const TIME_STEP As Long = 50

Private varTextureFile As String

Private varTexture As Graphics.Texture

Private Particle As Graphics.Sprite

Private WithEvents BBox As Physics.Body
Attribute BBox.VB_VarHelpID = -1

Public Visible As Boolean
Public Enabled As Boolean

Public Enum BehaviorMode
    StaticMode
    VerticalPath
    HorizontalPath
    CircularPath
    RandomPath
End Enum
    
Public Behavior As TLSA.BehaviorMode
Public Reborn As Boolean                    ' Indica si la particula resucita despues de muerta.
Public TriggerEventDuration As Long         ' Indica el tiempo que mantiene activas los estados de las plataformas asociadas a la figura de la particula al ser esta destruida.
Public varParticleType As Long              ' Codigo de la figura asociada a la particula y que se utilizara para asociar las plataformas al evento que disparara la particula al ser destruida.

Public Direction As Single

Private timerBehavior As Core.Timer
Private orientation As Long                 ' 0 o -1 : Direccion que tomara la entidad segun comportamiento (horizontal y vertical).

Private rotAxis As Core.Point               ' Se utilizara estas coordenadas para calcular la rotacion de las coordenadas de la entidad en el comportamiento de RotationPath.
Private rotRadius As Long
Private rotAngle As Single

Public Property Get Location() As Core.Point
    Location = Core.Generics.VECTOR2POINT(BBox.Location)
End Property

Public Property Let Location(value As Core.Point)
    BBox.Location = Core.Generics.POINT2VECTOR(value)
    If Behavior = 0 Then Particle.Location = Core.Generics.VECTOR2POINT(BBox.Location)
End Property

Public Property Get ParticleType() As Long
    ParticleType = varParticleType
End Property

Public Property Let ParticleType(value As Long)
    varParticleType = value
    Call Particle.SetCurrentAnimation(VBA.Mid("abcde", value + 1, 1))
    Call Particle.Update
End Property

Private Sub Class_Initialize()
    Set timerBehavior = New Core.Timer
    
    varTextureFile = App.Path & TLSA.ResourcePaths.Entities & "Particle\particulas.png"
    ' Cargamos la textura, si no esta cargada ya, y la asignamos:
    If Not GraphicEngine.Textures.Exists("particles") Then _
        Call Engine.GraphicEngine.Textures.LoadTexture(varTextureFile, "particles", False)
    
    Set Particle = Engine.GraphicEngine.Helper.CreateAnimatedSprite(Engine.GraphicEngine.Textures("particles"))
    Call Particle.SetCurrentAnimation("a")
    Call Particle.Update
    
    ' Creamos la instancia del actor en el simulador de fisicas:
    Set BBox = Engine.Scene.PhysicSimulator.Bodies.Add(0, 0, 0, 32, 32, Manual, 0, False, 16) ' PHYSIC_ENTITY_TYPE)
    Call BBox.SetColliders(AllEntireArea, 16)
    BBox.PhysicType = PhysicalEntity
    BBox.Color = Graphics.Color_Constant.DarkPurple
    
    Visible = True
    Enabled = True
    orientation = 1
End Sub

Private Sub Class_Terminate()
    Set timerBehavior = Nothing
    If Not Scene Is Nothing Then Call Scene.PhysicSimulator.Bodies.Remove(BBox)
    
    On Error Resume Next ' Evitamos la excepcion en caso de que la textura siga en uso:
    Call Engine.GraphicEngine.Textures.Unload("particles")
End Sub

Public Sub Draw()
    If Visible Then
        'Call BBox.Draw          ' Dibujamos el cuerpo fisico de la entidad para representarlo en pantalla.
        Call Particle.Draw
    End If
End Sub

Public Sub Update()
    If Enabled Then
        ' Comportamientos de movimiento:
        If timerBehavior.value >= TIME_STEP Then
            Select Case Behavior
                Case BehaviorMode.HorizontalPath
                    BBox.Location = Core.Generics.CreateVECTOR(BBox.Location.X + STEP * orientation, BBox.Location.Y, 0)
                Case BehaviorMode.VerticalPath
                    BBox.Location = Core.Generics.CreateVECTOR(BBox.Location.X, BBox.Location.Y + STEP * orientation, 0)
                Case BehaviorMode.CircularPath
                    BBox.Location = Core.Generics.POINT2VECTOR(Core.Math.TranslatePoint(Core.Generics.VECTOR2POINT(BBox.Location), rotRadius, rotAngle))
                    rotAngle = rotAngle + 0.5: If rotAngle >= 360 Then rotAngle = 0
                Case BehaviorMode.RandomPath
                    Call Core.Math.Random.SetSeed(VBA.Timer())
                    rotAngle = Core.Math.Random.GetRandomNumber2(359.99, 0)
                    BBox.Location = Core.Generics.POINT2VECTOR(Core.Math.TranslatePoint(Core.Generics.VECTOR2POINT(BBox.Location), rotRadius, rotAngle))
            End Select
            Particle.Location = Core.Generics.VECTOR2POINT(BBox.Location)
            Call timerBehavior.Reset
        End If
        Call Particle.Update
    End If
End Sub

' Crea una copia de la instancia:
Public Function Clone() As TLSA.ACT_ParticleEnemy
    Dim E As New TLSA.ACT_ParticleEnemy
    E.Behavior = Me.Behavior
    E.ParticleType = Me.ParticleType
    E.Enabled = Me.Enabled
    E.Location = Me.Location
    E.Reborn = Me.Reborn
    E.TriggerEventDuration = E.TriggerEventDuration
    E.Visible = Me.Visible
    
    Set Clone = E
    Set E = Nothing
End Function

Public Function ToString() As String
    ' <Key>:X;Y;Behavior;ParticleType;Reborn;TriggerEventDuration
    ToString = "ParticleEnemy=" & Me.Location.X & ";" & Me.Location.Y & ";" & Me.Behavior & ";" & Me.ParticleType & ";" & Me.Reborn & ";" & Me.TriggerEventDuration
End Function

Private Sub BBox_OnCollision(BodyCount As Long, E() As Physics.CollisionEventData)
    ' Alterna la direccion de la entidad:
    If Behavior = HorizontalPath Or VerticalPath Then orientation = orientation * -1
End Sub

Private Sub BBox_OnHit(Force As Single, eventData As Variant)
    ' ...
End Sub
