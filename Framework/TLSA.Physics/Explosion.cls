VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Explosion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Este objeto ejerce o aplica una cantidad de fuerza a todos los objetos fisicos que esten definidos como entidad y que esten dentro
' de su radio de onda expansiva y segun cercania al centro de la explosion.

Private varRadius As Long
Private varForce As Long
Private varLocation As Core.VECTOR
Private varFatherSim As Physics.Simulator

Public Property Get Radius() As Long
    Radius = varRadius
End Property

Public Property Let Radius(value As Long)
    varRadius = value
End Property

Public Property Get Force() As Single
    Force = varForce
End Property

Public Property Let Force(value As Single)
    varForce = value
End Property

Public Property Get Location() As Core.VECTOR
    Location = varLocation
End Property

Public Property Let Location(value As Core.VECTOR)
    varLocation = value
End Property

Public Sub SetLocation(X As Single, Y As Single, Z As Single)
    varLocation.X = X
    varLocation.Y = Y
    varLocation.Z = Z
End Sub

Friend Property Get FatherSim() As Physics.Simulator
    Set FatherSim = varFatherSim
End Property

Friend Property Set FatherSim(sim As Physics.Simulator)
    Set varFatherSim = sim
End Property

' Aplica la explosion con los parametros definidos:
Public Sub Explode()
    Dim distance As Long
    Dim Direction As Single
    Dim response As Single
    Dim b As Physics.Body
    Dim r As Physics.Ray
    
    ' Creamos un rayo para trazar trayectorias desde el epicentro a los posibles objetivos dentro del area de la onda expansiva:
    Set r = varFatherSim.CreateRayTracer(Core.Generics.NullPOINT)
    
    For Each b In varFatherSim.Bodies
        ' Si el objeto es una entidad fisica y no esta fija o desactivada:
        If b.Enabled And b.PhysicType = PhysicalEntity And Not b.Fixed Then
            ' Comprobamos pues si esta dentro del radio de la explosion:
            distance = Core.Math.GetDistance(Core.Generics.VECTOR2POINT(varLocation), Core.Generics.VECTOR2POINT(b.Location))
            If distance <= varRadius Then
                ' Comprobamos que no haya ningun obstaculo. Trazamos un segmento desde la explosion hacia el centro del objeto,
                ' y comprobamos que el objeto mas cercano es el analizado:
                If r.TraceSegment(Core.Generics.VECTOR2POINT(varLocation), Core.Generics.VECTOR2POINT(b.Location), Nothing, modPhysics.varHelper.EmptyZDiscardArray()) Is b Or _
                   Core.Math.IntersectRect(Core.Generics.CreateRECTANGLE(varLocation.X - varRadius \ 2, varLocation.Y - varRadius \ 2, Radius, Radius), b.Rect) Then
                    ' Si esta dentro del radio obtenemos el angulo desde el centro de la explision al centro del objeto
                    ' y hallamos el porcentaje de la fuerza que aplicaremos por la onda expansiva:
                    Direction = Core.Math.GetAngle(Core.Generics.VECTOR2POINT(varLocation), Core.Generics.VECTOR2POINT(b.Location))

                    ' Obtenemos la fuerza de respuesta segun distancia hacia el epicentro de la explosion:
                    If varForce > 0 Then
                        response = Core.Math.Invert(Core.Math.GetValueFromPercent(distance - varRadius, varForce))

                        ' Aplicamos la fuerza al objeto:
                        Call b.SetForce(response, Direction, True)
                    End If
                    
                    ' Invocamos el evento OnHit() del cuerpo.
                    ' Devolvemos la instancia de la explosion como infomarcion del evento:
                    Dim eventData As Variant: Set eventData = Me
                    Call b.RaiseOnHitEvent(response, eventData)
                End If
            End If
        End If
    Next
End Sub
