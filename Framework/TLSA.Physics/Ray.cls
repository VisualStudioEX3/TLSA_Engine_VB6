VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Ray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Codigo original de la funcion IntersectLine() encontrado en:
' http://workshop.evolutionzone.com/2007/09/10/code-2d-line-intersection/
'
' Creditos originales:
' --------------------------------------------------------------------------
' lineIntersect.pde
' Marius Watz - http://workshop.evolutionzone.com
'
' calculates valid intersection between two lines,
' so that the intersection will lie on the specified
' line segment.
'
' --------------------------------------------------------------------------
' Codigo traducido de C a Visual Basic 6.0 por
' José Miguel Sánchez Fernández

Private varSource As Core.POINT
Private varTarget As Core.POINT
Private varHit As Core.POINT
Private varDir As Single
Private varFatherSim As Physics.Simulator

Friend Property Get FatherSim() As Physics.Simulator
    Set FatherSim = varFatherSim
End Property

Friend Property Set FatherSim(sim As Physics.Simulator)
    Set varFatherSim = sim
End Property

' Angulo que define la direccion del rayo:
Public Property Get Direction() As Single
    Direction = varDir
End Property

Public Property Let Direction(value As Single)
    varDir = value
End Property

' Punto de origen del rayo a trazar:
Public Property Get Source() As Core.POINT
    Source = varSource
End Property

Public Property Let Source(value As Core.POINT)
    varSource = value
End Property

Public Sub SetSource(ByVal X As Long, ByVal Y As Long)
    varSource.X = X
    varSource.Y = Y
End Sub

' Punto de destino u objetivo hacia donde se trazara el rayo:
Public Property Get Target() As Core.POINT
    Target = varTarget
End Property

' Punto de corte del rayo:
Public Property Get Hit() As Core.POINT
    Hit = varHit
End Property

' Comprueba si un valor se encuentra en la lista:
Private Function IsDiscard(value As Variant, ZDiscard() As Long) As Boolean
    On Error GoTo ErrOut
    Dim i As Long
    For i = 0 To UBound(ZDiscard)
        If ZDiscard(i) = value Then
            IsDiscard = True
            Exit Function
        End If
    Next
ErrOut:
End Function

' Traza un segmento desde el punto origen hasta el punto destino y devuelve las coordenadas del punto de corte:
' El parametro SourceBody se utiliza para descartar las aristas de la caja del cuerpo que proyecta el segmento
' en el calculo del punto de corte mas cercano al origen. Si el origen de la proyeccion se realiza fuera de un
' cuerpo (desde vacio) el parametro debe ser Nothing. Si hay punto de corte se retorna la referencia al cuerpo
' intersectado:
Private Function TraceToTarget(SourceBody As Physics.Body, ZDiscard() As Long) As Physics.Body
    Dim pA As Physics.PHXVECTOR, pB As Physics.PHXVECTOR, pC As Physics.PHXVECTOR, pD As Physics.PHXVECTOR

    ' 1. Determinar el area rectangular donde se proyecta el rayo:
    pA.X = varSource.X: pA.Y = varSource.Y
    pB.X = varTarget.X: pB.Y = varTarget.Y

    If pA.X > pB.X Then
        pC.X = pA.X: pA.X = pB.X: pB.X = pC.X
    End If
    If pA.Y > pB.Y Then
        pC.Y = pA.Y: pA.Y = pB.Y: pB.Y = pC.Y
    End If
    
    Dim r As Core.RECTANGLE
    r.X = pA.X
    r.Y = pA.Y
    r.Width = pB.X - pA.X
    r.Height = pB.Y - pA.Y
    
    ' 2. Recorrer la lista de cuerpos del simulador asociado y obtener cuantas cajas intersectan con el area rectangular
    ' donde se proyecta el rayo:
    Dim boxes As New Collection, b As Physics.Body, sourceKey As String
    If Not SourceBody Is Nothing Then sourceKey = SourceBody.Key ' Si sourceBody es Nothing la clave estara vacia y no descartara ningun cuerpo.
    For Each b In varFatherSim.Bodies
        'If Not b Is Nothing Then   ' For Each devuelve un iterador nulo?
            If Not b.Key = sourceKey And b.Enabled And Not b.PhysicType = NotPhysicalEntity Then
                If Not IsDiscard(b.Z, ZDiscard()) Then ' Si no se descarta...
                    If Core.Math.IntersectRect(r, b.Rect) Then
                        Call boxes.Add(b)
                    End If
                End If
            End If
        'End If
    Next
    
    ' 3. Calculamos todos los puntos de corte del rayo con los segmentos que conforman las cajas de los cuerpos de la lista:
    Dim pts() As Physics.PHXVECTOR, pCount As Integer
    Dim ret As Core.VECTOR
    Dim i As Integer, j As Integer
    
    pA.X = varSource.X: pA.Y = varSource.Y: pB.X = varTarget.X: pB.Y = varTarget.Y
    For Each b In boxes
        ' Segmentos de la caja:
        For i = 0 To 3
            Select Case i
                Case 0: pC.X = b.Rect.X: pC.Y = b.Rect.Y: pD.X = b.Rect.X + b.Rect.Width: pD.Y = b.Rect.Y
                Case 1: pC.X = b.Rect.X: pC.Y = b.Rect.Y: pD.X = b.Rect.X: pD.Y = b.Rect.Y + b.Rect.Height
                Case 2: pC.X = b.Rect.X + b.Rect.Width: pC.Y = b.Rect.Y: pD.X = b.Rect.X + b.Rect.Width: pD.Y = b.Rect.Y + b.Rect.Height
                Case 3: pC.X = b.Rect.X: pC.Y = b.Rect.Y + b.Rect.Height: pD.X = b.Rect.X + b.Rect.Width: pD.Y = b.Rect.Y + b.Rect.Height
            End Select
    
            ' Si intersectan los dos segmentos entonces añadimos el punto de corte a la lista:
            If Core.Math.IntersectLine(modPhysics.PHXVECTOR2VECTOR(pA), modPhysics.PHXVECTOR2VECTOR(pB), modPhysics.PHXVECTOR2VECTOR(pC), modPhysics.PHXVECTOR2VECTOR(pD), ret) Then
                ReDim Preserve pts(pCount)
                pts(pCount) = modPhysics.VECTOR2PHXVECTOR(ret)
                pts(pCount).D = Core.Math.GetDistance(Core.Generics.VECTOR2POINT(ret), varSource)
                Set pts(pCount).b = b ' Agregamos la referencia del cuerpo del simulador.
                pCount = pCount + 1
            End If
        Next
    Next
    
    ' 4. Determinamos cual punto de corte de la lista esta mas proximo al punto de origen del rayo:
    If pCount > 0 Then
        Dim near As Physics.PHXVECTOR
        near = GetNear(pts()) ' Obtenemos el punto mas cercano.
        varHit = modPhysics.PHXVECTOR2POINT(near) ' Devolvemos las coordenadas de corte mas cercanas.
        Set TraceToTarget = near.b ' Devolvemos la rerferencia del cuerpo mas cercano.
    End If
End Function

' Obtiene el vector mas cercano al origen:
Private Function GetNear(list() As Physics.PHXVECTOR) As Physics.PHXVECTOR
    Dim near As Physics.PHXVECTOR
    Dim i As Long
    
    near = list(0)
    For i = 0 To UBound(list)
        If list(i).D < near.D Then near = list(i)
    Next
    
    GetNear = near
End Function

' Traza un segmento entre un punto origen y un punto destino y devuelve el punto de corte mas cercano al punto origen:
Public Function TraceSegment(Source As Core.POINT, Target As Core.POINT, SourceBody As Physics.Body, ZDiscard() As Long) As Physics.Body
    Dim remSource As Core.POINT, remTarget As Core.POINT
    
    ' Guardamos los valores actuales y aplicamos temporalmente los nuevos:
    remSource = varSource: varSource = Source
    remTarget = varTarget: varTarget = Target
    
    ' Trazamos el rayo:
    Set TraceSegment = TraceToTarget(SourceBody, ZDiscard)
    
    ' Restauramos los valores previos:
    varSource = remSource
    varTarget = remTarget
End Function

Public Function Trace(SourceBody As Physics.Body, ZDiscard() As Long) As Physics.Body
    ' Obtenemos el punto de destino proyectando la linea en un angulo con una longitud previa:
    varTarget = Core.Math.TranslatePoint(Source, Core.Math.MIN_INT_VALUE, varDir)

    ' Obtenemos el punto de corte con la caja que define el area de calculo del simulador para acotar lo
    ' maximo posible el area de calculo:
    Dim pA As Physics.PHXVECTOR, pB As Physics.PHXVECTOR, pC As Physics.PHXVECTOR, pD As Physics.PHXVECTOR, ret As Core.VECTOR ' Physics.PHXVECTOR
    Dim pts(3) As Physics.PHXVECTOR

    With FatherSim.WorkArea
        pA.X = Source.X: pA.Y = Source.Y: pB.X = Target.X: pB.Y = Target.Y

        Dim i As Integer
        For i = 0 To 3
            Select Case i
                Case 0: pC.X = .X: pC.Y = .Y: pD.X = .X + .Width: pD.Y = .Y
                Case 1: pC.X = .X: pC.Y = .Y: pD.X = .X: pD.Y = .Y + .Height
                Case 2: pC.X = .X + .Width: pC.Y = .Y: pD.X = .X + .Width: pD.Y = .Y + .Height
                Case 3: pC.X = .X: pC.Y = .Y + .Height: pD.X = .X + .Width: pD.Y = .Y + .Height
            End Select

            ' Si intersectan los dos segmentos entonces añadimos el punto de corte a la lista:
            If Core.Math.IntersectLine(modPhysics.PHXVECTOR2VECTOR(pA), modPhysics.PHXVECTOR2VECTOR(pB), modPhysics.PHXVECTOR2VECTOR(pC), modPhysics.PHXVECTOR2VECTOR(pD), ret) Then
                pts(i) = modPhysics.VECTOR2PHXVECTOR(ret)
                pts(i).D = Core.Math.GetDistance(Core.Generics.VECTOR2POINT(ret), Source)
            Else
                pts(i).D = Core.Math.MAX_INT_VALUE
            End If
        Next
    End With

    ' Obtenemos el punto mas cercano:
    varTarget = modPhysics.PHXVECTOR2POINT(GetNear(pts()))

    ' Trazamos un segmento para calcular la trayectoria y el punto de corte:
    Set Trace = TraceToTarget(SourceBody, ZDiscard) ' Me.TraceSegment(Source, varTarget, ZDiscard)

End Function

