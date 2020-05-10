VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EDIT_Brush"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event Hit(X As Long, Y As Long)
Public Event AltHit(X As Long, Y As Long)

Private varSize As Core.Size
Private varLocation As Core.Point
Public Visible As Boolean

Public Property Get Location() As Core.Point
    Location = varLocation
End Property

Public Property Get Size() As Core.Size
    Size = varSize
End Property

Public Property Let Size(value As Core.Size)
    varSize = value
End Property

Public Sub Draw()
    Const Size As Long = 8
    If Visible Then
        ' Cursor que representa la posicion en coordenadas de la pantalla o area de la ventana:
        Call GraphicEngine.Primitives.DrawLine(dev_input.ViewAxis.X, dev_input.ViewAxis.Y - Size, dev_input.ViewAxis.X, dev_input.ViewAxis.Y + Size, -8, Graphics.Color_Constant.Red, True)
        Call GraphicEngine.Primitives.DrawLine(dev_input.ViewAxis.X - Size, dev_input.ViewAxis.Y, dev_input.ViewAxis.X + Size, dev_input.ViewAxis.Y, -8, Graphics.Color_Constant.Red, True)
        
'        ' Dibujamos el area del pincel:
'        Call GraphicEngine.Primitives.DrawBox3(Location.X, Location.Y, varSize.width, varSize.height, -8, Graphics.Color_Constant.Red, True, &H77FF0000, True)
    End If
End Sub

' Sistema de coordenadas fijo pero con dimension de pincel variable:
Public Sub Update()
    ' Fijamos la localizacion del pincel:
    varLocation = Engine.dev_input.ViewAxis()
    
    Dim i As Long
    
    If Location.X > 0 Then
        For i = 1 To (Engine.GraphicEngine.CurrentDisplayMode.width \ 8) + 8
            If (Location.X >= (8 * (i - 1)) And (Location.X <= (8 * i))) Then
                varLocation.X = (8 * (i - 1))
                Exit For
            End If
        Next
    Else
        varLocation.X = 0
    End If
    
    If Location.Y > 0 Then
        For i = 1 To (Engine.GraphicEngine.CurrentDisplayMode.height \ 8) + 8
            If (Location.Y >= (8 * (i - 1)) And (Location.Y <= (8 * i))) Then
                varLocation.Y = (8 * (i - 1))
                Exit For
            End If
        Next
    Else
        varLocation.Y = 0
    End If
    
    ' Invocamos el evento Hit() en caso de presionarse el raton y devolvemos las coordenadas absolutas del centro de la celda:
    Dim pt As Core.Point: pt = Engine.Scene.ClientToScreen(Location)     ' Convertimos las coordenadas de la ventana en coordenadas de escena.
    If Engine.dev_input.Hit("edit_set") Then
        RaiseEvent Hit(pt.X + (varSize.width \ 2), pt.Y + (varSize.height \ 2))
    
    ' El evento AltHit lo usaremos para crear un evento donde implementar la eliminacion de objetos:
    ElseIf Engine.dev_input.Hit("edit_rem") Then
        RaiseEvent AltHit(pt.X + (varSize.width \ 2), pt.Y + (varSize.height \ 2))
    End If
    
    ' Informacion de depuracion:
    Call Engine.dev_debug.AddMessage("Posicion del pincel: " & Location.X & "x " & Location.Y & "y")
End Sub

Private Sub Class_Initialize()
    Visible = True
End Sub

' Sistema de coordenadas y pincel variable:
'Public Sub Update2()
'    ' Fijamos la localizacion del pincel:
'    varLocation = Engine.dev_input.ViewAxis()
'
'    Dim i As Long
'
'    If Location.X > 0 Then
'        For i = 1 To (Engine.GraphicEngine.CurrentDisplayMode.width \ varSize.width) + varSize.width
'            If (Location.X >= (varSize.width * (i - 1)) And (Location.X <= (varSize.width * i))) Then
'                varLocation.X = (varSize.width * (i - 1))
'                Exit For
'            End If
'        Next
'    Else
'        varLocation.X = 0
'    End If
'
'    If Location.Y > 0 Then
'        For i = 1 To (Engine.GraphicEngine.CurrentDisplayMode.height \ varSize.height) + varSize.height
'            If (Location.Y >= (varSize.height * (i - 1)) And (Location.Y <= (varSize.height * i))) Then
'                varLocation.Y = (varSize.height * (i - 1))
'                Exit For
'            End If
'        Next
'    Else
'        varLocation.Y = 0
'    End If
'
'    ' Invocamos el evento Hit() en caso de presionarse el raton y devolvemos las coordenadas absolutas del centro de la celda:
'    Dim pt As Core.Point: pt = Engine.Scene.ClientToScreen(Location)     ' Convertimos las coordenadas de la ventana en coordenadas de escena.
'    If Engine.dev_input.Hit("edit_set") Then RaiseEvent Hit(pt.X + (varSize.width \ 2), pt.Y + (varSize.height \ 2))
'
'    ' Informacion de depuracion:
'    Call Engine.dev_debug.AddMessage("Posicion del pincel: " & Location.X & "x " & Location.Y & "y")
'End Sub

