VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SimulatorMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private varActives As Long
Private varFixed As Long
Private varEntities As Long
Private varNotPhysical As Long
Private varTotal As Long                    ' Total de cuerpos en el simulador.

Public Property Get Fixed() As Long
    Fixed = varFixed
End Property

Public Property Get Entities() As Long
    Entities = varEntities
End Property

Public Property Get Actives() As Long
    Actives = varActives
End Property

Public Property Get NotPhysical() As Long
    NotPhysical = varNotPhysical
End Property

Public Property Let Fixed(value As Long)
    varFixed = value
End Property

Public Property Let Entities(value As Long)
    varEntities = value
End Property

Public Property Let Actives(value As Long)
    varActives = value
End Property

Public Property Let NotPhysical(value As Long)
    varNotPhysical = value
End Property


Public Property Get Total() As Long
    Total = varTotal
End Property

Public Property Let Total(value As Long)
    varTotal = value
End Property

Public Property Get InActives() As Long
    InActives = varTotal - varActives
End Property

' Devuelve un resumen con los valores de los elementos activos:
Public Function ToString() As String
    ToString = "Cuerpos fisicos en escena:" & vbNewLine & _
               " + Fijos:" & varFixed & vbNewLine & _
               " + Entidades: " & varEntities & vbNewLine & _
               " + No fisicos: " & varNotPhysical & vbNewLine & _
               " + Activos: " & varActives & vbNewLine & _
               " + Inactivos: " & (varTotal - varActives) & vbNewLine & _
               "TOTAL = " & varTotal
End Function
