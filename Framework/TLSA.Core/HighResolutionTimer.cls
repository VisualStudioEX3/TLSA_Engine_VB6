VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HighResolutionTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private varID As Long

Private Sub Class_Initialize()
    varID = sys.HITIMER_Create()
    If varID = Core.CORE_EXCEPTIONS.HIGH_RESOLUTION_TIMERS_NOT_SUPPORT Then
        Call Err.Raise(Core.CORE_EXCEPTIONS.HIGH_RESOLUTION_TIMERS_NOT_SUPPORT, , "El sistema no tiene soporte para cronometros de alta resolucion.")
    End If
End Sub

Public Property Get Value() As Currency
    Value = sys.HITIMER_GetValue(varID)
End Property

Public Sub Reset()
    Call sys.HITIMER_Reset(varID)
End Sub

Private Sub Class_Terminate()
    Call sys.HITIMER_Kill(varID)
End Sub
