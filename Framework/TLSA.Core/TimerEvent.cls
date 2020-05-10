VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TimerEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private id As Long

Public Sub SetEvent(Interval As Long, ProcessAddress As Long)
    id = sys.TIMER_CreateProcess(0, Interval, ProcessAddress)
End Sub

Private Sub Class_Terminate()
    Call sys.TIMER_KillProcess(0, id)
End Sub
