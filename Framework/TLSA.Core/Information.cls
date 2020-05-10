VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Information"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Type VersionValue
    High As Long
    Low As Long
    Revision As Long
End Type

Public Type MemoryStatus
    Total As Long
    Free As Long
    Used As Long
End Type

Public Property Get DirectXVersion() As VersionValue
    Dim v As VersionValue
    Call sys.SYS_GetDirectXVersion(v.High, v.Low)
    DirectXVersion = v
End Property

Public Property Get MemoryStatus() As Core.MemoryStatus
    Dim sta As Core.MemoryStatus
    Dim ret As dxlib32_221.Memory_Info
    Call sys.SYS_GetMemory(ret)
    sta.Total = ret.TotalPhys
    sta.Free = ret.AvailPhys
    sta.Used = sta.Total - sta.Free
    MemoryStatus = sta
End Property
