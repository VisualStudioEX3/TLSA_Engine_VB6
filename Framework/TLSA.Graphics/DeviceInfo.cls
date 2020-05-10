VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DeviceInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Clase para obtener la informacion del disposivito grafico:

Option Explicit

Dim info As dxlib32_221.Device_Info
'Dim dxVersion As Long, dxRevision As Long
Dim dxVersion As Core.VersionValue
Dim memTotal As Long, memFree As Long

Public Property Get AdapterName() As String
    AdapterName = info.AdapterName
End Property

Public Property Get DriverName() As String
    DriverName = info.DriverName
End Property

Public Property Get DriverVersion() As Core.VersionValue
    Dim v As Core.VersionValue
    v.High = info.DriverVersionHigh
    v.Low = info.DriverVersionLow
    v.Revision = info.Revision
    DriverVersion = v
End Property

'Public Property Get DriverVersionHigh() As Long
'    DriverVersionHigh = info.DriverVersionHigh
'End Property
'
'Public Property Get DriverVersionLow() As Long
'    DriverVersionLow = info.DriverVersionLow
'End Property

Public Property Get MaxBitsPerPixel() As Long
    MaxBitsPerPixel = info.MaxBitsPerPixel
End Property

Public Property Get MaxTextureHeight() As Long
    MaxTextureHeight = info.MaxTextureHeight
End Property

Public Property Get MaxTextureWidth() As Long
    MaxTextureWidth = info.MaxTextureWidth
End Property

Public Property Get Revision() As Long
    Revision = info.Revision
End Property

Public Property Get VendorID() As Long
    VendorID = info.VendorID
End Property

'Public Property Get DirectXVersion() As Long
'    DirectXVersion = dxVersion
'End Property
'
'Public Property Get DirectXRevision() As Long
'    DirectXRevision = dxRevision
'End Property

Public Property Get DirectXVersion() As Core.VersionValue
    DirectXVersion = Core.Information.DirectXVersion
End Property

Public Property Get TotalVideoMemory() As Long
    TotalVideoMemory = memTotal
End Property

Public Property Get FreeVideoMemory() As Long
    FreeVideoMemory = memFree
End Property

Private Sub Class_Initialize()
    Call gfx.DEVICE_GetInfo(info)
    'Call sys.SYS_GetDirectXVersion(dxVersion, dxRevision)
    memTotal = gfx.DEVICE_GetTotalVideoMem
    memFree = gfx.DEVICE_GetFreeVideoMem
End Sub

' Comprueba si un modo de video esta soportado por el dispositivo grafico:
Public Function ExistDisplayMode(Width As Long, Height As Long, Bpp As Long) As Boolean
    ExistDisplayMode = gfx.DEVICE_ExistsDisplayMode(Width, Height, Bpp)
End Function

' Devuelve una lista con todos los modos de video disponibles:
Public Function GetDisplayModes() As Graphics.DisplayMode()
    Dim Count As Long
    Dim data As Display_Mode
    Dim list() As Graphics.DisplayMode
    Dim i As Long
    
    Count = gfx.DEVICE_GetDisplayModeCount()
    ReDim list(Count - 1)
    For i = 0 To (Count - 1)
        Call gfx.DEVICE_GetDisplayMode(i, data)
        With list(i)
            .Width = data.Width
            .Height = data.Height
            .Bpp = data.Bpp
        End With
    Next
    
    GetDisplayModes = list
End Function
