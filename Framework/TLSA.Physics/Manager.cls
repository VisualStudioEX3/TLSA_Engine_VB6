VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Manager"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"BoxList"
Option Explicit

Public Property Get DEBUG_DrawColliders() As Boolean
    DEBUG_DrawColliders = modPhysics.debugDrawColliders
End Property

Public Property Let DEBUG_DrawColliders(value As Boolean)
    modPhysics.debugDrawColliders = value
End Property

' Configura el motor de fisica para poder renderizar las cajas de colision:
Public Sub SetGraphics(Instance As Graphics.Manager)
    Set modPhysics.gfx = Instance
End Sub

' Acceso a la clase de funciones auxiliares:
Public Property Get Helper() As Physics.Helper
    Set Helper = modPhysics.varHelper
End Property
