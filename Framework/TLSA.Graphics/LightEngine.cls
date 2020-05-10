VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LightEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Scene - Core.RECT - Diemnsiones de toda la escena que se utilizara con sprites.
' WorkArea - Core.RECT - Dimensiones del area visible de la escena.
' SpriteList - Collection - Lista con todos los sprites de la escena que pueden iluminarse.
' Cuadrants - Collection - Lista de cuadrantes. Cada cuadrante es una lista con los sprites que estan dentro de su area. Cada cuadrante tendra el tamaño de WorkArea.
