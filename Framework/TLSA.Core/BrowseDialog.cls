VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BrowseDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private varTitle As String
Private varPath As String

' Ruta del directorio seleccionado en el dialogo:
Public Property Get Path() As String
    Path = varPath
End Property

' Muestra el dialogo:
Public Function Show() As Boolean
    varPath = sys.DLG_BrowseFolder(0, varTitle)
    Show = (varPath <> "")
End Function
