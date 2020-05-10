VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OpenDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private varFilename As String
Private varTitle As String
Private varFilter As Collection, varFilterIndex As Long
Private varStartPath As String

' Devuelve el titulo del dialogo:
Public Property Get Title() As String
    Title = varTitle
End Property

' Establece el titulo del dialogo:
Public Property Let Title(value As String)
    varTitle = value
End Property

' Devuelve el nombre del archivo seleccionado:
Public Property Get Filename() As String
    Filename = varFilename
End Property

' Establece el valor del campo nombre de archivo al mostrar el dialogo:
Public Property Let Filename(value As String)
    varFilename = value
End Property

' Devuelve una cadena con los filtros agregados:
Public Property Get Filter() As String
    Dim ite As Variant, ret As String
    For Each ite In varFilter
        ret = ret & ite & "|"
    Next
    If Len(ret) > 1 Then Filter = Left(ret, Len(ret) - 1)
End Property

' Devuelve el filtro seleccionado por defecto:
Public Property Get FilterIndex() As Long
    FilterIndex = varFilterIndex
End Property

' Establece el filtro seleccionado por defecto:
Public Property Let FilterIndex(value As Long)
    If value < 1 Or value > varFilter.Count Then value = 1
    varFilterIndex = value
End Property

' Añade un filtro a la lista:
Public Sub AddFilter(Description As String, Extension As String)
    Dim f As String
    f = Description & "|" & Extension
    Call varFilter.Add(f, f)
End Sub

' Elimina un filtro de la lista:
Public Sub RemoveFilter(Index As Long)
    Call varFilter.Remove(Index)
End Sub

' Vacia la lista de filtros:
Public Sub ClearFilter()
    Set varFilter = New Collection
End Sub

' Devuelve el directorio que mostrara el dialogo al abrirse:
Public Property Get StartPath() As String
    StartPath = varStartPath
End Property

' Establece el directorio que mostrara el dialogo al abrirse:
Public Property Let StartPath(value As String)
    varStartPath = value
End Property

' Muestra el dialogo:
Public Function Show() As Boolean
    Filename = sys.DLG_OpenFile(0, Me.Filter, varTitle, varStartPath, varFilename, varFilterIndex - 1)
    Show = (Filename <> "")
End Function

Private Sub Class_Initialize()
    Set varFilter = New Collection
    varFilterIndex = 1
End Sub
