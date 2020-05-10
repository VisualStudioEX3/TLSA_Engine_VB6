VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Text"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Obtiene el texto desde la posicion indicada hasta encontrar la primera coincidencia desde el principio de la cadena:
Public Function GetFromLeft(Text As String, Match As String, Optional Start As Long = 1) As String
    GetFromLeft = VBA.Left$(Text, VBA.InStr(Start, Text, Match) - 1)
    'GetFromLeft = VBA.Left$(Text, VBA.InStrRev(Text, Match, Start, vbTextCompare))
End Function

' Obtiene el texto desde la posicion indicada hasta encontrar la primera coincidencia desde el final de la cadena:
Public Function GetFromRight(Text As String, Match As String, Optional Start As Long = -1) As String '-1
    GetFromRight = VBA.Right$(Text, Len(Text) - VBA.InStrRev(Text, Match, Start))
'    GetFromRight = VBA.Right$(Text, VBA.InStr(Start, Text, Match, vbTextCompare) - 1)
End Function

' Convierte un array de strings en un string con saltos de linea:
Public Function Array2String(StringArray() As String) As String
'    Dim ret As String, ite As Variant
'    For Each ite In StringArray
'        ret = ret & ite & vbNewLine
'    Next
'    Array2String = Left(ret, Len(ret) - Len(vbNewLine)) ' Devolvemos la cadena sin el ultimo salto de linea.
    Array2String = VBA.Join(StringArray(), vbNewLine)
End Function

' Comprueba si una cadena de texto empieza segun la cadena indicada indicada:
Public Function StartWith(Text As String, StartValue As String) As Boolean
    'If Len(Text) >= Len(StartValue) Then StartWith = (Left(Text, Len(StartValue)) = StartValue)
    StartWith = (VBA.InStr(1, Text, StartValue, vbTextCompare) > 0)
End Function

' Comprueba si una cadena de texto termina segun la cadena indicada indicada:
Public Function EndWith(Text As String, EndValue As String) As Boolean
    'If Len(Text) >= Len(EndValue) Then StartWith = (VBA.Right$(Text, Len(EndValue)) = EndValue)
    EndWith = (VBA.InStrRev(Text, EndValue, , vbTextCompare) > 0)
End Function

' Comprueba si el valor existe en la cadena y, si se desea, se devuelve la posicion donde se encuentra. Se puede realizar la busqueda en ambos sentidos:
Public Function Find(Text As String, Value As String, Optional Position As Long, Optional Inverse As Boolean = False) As Boolean
    If Inverse Then
        Position = VBA.InStrRev(Text, Value, , vbTextCompare)
    Else
        Position = VBA.InStr(1, Text, Value, vbTextCompare)
    End If
    Find = (Position > 0)
End Function
