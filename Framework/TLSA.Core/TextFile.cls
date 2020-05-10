VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TextFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private varBuffer As Collection
Private varFilename As String

Friend Sub LoadFile(Filename As String)

End Sub

Public Function ReadLine(Line As Long) As String
    If varBuffer.Count > 0 Then ReadLine = varBuffer(Line + 1)
End Function

Public Function WriteLine(Line As Long, Text As String)

End Function

Public Sub AppendLine(Text As String)
    Call varBuffer.Add(Text)
End Sub

Public Sub InsertLine(Line As Long, Text As String)
    Dim before As Long, after As Long
    If (Line + 1) = 0 Then
        before = 0
    ElseIf (Line + 1) = varBuffer.Count Then
        before = varBuffer.Count
    Else
    
    End If
    Call varBuffer.Add(Text, , before, after)
End Sub

Public Sub RemoveLine(Line As Long)
    Call varBuffer.Remove(Line)
End Sub

Public Sub Clear()
    Set varBuffer = New Collection
End Sub

Public Function GetEntireText() As String
    Dim buffer As String
    If varBuffer.Count > 0 Then
        Dim i As Long
        For i = 1 To varBuffer.Count
            buffer = buffer & varBuffer(i) & IIf(i < varLines - 1, vbNewLine, "")
        Next
        GetEntireText = buffer
    End If
End Function

Public Property Get Lines() As Long
    Lines = varBuffer.Count
End Property
