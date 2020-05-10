VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DEBUG_MessagePool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim label As Graphics.TextString
Dim Buffer As String

Private Sub Class_Initialize()
    If Engine.EDIT_MODE Then
        Call frmOutput.Show
    Else
        Set label = New TextString
        label.Color = Graphics.Color_Constant.White
        label.Fixed = True
        Set label.Font = GraphicEngine.Fonts("system")
        label.Location = Core.Generics.CreatePOINT(0, 0, -8)
    End If
End Sub

Public Sub AddMessage(message As String)
    Buffer = Buffer & message & vbNewLine
End Sub

' Dibujamos
Public Sub Draw()
    If Engine.EDIT_MODE Then
        frmOutput.txtOutput.Text = Buffer
        frmOutput.txtOutput.SelStart = Len(frmOutput.txtOutput.Text)
    Else
        ' Escribimos el buffer de mensajes en pantalla:
        label.Text = VBA.Left(Buffer, Len(Buffer) - 2)      ' Eliminamos el ultimo salto de linea.
        ' Dibujamos un fondo semitransparente para facilitar la lectura del texto:
        Call GraphicEngine.Primitives.DrawBox(0, 0, 800, label.Height, -8, 0, True, &H77777777, True)
        Call label.Draw
    End If
    Buffer = ""
End Sub

Private Sub Class_Terminate()
    If Engine.EDIT_MODE Then
        Call Unload(frmOutput)
    Else
        Set label = Nothing
    End If
End Sub
