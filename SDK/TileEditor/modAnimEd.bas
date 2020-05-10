Attribute VB_Name = "modAnimEd"
Option Explicit

Public Sub ShowData(Key As String)
    Dim sec As Graphics.Animation
    
    Set sec = Spr.Animations(Key)

    ' Mostramos las propiedades:
    frmQuickAnimEd.txtDelay.Text = sec.FrameDelay
    frmQuickAnimEd.txtKey.Text = sec.Key
    frmQuickAnimEd.txtKey.Tag = sec.Key
    frmQuickAnimEd.chkLoop.value = Abs(CInt(sec.Looping))
End Sub

' Cambia la clave de la secuencia:
Public Sub ChangeKey(OldKey As String, NewKey As String)
    ' Primero comprobamos que la clave no este en uso por otro tile:
    If Not Spr.Animations.Exists(NewKey) Then
        ' Creamos el tile en base al actual pero usando la nueva clave:
        Call Spr.Animations.Add(Spr.Animations(OldKey), NewKey)
        Call Spr.Animations.Remove(OldKey)

        Call LoadSecuences ' Actualizamos la lista.
    Else
        Call MsgBox("La clave ya existe en la lista.", vbCritical, "Clave duplicada")
    End If
End Sub

Public Sub LoadSecuences()
    Dim state As Boolean
    Dim i As Graphics.Animation
    
    Call frmQuickAnimEd.cbxAnimSec.Clear
    For Each i In Spr.Animations
        Call frmQuickAnimEd.cbxAnimSec.AddItem(i.Key)
    Next
    
    state = (frmQuickAnimEd.cbxAnimSec.ListCount > 0)
    
    If state Then
        frmQuickAnimEd.cbxAnimSec.ListIndex = 0
        frmQuickAnimEd.cmdPreview.Enabled = True
    End If
    
    frmQuickAnimEd.txtKey.Enabled = state
    frmQuickAnimEd.txtDelay.Enabled = state
    frmQuickAnimEd.cmdUpdate.Enabled = state
    frmQuickAnimEd.cmdDelete.Enabled = state
    frmQuickAnimEd.cmdPreview.Enabled = state
End Sub

Public Sub AddSecuence(Key As String, TileName As String, CounterFormat As String)
    Dim a As New Graphics.Animation
    Dim i As Long, value As String, counter As Long
    For i = 0 To frmMain.lstTiles.ListCount - 1
        value = frmMain.lstTiles.List(i)
        If value = TileName & Format(counter, CounterFormat) Then
            Call a.Tiles.AddRef(Tex.Tiles(value), value)
            counter = counter + 1
        End If
    Next

    Call Spr.Animations.Add(a, Key)
    Set a = Nothing
    
    Call MsgBox("Se agregaron " & counter & " tiles a la secuencia.", vbInformation, "Secuencia completada")
End Sub
