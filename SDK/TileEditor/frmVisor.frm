VERSION 5.00
Begin VB.Form frmVisor 
   BackColor       =   &H8000000C&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Textura"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   MouseIcon       =   "frmVisor.frx":0000
   MousePointer    =   99  'Custom
   Moveable        =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
End
Attribute VB_Name = "frmVisor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    'Call Me.ZOrder(1)
    'If Me.Timer1.Enabled = False Then Me.Caption = Me.Caption & " *** Render desactivado! ***"
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not Tex Is Nothing Then
        If renderMode = 0 Then
            If Tex.Tiles.Count > 0 Then
                If Button = vbLeftButton Then
                    If Not selAset And Not selCset And Not selPset Then
                        selA.X = (CLng(X) \ Screen.TwipsPerPixelX) \ zoomScale
                        selA.Y = (CLng(Y) \ Screen.TwipsPerPixelY) \ zoomScale
                        selB = selA
                        selAset = True
                        Me.MousePointer = 2
                        
                        ' Vamos mostrando el resultado en los campos:
                        frmMain.txtX.Text = selA.X
                        frmMain.txtY.Text = selA.Y
                        
                    ElseIf selAset Or selCset Or selPset Then
                        Dim Key As String
                        Key = frmMain.txtKey.Text
                        
                        Dim t As Graphics.Tile
                        Set t = Tex.Tiles(Key)
                        
                        If selAset Then
                            Call t.SetRegion(CLng(frmMain.txtX.Text), CLng(frmMain.txtY.Text), CLng(frmMain.txtWidth.Text), CLng(frmMain.txtHeight.Text))
                        ElseIf selCset Then
                            Call t.SetCenter(CLng(frmMain.txtCenterX.Text), CLng(frmMain.txtCenterY.Text))
                        ElseIf selPset Then
                            Call t.ControlPoints.SetPoint(frmPointControl.txtKey.Text, CLng(frmPointControl.txtX.Text), CLng(frmPointControl.txtY.Text))
                        End If
                        
                        Set t = Nothing
                        
                        selAset = False
                        selCset = False
                        selPset = False
                        Me.MousePointer = 99
                    End If
                    
                ElseIf Button = vbRightButton Then
                    ' Cancelamos la seleccion:
                    Call modMain.ShowTileData(frmMain.lstTiles.List(frmMain.lstTiles.ListIndex))
                    
                    selAset = False
                    selCset = False
                    selPset = False
                    
                    Me.MousePointer = 99
                End If
            End If
        Else
            'Call Spr.SetLocation(CLng(X / Screen.TwipsPerPixelX), CLng(Y / Screen.TwipsPerPixelY), 0)
        End If
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    visorMouseCoord.X = CLng(X) \ Screen.TwipsPerPixelX
    visorMouseCoord.Y = CLng(Y) \ Screen.TwipsPerPixelY
        
    If selAset Then
        selB.X = visorMouseCoord.X \ zoomScale
        selB.Y = visorMouseCoord.Y \ zoomScale
        
        ' Vamos mostrando el resultado en los campos:
        frmMain.txtWidth.Text = (selB.X - selA.X) + 1
        frmMain.txtHeight.Text = (selB.Y - selA.Y) + 1
        
    ElseIf selCset Then
        selC.X = visorMouseCoord.X \ zoomScale
        selC.Y = visorMouseCoord.Y \ zoomScale
        
        ' Vamos mostrando el resultado en los campos:
        frmMain.txtCenterX.Text = selC.X - selA.X
        frmMain.txtCenterY.Text = selC.Y - selA.Y
    
    ElseIf selPset Then
        selP.X = visorMouseCoord.X \ zoomScale
        selP.Y = visorMouseCoord.Y \ zoomScale
        
        ' Vamos mostrando el resultado en los campos:
        frmPointControl.txtX.Text = selP.X - selA.X
        frmPointControl.txtY.Text = selP.Y - selA.Y
    End If

End Sub
