VERSION 5.00
Begin VB.Form frmPrevAnim 
   BackColor       =   &H8000000C&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Previsualizacion"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmPrevAnim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    frmQuickAnimEd.Enabled = False
    
    Me.Show
    
    Set modAnimEd.prevGfx = New Graphics.Manager
    Call modAnimEd.prevGfx.Initialize(Me.hwnd, 256, 256, 32, True, True)
    modAnimEd.prevGfx.MaxFrames = 60
    
    Call modAnimEd.prevGfx.Textures.LoadTexture(App.Path & "\alpha.bmp", "alpha", False)
    Set modAnimEd.prevAlpha = New Graphics.Sprite
    Call modAnimEd.prevAlpha.SetTexture(modAnimEd.prevGfx.Textures("alpha"))
    
    ' Cargamos la textura, con sus tiles y las animaciones definidas:
    Call modAnimEd.LoadSpriteAnimation
    
    ' Cargamos el panel de reproduccion:
    Call frmPrevCtrl.Show
    
    Do While Not modAnimEd.prevGfx Is Nothing
        Call modAnimEd.PrevDrawAlphaBackground
        Call modAnimEd.prevSpr.Update
        Call modAnimEd.prevSpr.Draw
        
        Call modAnimEd.prevGfx.Render
    Loop
    
    Call Unload(frmPrevCtrl)
    
    frmQuickAnimEd.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set modAnimEd.prevGfx = Nothing
End Sub
