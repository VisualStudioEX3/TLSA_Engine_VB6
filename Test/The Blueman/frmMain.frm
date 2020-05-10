VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TLSA Engine - The Blueman"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private looping As Boolean

Private Sub Form_Load()
    Call Me.Show
    Call modMain.Initialize
    
    Call CreateObjects
        
    looping = True
    Do While looping
        Call modMain.Update
        Call modMain.Draw
    Loop
    
    Call modMain.Terminate
    End
End Sub

Private Sub Form_Unload(Cancel As Integer)
    looping = False
End Sub

Private Sub CreateObjects()
    Dim b As Block
    
    Set b = New Block
    Call b.SetLocation(50, 50): Call modMain.Entities.Add(b)
    Set b = Nothing
    
    Set b = New Block
    Call b.SetLocation(50, 150): Call modMain.Entities.Add(b)
    Set b = Nothing
    
    Set b = New Block
    Call b.SetLocation(50, 250): Call modMain.Entities.Add(b)
    Set b = Nothing
    
    Set b = New Block
    Call b.SetLocation(50, 350): Call modMain.Entities.Add(b)
    Set b = Nothing
    
    
    Set b = New Block
    Call b.SetLocation(150, 350): Call modMain.Entities.Add(b)
    Set b = Nothing
    
    Set b = New Block
    Call b.SetLocation(250, 350): Call modMain.Entities.Add(b)
    Set b = Nothing
    
    Set b = New Block
    Call b.SetLocation(350, 350): Call modMain.Entities.Add(b)
    Set b = Nothing
    
    Set b = New Block
    Call b.SetLocation(450, 350): Call modMain.Entities.Add(b)
    Set b = Nothing
    
    Set b = New Block
    Call b.SetLocation(450, 250): Call modMain.Entities.Add(b)
    Set b = Nothing
    
    
    Dim player As New Blueman: Call player.SetLocation(350, 50): Call modMain.Entities.Add(player, "player")
End Sub
