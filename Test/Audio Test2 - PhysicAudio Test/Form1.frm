VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Audio Test2 - PhysicAudio test"
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picScene 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   9000
      Left            =   0
      ScaleHeight     =   598
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   798
      TabIndex        =   0
      Top             =   0
      Width           =   12000
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   5790
         Top             =   4290
      End
      Begin VB.Shape shpGargle 
         BorderColor     =   &H00FF0000&
         Height          =   2655
         Left            =   6810
         Top             =   3270
         Width           =   2835
      End
      Begin VB.Shape shpEcho 
         BorderColor     =   &H0000FF00&
         Height          =   2655
         Left            =   2640
         Top             =   3150
         Width           =   2835
      End
      Begin VB.Image imgListener 
         Height          =   480
         Left            =   2400
         Picture         =   "Form1.frx":0000
         Top             =   270
         Width           =   480
      End
      Begin VB.Image imgSound 
         Height          =   240
         Left            =   600
         Picture         =   "Form1.frx":08CA
         Top             =   840
         Width           =   240
      End
      Begin VB.Shape shpRadius 
         BorderColor     =   &H000000FF&
         Height          =   1905
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Width           =   1455
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim AudioEngine As New Audio.Manager
Dim Sound As Audio.AudioPhysicEmitter
Dim Ok As Boolean

Private Sub Form_Load()
    Me.AutoRedraw = True
    Me.picScene.AutoRedraw = True
    
    Call AudioEngine.SetWindowHandle(Me.hWnd)
    Call AudioEngine.SoundMixer.Samples.LoadSample(App.Path & "\loop.wav", "sample")
    AudioEngine.SoundMixer.PhysicEngine.Enabled = True
    
    ' Creamos un emisor de sonido:
    Set Sound = AudioEngine.SoundMixer.PhysicEngine.Emitters.Create("sample", AudioEngine.SoundMixer.Samples("sample"), 400, 300, 600)
    
    ' Posicionamos al escuchador en el centro del escenario:
    Call AudioEngine.SoundMixer.PhysicEngine.SetListener(0, 300) 'Me.picScene.Width / 2, Me.picScene.Height / 2)
    
    Call AudioEngine.SoundMixer.PhysicEngine.EffectRegions.Create("echo", shpEcho.Left, shpEcho.Top, shpEcho.Width, shpEcho.Height)
    Call AudioEngine.SoundMixer.PhysicEngine.EffectRegions("echo").SetEffects(False, False, False, True, False, False, False)
    
    Call AudioEngine.SoundMixer.PhysicEngine.EffectRegions.Create("gargle", shpGargle.Left, shpGargle.Top, shpGargle.Width, shpGargle.Height)
    Call AudioEngine.SoundMixer.PhysicEngine.EffectRegions("gargle").SetEffects(False, False, False, False, False, True, False)
    
    ' Reproducimos la muestra en bucle:
    Call Sound.Play(True)
    
    Call UpdateControls
    
    Ok = True
End Sub

Private Sub UpdateControls()
    Me.imgListener.Left = (AudioEngine.SoundMixer.PhysicEngine.Listener.X - (Me.imgListener.Width / 2))
    Me.imgListener.Top = (AudioEngine.SoundMixer.PhysicEngine.Listener.Y - (Me.imgListener.Height / 2))
    
    Me.imgSound.Left = (Sound.Location.X - (Me.imgSound.Width / 2))
    Me.imgSound.Top = (Sound.Location.Y - (Me.imgSound.Height / 2))
    
    Me.shpRadius.Width = Sound.Radius
    Me.shpRadius.Height = Sound.Radius
    Me.shpRadius.Left = (Sound.Location.X - (Me.shpRadius.Width / 2))
    Me.shpRadius.Top = (Sound.Location.Y - (Me.shpRadius.Height / 2))
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Ok = False
    Call AudioEngine.SoundMixer.Samples.Remove("sample")
    Set AudioEngine = Nothing
End Sub

Private Sub picScene_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Call AudioEngine.SoundMixer.PhysicEngine.SetListener(CLng(X), CLng(Y))
    ElseIf Button = 2 Then
        Call Sound.SetLocation(CLng(X), CLng(Y))
    End If
    Call UpdateControls
End Sub

Private Sub Timer1_Timer()
    If Ok Then
        Me.imgListener.Left = Me.imgListener.Left + 1
        Call AudioEngine.SoundMixer.PhysicEngine.SetListener(Me.imgListener.Left, Me.imgListener.Top)
        Call AudioEngine.Update
    End If
End Sub
