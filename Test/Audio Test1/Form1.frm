VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4665
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4665
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "MusicMixer.Stop_"
      Height          =   405
      Left            =   150
      TabIndex        =   5
      Top             =   2100
      Width           =   1515
   End
   Begin VB.CommandButton Command5 
      Caption         =   "MusicMixer.Pause_"
      Height          =   375
      Left            =   150
      TabIndex        =   4
      Top             =   1740
      Width           =   1515
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Play Sound 2"
      Height          =   525
      Left            =   1740
      TabIndex        =   3
      Top             =   810
      Width           =   1515
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Play Sound 1"
      Height          =   525
      Left            =   120
      TabIndex        =   2
      Top             =   810
      Width           =   1515
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Play Track 2"
      Height          =   525
      Left            =   1740
      TabIndex        =   1
      Top             =   120
      Width           =   1515
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Play Track 1"
      Height          =   525
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1515
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4080
      Top             =   2640
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim AudioEngine As New Audio.Manager
Dim Track(1) As Audio.MusicSample
Dim fx As Audio.SoundEffects

Private Sub Command1_Click()
    Call AudioEngine.MusicMixer.Play_(Track(0), True, True)
End Sub

Private Sub Command2_Click()
    Call AudioEngine.MusicMixer.Play_(Track(1), True, True)
End Sub

Private Sub Command3_Click()
    Call AudioEngine.Soundmixer.Samples("B").Play(False, Audio_Type_Effect)
End Sub

Private Sub Command4_Click()
    Call AudioEngine.Soundmixer.Samples("C").Play(True, Audio_Type_Effect)
End Sub

Private Sub Command5_Click()
    Call AudioEngine.MusicMixer.Pause_
End Sub

Private Sub Command6_Click()
    Call AudioEngine.MusicMixer.Stop_
End Sub

Private Sub Form_Load()
    Call AudioEngine.SetWindowHandle(Me.hWnd)
    AudioEngine.VolumeControl.Music = 100
    AudioEngine.VolumeControl.Effects = 100
    AudioEngine.MusicMixer.FadeDelay = 50
    
    Set Track(0) = AudioEngine.MusicMixer.Samples.LoadSample(App.Path & "\Black Vortex.mp3", "Track1")
    Set Track(1) = AudioEngine.MusicMixer.Samples.LoadSample(App.Path & "\The Complex.mp3", "Track2")
    
    With AudioEngine.Soundmixer.Samples
        Call .LoadSample(App.Path & "\fx1.wav", "A")
        Call .LoadSample(App.Path & "\fx2.wav", "B")
        Call .LoadSample(App.Path & "\fx3.wav", "C")
    End With
    
    fx.Echo = True
    AudioEngine.Soundmixer.Globaleffects = fx
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set AudioEngine = Nothing
End Sub

Private Sub Timer1_Timer()
    If Not AudioEngine Is Nothing Then
        Call AudioEngine.Update
    End If
End Sub
