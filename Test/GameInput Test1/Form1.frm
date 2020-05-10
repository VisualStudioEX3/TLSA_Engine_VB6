VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command1"
      Height          =   555
      Left            =   1920
      TabIndex        =   1
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   555
      Left            =   180
      TabIndex        =   0
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2040
      Top             =   1230
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim inputManager As New gameinput.Manager
Dim prf As gameinput.Profile

Private Sub Command1_Click()
    Call prf.SetVibration(-1, 100)
End Sub

Private Sub Command2_Click()
    Call prf.SetVibration(2000, 50)
End Sub

Private Sub Form_Load()
    Call inputManager.SetWindowHandle(Me.hWnd)
    
    Set prf = inputManager.Profiles.Create("Test", Player1, Gamepad)
    
    prf.GamepadUsed = 0
    
    Call MsgBox(inputManager.GetGamepadName(prf.GamepadUsed))
    
    With prf.Actions
        Call .Add("jump")
        Call .Add("shoot")
        Call .Add("run")
        Call .Add("Next")
        Call .Add("Previous")
    End With
    
    Call prf.SetActionButton("jump", gameinput.KeyboardMouseButtons.Key_Space, KeybAndMouse)
    Call prf.SetActionButton("jump", gameinput.GamepadButtons.Joy_Button11, Gamepad)
    
    Call prf.SetActionButton("shoot", gameinput.KeyboardMouseButtons.Key_LControl, KeybAndMouse)
    Call prf.SetActionButton("shoot", gameinput.GamepadButtons.Joy_Button12, Gamepad)
    
    Call prf.SetActionButton("run", gameinput.KeyboardMouseButtons.Mouse_Left, KeybAndMouse)
    Call prf.SetActionButton("run", gameinput.KeyboardMouseButtons.Key_Space, KeybAndMouse)
    Call prf.SetActionButton("Next", gameinput.KeyboardMouseButtons.Mouse_Wheel_Up, KeybAndMouse)
    Call prf.SetActionButton("Previous", gameinput.KeyboardMouseButtons.Mouse_Wheel_Down, KeybAndMouse)
    
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    Cls
    Print "Actions:"
    Print "jump - " & prf.Hit("jump")
    Print "shoot - " & prf.Press("shoot")
    Print "run - " & prf.Press("run")
    Print "GamepadAxis - x" & prf.GamepadAxis.X & " y" & prf.GamepadAxis.Y
    If prf.Hit("Next") Then
        Print "Next"
    ElseIf prf.Hit("Previous") Then
        Print "Previous"
    End If
    
    Print "ViewAxis - x" & prf.viewAxis.X & "  y" & prf.viewAxis.Y
    Print "POV angle: " & prf.POV
    
    Print "Vibration support: " & prf.VibrationSupport
    
    Call inputManager.Update
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set prf = Nothing
    Set inputManager = Nothing
End Sub
