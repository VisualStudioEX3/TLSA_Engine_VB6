VERSION 5.00
Begin VB.Form frmOutput 
   Caption         =   "Salida de depuracion"
   ClientHeight    =   3855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6810
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   257
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   454
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtOutput 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   2565
   End
End
Attribute VB_Name = "frmOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Me.Top = Screen.Height - Me.Height - 512
    Me.Width = 800 * Screen.TwipsPerPixelX
    Me.Height = 256 * Screen.TwipsPerPixelY
    Call Form_Resize
End Sub

Private Sub Form_Resize()
    Dim Client As Win32API.RECT
    Call Win32API.GetClientRect(Me.hwnd, Client)
    Me.txtOutput.Width = Client.Right
    Me.txtOutput.Height = Client.Bottom
End Sub
