VERSION 5.00
Begin VB.Form frmDebug 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TLSA Engine: Opciones de depuracion"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5250
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   5250
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check5 
      Caption         =   "Activar el editor de niveles"
      Height          =   315
      Left            =   30
      TabIndex        =   9
      Top             =   2130
      Value           =   1  'Checked
      Width           =   5175
   End
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   5250
      TabIndex        =   6
      Top             =   2445
      Width           =   5250
      Begin VB.CommandButton Command2 
         Cancel          =   -1  'True
         Caption         =   "Salir"
         Height          =   435
         Left            =   3690
         TabIndex        =   8
         Top             =   60
         Width           =   1485
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Iniciar TLSA"
         Default         =   -1  'True
         Height          =   435
         Left            =   2160
         TabIndex        =   7
         Top             =   60
         Width           =   1485
      End
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Mostrar guias de objetos fisicos y escenario."
      Height          =   315
      Left            =   30
      TabIndex        =   4
      Top             =   1800
      Value           =   1  'Checked
      Width           =   5175
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Separar actualizacion de estados de fisica en hilo independiente."
      Height          =   315
      Left            =   30
      TabIndex        =   3
      Top             =   1470
      Width           =   5175
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Limitar a 30fps maximo (desactivado son 60fps maximo)"
      Height          =   315
      Left            =   30
      TabIndex        =   2
      Top             =   1140
      Value           =   1  'Checked
      Width           =   5175
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Modo pantalla completa"
      Height          =   315
      Left            =   30
      TabIndex        =   1
      Top             =   810
      Width           =   5175
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   765
      Left            =   0
      Picture         =   "frmDebug.frx":0000
      ScaleHeight     =   735
      ScaleWidth      =   5220
      TabIndex        =   0
      Top             =   0
      Width           =   5250
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmDebug.frx":1B42
         Height          =   615
         Left            =   780
         TabIndex        =   5
         Top             =   60
         Width           =   4395
      End
   End
End
Attribute VB_Name = "frmDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Engine.FULL_SCREEN = CBool(Check1.value)
    Engine.FPS_MAX = IIf(CBool(Check2.value), 30, 60)
    Engine.PHYSICS_SEPARATE_THREAD = CBool(Check3.value)
    Engine.PHYSICS_DRAW_GUIDES = CBool(Check4.value)
    Engine.EDIT_MODE = CBool(Check5.value)
    Call Unload(Me)
End Sub

Private Sub Command2_Click()
    End
End Sub
