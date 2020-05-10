VERSION 5.00
Begin VB.Form frmEntityList 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Entidades"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2610
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   2610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstEntities 
      Height          =   2985
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2595
   End
End
Attribute VB_Name = "frmEntityList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Call Me.lstEntities.AddItem("Jugador")
    Call Me.lstEntities.AddItem("Particula")
    Call Me.lstEntities.AddItem("Plataforma")
    Call Me.lstEntities.AddItem("Salida")
End Sub

Private Sub lstEntities_Click()
    Call Engine.Scene.LevelEditor.EntityEditor.SetEntity(Me.lstEntities.Text)
End Sub
