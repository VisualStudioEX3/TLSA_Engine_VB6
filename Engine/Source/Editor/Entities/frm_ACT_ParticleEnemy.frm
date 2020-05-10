VERSION 5.00
Begin VB.Form frm_ACT_ParticleEnemy 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Particle Enemy"
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2625
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   2625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtTriggerEventDuration 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1950
      TabIndex        =   10
      Text            =   "0"
      Top             =   1560
      Width           =   645
   End
   Begin VB.CheckBox chkReborn 
      Alignment       =   1  'Right Justify
      Caption         =   "Regenerar particula"
      Height          =   315
      Left            =   0
      TabIndex        =   8
      Top             =   1290
      Width           =   2595
   End
   Begin VB.Frame Frame3 
      Caption         =   "Entidad objetivo"
      ForeColor       =   &H8000000D&
      Height          =   645
      Left            =   0
      TabIndex        =   2
      Top             =   660
      Width           =   2625
      Begin VB.OptionButton optTarget 
         Height          =   405
         Index           =   4
         Left            =   1740
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   180
         Width           =   405
      End
      Begin VB.OptionButton optTarget 
         Height          =   405
         Index           =   3
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   180
         Width           =   405
      End
      Begin VB.OptionButton optTarget 
         Height          =   405
         Index           =   2
         Left            =   900
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   180
         Width           =   405
      End
      Begin VB.OptionButton optTarget 
         Height          =   405
         Index           =   1
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   180
         Width           =   405
      End
      Begin VB.OptionButton optTarget 
         Height          =   405
         Index           =   0
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   180
         Value           =   -1  'True
         Width           =   405
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Comportamiento de movimiento"
      ForeColor       =   &H8000000D&
      Height          =   645
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2625
      Begin VB.ComboBox cbxBehavior 
         Height          =   315
         Left            =   60
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2475
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Duracion del evento"
      Height          =   195
      Left            =   30
      TabIndex        =   9
      Top             =   1620
      Width           =   1440
   End
End
Attribute VB_Name = "frm_ACT_ParticleEnemy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cbxBehavior_Click()
    Engine.Scene.LevelEditor.EntityEditor.Entity.Behavior = Me.cbxBehavior.ListIndex
End Sub

Private Sub chkReborn_Click()
    Engine.Scene.LevelEditor.EntityEditor.Entity.Reborn = CBool(Me.chkReborn.value)
End Sub

Private Sub Form_Load()
    Call Me.cbxBehavior.AddItem("Estatico")
    Call Me.cbxBehavior.AddItem("Horizontal")
    Call Me.cbxBehavior.AddItem("Vertical")
    Call Me.cbxBehavior.AddItem("Rotacion")
    Call Me.cbxBehavior.AddItem("Aleatorio")
    Me.cbxBehavior.ListIndex = 0
End Sub

Private Sub optTarget_Click(Index As Integer)
    Engine.Scene.LevelEditor.EntityEditor.Entity.ParticleType = Index
End Sub

Private Sub txtTriggerEventDuration_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call txtTriggerEventDuration_LostFocus
End Sub

Private Sub txtTriggerEventDuration_LostFocus()
    If IsNumeric(Me.txtTriggerEventDuration.Text) Then
        Engine.Scene.LevelEditor.EntityEditor.Entity.TriggerEventDuration = CLng(Me.txtTriggerEventDuration.Text)
    Else
        Call MsgBox("Valor incorrecto para la propiedad TriggerEventDuration de ACT_ParticleEnemy.", vbCritical, "Valor no valido")
        Me.txtTriggerEventDuration.Text = "0"
    End If
End Sub
