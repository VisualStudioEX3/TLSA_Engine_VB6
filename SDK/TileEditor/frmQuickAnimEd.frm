VERSION 5.00
Begin VB.Form frmQuickAnimEd 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Editor de secuencias de animacion"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3510
   Icon            =   "frmQuickAnimEd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   3510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Secuencias de animacion"
      ForeColor       =   &H8000000D&
      Height          =   1065
      Left            =   30
      TabIndex        =   11
      Top             =   0
      Width           =   3465
      Begin VB.ComboBox cbxAnimSec 
         Height          =   315
         Left            =   90
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   3285
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "Nueva"
         Height          =   345
         Left            =   90
         TabIndex        =   1
         Top             =   600
         Width           =   885
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Eliminar"
         Height          =   345
         Left            =   990
         TabIndex        =   2
         Top             =   600
         Width           =   885
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Propiedades"
      ForeColor       =   &H8000000D&
      Height          =   1905
      Left            =   30
      TabIndex        =   8
      Top             =   1080
      Width           =   3465
      Begin VB.CheckBox chkLoop 
         Alignment       =   1  'Right Justify
         Caption         =   "Animacion en bucle"
         Height          =   315
         Left            =   90
         TabIndex        =   12
         Top             =   1200
         Width           =   3285
      End
      Begin VB.TextBox txtKey 
         Height          =   315
         Left            =   90
         TabIndex        =   3
         Top             =   480
         Width           =   3285
      End
      Begin VB.TextBox txtDelay 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2550
         MaxLength       =   6
         TabIndex        =   4
         Text            =   "0"
         Top             =   870
         Width           =   825
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Actualizar"
         Height          =   315
         Left            =   60
         TabIndex        =   5
         Top             =   1530
         Width           =   3345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre de la secuencia de animacion"
         Height          =   195
         Left            =   90
         TabIndex        =   10
         Top             =   240
         Width           =   2715
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tiempo de actualizacion (mlSeg)"
         Height          =   195
         Left            =   90
         TabIndex        =   9
         Top             =   930
         Width           =   2295
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Opciones"
      ForeColor       =   &H8000000D&
      Height          =   585
      Left            =   30
      TabIndex        =   7
      Top             =   3000
      Width           =   3465
      Begin VB.CommandButton cmdPreview 
         Caption         =   "Previsualizar"
         Enabled         =   0   'False
         Height          =   315
         Left            =   60
         TabIndex        =   6
         Top             =   210
         Width           =   3345
      End
   End
End
Attribute VB_Name = "frmQuickAnimEd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cbxAnimSec_Click()
    Call modAnimEd.ShowData(Me.cbxAnimSec.List(Me.cbxAnimSec.ListIndex))
End Sub

Private Sub cmdDelete_Click()
    If MsgBox("¿Eliminar la secuencia de la lista?", vbExclamation + vbYesNo, "Eliminar secuencia de animacion") = vbYes Then
        Call Spr.Animations.Remove(Me.cbxAnimSec.List(Me.cbxAnimSec.ListIndex))
        Call Me.cbxAnimSec.RemoveItem(Me.cbxAnimSec.ListIndex)
        If Me.cbxAnimSec.ListCount > 0 Then Me.cbxAnimSec.ListIndex = 0
        Call modAnimEd.LoadSecuences
    End If
End Sub

Private Sub cmdNew_Click()
    Call frmSecuenceDialog.Show(vbModal)
End Sub

Private Sub cmdPreview_Click()
    ' Cambiamos a modo reproduccion:
    Call modMain.SetAnimationMode
End Sub

Private Sub cmdUpdate_Click()
    Dim sec As Graphics.Animation
    Set sec = Spr.Animations(Me.txtKey.Tag)
    
    sec.FrameDelay = Me.txtDelay.Text
    sec.Looping = CBool(Me.chkLoop.value)
    
    Set sec = Nothing
    
    If Me.txtKey.Text <> Me.txtKey.Tag Then Call modAnimEd.ChangeKey(Me.txtKey.Tag, Me.txtKey.Text)
    
    Call modAnimEd.ShowData(Me.txtKey.Text)
End Sub

Private Sub Form_Load()
    Call modAnimEd.LoadSecuences
    Me.Left = frmPointControl.Left
    Me.Top = frmPointControl.Top + Me.Height
End Sub

Private Sub txtDelay_LostFocus()
    If Not IsNumeric(Me.txtDelay.Text) Then Me.txtDelay.Text = 0
End Sub

Private Sub txtKey_LostFocus()
    If Me.txtKey.Text = "" Then Me.txtKey.Text = Me.txtKey.Tag
End Sub
