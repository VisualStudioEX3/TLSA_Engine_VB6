VERSION 5.00
Begin VB.Form frmPrevCtrl 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Controles de reproduccion"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3540
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   3540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option1 
      Caption         =   "4x"
      Height          =   225
      Index           =   2
      Left            =   3000
      TabIndex        =   12
      Top             =   2100
      Width           =   615
   End
   Begin VB.OptionButton Option1 
      Caption         =   "2x"
      Height          =   225
      Index           =   1
      Left            =   2370
      TabIndex        =   11
      Top             =   2100
      Width           =   615
   End
   Begin VB.OptionButton Option1 
      Caption         =   "1x"
      Height          =   225
      Index           =   0
      Left            =   1740
      TabIndex        =   10
      Top             =   2100
      Value           =   -1  'True
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Caption         =   "Propiedades"
      ForeColor       =   &H8000000D&
      Height          =   1695
      Left            =   30
      TabIndex        =   3
      Top             =   330
      Width           =   3495
      Begin VB.CheckBox chkLoop 
         Alignment       =   1  'Right Justify
         Caption         =   "Animacion en bucle"
         Height          =   255
         Left            =   90
         TabIndex        =   8
         Top             =   1170
         Width           =   3285
      End
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   1140
         Top             =   210
      End
      Begin VB.CheckBox chkPoints 
         Alignment       =   1  'Right Justify
         Caption         =   "Mostrar puntos de control"
         Height          =   255
         Left            =   90
         TabIndex        =   9
         Top             =   1410
         Width           =   3285
      End
      Begin VB.CheckBox chkRev 
         Alignment       =   1  'Right Justify
         Caption         =   "Invertir sentido de la animacion"
         Height          =   255
         Left            =   90
         TabIndex        =   7
         Top             =   930
         Width           =   3285
      End
      Begin VB.TextBox txtDelay 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2550
         MaxLength       =   6
         TabIndex        =   4
         Text            =   "0"
         Top             =   600
         Width           =   825
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   3360
         Y1              =   450
         Y2              =   450
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0 de 0"
         Height          =   195
         Left            =   2895
         TabIndex        =   6
         Top             =   240
         Width           =   450
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tiempo de actualizacion (mlSeg)"
         Height          =   195
         Left            =   90
         TabIndex        =   5
         Top             =   660
         Width           =   2295
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   ";"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   3060
      TabIndex        =   1
      ToolTipText     =   "Reproducir"
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   2550
      TabIndex        =   0
      ToolTipText     =   "Reiniciar la secuencia"
      Top             =   0
      Width           =   495
   End
   Begin VB.ComboBox cbxSecuence 
      Height          =   315
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   0
      Width           =   2565
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Zoom:"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   2100
      Width           =   450
   End
End
Attribute VB_Name = "frmPrevCtrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim secuence As String
Dim reverb As Boolean

Private Sub cbxSecuence_Click()
    secuence = Me.cbxSecuence.List(Me.cbxSecuence.ListIndex)
    Call Spr.SetCurrentAnimation(secuence)
    Call Spr.CurrentAnimation.Reset
    Call Spr.CurrentAnimation.Play
    Me.Command1(1).Caption = ";"
    Me.txtDelay.Text = Spr.CurrentAnimation.FrameDelay
    Me.chkRev.value = vbUnchecked
    Me.chkLoop.value = Abs(CInt(Spr.CurrentAnimation.Looping)): reverb = False
    Me.chkPoints.value = vbUnchecked
    Select Case zoomScale
        Case 1: Me.Option1(0).value = True
        Case 2: Me.Option1(1).value = True
        Case 4: Me.Option1(2).value = True
    End Select
End Sub

Private Sub chkRev_Click()
    reverb = Not reverb
    
    If reverb Then
        Spr.CurrentAnimation.AnimatePath = graphics.AnimationPath.Reverse
    Else
        Spr.CurrentAnimation.AnimatePath = graphics.AnimationPath.Foward
    End If
End Sub

Private Sub chkPoints_Click()
    modMain.showControlPointsInAnimationMode = Not modMain.showControlPointsInAnimationMode
End Sub

Private Sub chkLoop_Click()
    Spr.CurrentAnimation.Looping = CBool(Me.chkLoop.value)
End Sub

Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case 0
            Call Spr.CurrentAnimation.Reset
        
        Case 1
            If Spr.CurrentAnimation.IsPaused Or Spr.CurrentAnimation.IsAnimationEnded Then
                Call Spr.CurrentAnimation.Play
                Me.Command1(1).Caption = ";"
            Else
                Call Spr.CurrentAnimation.Pause
                Me.Command1(1).Caption = "4"
            End If
    End Select
End Sub

Private Sub Form_Load()
    Call modMain.OnTop(Me)
    
    Dim i As Long
    For i = 0 To frmQuickAnimEd.cbxAnimSec.ListCount - 1
        Call Me.cbxSecuence.AddItem(frmQuickAnimEd.cbxAnimSec.List(i))
    Next
    
    Me.cbxSecuence.ListIndex = 0

    Call Spr.SetCurrentAnimation(Me.cbxSecuence.List(0))
    
    Spr.CurrentAnimation.AnimatePath = graphics.AnimationPath.Foward
    modMain.showControlPointsInAnimationMode = False
    
    'Me.Option1(0).value = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Cambiamos al modo edicion:
    Call modMain.SetEditorMode
End Sub

Private Sub Option1_Click(Index As Integer)
    Select Case Index
        Case 0: Call Spr.SetScale(1)
        Case 1: Call Spr.SetScale(2)
        Case 2: Call Spr.SetScale(4)
    End Select
'    MsgBox Spr.Location.X & ", " & Spr.Location.Y
'    Call Spr.SetLocation(256, 256, 0)
End Sub

Private Sub Timer1_Timer()
    Me.lblInfo.Caption = Spr.CurrentTile.Key & " - " & Spr.CurrentAnimation.CurrentTileIndex & " de " & Spr.CurrentAnimation.Tiles.Count
    If Spr.CurrentAnimation.IsAnimationEnded Then
        'Call Command1_Click(1)
        Me.Command1(1).Caption = "4"
    End If
End Sub

Private Sub txtDelay_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not IsNumeric(Me.txtDelay.Text) Then
            Me.txtDelay.Text = 0
            Exit Sub
        Else
            Spr.CurrentAnimation.FrameDelay = Me.txtDelay.Text
        End If
    End If
End Sub
