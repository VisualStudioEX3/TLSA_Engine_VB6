VERSION 5.00
Begin VB.Form frmPhysParams 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Editor de fisicas de escenario"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   2700
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   2700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkEnabled 
      Alignment       =   1  'Right Justify
      Caption         =   "Habilitado"
      Enabled         =   0   'False
      ForeColor       =   &H8000000D&
      Height          =   225
      Left            =   0
      TabIndex        =   11
      Top             =   2610
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox txtParam 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   2250
      MaxLength       =   3
      TabIndex        =   10
      Text            =   "32"
      Top             =   2340
      Width           =   405
   End
   Begin VB.ComboBox cbxTipo 
      Height          =   315
      Left            =   30
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1980
      Width           =   2625
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tamaño"
      ForeColor       =   &H8000000D&
      Height          =   1695
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   2625
      Begin VB.ComboBox cbxDefSizes 
         Height          =   315
         Left            =   450
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   420
         Width           =   2085
      End
      Begin VB.TextBox txtHeight 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "64"
         Top             =   1080
         Width           =   405
      End
      Begin VB.TextBox txtWidth 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "64"
         Top             =   840
         Width           =   405
      End
      Begin VB.HScrollBar scrWidth 
         Height          =   255
         LargeChange     =   16
         Left            =   330
         Max             =   768
         Min             =   16
         SmallChange     =   16
         TabIndex        =   2
         Top             =   1350
         Value           =   64
         Width           =   2205
      End
      Begin VB.VScrollBar scrHeight 
         Height          =   1155
         LargeChange     =   16
         Left            =   90
         Max             =   576
         Min             =   16
         SmallChange     =   16
         TabIndex        =   1
         Top             =   210
         Value           =   64
         Width           =   255
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tamaños predefinidos"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   450
         TabIndex        =   13
         Top             =   210
         Width           =   1560
      End
      Begin VB.Label Label1 
         Caption         =   "Alto"
         Height          =   225
         Left            =   480
         TabIndex        =   5
         Top             =   1080
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Ancho"
         Height          =   225
         Left            =   480
         TabIndex        =   3
         Top             =   840
         Width           =   465
      End
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Grosor de los colisionadores"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   30
      TabIndex        =   9
      Top             =   2370
      Width           =   1980
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Tipo de superficie"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   30
      TabIndex        =   8
      Top             =   1740
      Width           =   1260
   End
End
Attribute VB_Name = "frmPhysParams"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cbxDefSizes_Click()
    Dim value() As String: value = Split(Me.cbxDefSizes.Text, "x")
    Me.scrWidth.value = CLng(value(0))
    Me.scrHeight.value = CLng(value(1))
End Sub

Private Sub cbxTipo_Click()
    On Error Resume Next
    Engine.Scene.LevelEditor.PhysicEditor.SelectColliderSpecification = Me.cbxTipo.ListIndex
End Sub

Private Sub chkEnabled_Click()
    'Engine.Scene.LevelEditor.PhysicEditor.BlockEnable = CBool(chkEnabled.value)
End Sub

Private Sub Form_Load()
    Call Me.cbxTipo.Clear

    Call Me.cbxTipo.AddItem("Suelo")
    Call Me.cbxTipo.AddItem("Techo")
    Call Me.cbxTipo.AddItem("Pared Izquierda")
    Call Me.cbxTipo.AddItem("Pared Derecha")
    Call Me.cbxTipo.AddItem("Suelo y techo (toda el area)")
    Call Me.cbxTipo.AddItem("Paredes (toda el area)")
    Call Me.cbxTipo.AddItem("Suelo y techo (media area)")
    Call Me.cbxTipo.AddItem("Paredes (media area)")
    Call Me.cbxTipo.AddItem("Todos los colisionadores")
    Call Me.cbxTipo.AddItem("Todos con prioridad suelo y techo")
    Call Me.cbxTipo.AddItem("Todos con esquinas libres")
    'Call Me.cbxTipo.AddItem("Manual")
    
    Me.cbxTipo.ListIndex = 0
    
    
    Call Me.cbxDefSizes.Clear
    
    Call Me.cbxDefSizes.AddItem("16x16")
    Call Me.cbxDefSizes.AddItem("32x32")
    Call Me.cbxDefSizes.AddItem("64x64")
    Call Me.cbxDefSizes.AddItem("128x128")
    Call Me.cbxDefSizes.AddItem("256x256")
    Call Me.cbxDefSizes.AddItem("512x512")
    Call Me.cbxDefSizes.AddItem("64x32")
    Call Me.cbxDefSizes.AddItem("128x32")
    Call Me.cbxDefSizes.AddItem("256x32")
    Call Me.cbxDefSizes.AddItem("512x32")
    Call Me.cbxDefSizes.AddItem("32x64")
    Call Me.cbxDefSizes.AddItem("32x128")
    Call Me.cbxDefSizes.AddItem("32x256")
    Call Me.cbxDefSizes.AddItem("32x512")
End Sub

Private Sub scrHeight_Change()
    Me.txtHeight.Text = Me.scrHeight.value
End Sub

Private Sub scrHeight_Scroll()
    Call scrHeight_Change
End Sub

Private Sub scrWidth_Change()
    Me.txtWidth.Text = Me.scrWidth.value
End Sub

Private Sub scrWidth_Scroll()
    Call scrWidth_Change
End Sub

Private Sub txt_Change()
    Engine.Scene.LevelEditor.PhysicEditor.Param = txtParam.Text
End Sub

Private Sub txtHeight_Change()
    Engine.Scene.LevelEditor.Brush.Size = _
        Core.Generics.CreateSIZE(Engine.Scene.LevelEditor.Brush.Size.Width, CLng(Me.txtHeight.Text))
End Sub

Private Sub txtParam_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call txtParam_LostFocus
End Sub

Private Sub txtParam_LostFocus()
    If IsNumeric(Me.txtParam.Text) Then Engine.Scene.LevelEditor.PhysicEditor.Param = CLng(Me.txtParam.Text)
End Sub

Private Sub txtWidth_Change()
    Engine.Scene.LevelEditor.Brush.Size = _
        Core.Generics.CreateSIZE(CLng(Me.txtWidth.Text), Engine.Scene.LevelEditor.Brush.Size.Height)
End Sub
