VERSION 5.00
Begin VB.Form frmTileEditor 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Editor de tiles"
   ClientHeight    =   7905
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   1965
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   527
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   131
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.VScrollBar scrScale 
      Height          =   255
      Left            =   1290
      Max             =   8
      Min             =   1
      TabIndex        =   34
      Top             =   2280
      Value           =   1
      Width           =   285
   End
   Begin VB.TextBox txtScale 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   1530
      MaxLength       =   3
      TabIndex        =   32
      Text            =   "1"
      Top             =   2280
      Width           =   405
   End
   Begin VB.Frame Frame6 
      Caption         =   "Capa de dibujo"
      ForeColor       =   &H8000000D&
      Height          =   585
      Left            =   30
      TabIndex        =   30
      Top             =   6180
      Width           =   1920
      Begin VB.ComboBox cbxLayer 
         Height          =   315
         Left            =   60
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   210
         Width           =   1815
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Visualizar capas"
      ForeColor       =   &H8000000D&
      Height          =   1095
      Left            =   30
      TabIndex        =   26
      Top             =   6780
      Width           =   1920
      Begin VB.CheckBox chkLayerFront 
         Caption         =   "Primer plano"
         Height          =   285
         Left            =   60
         TabIndex        =   29
         Top             =   750
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.CheckBox chkLayerMain 
         Caption         =   "Principal"
         Height          =   285
         Left            =   60
         TabIndex        =   28
         Top             =   480
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.CheckBox chkLayerBackground 
         Caption         =   "Fondo"
         Height          =   285
         Left            =   60
         TabIndex        =   27
         Top             =   210
         Value           =   1  'Checked
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdTexLib 
      Caption         =   "Seleccionar biblioteca"
      Height          =   315
      Left            =   30
      TabIndex        =   1
      Top             =   1950
      Width           =   1920
   End
   Begin VB.Frame Frame4 
      Caption         =   "Espejado"
      ForeColor       =   &H8000000D&
      Height          =   585
      Left            =   30
      TabIndex        =   24
      Top             =   5580
      Width           =   1905
      Begin VB.ComboBox cbxMirror 
         Height          =   315
         Left            =   60
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   210
         Width           =   1815
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Efecto"
      ForeColor       =   &H8000000D&
      Height          =   585
      Left            =   30
      TabIndex        =   22
      Top             =   4980
      Width           =   1905
      Begin VB.ComboBox cbxFX 
         Height          =   315
         Left            =   60
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   210
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Color"
      ForeColor       =   &H8000000D&
      Height          =   1875
      Left            =   30
      TabIndex        =   5
      Top             =   3090
      Width           =   1920
      Begin VB.PictureBox picColorPreview 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   900
         ScaleHeight     =   195
         ScaleWidth      =   915
         TabIndex        =   20
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox txtColorValue 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   900
         TabIndex        =   19
         Text            =   "-1"
         Top             =   1290
         Width           =   975
      End
      Begin VB.TextBox txtBlue 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   1470
         MaxLength       =   3
         TabIndex        =   16
         Text            =   "255"
         Top             =   1020
         Width           =   405
      End
      Begin VB.HScrollBar scrBlue 
         Height          =   255
         Left            =   240
         Max             =   255
         TabIndex        =   15
         Top             =   1020
         Value           =   255
         Width           =   1275
      End
      Begin VB.TextBox txtGreen 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   1470
         MaxLength       =   3
         TabIndex        =   13
         Text            =   "255"
         Top             =   750
         Width           =   405
      End
      Begin VB.HScrollBar scrGreen 
         Height          =   255
         Left            =   240
         Max             =   255
         TabIndex        =   12
         Top             =   750
         Value           =   255
         Width           =   1275
      End
      Begin VB.TextBox txtRed 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   1470
         MaxLength       =   3
         TabIndex        =   10
         Text            =   "255"
         Top             =   480
         Width           =   405
      End
      Begin VB.HScrollBar scrRed 
         Height          =   255
         Left            =   240
         Max             =   255
         TabIndex        =   9
         Top             =   480
         Value           =   255
         Width           =   1275
      End
      Begin VB.TextBox txtAlpha 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   1470
         MaxLength       =   3
         TabIndex        =   7
         Text            =   "255"
         Top             =   210
         Width           =   405
      End
      Begin VB.HScrollBar scrAlpha 
         Height          =   255
         Left            =   240
         Max             =   255
         TabIndex        =   6
         Top             =   210
         Value           =   255
         Width           =   1275
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Muestra"
         Height          =   195
         Left            =   60
         TabIndex        =   21
         Top             =   1560
         Width           =   570
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Valor final"
         Height          =   195
         Left            =   60
         TabIndex        =   18
         Top             =   1320
         Width           =   690
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "B"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   60
         TabIndex        =   17
         Top             =   1020
         Width           =   195
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "G"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   60
         TabIndex        =   14
         Top             =   750
         Width           =   195
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "R"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   60
         TabIndex        =   11
         Top             =   480
         Width           =   195
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "A"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   60
         TabIndex        =   8
         Top             =   210
         Width           =   195
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Angulo"
      ForeColor       =   &H8000000D&
      Height          =   525
      Left            =   30
      TabIndex        =   2
      Top             =   2550
      Width           =   1920
      Begin VB.HScrollBar scrAngle 
         Height          =   255
         LargeChange     =   30
         Left            =   60
         Max             =   359
         TabIndex        =   4
         Top             =   210
         Width           =   1450
      End
      Begin VB.TextBox txtAngle 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   1470
         MaxLength       =   3
         TabIndex        =   3
         Text            =   "0"
         Top             =   210
         Width           =   405
      End
   End
   Begin VB.CommandButton cmdExplorer 
      Caption         =   "Explorador de tiles"
      Enabled         =   0   'False
      Height          =   315
      Left            =   30
      TabIndex        =   0
      Top             =   1980
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Escala (1 a 8)"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   30
      TabIndex        =   33
      Top             =   2310
      Width           =   975
   End
   Begin VB.Image imgPreview 
      BorderStyle     =   1  'Fixed Single
      Height          =   1905
      Left            =   30
      Stretch         =   -1  'True
      Top             =   30
      Width           =   1905
   End
End
Attribute VB_Name = "frmTileEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ResetParams()
    scrScale.value = 1
    scrAngle.value = 0
    scrAlpha.value = 255
    scrRed.value = 255
    scrGreen.value = 255
    scrBlue.value = 255
    cbxFX.ListIndex = 0
    cbxMirror.ListIndex = 0
End Sub

Private Sub cbxFX_Click()
    If Not Engine.Scene Is Nothing Then
        Select Case cbxFX.ListIndex
            Case 0: Engine.Scene.LevelEditor.TileEditor.Effect = Default
            Case 1: Engine.Scene.LevelEditor.TileEditor.Effect = Aditive
            Case 2: Engine.Scene.LevelEditor.TileEditor.Effect = Sustrative
            Case 3: Engine.Scene.LevelEditor.TileEditor.Effect = Negative
            Case 4: Engine.Scene.LevelEditor.TileEditor.Effect = XOR_Exclusion
            Case 5: Engine.Scene.LevelEditor.TileEditor.Effect = Greyscale
            Case 6: Engine.Scene.LevelEditor.TileEditor.Effect = Crystaline
        End Select
    End If
End Sub

Private Sub cbxLayer_Click()
    If Not Engine.Scene Is Nothing Then
        Select Case cbxLayer.ListIndex
            Case 0: Engine.Scene.LevelEditor.TileEditor.Layer = 4
            Case 1: Engine.Scene.LevelEditor.TileEditor.Layer = 0
            Case 2: Engine.Scene.LevelEditor.TileEditor.Layer = -4
        End Select
    End If
End Sub

Private Sub cbxMirror_Click()
    If Not Engine.Scene Is Nothing Then
        Select Case cbxMirror.ListIndex
            Case 0: Engine.Scene.LevelEditor.TileEditor.Mirror = WithoutMirroring
            Case 1: Engine.Scene.LevelEditor.TileEditor.Mirror = Horizontal
            Case 2: Engine.Scene.LevelEditor.TileEditor.Mirror = Vertical
            Case 3: Engine.Scene.LevelEditor.TileEditor.Mirror = Both
        End Select
    End If
End Sub

Private Sub chkLayerBackground_Click()
    Engine.Scene.ShowBackLayer = CBool(Me.chkLayerBackground.value)
End Sub

Private Sub chkLayerFront_Click()
    Engine.Scene.ShowFrontLayer = CBool(Me.chkLayerFront.value)
End Sub

Private Sub chkLayerMain_Click()
    Engine.Scene.ShowMainLayer = CBool(Me.chkLayerMain.value)
End Sub

Private Sub cmdExplorer_Click()
    If Engine.Scene.LevelEditor.TileEditor.Sprite.Animations.Count > 0 Then
        Call MsgBox("El sprite contiene animaciones. Se utilizara la primera animacion por defecto como tile animado.", vbExclamation, "No se puede seleccionar tiles")
    Else
        Call frmTileExplorer.Show(vbModal)
    End If
End Sub

Private Sub Command1_Click()
    Call frmTextureBrowser.Show(vbModal)
End Sub

Private Sub cmdTexLib_Click()
    Call frmTextureBrowser.Show(vbModal)
    If Engine.Scene.LevelEditor.TileEditor.Texture Is Nothing Then
        Me.cmdExplorer.Enabled = False
        Set Me.imgPreview.Picture = Nothing
        
        ' Reiniciamos los parametros del editor:
        Call ResetParams
    Else
        Me.cmdExplorer.Enabled = True
        
        ' Aplicamos los valores al tile seleccionado:
        Call scrAngle_Change
        Call scrAlpha_Change
        Call scrRed_Change
        Call scrGreen_Change
        Call scrBlue_Change
        Call cbxFX_Click
        Call cbxMirror_Click
    End If
End Sub

Private Sub Form_Load()
    Call SetColorPreview
    
    ' Lista de efectos:
    Call Me.cbxFX.Clear
    Call Me.cbxFX.AddItem("Ninguno")
    Call Me.cbxFX.AddItem("Aditivo")
    Call Me.cbxFX.AddItem("Sustrativo")
    Call Me.cbxFX.AddItem("Negativo")
    Call Me.cbxFX.AddItem("XOR")
    Call Me.cbxFX.AddItem("Escala de grises")
    Call Me.cbxFX.AddItem("Cristalino")
    Me.cbxFX.ListIndex = 0
    
    ' Lista de espejados:
    Call Me.cbxMirror.Clear
    Call Me.cbxMirror.AddItem("Ninguno")
    Call Me.cbxMirror.AddItem("Horizontal")
    Call Me.cbxMirror.AddItem("Vertical")
    Call Me.cbxMirror.AddItem("Ambos")
    Me.cbxMirror.ListIndex = 0
    
    ' Lista de capas del nivel:
    Call Me.cbxLayer.Clear
    Call Me.cbxLayer.AddItem("Fondo")
    Call Me.cbxLayer.AddItem("Principal")
    Call Me.cbxLayer.AddItem("Primer plano")
    Me.cbxLayer.ListIndex = 1
End Sub


Private Sub PreviewTexture1_Click()
    Call cmdExplorer_Click
End Sub

'

Private Sub scrScale_Change()
    Me.txtScale.Text = scrScale.value
    If Not Engine.Scene Is Nothing Then Engine.Scene.LevelEditor.TileEditor.ScaleFactor = CSng(scrScale.value)
End Sub

Private Sub scrScale_Scroll()
    Call scrScale_Change
End Sub

Private Sub txtScale_LostFocus()
    If Not IsNumeric(txtScale.Text) Then
        txtScale.Text = "1"
    Else
        If CSng(txtScale.Text) = 0 Or CSng(txtScale.Text) > 8 Then txtScale.Text = "1"
    End If
    Engine.Scene.LevelEditor.TileEditor.ScaleFactor = CSng(txtScale.Text)
End Sub

'

Private Sub scrAngle_Change()
    Me.txtAngle.Text = scrAngle.value
    If Not Engine.Scene Is Nothing Then Engine.Scene.LevelEditor.TileEditor.Angle = scrAngle.value
End Sub

Private Sub scrAngle_Scroll()
    Call scrAngle_Change
End Sub

Private Sub txtAngle_LostFocus()
   If IsNumeric(txtAngle.Text) Then
        Dim value As Integer: value = Abs(CInt(txtAngle.Text))
        If (value >= 0 And value <= 359) Then
            Me.scrAngle.value = value
        End If
    End If
    Me.txtAngle.Text = Me.scrAngle.value
End Sub

Private Sub txtAngle_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call txtAngle_LostFocus
        KeyAscii = 0
    End If
End Sub

'

Private Sub scrAlpha_Change()
    Me.txtAlpha.Text = scrAlpha.value
    Call SetColorPreview
End Sub

Private Sub scrAlpha_Scroll()
    Call scrAlpha_Change
End Sub

Private Sub txtAlpha_LostFocus()
    If IsNumeric(txtAlpha.Text) Then
        Dim value As Integer: value = Abs(CInt(txtAlpha.Text))
        If (value >= 0 And value <= 255) Then
            Me.scrAlpha.value = value
            Call SetColorPreview
        End If
    End If
    Me.txtAlpha.Text = Me.scrAlpha.value
End Sub

Private Sub txtAlpha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call txtAlpha_LostFocus
        KeyAscii = 0
    End If
End Sub

'

Private Sub scrRed_Change()
    Me.txtRed.Text = scrRed.value
    Call SetColorPreview
End Sub

Private Sub scrRed_Scroll()
    Call scrRed_Change
    Call SetColorPreview
End Sub

Private Sub txtRed_LostFocus()
    If IsNumeric(txtRed.Text) Then
        Dim value As Integer: value = Abs(CInt(txtRed.Text))
        If (value >= 0 And value <= 255) Then
            Me.scrRed.value = value
            Call SetColorPreview
        End If
    End If
    Me.txtRed.Text = Me.scrRed.value
End Sub

Private Sub txtRed_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call txtRed_LostFocus
        KeyAscii = 0
    End If
End Sub

'

Private Sub scrGreen_Change()
    Me.txtGreen.Text = scrGreen.value
    Call SetColorPreview
End Sub

Private Sub scrGreen_Scroll()
    Call scrGreen_Change
End Sub

Private Sub txtGreen_LostFocus()
    If IsNumeric(txtGreen.Text) Then
        Dim value As Integer: value = Abs(CInt(txtGreen.Text))
        If (value >= 0 And value <= 255) Then
            Me.scrGreen.value = value
            Call SetColorPreview
        End If
    End If
    Me.txtGreen.Text = Me.scrGreen.value
End Sub

Private Sub txtGreen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call txtGreen_LostFocus
        KeyAscii = 0
    End If
End Sub

'

Private Sub scrBlue_Change()
    Me.txtBlue.Text = scrBlue.value
    Call SetColorPreview
End Sub

Private Sub scrBlue_Scroll()
    Call scrBlue_Change
End Sub

Private Sub txtBlue_LostFocus()
    If IsNumeric(txtBlue.Text) Then
        Dim value As Integer: value = Abs(CInt(txtBlue.Text))
        If (value >= 0 And value <= 255) Then
            Me.scrBlue.value = value
            Call SetColorPreview
        End If
    End If
    Me.txtBlue.Text = Me.scrBlue.value
End Sub

Private Sub txtBlue_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call txtBlue_LostFocus
        KeyAscii = 0
    End If
End Sub

'

Private Sub txtColorValue_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If IsNumeric(Me.txtColorValue.Text) Then
            Dim value As Long: value = CLng(Me.txtColorValue.Text)
            Dim Color As Graphics.ARGBColor: Color = GraphicEngine.Helper.GetARGB(value)
            Me.scrAlpha.value = Color.alpha
            Me.scrRed.value = Color.Red
            Me.scrGreen.value = Color.Green
            Me.scrBlue.value = Color.Blue
        Else
            Me.txtColorValue.Text = "0"
        End If
    End If
End Sub

Private Sub SetColorPreview()
    Me.txtColorValue.Text = CLng(GraphicEngine.Helper.SetARGB(Me.scrAlpha.value, Me.scrRed.value, Me.scrGreen.value, Me.scrBlue.value))
    Me.picColorPreview.BackColor = RGB(Me.scrRed.value, Me.scrGreen.value, Me.scrBlue.value)
    If Not Engine.Scene Is Nothing Then Engine.Scene.LevelEditor.TileEditor.Color = CLng(Me.txtColorValue.Text)
End Sub
