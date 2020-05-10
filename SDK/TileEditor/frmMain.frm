VERSION 5.00
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "TLSA SDK: Texture Studio"
   ClientHeight    =   8430
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   12495
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picMainPanel 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   8055
      Left            =   0
      ScaleHeight     =   8055
      ScaleWidth      =   3645
      TabIndex        =   24
      Top             =   375
      Width           =   3645
      Begin VB.CommandButton Command1 
         Caption         =   "Editor de puntos de control"
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   60
         TabIndex        =   37
         Top             =   6480
         Width           =   3525
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Editor de secuencias de animacion"
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   60
         TabIndex        =   19
         Top             =   6780
         Width           =   3525
      End
      Begin VB.Frame Frame2 
         Caption         =   "Propiedades del tile"
         ForeColor       =   &H8000000D&
         Height          =   2295
         Left            =   60
         TabIndex        =   27
         Top             =   4170
         Width           =   3555
         Begin VB.Frame Frame4 
            Caption         =   "Region de la textura"
            ForeColor       =   &H8000000D&
            Height          =   645
            Left            =   60
            TabIndex        =   31
            Top             =   570
            Width           =   3435
            Begin VB.TextBox txtX 
               Enabled         =   0   'False
               Height          =   285
               Left            =   240
               TabIndex        =   11
               Top             =   240
               Width           =   500
            End
            Begin VB.TextBox txtY 
               Enabled         =   0   'False
               Height          =   285
               Left            =   990
               TabIndex        =   12
               Top             =   240
               Width           =   500
            End
            Begin VB.TextBox txtWidth 
               Enabled         =   0   'False
               Height          =   285
               Left            =   2010
               TabIndex        =   13
               Top             =   240
               Width           =   500
            End
            Begin VB.TextBox txtHeight 
               Enabled         =   0   'False
               Height          =   285
               Left            =   2880
               TabIndex        =   14
               Top             =   240
               Width           =   500
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "X."
               Height          =   195
               Left            =   90
               TabIndex        =   35
               Top             =   300
               Width           =   150
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Y."
               Height          =   195
               Left            =   840
               TabIndex        =   34
               Top             =   300
               Width           =   150
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Ancho."
               Height          =   195
               Left            =   1515
               TabIndex        =   33
               Top             =   300
               Width           =   510
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Alto."
               Height          =   195
               Left            =   2580
               TabIndex        =   32
               Top             =   300
               Width           =   315
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Origen de dibujo del tile (OffSet)"
            ForeColor       =   &H8000000D&
            Height          =   645
            Left            =   60
            TabIndex        =   28
            Top             =   1260
            Width           =   3435
            Begin VB.CommandButton cmdSetCenter 
               Caption         =   "Marcar"
               Enabled         =   0   'False
               Height          =   315
               Left            =   1590
               TabIndex        =   17
               Top             =   240
               Width           =   1755
            End
            Begin VB.TextBox txtCenterX 
               Enabled         =   0   'False
               Height          =   285
               Left            =   240
               TabIndex        =   15
               Top             =   240
               Width           =   500
            End
            Begin VB.TextBox txtCenterY 
               Enabled         =   0   'False
               Height          =   285
               Left            =   990
               TabIndex        =   16
               Top             =   240
               Width           =   500
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "X."
               Height          =   195
               Left            =   90
               TabIndex        =   30
               Top             =   300
               Width           =   150
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Y."
               Height          =   195
               Left            =   840
               TabIndex        =   29
               Top             =   300
               Width           =   150
            End
         End
         Begin VB.TextBox txtKey 
            Enabled         =   0   'False
            Height          =   285
            Left            =   510
            TabIndex        =   10
            Top             =   240
            Width           =   2985
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Actualizar"
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   60
            TabIndex        =   18
            Top             =   1950
            Width           =   3435
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Clave"
            Height          =   195
            Left            =   60
            TabIndex        =   36
            Top             =   300
            Width           =   405
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Lista de tiles"
         ForeColor       =   &H8000000D&
         Height          =   4155
         Left            =   60
         TabIndex        =   25
         Top             =   0
         Width           =   3555
         Begin VB.CommandButton cmdRemove 
            Caption         =   "Eliminar"
            Height          =   435
            Left            =   780
            TabIndex        =   9
            Top             =   3660
            Width           =   700
         End
         Begin VB.ListBox lstTiles 
            Height          =   3375
            Left            =   60
            TabIndex        =   7
            Top             =   240
            Width           =   3435
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Añadir"
            Height          =   435
            Left            =   60
            TabIndex        =   8
            Top             =   3660
            Width           =   700
         End
         Begin VB.Label lblCountIndex 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "0 de 0"
            Height          =   195
            Left            =   2985
            TabIndex        =   26
            Top             =   3750
            Width           =   450
         End
      End
   End
   Begin VB.PictureBox picToolBar 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   12495
      TabIndex        =   20
      Top             =   0
      Width           =   12495
      Begin VB.CommandButton cmdHelp 
         Height          =   375
         Left            =   12120
         Picture         =   "frmMain.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         Width           =   375
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   7440
         ScaleHeight     =   315
         ScaleWidth      =   4695
         TabIndex        =   22
         Top             =   60
         Width           =   4695
         Begin VB.OptionButton Option1 
            Caption         =   "4x"
            Height          =   225
            Index           =   2
            Left            =   1800
            TabIndex        =   5
            Top             =   0
            Width           =   615
         End
         Begin VB.OptionButton Option1 
            Caption         =   "2x"
            Height          =   225
            Index           =   1
            Left            =   1230
            TabIndex        =   4
            Top             =   0
            Width           =   615
         End
         Begin VB.CommandButton cmdBackColor 
            Height          =   375
            Index           =   3
            Left            =   4260
            Picture         =   "frmMain.frx":06EA
            Style           =   1  'Graphical
            TabIndex        =   41
            Top             =   -60
            Width           =   375
         End
         Begin VB.CommandButton cmdBackColor 
            Height          =   375
            Index           =   2
            Left            =   3900
            Picture         =   "frmMain.frx":0A2C
            Style           =   1  'Graphical
            TabIndex        =   40
            Top             =   -60
            Width           =   375
         End
         Begin VB.CommandButton cmdBackColor 
            Height          =   375
            Index           =   1
            Left            =   3540
            Picture         =   "frmMain.frx":0D6E
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   -60
            Width           =   375
         End
         Begin VB.CommandButton cmdBackColor 
            Height          =   375
            Index           =   0
            Left            =   3180
            Picture         =   "frmMain.frx":10B0
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   -60
            Width           =   375
         End
         Begin VB.OptionButton Option1 
            Caption         =   "1x"
            Height          =   225
            Index           =   0
            Left            =   660
            TabIndex        =   3
            Top             =   0
            Value           =   -1  'True
            Width           =   615
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Fondo:"
            Height          =   195
            Left            =   2640
            TabIndex        =   42
            Top             =   0
            Width           =   495
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Zoom:"
            Height          =   195
            Left            =   120
            TabIndex        =   23
            Top             =   0
            Width           =   450
         End
      End
      Begin VB.CommandButton cmdLoad 
         Caption         =   "Cargar"
         Height          =   315
         Left            =   0
         TabIndex        =   0
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Salvar"
         Enabled         =   0   'False
         Height          =   315
         Left            =   810
         TabIndex        =   1
         Top             =   0
         Width           =   795
      End
      Begin VB.TextBox txtFilename 
         BackColor       =   &H8000000F&
         Height          =   345
         Left            =   1620
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   0
         Width           =   4305
      End
      Begin VB.Label lblInfo 
         Caption         =   "Tamaño:  "
         Height          =   255
         Left            =   6000
         TabIndex        =   21
         Top             =   60
         Width           =   1875
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
    selCset = False
    If Not Tex Is Nothing Then
        Dim Key As String
        Dim t As New Graphics.Tile
        
        Key = InputBox("Introduzca un nombre para identificar al tile en la lista:", "Añadir nuevo tile", "Nuevo tile " & Me.lstTiles.ListCount)
        If Not Key = "" Then
            If Not Tex.Tiles.Exists(Key) Then
                'Call t.SetTexture(Tex)
                'Call Tex.Tiles.Add(t, Key)
                Call Tex.Tiles.Create(Key, Tex, 0, 0, Tex.Information.Texture.Width, Tex.Information.Texture.Height, 0, 0)
                Call Me.lstTiles.AddItem(Key) ' Añadimos el nombre a la lista del editor.
                Me.lstTiles.ListIndex = Me.lstTiles.ListCount - 1
            End If
        End If
    End If
    
    Call frmVisor.SetFocus
End Sub

Private Sub cmdBackColor_Click(Index As Integer)
    Select Case Index
        Case 0: Alpha.Color = Graphics.Color_Constant.White
        Case 1: Alpha.Color = Graphics.Color_Constant.Blue
        Case 2: Alpha.Color = Graphics.Color_Constant.Green
        Case 3: Alpha.Color = Graphics.Color_Constant.Magenta
    End Select
End Sub

Private Sub cmdHelp_Click()
    Call frmHelp.Show
    Call frmVisor.SetFocus
End Sub

Private Sub cmdRemove_Click()
    selCset = False
    If Not Tex Is Nothing Then
        If MsgBox("¿Borrar el tile de la lista?", vbExclamation + vbYesNo, "Borrar tile") = vbYes Then
            Call Tex.Tiles.Remove(Me.lstTiles.List(Me.lstTiles.ListIndex))
            Call modMain.LoadTiles
            If Tex.Tiles.Count = 0 Then
                Me.txtKey.Text = ""
                Me.txtX.Text = "0"
                Me.txtY.Text = "0"
                Me.txtWidth.Text = "0"
                Me.txtHeight.Text = "0"
                Me.txtCenterX.Text = "0"
                Me.txtCenterY.Text = "0"
            Else
                Me.lstTiles.ListIndex = Me.lstTiles.ListCount - 1
            End If
        End If
    End If
    
    Call frmVisor.SetFocus
End Sub

Private Sub cmdSave_Click()
    selCset = False
    If Not Tex Is Nothing Then
        Call Tex.ExportTiles ' Generamos el archivo con las definiciones de tiles y sus puntos de control.
        Call MsgBox("Archivo de definicion de tiles generado.", vbInformation, "Tiles generados")
        
        If Spr.Animations.Count > 0 Then
            Call Gfx.Helper.SaveAnimatedSprite(Spr)
            Call MsgBox("Animaciones exportadas con exito.", vbInformation, "Secuencias de animacion")
        End If
    End If
End Sub

Private Sub cmdSetCenter_Click()
    selCset = True
    Call frmVisor.SetFocus
End Sub

' Actualiza los valores introducidos manualmente:
Private Sub Command1_Click(Index As Integer)
    If Not Tex Is Nothing Then
        Select Case Index
            Case 0
                selCset = False
                
                Dim t As Graphics.Tile
                Set t = Tex.Tiles(Me.txtKey.Tag)
                
                Call t.SetRegion(CLng(Me.txtX.Text), CLng(Me.txtY.Text), CLng(Me.txtWidth.Text), CLng(Me.txtHeight.Text))
                Call t.SetCenter(CLng(Me.txtCenterX.Text), CLng(Me.txtCenterY.Text))
                
                Set t = Nothing
                
                If Me.txtKey.Text <> Me.txtKey.Tag Then Call modMain.TileChangeKey(Me.txtKey.Text)
                
                Call modMain.ShowTileData(lstTiles.List(lstTiles.ListIndex))
            Case 1
                Call frmPointControl.Show
            Case 2
'                Me.Enabled = False
                Call frmQuickAnimEd.Show
        End Select
    End If
End Sub

Private Sub lstTiles_Click()
    Dim isDef As Boolean
    isDef = (lstTiles.ListIndex = 0)
    
    Me.cmdRemove.Enabled = Not isDef
    Me.cmdSetCenter.Enabled = Not isDef
    Me.txtKey.Enabled = Not isDef
    Me.txtX.Enabled = Not isDef
    Me.txtY.Enabled = Not isDef
    Me.txtWidth.Enabled = Not isDef
    Me.txtHeight.Enabled = Not isDef
    Me.txtCenterX.Enabled = Not isDef
    Me.txtCenterY.Enabled = Not isDef
    
    frmPointControl.lstPoints.Enabled = Not isDef
    frmPointControl.cmdAdd.Enabled = Not isDef
    frmPointControl.cmdRemove.Enabled = Not isDef
    frmPointControl.cmdClear.Enabled = Not isDef
    frmPointControl.cmdSetPoint.Enabled = Not isDef
    frmPointControl.cmdTools(0).Enabled = Not isDef
    frmPointControl.cmdUpdate.Enabled = Not isDef
    frmPointControl.txtKey.Enabled = Not isDef
    frmPointControl.txtX.Enabled = Not isDef
    frmPointControl.txtY.Enabled = Not isDef
    
    If isDef Then
        Me.txtKey.Text = "Default"
        Me.txtX.Text = "0"
        Me.txtY.Text = "0"
        Me.txtWidth.Text = Tex.Information.Image.Width
        Me.txtHeight.Text = Tex.Information.Image.Height
        Me.txtCenterX.Text = "0"
        Me.txtCenterY.Text = "0"
        
        selA.X = 0: selA.Y = 0
        selB.X = 0: selB.Y = 0
        selC.X = 0: selC.Y = 0
        
        Call frmPointControl.lstPoints.Clear
        Call frmPointControl.lstPoints.AddItem("Default")
        frmPointControl.txtKey.Text = ""
        frmPointControl.txtX.Text = ""
        frmPointControl.txtY.Text = ""
        
    Else
        If Not selAset And Not selCset And Not selPset And lstTiles.ListIndex > -1 Then
            Call modMain.ShowTileData(lstTiles.List(lstTiles.ListIndex))
            Call frmPointControl.SelDefault
        End If
    End If
    
'    If Not selAset And Not selCset And Not selPset And lstTiles.ListIndex > -1 Then
'        Call modMain.ShowTileData(lstTiles.List(lstTiles.ListIndex))
'        Call frmPointControl.SelDefault
'    End If
    
    frmMain.lblCountIndex.Caption = (frmMain.lstTiles.ListIndex + 1) & " de " & frmMain.lstTiles.ListCount
    
End Sub

Private Sub MDIForm_Load()
    Me.Caption = modMain.AppTitle
    
    Call frmVisor.Show
    
    Set Gfx = New Graphics.Manager
    Call Gfx.Initialize(frmVisor.hwnd, 256, 256, 32, True, True)
    Gfx.MaxFrames = 60
    Call CenterVisor
    
    Call Gfx.Textures.LoadTexture(App.Path & "\alpha.bmp", "alpha", False)
    Set Alpha = New Graphics.Sprite
    Call Alpha.SetTexture(Gfx.Textures("alpha"))
    
    cmdHelp.Left = Me.picToolBar.Width - cmdHelp.Width
    
    zoomScale = 1
    
    Do While Not Gfx Is Nothing
        Call modMain.Render
    Loop
End Sub

Private Sub MDIForm_Resize()
    Dim f As Form
    If Me.WindowState = 1 Then
        For Each f In Forms
            If Not f Is Me Then Call f.Hide
        Next
    Else
        For Each f In Forms
            If Not f Is Me Then Call f.Show
        Next
    End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    Call Gfx.Textures.UnloadAll
    Set Spr = Nothing
    Set Gfx = Nothing
    End
End Sub

Private Sub cmdLoad_Click()
    selAset = False
    selCset = False
    selPset = False
    
    Call LoadTexture
    
    If Not Tex Is Nothing Then
        On Error GoTo ErrOut
        Set Spr = Gfx.Helper.CreateAnimatedSprite(modMain.Tex)
        Exit Sub
    Else
        Exit Sub
    End If
    
ErrOut: ' Si fallo la carga del archivo de animacion es que este posiblemente no exista. Asociaremos la textura por defecto:
    Set Spr = New Graphics.Sprite
    Call Spr.SetTexture(modMain.Tex)
End Sub

Private Sub Option1_Click(Index As Integer)
    If Not Tex Is Nothing And Index = 0 Then Call SetZoom(1) Else Call SetZoom(Index * 2)
    Call frmVisor.SetFocus
End Sub

Private Sub Picture1_Resize()
    cmdHelp.Left = Me.picToolBar.Width - cmdHelp.Width
End Sub

Private Sub picToolBar_Resize()
    Me.cmdHelp.Left = Me.picToolBar.Width - Me.cmdHelp.Width
End Sub

Private Sub txtX_LostFocus()
    If Not IsNumeric(Me.txtX.Text) Then Me.txtX.Text = 0
End Sub

Private Sub txtY_LostFocus()
    If Not IsNumeric(Me.txtY.Text) Then Me.txtY.Text = 0
End Sub

Private Sub txtCenterX_LostFocus()
    If Not IsNumeric(Me.txtCenterX.Text) Then Me.txtCenterX.Text = 0
End Sub

Private Sub txtCenterY_LostFocus()
    If Not IsNumeric(Me.txtCenterY.Text) Then Me.txtCenterY.Text = 0
End Sub
