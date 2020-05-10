VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TLSA SDK: Editor de definicion de fuentes de texto"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   299
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   427
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Height          =   1275
      Left            =   60
      TabIndex        =   16
      Top             =   600
      Width           =   375
      Begin VB.OptionButton optFont 
         Caption         =   "Option1"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   18
         Top             =   930
         Width           =   195
      End
      Begin VB.OptionButton optFont 
         Caption         =   "Option1"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   17
         Top             =   330
         Value           =   -1  'True
         Width           =   195
      End
   End
   Begin VB.PictureBox picToolBar 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   570
      Left            =   0
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   425
      TabIndex        =   13
      Top             =   0
      Width           =   6405
      Begin VB.CommandButton cmdToolBar 
         BackColor       =   &H00E1DAD0&
         Height          =   570
         Index           =   3
         Left            =   5820
         Picture         =   "frmMain.frx":09EE
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Acerca de esta aplicacion"
         Top             =   0
         Width           =   570
      End
      Begin VB.CommandButton cmdToolBar 
         BackColor       =   &H00FFFFFF&
         Height          =   570
         Index           =   0
         Left            =   0
         Picture         =   "frmMain.frx":12B8
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Crear nueva fuente"
         Top             =   0
         Width           =   570
      End
      Begin VB.CommandButton cmdToolBar 
         BackColor       =   &H00FFFFFF&
         Height          =   570
         Index           =   1
         Left            =   540
         Picture         =   "frmMain.frx":1B82
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Importar fuente"
         Top             =   0
         Width           =   570
      End
      Begin VB.CommandButton cmdToolBar 
         BackColor       =   &H00FFFFFF&
         Height          =   570
         Index           =   2
         Left            =   1080
         Picture         =   "frmMain.frx":284C
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Guardar fuente"
         Top             =   0
         Width           =   570
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3000
      Top             =   2032
   End
   Begin VB.Frame Frame3 
      Caption         =   "Estilo"
      ForeColor       =   &H8000000D&
      Height          =   1275
      Left            =   3540
      TabIndex        =   4
      Top             =   600
      Width           =   2805
      Begin VB.CommandButton cmdApply 
         Caption         =   "&Aplicar"
         Height          =   435
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Width           =   1275
      End
      Begin VB.CheckBox chkStrikethrough 
         Caption         =   "Tachado"
         Height          =   195
         Left            =   1560
         TabIndex        =   11
         Top             =   960
         Width           =   1095
      End
      Begin VB.CheckBox chkUnderline 
         Caption         =   "Subrayado"
         Height          =   195
         Left            =   1560
         TabIndex        =   10
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox chkItalic 
         Caption         =   "Cursiva"
         Height          =   195
         Left            =   1560
         TabIndex        =   9
         Top             =   480
         Width           =   1095
      End
      Begin VB.CheckBox chkBold 
         Caption         =   "Negrita"
         Height          =   195
         Left            =   1560
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.VScrollBar vScrollSize 
         Height          =   315
         Left            =   1140
         Max             =   6
         Min             =   96
         TabIndex        =   7
         Top             =   240
         Value           =   10
         Width           =   255
      End
      Begin VB.TextBox txtSize 
         Height          =   315
         Left            =   780
         TabIndex        =   6
         Text            =   "10"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Tamaño"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   270
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Fuentes externas"
      Enabled         =   0   'False
      ForeColor       =   &H8000000D&
      Height          =   675
      Left            =   390
      TabIndex        =   2
      Top             =   1200
      Width           =   3195
      Begin VB.CommandButton cmdLoadFile 
         Caption         =   "..."
         Height          =   305
         Left            =   2760
         TabIndex        =   22
         Top             =   270
         Width           =   315
      End
      Begin VB.TextBox txtFilename 
         Height          =   285
         Left            =   120
         TabIndex        =   21
         Top             =   270
         Width           =   2685
      End
   End
   Begin VB.PictureBox picPreview 
      BackColor       =   &H8000000C&
      Height          =   2340
      Left            =   60
      ScaleHeight     =   2280
      ScaleWidth      =   6240
      TabIndex        =   1
      Top             =   2100
      Width           =   6300
   End
   Begin VB.Frame Frame1 
      Caption         =   "Fuentes del sistema"
      ForeColor       =   &H8000000D&
      Height          =   675
      Left            =   390
      TabIndex        =   0
      Top             =   600
      Width           =   3195
      Begin VB.ComboBox cbxSystemFonts 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   2955
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Previsualizacion"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   60
      TabIndex        =   12
      Top             =   1890
      Width           =   1335
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const stringSample As String = "abcdedgh" & vbNewLine & "ABCDEFGH" & vbNewLine & "0123456789" & vbNewLine & "ºª\!|""@·#$~%€&¬"

Private Sub cmdApply_Click()
    If gfx.Fonts.Exists("sample") Then Call gfx.Fonts.UnloadFont("sample")
    If optFont(0).Value Then
        Call modFontDefEditor.LoadSystemFont
    Else
        Call modFontDefEditor.LoadFileFont(Me.txtFilename.Text)
    End If
    
    Call sample.Initialize(gfx.Fonts("sample"), stringSample)
End Sub

Private Sub cmdLoadFile_Click()
    'Dim Filename As String
    'Filename = sys.DLG_OpenFile(Me.hWnd, "Archivo de fuentes TrueType (*.TTF)|*.ttf", "Importar fuente externa al sistema", "")
    Dim openFile As New Core.OpenDialog
    openFile.Title = "Importar fuente externa al sistema"
    Call openFile.AddFilter("Archivo de fuentes TrueType (*.TTF)", "*.ttf")
    'If Not Filename = "" Then
    If openFile.Show() Then
        Call modFontDefEditor.LoadFileFont(openFile.Filename)
        Call cmdApply_Click
    End If
    Set openFile = Nothing
End Sub

Private Sub cmdToolBar_Click(Index As Integer)
    Call picToolBar.SetFocus
    
    Dim Filename As String
    Select Case Index
        Case 0
            Me.optFont(0).Value = True
            Me.optFont(1).Value = False
            Me.cbxSystemFonts.ListIndex = IndexOf("System")
            Me.txtFilename.Text = ""
            Me.txtSize.Text = "10"
            Me.chkBold.Value = False
            Me.chkItalic.Value = False
            Me.chkUnderline.Value = False
            Me.chkStrikethrough.Value = False
        Case 1
            'Filename = sys.DLG_OpenFile(Me.hWnd, "Archivo de definicion de fuentes (*.FNT)|*.fnt", "Cargar defincion de fuente", "")
            Dim openFile As New Core.OpenDialog
            openFile.Title = "Cargar definicion de fuente"
            Call openFile.AddFilter("Archivo de definicion de fuentes (*.FNT)", "*.fnt")
            If openFile.Show() Then
            'If Not Filename = "" Then
                Call modFontDefEditor.ImportFont(openFile.Filename)
                With gfx.Fonts("sample")
                    Me.optFont(0).Value = CInt(Not .LoadFromFile)
                    Me.optFont(1).Value = CInt(.LoadFromFile)
                    If Me.optFont(0).Value Then Me.cbxSystemFonts.ListIndex = IndexOf(.Name)
                    Me.txtFilename.Text = .Filename
                    Me.txtSize.Text = CStr(.Size)
                    Me.chkBold.Value = Abs(CInt(.Bold))
                    Me.chkItalic.Value = Abs(CInt(.Italic))
                    Me.chkUnderline.Value = Abs(CInt(.Underline))
                    Me.chkStrikethrough.Value = Abs(CInt(.Strikethrough))
                End With
            End If
            Set openFile = Nothing
        Case 2
            Dim saveFile As New Core.SaveDialog
            saveFile.Title = "Crear definicion de fuente"
            Call saveFile.AddFilter("Archivo de definicion de fuentes (*.FNT)", "*.fnt")
            'Filename = sys.DLG_SaveFile(Me.hWnd, "Archivo de definicion de fuentes (*.FNT)|*.fnt", "Crear defincion de fuente", "")
            If saveFile.Show() Then Call gfx.Fonts("sample").Export(saveFile.Filename)
            Set saveFile = Nothing
        Case 3
            Const txtAbout As String = "Acerca del proyecto TLSA y sus componentes:" & vbNewLine & vbNewLine & _
                               "TLSA SDK © José Miguel Sánchez Fernández - 2006/2009" & vbNewLine & _
                               "TLSA Engine © José Miguel Sánchez Fernández - 2004/2009" & vbNewLine & _
                               "TLSA © José Miguel Sánchez Fernández - 2001/2009" & vbNewLine & _
                               "dx_lib32 © José Miguel Sánchez Fernández - 2001/2004/2009"
            Call MsgBox(txtAbout, vbInformation, "Acerca de " & Me.Caption)
    End Select
    Call cmdApply_Click
End Sub

' Devuelve el indice de la fuente en la lista:
Private Function IndexOf(FontName As String)
    Dim i As Integer
    For i = 0 To Me.cbxSystemFonts.ListCount - 1
        If Me.cbxSystemFonts.List(i) = FontName Then
            IndexOf = i
            Exit For
        End If
    Next
End Function

Private Sub Form_Load()
    ' Instanciamos dx_lib32:
    Set gfx = New Graphics.Manager
    'Set sys = New dx_System_Class
    
    ' Inicializamos el framework:
    Call gfx.Initialize(Me.picPreview.hWnd, 420, 128, 32, True, True)
    
    ' Corregimos la posicion del control de previsualizacion:
    With Me.picPreview
        .Left = 2: .Top = 140
    End With
    
    ' Cargamos la lista de fuentes del sistema que esten instaladas:
    Call LoadSystemFonts
    
    ' Aplicamos la fuente System:
    Me.cbxSystemFonts.ListIndex = IndexOf("System")
    Call cmdApply_Click
End Sub

Private Sub LoadSystemFonts()
    Dim i As Integer
    For i = 0 To Screen.FontCount - 1
        Call cbxSystemFonts.AddItem(Screen.Fonts(i))
    Next
    cbxSystemFonts.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set gfx = Nothing
    'Set sys = Nothing
End Sub

Private Sub optFont_Click(Index As Integer)
    Select Case Index
        Case 0
            Frame1.Enabled = True
            Frame2.Enabled = False
        Case 1
            Frame1.Enabled = False
            Frame2.Enabled = True
    End Select
End Sub

Private Sub Timer1_Timer()
    Call sample.Draw
    Call gfx.Render
End Sub

Private Sub txtSize_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = 8 Then KeyAscii = 0
End Sub

Private Sub txtSize_LostFocus()
    If Not IsNumeric(txtSize.Text) Or txtSize.Text = "" Or CInt(txtSize.Text) < 6 Or CInt(txtSize.Text) > 96 Then txtSize.Text = 10
End Sub

Private Sub vScrollSize_Change()
    txtSize.Text = vScrollSize.Value
End Sub

Private Sub vScrollSize_GotFocus()
    Call txtSize.SetFocus
End Sub
