VERSION 5.00
Begin VB.Form frmMainMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TLSA Engine - Menu de edicion de niveles"
   ClientHeight    =   705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   47
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   600
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Control de ejecucion"
      ForeColor       =   &H8000000D&
      Height          =   735
      Left            =   6930
      TabIndex        =   10
      Top             =   0
      Width           =   2055
      Begin VB.CommandButton cmdRestart 
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   20.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   1500
         TabIndex        =   14
         Top             =   210
         Width           =   500
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   20.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   1020
         TabIndex        =   13
         Top             =   210
         Width           =   500
      End
      Begin VB.CommandButton cmdPause 
         Caption         =   ";"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   20.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   540
         TabIndex        =   12
         Top             =   210
         Width           =   500
      End
      Begin VB.CommandButton cmdRun 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   20.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   60
         TabIndex        =   11
         Top             =   210
         Width           =   500
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Visualizar"
      ForeColor       =   &H8000000D&
      Height          =   735
      Left            =   4380
      TabIndex        =   6
      Top             =   0
      Width           =   2535
      Begin VB.CheckBox chkShowBodies 
         Caption         =   "Cuerpos fisicos"
         Height          =   255
         Left            =   900
         TabIndex        =   15
         Top             =   450
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox chkViewGrid 
         Caption         =   "Rejilla de diseño"
         Height          =   255
         Left            =   900
         TabIndex        =   9
         Top             =   180
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox chkViewPhysics 
         Caption         =   "Fisica"
         Height          =   255
         Left            =   60
         TabIndex        =   8
         Top             =   450
         Value           =   1  'Checked
         Width           =   825
      End
      Begin VB.CheckBox chkViewTiles 
         Caption         =   "Tiles"
         Height          =   255
         Left            =   60
         TabIndex        =   7
         Top             =   180
         Value           =   1  'Checked
         Width           =   825
      End
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Entidad"
      Height          =   720
      Index           =   2
      Left            =   3630
      Picture         =   "frmMainMenu.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   0
      Width           =   720
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Fisicas"
      Height          =   720
      Index           =   1
      Left            =   2910
      Picture         =   "frmMainMenu.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   0
      Width           =   720
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Tiles"
      Height          =   720
      Index           =   0
      Left            =   2190
      Picture         =   "frmMainMenu.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Value           =   -1  'True
      Width           =   720
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Guardar"
      Height          =   720
      Index           =   2
      Left            =   1440
      Picture         =   "frmMainMenu.frx":1E5E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   720
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Abrir"
      Height          =   720
      Index           =   1
      Left            =   720
      Picture         =   "frmMainMenu.frx":2B28
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   720
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Nuevo"
      Height          =   720
      Index           =   0
      Left            =   0
      Picture         =   "frmMainMenu.frx":37F2
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   720
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkShowBodies_Click()
    Engine.Scene.ShowPhysicBodies = CBool(chkShowBodies.value)
End Sub

Private Sub chkViewGrid_Click()
    Engine.Scene.LevelEditor.Grid.Visible = CBool(chkViewGrid.value)
End Sub

Private Sub chkViewPhysics_Click()
    Engine.PhysicEngine.DEBUG_DrawColliders = CBool(chkViewPhysics.value)
End Sub

Private Sub chkViewTiles_Click()
    Engine.Scene.ShowTiles = CBool(chkViewTiles.value)
End Sub

Private Sub cmdPause_Click()
    Call Editor.Pause
End Sub

Private Sub cmdRestart_Click()
    Call Editor.Restart
End Sub

Private Sub cmdRun_Click()
    Call Editor.Run
End Sub

Private Sub cmdStop_Click()
    Call Editor.Terminate
End Sub

Private Sub Command1_Click(Index As Integer)
    
    Select Case Index
        Case 0 ' Nuevo escenario.
            If MsgBox("¿Estas seguro de iniciar un nuevo escenario? Todo el contenido del escenario actual se perdera si no lo has guardado previamente.", vbExclamation + vbYesNo, "Nuevo escenario") = vbYes Then
                Call Engine.Scene.Clear
            End If
        Case 1 ' Cargar escenario.
            Dim openDlg As New core.OpenDialog
            openDlg.Title = "Cargar escenario"
            Call openDlg.AddFilter("Escenario de TLSA Engine (*.tlv)", "*.TLV")
            openDlg.StartPath = App.Path & TLSA.ResourcePaths.Levels
            If openDlg.Show() Then Call Engine.Scene.LoadScene(openDlg.filename)
            Set openDlg = Nothing
            '--
            Engine.Scene.PhysicSimulator.Enabled = False
            '--
        Case 2 ' Guardar escenario.
            Dim saveDlg As New core.SaveDialog
            saveDlg.Title = "Guardar escenario"
            Call saveDlg.AddFilter("Escenario de TLSA Engine (*.tlv)", "*.TLV")
            saveDlg.StartPath = App.Path & TLSA.ResourcePaths.Levels
            If saveDlg.Show() Then Call Engine.Scene.LevelEditor.Export(saveDlg.filename)
            Call MsgBox("Escenario guardado satisfactoriamente.", vbInformation, "Guardar escenario")
            Set saveDlg = Nothing
    End Select
End Sub

Private Sub Option1_Click(Index As Integer)
    Select Case Index
        Case 0
            Call Engine.Scene.LevelEditor.SetEditor(TileEdition)
        Case 1
            Call Engine.Scene.LevelEditor.SetEditor(PhysicEdition)
        Case 2
            Call Engine.Scene.LevelEditor.SetEditor(EntityEdition)
    End Select
End Sub
