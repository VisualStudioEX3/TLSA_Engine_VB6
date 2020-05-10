VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TLSA SDK: Input Profile Editor"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   6375
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtUser 
      Height          =   315
      Left            =   1980
      TabIndex        =   4
      Top             =   600
      Width           =   4365
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
      ScaleWidth      =   423
      TabIndex        =   13
      Top             =   0
      Width           =   6375
      Begin VB.CommandButton cmdToolBar 
         BackColor       =   &H00FFFFFF&
         Height          =   570
         Index           =   2
         Left            =   1080
         Picture         =   "frmMain.frx":09EE
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Guardar fuente"
         Top             =   0
         Width           =   570
      End
      Begin VB.CommandButton cmdToolBar 
         BackColor       =   &H00FFFFFF&
         Height          =   570
         Index           =   1
         Left            =   540
         Picture         =   "frmMain.frx":16B8
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Importar fuente"
         Top             =   0
         Width           =   570
      End
      Begin VB.CommandButton cmdToolBar 
         BackColor       =   &H00FFFFFF&
         Height          =   570
         Index           =   0
         Left            =   0
         Picture         =   "frmMain.frx":2382
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Crear nueva fuente"
         Top             =   0
         Width           =   570
      End
      Begin VB.CommandButton cmdToolBar 
         BackColor       =   &H00E1DAD0&
         Height          =   570
         Index           =   3
         Left            =   5820
         Picture         =   "frmMain.frx":2C4C
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Acerca de esta aplicacion"
         Top             =   0
         Width           =   570
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Mapa de acciones"
      ForeColor       =   &H8000000D&
      Height          =   5235
      Left            =   0
      TabIndex        =   12
      Top             =   2580
      Width           =   6375
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   4140
         Top             =   2220
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Elimina&r"
         Height          =   315
         Left            =   1620
         TabIndex        =   8
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Editar"
         Height          =   315
         Left            =   870
         TabIndex        =   7
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Añadir"
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   735
      End
      Begin VB.ListBox lstActionMap 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4185
         Left            =   120
         TabIndex        =   9
         Top             =   900
         Width           =   6135
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Joystick/Gamepad"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   4800
         TabIndex        =   16
         Top             =   690
         Width           =   1335
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Teclado/Raton"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   3090
         TabIndex        =   15
         Top             =   690
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nombre de la accion"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   150
         TabIndex        =   14
         Top             =   690
         Width           =   1470
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Seleccione el gamepad que utilizara para configurar el perfil"
      ForeColor       =   &H8000000D&
      Height          =   1575
      Left            =   0
      TabIndex        =   10
      Top             =   990
      Width           =   6375
      Begin VB.ComboBox cbxGamepad 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1140
         Width           =   6135
      End
      Begin VB.Label Label1 
         Caption         =   $"frmMain.frx":3516
         Height          =   975
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   6105
      End
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Nombre del usuario o perfil"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   60
      TabIndex        =   17
      Top             =   660
      Width           =   1875
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const strAbout As String = "Esta herramienta le permite definir perfiles de mapas de acciones para los proyectos que utilicen " & _
                           "TLSA.Gameinput.dll. Con esta herramienta podra configurar la teclas del teclado, raton y/o joystick " & _
                           "o gamepad a las acciones que defina. Esto le permitira probar configuraciones distintas de perfiles de " & _
                           "forma rapida sin programar una sola linea de codigo adicional en su programa, simplemente cargar y listo."
                       
Const strCopyright As String = "TLSA.GameInput & TLSA SDK: Input profile editor © 2010 - José Miguel Sánchez Fernández"

Dim boolChange As Boolean ' Indica si se han realizado cambios en el perfil.

Private Sub cbxGamepad_Click()
    modMain.varProfile.GamepadUsed = cbxGamepad.ListIndex
End Sub

Private Sub cmdAdd_Click()
    If Me.lstActionMap.ListCount < 20 Then
        Dim f As New frmSetup
        Call f.CreateAction
        Set f = Nothing
    Else
        Call MsgBox("No se pueden agregar mas acciones al perfil.", vbExclamation, "Mapa de acciones completo")
    End If
End Sub

Private Sub cmdEdit_Click()
    Call lstActionMap_DblClick
End Sub

Private Sub cmdRemove_Click()
    Call modMain.varProfile.Actions.Remove(modMain.varProfile.Actions.GetValue(Me.lstActionMap.ListIndex + 1).Name)
    Call modMain.LoadActionMap(Me.lstActionMap)
End Sub

Private Sub cmdToolBar_Click(Index As Integer)
    Dim ret As Long, strFilename As String
    Select Case Index
        Case 0 ' Nuevo perfil.
            ret = MsgBox("¿Crear nuevo perfil? Para continuar editando el actual pulse No.", vbYesNo + vbExclamation, "Nuevo perfil")
            If ret = vbYes Then
                Set modMain.varProfile = modMain.varInput.Profiles.Create("Noname", gameinput.PlayerIndex.Player1, gameinput.InputDevice.KeybAndMouse)
                Me.txtUser.Text = modMain.varProfile.UserName
            End If
            
        Case 1 ' Cargar perfil.
            Dim openDLG As New Core.OpenDialog
            With openDLG
                Call .AddFilter("TLSA Game Input Profile (*.PRF)", "*.prf")
                Call .AddFilter("Todos los archivos (*.*)", "*.*")
                .Title = "Cargar perfil de TLSA.GameInput"
                Call .Show
                If .FileName <> "" Then
                    Call modMain.varProfile.Import(.FileName)
                    Me.txtUser.Text = modMain.varProfile.UserName
                End If
            End With
            Set openDLG = Nothing
        Case 2 ' Guardar perfil.
            Dim saveDLG As New Core.SaveDialog
            With saveDLG
                Call .AddFilter("TLSA Game Input Profile (*.PRF)", "*.prf")
                Call .AddFilter("Todos los archivos (*.*)", "*.*")
                .Title = "Guardar perfil de TLSA.GameInput"
                Call .Show
                If .FileName <> "" Then
                    Call modMain.varProfile.Export(.FileName)
                End If
            End With
            Set saveDLG = Nothing
        Case 3 ' Acerca de...
            Call MsgBox(strAbout & vbNewLine & vbNewLine & strCopyright, vbInformation, "Acerca de TLSA SDK: Input profile editor")
    End Select
    Call modMain.LoadActionMap(Me.lstActionMap)
End Sub

Private Sub Form_Load()
    ' Creamos la instancia del gestor de dispositivos de entrada y la inicializamos:
    Set modMain.varInput = New gameinput.Manager
    Call modMain.varInput.SetWindowHandle(Me.hWnd)
    
    ' Cargamos en la lista los gamepads que esten conectados:
    Call modMain.LoadGamepads(Me.cbxGamepad)
        
    ' Creamos un nuevo perfil:
    Set modMain.varProfile = modMain.varInput.Profiles.Create("Noname", gameinput.PlayerIndex.Player1, gameinput.InputDevice.KeybAndMouse)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set modMain.varProfile = Nothing
    Set modMain.varInput = Nothing
    End
End Sub

Private Sub lstActionMap_DblClick()
    If Not lstActionMap.ListIndex = -1 Then
        Dim f As New frmSetup
        Call f.EditAction(modMain.varProfile.Actions.GetValue(lstActionMap.ListIndex + 1).Name)
        Set f = Nothing
    End If
End Sub

Private Sub Timer1_Timer()
    On Error Resume Next
    modMain.varProfile.Device = gameinput.InputDevice.Gamepad
    Call modMain.varInput.Update
'    Debug.Print modMain.varProfile.ViewAngle()
    If Not modMain.varProfile.VirtualCursor Then
        modMain.varProfile.VirtualCursor = True
        modMain.varProfile.VirtualCursorSensitivity = 1
    End If
    Debug.Print "x" & modMain.varProfile.ViewAxis.X & " y" & modMain.varProfile.ViewAxis.Y
End Sub

Private Sub txtUser_LostFocus()
    modMain.varProfile.UserName = Me.txtUser.Text
End Sub
