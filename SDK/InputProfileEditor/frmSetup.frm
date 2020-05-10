VERSION 5.00
Begin VB.Form frmSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configurar accion"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4755
   Icon            =   "frmSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   4755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picJoy 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2610
      ScaleHeight     =   255
      ScaleWidth      =   1665
      TabIndex        =   9
      Top             =   780
      Width           =   1725
   End
   Begin VB.PictureBox picKeyb 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2610
      ScaleHeight     =   255
      ScaleWidth      =   1665
      TabIndex        =   8
      Top             =   420
      Width           =   1725
   End
   Begin VB.CommandButton cmdResetJoy 
      Caption         =   "r"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4380
      TabIndex        =   7
      Top             =   780
      Width           =   315
   End
   Begin VB.CommandButton cmdResetKeyb 
      Caption         =   "r"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4380
      TabIndex        =   6
      Top             =   420
      Width           =   315
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancelar"
      Height          =   435
      Left            =   3480
      TabIndex        =   5
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Guardar"
      Height          =   435
      Left            =   2220
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtAction 
      Height          =   315
      Left            =   2610
      TabIndex        =   0
      Top             =   60
      Width           =   2085
   End
   Begin VB.Image Image1 
      Height          =   15
      Left            =   90
      Top             =   90
      Width           =   15
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Boton del joystick o gamepad"
      Height          =   195
      Left            =   75
      TabIndex        =   3
      Top             =   840
      Width           =   2085
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Tecla del teclado o boton del raton"
      Height          =   195
      Left            =   75
      TabIndex        =   2
      Top             =   510
      Width           =   2475
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nombre de la accion"
      Height          =   195
      Left            =   75
      TabIndex        =   1
      Top             =   150
      Width           =   1470
   End
End
Attribute VB_Name = "frmSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private varName As String                                                   ' Nombre original de la accion.
Private valKeyb As Long, valJoy As Long                                     ' Valores capturados.
Private namChange As Boolean, keybChange As Boolean, joyChange As Boolean   ' Estados de cambio.

Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)


Private Sub SetKeybConstant(Caption As String)
    With Me.picKeyb
        .Tag = Caption
        .Cls: Me.picKeyb.Print .Tag
        modMain.Keyb = .Tag
    End With
End Sub

Private Sub MarkKeybConstant(Mark As Boolean)
    With Me.picKeyb
        If Mark Then
            .BackColor = vbYellow
            .ForeColor = vbRed
        Else
            .BackColor = vbWhite
            .ForeColor = vbBlack
        End If
        .Cls: Me.picKeyb.Print .Tag
    End With
    
    Me.txtAction.Enabled = Not Mark
    Me.picJoy.Enabled = Not Mark
    Me.cmdSave.Enabled = Not Mark
    Me.cmdResetJoy.Enabled = Not Mark
End Sub

Private Sub SetJoyConstant(Caption As String)
    With Me.picJoy
        .Tag = Caption
        .Cls: Me.picJoy.Print .Tag
        modMain.Joy = .Tag
    End With
End Sub

Private Sub MarkJoyConstant(Mark As Boolean)
    With Me.picJoy
        If Mark Then
            .BackColor = vbYellow
            .ForeColor = vbRed
        Else
            .BackColor = vbWhite
            .ForeColor = vbBlack
        End If
        .Cls: Me.picJoy.Print .Tag
    End With
    
    Me.txtAction.Enabled = Not Mark
    Me.picKeyb.Enabled = Not Mark
    Me.cmdSave.Enabled = Not Mark
    Me.cmdResetKeyb.Enabled = Not Mark
End Sub

' Crea una nueva accion:
Public Sub CreateAction()
    modMain.Action = ""
    Call SetKeybConstant(modMain.varInput.KeyboadDictionary.GetKey(0))
    Call SetJoyConstant(modMain.varInput.GamepadDictionary.GetButton(0))
    modMain.NewAction = True
    
    Call frmSetup.Show(vbModal)
End Sub

' Carga los parametros de una accion para ser modificados:
Public Sub EditAction(Action As String)
    modMain.Action = Action
    Call SetKeybConstant(modMain.varInput.KeyboadDictionary.GetKey(modMain.varProfile.GetActionButton(Action, gameinput.InputDevice.KeybAndMouse)))
    Call SetJoyConstant(modMain.varInput.GamepadDictionary.GetButton(modMain.varProfile.GetActionButton(Action, gameinput.InputDevice.Gamepad)))
    modMain.NewAction = False
    
    Call frmSetup.Show(vbModal)
End Sub

Private Sub cmdResetKeyb_Click()
    valKeyb = 0
    Call SetKeybConstant(modMain.varInput.KeyboadDictionary.GetKey(0))
    Call MarkKeybConstant(False)
    keybChange = True
End Sub

Private Sub cmdResetJoy_Click()
    valJoy = 0
    Call SetJoyConstant(modMain.varInput.GamepadDictionary.GetButton(0))
    Call MarkJoyConstant(False)
    joyChange = True
End Sub

Private Sub cmdCancel_Click()
    Call Unload(Me)
End Sub

Private Sub cmdSave_Click()
    ' Creamos la accion:
    If modMain.NewAction Then
        On Error Resume Next    ' Try
        Call modMain.varProfile.Actions.Add(modMain.Action)
        If Err.Number <> 0 Then ' Catch
            Call MsgBox(Err.Description, vbCritical, "Error al crear accion")
            Exit Sub
        End If
        On Error GoTo 0
        
    ' Si el nombre cambio se intenta renombrar la accion:
    ElseIf namChange Then
        Call modMain.varProfile.Actions.Rename(varName, Me.txtAction.Text)
        modMain.Action = Me.txtAction.Text
    End If
    
    ' Modificamos los valores de las teclas asignadas:
    If keybChange Then Call modMain.varProfile.SetActionButton(modMain.Action, valKeyb, gameinput.InputDevice.KeybAndMouse)
    If joyChange Then Call modMain.varProfile.SetActionButton(modMain.Action, valJoy, gameinput.InputDevice.Gamepad)
    
    ' Actualizamos la lista:
    Call modMain.LoadActionMap(frmMain.lstActionMap)
    
    Call Unload(Me)
End Sub

Private Sub Form_Load()
    varName = modMain.Action
    namChange = False
    Me.txtAction.Text = modMain.Action
    Call SetKeybConstant(modMain.Keyb)
    Call SetJoyConstant(modMain.Joy)
    
    ' Cambiamos el foco del gestor de entrada al formulario de configuracion para realizar la captura:
    Call modMain.varInput.SetWindowHandle(Me.hWnd)
End Sub

Private Sub picKeyb_DblClick()
    Call MarkKeybConstant(True)
    Call Sleep(500) ' Evita que el eco (tiempo prolongado de pulsacion) del boton primario del raton se asigne como entrada en la accion por error.
    Do
        Call modMain.varInput.Update
        valKeyb = modMain.varProfile.Capture(gameinput.InputDevice.KeybAndMouse)
        If valKeyb <> 0 Then
            Call SetKeybConstant(modMain.varInput.KeyboadDictionary.GetKey(valKeyb))
            keybChange = True
            Exit Do
        End If
        DoEvents
    Loop
    Call MarkKeybConstant(False)
End Sub

Private Sub picJoy_DblClick()
    Call MarkJoyConstant(True)
    Do
        Call modMain.varInput.Update
        valJoy = modMain.varProfile.Capture(gameinput.InputDevice.Gamepad)
        If valJoy <> 0 Then
            Call SetJoyConstant(modMain.varInput.GamepadDictionary.GetButton(valJoy))
            joyChange = True
            Exit Do
        End If
        DoEvents
    Loop
    Call MarkJoyConstant(False)
End Sub

Private Sub txtAction_LostFocus()
    If Me.txtAction.Text <> modMain.Action Then
        modMain.Action = Me.txtAction.Text
        namChange = True
    End If
End Sub

