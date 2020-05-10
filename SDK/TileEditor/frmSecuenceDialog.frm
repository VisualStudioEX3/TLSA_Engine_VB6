VERSION 5.00
Begin VB.Form frmSecuenceDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Secuencia de animacion"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4905
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdHelp 
      Height          =   375
      Left            =   90
      Picture         =   "frmSecuenceDialog.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1350
      Width           =   375
   End
   Begin VB.TextBox txtSecuenceName 
      Height          =   315
      Left            =   2520
      TabIndex        =   1
      Top             =   450
      Width           =   2265
   End
   Begin VB.ComboBox cbxFormat 
      Height          =   315
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   810
      Width           =   2265
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   405
      Left            =   3720
      TabIndex        =   4
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   405
      Left            =   2520
      TabIndex        =   3
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox txtKey 
      Height          =   315
      Left            =   2520
      TabIndex        =   0
      Top             =   90
      Width           =   2265
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Nombre de la secuencia de tiles"
      Height          =   195
      Left            =   90
      TabIndex        =   8
      Top             =   510
      Width           =   2265
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Formato del contador"
      Height          =   195
      Left            =   90
      TabIndex        =   7
      Top             =   900
      Width           =   1500
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nombre de la animacion"
      Height          =   195
      Left            =   90
      TabIndex        =   6
      Top             =   150
      Width           =   1710
   End
End
Attribute VB_Name = "frmSecuenceDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Call Unload(Me)
End Sub

Private Sub cmdHelp_Click()
    Call MsgBox("Esta utilidad permite importar una serie de tiles con el mismo nombre usando un contador consecutivo. Si por ejemplo tuviesemos la secuencia de animacion de caminar definida con los tiles: 'saltar_00', 'saltar_01', ..., 'saltar_09' y 'saltar_10' los parametros a introducir serian 'saltar_' como nombre de la secuencia y el formato '0#' para identificar el contador en el nombre de los tiles. La serie siempre ha de empezar por 0, de lo contrario no se importara ningun tile a la secuencia de animacion.", vbInformation, "Ayuda rapida")
End Sub

Private Sub cmdOk_Click()
    If Spr.Animations.Exists(Me.txtKey.Text) Then
        Call MsgBox("Ya existe una secuencia con el mismo nombre.", vbCritical, "Nombre en uso")
    Else
        Call modAnimEd.AddSecuence(Me.txtKey.Text, Me.txtSecuenceName.Text, Me.cbxFormat.Text)
        Call modAnimEd.LoadSecuences
        Call Unload(Me)
    End If
End Sub

Private Sub Form_Load()
    Call Me.cbxFormat.AddItem("#")
    Call Me.cbxFormat.AddItem("0#")
    Call Me.cbxFormat.AddItem("00#")
    Me.cbxFormat.ListIndex = 0
End Sub
