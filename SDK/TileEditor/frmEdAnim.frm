VERSION 5.00
Begin VB.Form frmEdAnim 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Editor de animaciones"
   ClientHeight    =   3705
   ClientLeft      =   150
   ClientTop       =   180
   ClientWidth     =   10230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   10230
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   "Opciones"
      Enabled         =   0   'False
      ForeColor       =   &H8000000D&
      Height          =   915
      Left            =   6750
      TabIndex        =   18
      Top             =   2760
      Width           =   3465
      Begin VB.CommandButton Command3 
         Caption         =   "Previsualizar"
         Height          =   315
         Index           =   0
         Left            =   60
         TabIndex        =   19
         Top             =   210
         Width           =   3345
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Propiedades"
      Enabled         =   0   'False
      ForeColor       =   &H8000000D&
      Height          =   1635
      Left            =   6750
      TabIndex        =   9
      Top             =   1110
      Width           =   3465
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Actualizar"
         Height          =   315
         Left            =   60
         TabIndex        =   21
         Top             =   1260
         Width           =   3345
      End
      Begin VB.TextBox txtDelay 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2550
         MaxLength       =   6
         TabIndex        =   13
         Text            =   "0"
         Top             =   870
         Width           =   825
      End
      Begin VB.TextBox txtKey 
         Height          =   315
         Left            =   90
         TabIndex        =   11
         Top             =   480
         Width           =   3285
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tiempo de actualizacion (mlSeg)"
         Height          =   195
         Left            =   90
         TabIndex        =   14
         Top             =   930
         Width           =   2295
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre de la secuencia de animacion"
         Height          =   195
         Left            =   90
         TabIndex        =   12
         Top             =   240
         Width           =   2715
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Secuencias de animacion"
      ForeColor       =   &H8000000D&
      Height          =   1065
      Left            =   6750
      TabIndex        =   5
      Top             =   30
      Width           =   3465
      Begin VB.CommandButton Command2 
         Caption         =   "Eliminar"
         Height          =   345
         Index           =   1
         Left            =   990
         TabIndex        =   8
         Top             =   600
         Width           =   885
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Nueva"
         Height          =   345
         Index           =   0
         Left            =   90
         TabIndex        =   7
         Top             =   600
         Width           =   885
      End
      Begin VB.ComboBox cbxAnimSec 
         Height          =   315
         Left            =   90
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   240
         Width           =   3285
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tiles de la textura"
      Enabled         =   0   'False
      ForeColor       =   &H8000000D&
      Height          =   3645
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   6705
      Begin VB.CommandButton Command1 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   14.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   6
         Left            =   3060
         TabIndex        =   20
         ToolTipText     =   "Añadir tiles mediante filtro"
         Top             =   3090
         Width           =   585
      End
      Begin VB.CommandButton Command1 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   14.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   5
         Left            =   3060
         TabIndex        =   17
         ToolTipText     =   "Bajar posicion"
         Top             =   2610
         Width           =   585
      End
      Begin VB.CommandButton Command1 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   14.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   4
         Left            =   3060
         TabIndex        =   16
         ToolTipText     =   "Subir posicion"
         Top             =   2130
         Width           =   585
      End
      Begin VB.CommandButton Command1 
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   14.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   3
         Left            =   3060
         TabIndex        =   15
         ToolTipText     =   "Eliminar todos los tiles"
         Top             =   1650
         Width           =   585
      End
      Begin VB.CommandButton Command1 
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   14.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   2
         Left            =   3060
         TabIndex        =   10
         ToolTipText     =   "Añadir todos los tiles"
         Top             =   1170
         Width           =   585
      End
      Begin VB.CommandButton Command1 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   14.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   1
         Left            =   3060
         TabIndex        =   4
         ToolTipText     =   "Eliminar tile"
         Top             =   690
         Width           =   585
      End
      Begin VB.CommandButton Command1 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   14.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   0
         Left            =   3060
         TabIndex        =   3
         ToolTipText     =   "Añadir tile"
         Top             =   210
         Width           =   585
      End
      Begin VB.ListBox lstAnimTiles 
         Height          =   3375
         Left            =   3660
         TabIndex        =   2
         Top             =   210
         Width           =   2985
      End
      Begin VB.ListBox lstTexTiles 
         Height          =   3375
         Left            =   60
         TabIndex        =   1
         Top             =   210
         Width           =   2985
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tiles de la animacion"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   3690
         TabIndex        =   22
         Top             =   0
         Width           =   1485
      End
   End
End
Attribute VB_Name = "frmEdAnim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cbxAnimSec_Change()
    Call modAnimEd.ShowData(Me.txtKey.Text)
End Sub

'Private Sub cbxAnimSec_Click()
'
'End Sub

Private Sub Command2_Click(Index As Integer)
    Select Case Index
        Case 0
            Dim key As String
            Dim a As Ed_Animation
                
            key = InputBox("Introduzca un nombre para identificar la secuencia en la lista:", "Añadir nueva secuencia de animacion", "Nueva secuencia " & Me.cbxAnimSec.ListCount)
            If Not key = "" Then
                ' *** Comprobar si la clave existe ***
                ' ...
                a.key = key
                a.Delay = 0
                Call modAnimEd.Secuences.Add(a, key)
                Call Me.cbxAnimSec.AddItem(key) ' Añadimos el nombre a la lista del editor.
                Me.cbxAnimSec.ListIndex = Me.cbxAnimSec.ListCount - 1
            End If
        Case 1
            If MsgBox("¿Borrar secuencia de animacion de la lista?", vbExclamation + vbYesNo, "Eliminar secuencia de animacion") = vbYes Then
                Call modAnimEd.Secuences.Remove(Me.txtKey.Text)
                If modAnimEd.Secuences.Count = 0 Then
                    Me.txtDelay.Text = 0
                    Me.txtKey.Text = ""
                Else
                    frmEdAnim.cbxAnimSec.ListIndex = 0
                End If
            End If
    End Select
    
    Me.Frame1.Enabled = (Me.cbxAnimSec.ListCount > 0)
    Me.Frame3.Enabled = (Me.cbxAnimSec.ListCount > 0)
    Me.Frame4.Enabled = (Me.cbxAnimSec.ListCount > 0)
End Sub

Private Sub Form_Load()
    Dim i As Long
    For i = 0 To frmMain.lstTiles.ListCount - 1
        Me.lstTexTiles.AddItem (frmMain.lstTiles.List(i))
    Next
End Sub
