VERSION 5.00
Begin VB.Form frmTileSelector 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleccione un tile de la lista"
   ClientHeight    =   1140
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3150
   Icon            =   "frmTileSelector.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1140
   ScaleWidth      =   3150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Seleccionar"
      Default         =   -1  'True
      Height          =   435
      Left            =   30
      TabIndex        =   2
      Top             =   690
      Width           =   3105
   End
   Begin VB.Frame Frame1 
      Caption         =   "Lista de tiles definidos"
      ForeColor       =   &H8000000D&
      Height          =   645
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   3105
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   90
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2925
      End
   End
End
Attribute VB_Name = "frmTileSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private varCancelled As Boolean
Private varSelection As String

Public Property Get Cancelled() As Boolean
    Cancelled = varCancelled
End Property

Public Property Get Selection() As String
    Selection = varSelection
End Property

Private Sub Command1_Click()
    varSelection = Me.Combo1.List(Me.Combo1.ListIndex)
    Call Unload(Me)
End Sub

Private Sub Form_Load()
    Dim i As Long
    If frmMain.lstTiles.ListCount > 0 Then
        For i = 1 To frmMain.lstTiles.ListCount - 1
            Call Me.Combo1.AddItem(frmMain.lstTiles.List(i))
        Next
        Me.Combo1.ListIndex = 0
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    varCancelled = (UnloadMode = 0) ' UnloadMode = 0 -> Se cerro mediante el comando de la ventana (1 mediante Unload())
End Sub
