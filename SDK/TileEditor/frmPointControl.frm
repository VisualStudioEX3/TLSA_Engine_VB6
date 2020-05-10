VERSION 5.00
Begin VB.Form frmPointControl 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Editor de puntos de control"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3180
   Icon            =   "frmPointControl.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   3180
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Opciones"
      ForeColor       =   &H8000000D&
      Height          =   645
      Left            =   30
      TabIndex        =   16
      Top             =   4500
      Width           =   3135
      Begin VB.CommandButton cmdTools 
         Caption         =   "Copiar puntos de control desde otro tile"
         Height          =   345
         Index           =   0
         Left            =   60
         TabIndex        =   8
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Actualizar"
      Height          =   345
      Left            =   30
      TabIndex        =   9
      Top             =   5190
      Width           =   3135
   End
   Begin VB.Frame Frame5 
      Caption         =   "Coordenadas del punto de control"
      ForeColor       =   &H8000000D&
      Height          =   645
      Left            =   30
      TabIndex        =   12
      Top             =   3840
      Width           =   3135
      Begin VB.TextBox txtY 
         Height          =   285
         Left            =   990
         TabIndex        =   6
         Top             =   240
         Width           =   500
      End
      Begin VB.TextBox txtX 
         Height          =   285
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   500
      End
      Begin VB.CommandButton cmdSetPoint 
         Caption         =   "Marcar"
         Height          =   345
         Left            =   1530
         TabIndex        =   7
         Top             =   240
         Width           =   1545
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Y."
         Height          =   195
         Left            =   840
         TabIndex        =   14
         Top             =   300
         Width           =   150
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "X."
         Height          =   195
         Left            =   90
         TabIndex        =   13
         Top             =   300
         Width           =   150
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Puntos de control del tile seleccionado"
      ForeColor       =   &H8000000D&
      Height          =   3195
      Left            =   30
      TabIndex        =   11
      Top             =   0
      Width           =   3135
      Begin VB.CommandButton cmdClear 
         Caption         =   "Limpiar"
         Height          =   435
         Left            =   1500
         TabIndex        =   3
         Top             =   2700
         Width           =   700
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Eliminar"
         Height          =   435
         Left            =   780
         TabIndex        =   2
         Top             =   2700
         Width           =   700
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Añadir"
         Height          =   435
         Left            =   60
         TabIndex        =   1
         Top             =   2700
         Width           =   700
      End
      Begin VB.ListBox lstPoints 
         Height          =   2400
         Left            =   60
         TabIndex        =   0
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label lblCountIndex 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0 de 0"
         Height          =   195
         Left            =   2580
         TabIndex        =   15
         Top             =   2820
         Width           =   450
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Clave del punto de control"
      ForeColor       =   &H8000000D&
      Height          =   615
      Left            =   30
      TabIndex        =   10
      Top             =   3210
      Width           =   3135
      Begin VB.TextBox txtKey 
         Height          =   285
         Left            =   60
         TabIndex        =   4
         Top             =   240
         Width           =   2995
      End
   End
End
Attribute VB_Name = "frmPointControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
    selPset = False
    If Not CurrentTile Is Nothing Then
        Dim Key As String

        Key = InputBox("Introduzca un nombre para identificar al punto de control en la lista:", "Añadir nuevo punto de control", "Nuevo punto de control " & Me.lstPoints.ListCount)
        If Not Key = "" Then
            If Not CurrentTile.ControlPoints.Exists(Key) Then
                Call CurrentTile.ControlPoints.Add(Key, 0, 0)
                Call Me.lstPoints.AddItem(Key) ' Añadimos el nombre a la lista del editor.
                Me.lstPoints.ListIndex = Me.lstPoints.ListCount - 1
                Me.txtKey.Tag = Key
            End If
        End If
    End If

    Call frmVisor.SetFocus
End Sub

Private Sub cmdClear_Click()
    If MsgBox("Esto borrara todos los puntos de control de la lista " & _
              "(excepto el punto por defecto). ¿Desea borrar todos los puntos de control de la lista? " & _
              "(Esta operacion no se puede deshacer)" _
              , vbExclamation + vbYesNo, "Borrar todos los puntos de control") = vbYes Then
              
              Call CurrentTile.ControlPoints.Clear
              Call modMain.ShowTileData(CurrentTile.Key)
    End If
End Sub

Private Sub cmdRemove_Click()
    If MsgBox("¿Borrar seguro el tile de la lista?", vbExclamation + vbYesNo, "Borrar tile") = vbYes Then
        Call CurrentTile.ControlPoints.Remove(Me.txtKey.Text)
        Call modMain.ShowTileData(CurrentTile.Key)
    End If
End Sub

Private Sub cmdSetPoint_Click()
    selPset = True
    Call frmVisor.SetFocus
End Sub

Private Sub cmdTools_Click(Index As Integer)
    Select Case Index
        Case 0
            Call frmTileSelector.Show(vbModal)
            If Not frmTileSelector.Cancelled Then
                Call CopyFromTile(frmTileSelector.Selection)
                Call modMain.ShowTileData(CurrentTile.Key)
            End If
    End Select
End Sub

Private Sub cmdUpdate_Click()
    selPset = False
    
    If Me.txtKey.Text <> Me.txtKey.Tag Then Call modMain.ControlPointChangeKey(Me.txtKey.Tag, Me.txtKey.Text)
    Call CurrentTile.ControlPoints.SetPoint(Me.txtKey.Tag, CLng(Me.txtX.Text), CLng(Me.txtY.Text))
    Call modMain.ShowTileData(frmMain.lstTiles.List(frmMain.lstTiles.ListIndex))
End Sub

Private Sub Form_Load()
    Call modMain.OnTop(Me)
    Me.Left = frmMain.Left + (frmMain.Width - Me.Width) - 250
    Me.Top = frmMain.Top + frmMain.Top + 950
End Sub

Public Sub SelDefault()
    Me.lstPoints.ListIndex = 0
    Call lstPoints_Click
End Sub

Private Sub lstPoints_Click()
    If lstPoints.ListIndex = 0 Then
        showControlPoint = False
    Else
        If Not selAset And Not selCset And Not selPset And lstPoints.ListIndex > -1 Then Call modMain.ShowControlPointData(lstPoints.List(lstPoints.ListIndex))
               
        showControlPoint = True
    End If
    
    Me.txtKey.Tag = Me.txtKey.Text
    Me.cmdRemove.Enabled = showControlPoint
    Me.cmdSetPoint.Enabled = showControlPoint
    Me.cmdUpdate.Enabled = showControlPoint
    Me.txtKey.Enabled = showControlPoint
    Me.txtX.Enabled = showControlPoint
    Me.txtY.Enabled = showControlPoint
    
    frmPointControl.lblCountIndex.Caption = (frmPointControl.lstPoints.ListIndex + 1) & " de " & frmPointControl.lstPoints.ListCount
End Sub

Private Sub txtX_LostFocus()
    If Not IsNumeric(Me.txtX.Text) Then Me.txtX.Text = 0
End Sub

Private Sub txtY_LostFocus()
    If Not IsNumeric(Me.txtY.Text) Then Me.txtY.Text = 0
End Sub
