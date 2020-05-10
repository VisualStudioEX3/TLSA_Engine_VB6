VERSION 5.00
Begin VB.Form frmKeys 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Editor de claves de puntos de control"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4275
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   4275
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0C0&
      Height          =   315
      Index           =   7
      Left            =   2190
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   20
      Top             =   630
      Width           =   315
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   0
      Left            =   2520
      TabIndex        =   19
      Text            =   "Punto 0"
      Top             =   630
      Width           =   1725
   End
   Begin VB.TextBox txtTileKey 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   1170
      Locked          =   -1  'True
      TabIndex        =   16
      Text            =   "Nombre Tile"
      Top             =   0
      Width           =   3075
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Guardar"
      Height          =   435
      Left            =   2820
      TabIndex        =   14
      Top             =   2490
      Width           =   1425
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   7
      Left            =   360
      TabIndex        =   13
      Text            =   "Punto 7"
      Top             =   2610
      Width           =   1725
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FF00FF&
      Height          =   315
      Index           =   6
      Left            =   30
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   12
      Top             =   2610
      Width           =   315
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   6
      Left            =   360
      TabIndex        =   11
      Text            =   "Punto 6"
      Top             =   2280
      Width           =   1725
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FF0000&
      Height          =   315
      Index           =   5
      Left            =   30
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   10
      Top             =   2280
      Width           =   315
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   5
      Left            =   360
      TabIndex        =   9
      Text            =   "Punto 5"
      Top             =   1950
      Width           =   1725
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFF00&
      Height          =   315
      Index           =   4
      Left            =   30
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   8
      Top             =   1950
      Width           =   315
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   4
      Left            =   360
      TabIndex        =   7
      Text            =   "Punto 4"
      Top             =   1620
      Width           =   1725
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0000FF00&
      Height          =   315
      Index           =   3
      Left            =   30
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   6
      Top             =   1620
      Width           =   315
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   3
      Left            =   360
      TabIndex        =   5
      Text            =   "Punto 3"
      Top             =   1290
      Width           =   1725
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0000FFFF&
      Height          =   315
      Index           =   2
      Left            =   30
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   4
      Top             =   1290
      Width           =   315
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   2
      Left            =   360
      TabIndex        =   3
      Text            =   "Punto 2"
      Top             =   960
      Width           =   1725
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H000080FF&
      Height          =   315
      Index           =   1
      Left            =   30
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   2
      Top             =   960
      Width           =   315
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Text            =   "Punto 1"
      Top             =   630
      Width           =   1725
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H000000FF&
      Height          =   315
      Index           =   0
      Left            =   30
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   0
      Top             =   630
      Width           =   315
   End
   Begin VB.Label Label3 
      Caption         =   "El punto de control primario esta representado por el Offset del tile definido en el editor de tiles."
      Height          =   825
      Left            =   2190
      TabIndex        =   21
      Top             =   990
      Width           =   2085
   End
   Begin VB.Line Line1 
      X1              =   4470
      X2              =   -30
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label2 
      Caption         =   "Punto de control primario:"
      Height          =   225
      Left            =   2190
      TabIndex        =   18
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Lista de puntos de control:"
      Height          =   225
      Index           =   0
      Left            =   30
      TabIndex        =   17
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nombre del tile"
      Height          =   195
      Index           =   1
      Left            =   60
      TabIndex        =   15
      Top             =   60
      Width           =   1050
   End
End
Attribute VB_Name = "frmKeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

