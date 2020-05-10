VERSION 5.00
Begin VB.Form frm_ENT_ParticlePlatform 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Plataforma"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2640
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   2640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cbxTipo 
      Height          =   315
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1560
      Width           =   2625
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tile"
      ForeColor       =   &H8000000D&
      Height          =   645
      Left            =   0
      TabIndex        =   7
      Top             =   660
      Width           =   2625
      Begin VB.ComboBox cbxTile 
         Height          =   315
         Left            =   60
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   240
         Width           =   2475
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Entidad objetivo"
      ForeColor       =   &H8000000D&
      Height          =   645
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2625
      Begin VB.OptionButton optTarget 
         Height          =   405
         Index           =   0
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   180
         Value           =   -1  'True
         Width           =   405
      End
      Begin VB.OptionButton optTarget 
         Height          =   405
         Index           =   1
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   180
         Width           =   405
      End
      Begin VB.OptionButton optTarget 
         Height          =   405
         Index           =   2
         Left            =   900
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   180
         Width           =   405
      End
      Begin VB.OptionButton optTarget 
         Height          =   405
         Index           =   3
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   180
         Width           =   405
      End
      Begin VB.OptionButton optTarget 
         Height          =   405
         Index           =   4
         Left            =   1740
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   180
         Width           =   405
      End
      Begin VB.OptionButton optTarget 
         Height          =   405
         Index           =   5
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   180
         Width           =   405
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Tipo de superficie"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   0
      TabIndex        =   10
      Top             =   1320
      Width           =   1260
   End
End
Attribute VB_Name = "frm_ENT_ParticlePlatform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

