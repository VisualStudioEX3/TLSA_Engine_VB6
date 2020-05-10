VERSION 5.00
Begin VB.Form frmHelp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "..."
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6180
   Icon            =   "frmHelp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   6180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Acerca del proyecto TLSA y sus componentes"
      ForeColor       =   &H8000000D&
      Height          =   1185
      Left            =   30
      TabIndex        =   3
      Top             =   2220
      Width           =   6135
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   795
         Left            =   90
         TabIndex        =   4
         Top             =   240
         Width           =   5985
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   1755
      Left            =   0
      Picture         =   "frmHelp.frx":000C
      ScaleHeight     =   1695
      ScaleWidth      =   6120
      TabIndex        =   2
      Top             =   0
      Width           =   6180
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ayuda rapida de comandos y acciones del editor"
      ForeColor       =   &H8000000D&
      Height          =   2565
      Left            =   30
      TabIndex        =   0
      Top             =   3420
      Width           =   6135
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   2205
         Left            =   120
         TabIndex        =   1
         Top             =   210
         Width           =   5865
      End
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "http://perso.wanadoo.es/manuelsanchezromero/index2.htm"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   945
      MouseIcon       =   "frmHelp.frx":223A6
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   1980
      Width           =   4305
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Los sprites de ""Blueman"" son propiedad de Ruben Sanchez / Ed_Hunter"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   488
      TabIndex        =   5
      Top             =   1740
      Width           =   5205
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Definición de la API llamar al explorador por defecto
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'Definicion de las constantes
Private Const SW_SHOWNORMAL = 1

Private Sub Form_Load()
    Call modMain.OnTop(Me)
    
    Me.Caption = "Acerca de '" & modMain.AppTitle & "' y ayuda rapida"
    
    Const txtAbout As String = "TLSA SDK © José Miguel Sánchez Fernández - 2006/2009" & vbNewLine & _
                               "TLSA Engine © José Miguel Sánchez Fernández - 2004/2009" & vbNewLine & _
                               "TLSA © José Miguel Sánchez Fernández - 2001/2009" & vbNewLine & _
                               "dx_lib32 © José Miguel Sánchez Fernández - 2001/2004/2009"
                               
    Me.Label2.Caption = txtAbout
    
    Const txtHelp As String = "Para seleccionar pulse con el boton izquierdo del raton el punto de origen y despues arrastre el cursor para desplazar el punto del extremo de la seleccion." & vbNewLine & vbNewLine & _
                      "Pulse el boton derecho para cancelar seleccion." & vbNewLine & vbNewLine & _
                      "Para seleccionar el punto de origen o centro del sprite (spritesheet) pulse primero sobre el boton Marcar y despues indique la ubicacion con el raton desplazando la guia de color rojo." & vbNewLine & vbNewLine & _
                      "Puede modificar los valores de las selecciones o el centro del sprite a mano. Para aplicar cambios pulse el boton Actualizar."
    
    Me.Label1.Caption = txtHelp
End Sub

Private Sub Label4_Click()
    Call ShellExecute(Me.hwnd, vbNullString, Me.Label4.Caption, vbNullString, "c:\", SW_SHOWNORMAL)
End Sub
