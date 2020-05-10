VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H8000000C&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TLSA Engine"
   ClientHeight    =   3030
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   4560
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "frmMain.frx":70CA
   MousePointer    =   99  'Custom
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Unload(Cancel As Integer)
    Engine.Ready = False
End Sub
