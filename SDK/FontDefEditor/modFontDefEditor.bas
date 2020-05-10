Attribute VB_Name = "modFontDefEditor"
Option Explicit

Public gfx As graphics.Manager
Public sample As New graphics.TextString
'Public sys As dx_System_Class

' Carga una fuente de sistema:
' *** No importa archivo de definicion alguno ***
Public Sub LoadSystemFont()
    On Error GoTo ErrOut
    
    With frmMain
        Call gfx.Fonts.UnloadAll
        Call gfx.Fonts.LoadFont(.cbxSystemFonts.Text, "sample", CInt(.txtSize.Text), CBool(.chkBold.Value), CBool(.chkItalic.Value), CBool(.chkUnderline.Value), CBool(.chkStrikethrough.Value))
        Call sample.Initialize(gfx.Fonts("sample"))
    End With
    
    Exit Sub
    
ErrOut:
    Call MsgBox(Err.Description, vbCritical, "Error al cargar fuente del sistema")
End Sub

' Carga una fuente desde archivo, externo al sistema:
' *** No importa archivo de definicion alguno ***
Public Sub LoadFileFont(Filename As String)
    On Error GoTo ErrOut
    
    frmMain.txtFilename.Text = Filename
    With frmMain
        Call gfx.Fonts.UnloadAll
        Call gfx.Fonts.LoadFontFromFile(.txtFilename.Text, "sample", CInt(.txtSize.Text), CBool(.chkBold.Value), CBool(.chkItalic.Value), CBool(.chkUnderline.Value), CBool(.chkStrikethrough.Value))
        Call sample.Initialize(gfx.Fonts("sample"))
    End With
    
    Exit Sub
    
ErrOut:
    Call MsgBox(Err.Description, vbCritical, "Error al cargar archivo fuente")
End Sub

' Importa un archivo de definicion:
Public Sub ImportFont(Filename As String)
    On Error GoTo ErrOut
    
    Call gfx.Fonts.UnloadAll
    Call gfx.Fonts.LoadFontFileDefinition(Filename, "sample")
    Call sample.Initialize(gfx.Fonts("sample"))
    
    Exit Sub
    
ErrOut:
    Call MsgBox(Err.Description, vbCritical, "Error al importar fuente")
End Sub
