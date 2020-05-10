VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ENG_Script"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Objeto Script que contiene una secuencia de instrucciones para ejecutar acciones contra los objetos del motor.

Option Explicit

' Lista de comandos disponibles:
Private Const Commands As String = "EXIT BREAK WAIT GOTO MARK VAR LET IF SELECT SETACTION SETPLAYERINPUT " & _
                                   "SETCAMERATARGET ACTIVATE DEACTIVATE DEBUG GET GETACTION GETANGLE " & _
                                   "ISFIXEDANGLE SETFIXEDANGLE UNSETFIXEDANGLE GETX GETY GETZ " & _
                                   "GETCURRENTOBJECT GETPLAYERINPUT GETCAMERAX GETCAMERAY GETCAMERATARGET"
' 28 instrucciones actualmente.

Private Type ENG_InstructionData
    Command As String   ' Clave de la instruccion a ejecutar.
    params() As String  ' Parametros de la instruccion.
    ParamCount As Long  ' Numero de parametros.
End Type

Private Code() As ENG_InstructionData       ' Lista de instrucciones del script con sus parametros.
Private Codes As Long                       ' Numero de instrucciones.
Private Vars As Collection                  ' Lista de variables del script.
Private BookMarks As Collection             ' Lista de marcadores del script.
Private CurrentCode As Long                 ' Instruccion actual.
Private CurrentObject As Object             ' Entidad actual a la que llamara el script.
Private CurrentObjectName As String         ' Clave de la entidad actual.
Private Running As Boolean                  ' Indica si el script esta en ejecucion.
Private WaitStatus As Boolean               ' Indica si se esta ejecutando una espera.
Private Timer As New System.Timer           ' Cronometro para controlar las esperas.
Private WaitDelay As Long                   ' Tiempo de espera del cronometro.

' Carga un script en memoria:
Public Sub Import(Filename As String)
    On Error GoTo ErrOut
    
    Dim errDesc As String, errLine As Long, errCode As String
    
    ' Leemos el contenido del archivo:
    Dim buffer() As String: Call System.IO.ReadTextFile(Filename, buffer())
    
    ' Parseamos las lineas de texto en busca de instrucciones validas:
    Dim i As Long, j As Long
    Dim ins As ENG_InstructionData
    Dim ignore As Boolean           ' Indica si la instruccion no ha de agregarse a la pila (la instruccion Mark solo se procesa en tiempo de carga)
    Dim line As String              ' Linea de codigo que se analizara.
    Dim args() As String            ' Para realizar los analisis.
    
    For i = 0 To UBound(buffer)
        line = buffer(i)
        ins.Command = UCase(System.Text.GetFromLeft(line, "(")) ' Obtenemos el nombre de la instruccion.
        
        ' Comprobamos si la instruccion es valida:
        If System.Text.Find(Commands, ins.Command) Then
            ' Extraemos los parametros segun la instruccion:
            With ins
                Select Case .Command
                    ' Instrucciones sin parametros:
                    Case "EXIT", "BREAK", "ACTIVATE", "DEACTIVATE", "UNSETFIXEDANGLE"
                        ignore = False
                    
                    ' Parametro simple:
                    Case "WAIT", "GOTO", "SELECT", "SETPLAYERINPUT", "SETCAMERATARGET", "DEBUG", "SETFIXEDANGLE"
                        .ParamCount = 1: ReDim .params(0)
                        args = Split(line, " ")
                        .params(0) = Trim(args(1))
                        ignore = False
                    
                    ' Instrucciones con varios parametros:
                    Case "LET", "SETACTION"
                        .ParamCount = 2: ReDim .params(1)
                        args = Split(line, " ")
                        For j = 0 To UBound(args)
                            .params(j) = Trim(args(j + 1))
                        Next
                        ignore = False
                        
                    ' Instrucciones con varios parametros:
                    Case "IF"
                        .ParamCount = 3: ReDim .params(2)
                        args = Split(line, " ")
                        .params(0) = Trim(args(1))
                        ignore = False
                    
                    ' Instrucciones en tiempo de compilacion:
                    Case "MARK"
                        args = Split(line, " ")
                        Call BookMarks.Add(i, UCase(Trim(args(1))))
                        ignore = True
                    
                    Case "VAR"
                        args = Split(line, " ")
                        Call Vars.Add("", UCase(Trim(args(1))))
                        ignore = True
                End Select
            End With
            ' Si no se ignora, se añade la instruccion y sus parametros a la pila de llamadas:
            If Not ignore Then
                ReDim Preserve Code(Codes)
                Code(Codes) = ins
                Codes = Codes + 1
            End If
        Else
            ' La instruccion no es valida:
            Call Err.Raise(TLSA_EXCEPTIONS.SCRIPT_PARSE_ERROR, , "Error al parsear script: " & vbNewLine & _
                                                                 "Descripcion: La instruccion no es valida. " & vbNewLine & vbNewLine & _
                                                                 "Archivo: '" & Filename & "'" & vbNewLine & _
                                                                 "Linea: " & i & vbNewLine & _
                                                                 "Codigo: '" & buffer(i) & "'")
        End If
    Next
    
    Set Timer = New System.Timer
    
    Exit Sub
    
ErrOut:
   ' La instruccion no es valida:
    Call Err.Raise(TLSA_EXCEPTIONS.SCRIPT_PARSE_ERROR, , "Error al parsear script: " & vbNewLine & _
                                                         "Descripcion: " & Err.Description & vbNewLine & vbNewLine & _
                                                         "Archivo: '" & Filename & "'")
End Sub

' Ejecuta el script:
Public Sub Run()
    If CurrentCode > Codes Then
        Running = False
        CurrentCode = 0
    End If
    
    ' Si el script esta activo se ejecutan las instrucciones:
    If Running Then
        ' Si no hay espera programada o en caso contrario la espera ha finalizado ejecutamos el codigo:
        If (Not WaitStatus) Or (WaitStatus And Timer.value >= WaitDelay) Then
            Dim i As Long
            For i = CurrentCode To Codes
                CurrentCode = i ' Se marca la linea de la instruccion actual a ejecutar.
                ' Ejecutamos la instruccion correspondiente:
                Select Case Code(i).Command
                    Case "EXIT"             ' Detiene la ejecucion del script.
                        CurrentCode = 0
                        Running = False
                        Exit Sub
                    Case "BREAK"            ' Detiene la ejecucion hasta el proximo ciclo.
                        CurrentCode = i + 1
                        Exit Sub
                    Case "WAIT"             ' Detiene la ejecucion hasta que el proximo ciclo cumpla el tiempo de espera establecido.
                        WaitStatus = True
                        WaitDelay = Code(i).params(0)
                        Call Timer.Reset
                        CurrentCode = i + 1
                        Exit Sub
                    Case "GOTO"             ' Dirigimos la ejecucion al marcador designado.
                        i = BookMarks(Code(i).params(0))
                    Case "VAR"              ' Crea una variable.
                        Call Vars.Add(Code(i).params(0))
                    Case "LET"              ' Escribe un valor a una variable. LET "A", <valor o expresion>
                        Vars(Code(i).params(0)) = GetValue(CStr(Code(i).params(1)))
                    Case "IF"               ' Dirige el flujo de ejecucion segun condicion entre dos variables.
                        If Not Vars(Code(i).params(0)) = Vars(Code(i).params(1)) Then i = BookMarks(Code(i).params(2))
                    Case "SELECT"           ' Selecciona la entidad u objeto al que realizaremos llamadas.
                        Set CurrentObject = Engine.Scene.Entities(Code(i).params(0))
                    Case "SETACTION"        ' Activa una accion en el canal especificado del objeto actualmente seleccionado:
                        Call CurrentObject.SetAction(CLng(Code(i).params(0)), CStr(Code(i).params(1)))
                    Case "SETPLAYERINPUT"   ' Activa o desactiva la entrada de los controles del jugador:
                        Engine.PlayerInputEnabled = CBool(Code(i).params(0))
                    Case "SETCAMERATARGET"  ' Selecciona el objeto que la camara seguira:
                    
                    Case "SETFIXEDANGLE"    ' Fija el angulo del objeto:
                    
                    Case "UNSETFIXEDANGLE"  ' Libera el angulo del objeto:
                    
                    Case "ACTIVATE"         ' Activa un objeto de la escena:
                    
                    Case "DEACTIVATE"       ' Desactiva un objeto de la escena:
                    
                    Case "DEBUG"            ' Envia una cadena de texto a la salida de depuracion:
                    
                    Case Else               ' No se ha encontrado ninguna instruccion en el codigo:
                        
                End Select
            Next
        End If
    End If
End Sub

' Analiza la expresion y devuelve el valor que corresponda. Se utiliza para parsear llamadas a funciones que devuelvan valores:
' Solo se utiliza con la instruccion LET:
Private Function GetValue(Expression As String) As Variant
    If UCase(Expression) = "TRUE" Or UCase(Expression) = "FALSE" Then ' Constante booleana.
        GetValue = Expression
    ElseIf IsNumeric(Expression) Then   ' Numero.
        GetValue = Val(Expression)
    ElseIf VBA.Left(Expression, 1) = Chr(34) And VBA.Right(Expression, 1) = Chr(34) Then ' Cadena de texto.
        GetValue = Replace(Expression, Chr(34), "")
    Else ' Se esta intentando llamar a una instruccion que devuelve un valor:
        Dim arr() As String, cmd As String, param As String, i As Long
        
        ' Extraemos el nombre de la instruccion y el parametro:
        arr = Split(Replace(Expression, ")", ""), "(")
        cmd = arr(0) ' Instruccion.
        If UBound(arr) = 1 Then param = arr(1) ' Parametro unico.
        
        ' Buscamos la instruccion en la lista de posibles funciones:
        Select Case cmd
            Case "GET" ' Lectura de variable.
                GetValue = Vars(Replace(cmd, Chr(34), ""))
            Case "GETACTION" ' Devuelve la accion del canal indicado del objeto actual:
                GetValue = CurrentObject.GetAction(CLng(param))
            Case "GETANGLE" ' Devuelve el angulo del objeto.
                GetValue = CurrentObject.Angle
            Case "ISFIXEDANGLE" ' Indica si el angulo esta fijado o no.
                GetValue = CurrentObject.FixedAngle
            Case "GETX" ' Devuelve la coordenada X:
                GetValue = CurrentObject.Location.x
            Case "GETY" ' Devuelve la coordenada Y:
                GetValue = CurrentObject.Location.y
            Case "GETZ" ' Devuelve la coordenada Z:
                GetValue = CurrentObject.Location.Z
            Case "GETCURRENTOBJECT" ' Devuelve la clave del objeto actual:
                GetValue = CurrentObjectName
            Case "GETPLAYERINPUT" ' Devuelve si el jugador tiene activado los controles de entrada:
                GetValue = Engine.PlayerInputEnabled
            Case "GETCAMERAX" ' Devuelve la coordenada X de la camara:
                GetValue = Engine.Scene.Camera.x
            Case "GETCAMERAY" ' Devuelve la coordenada Y de la camara:
                GetValue = Engine.Scene.Camera.y
            Case "GETCAMERATARGET" ' Devuelve la clave del objeto al que sigue la camara:
            
            Case Else
                ' Error.
        End Select
    End If
End Function

' Detiene la ejecucion del script:
Public Sub Terminate()
    Set Timer = Nothing
End Sub
