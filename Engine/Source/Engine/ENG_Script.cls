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
'' Objeto Script que contiene una secuencia de instrucciones para ejecutar acciones contra los objetos del motor.
'
'Option Explicit
'
'' Lista de comandos disponibles:
'Private Const Commands As String = "EXIT BREAK WAIT GOTO MARK VAR LET IF SELECT SETACTION SETPLAYERINPUT " & _
'                                   "SETCAMERATARGET ACTIVATE DEACTIVATE DEBUG GET GETACTION GETANGLE " & _
'                                   "ISFIXEDANGLE SETFIXEDANGLE UNSETFIXEDANGLE GETX GETY GETZ " & _
'                                   "GETCURRENTOBJECT GETPLAYERINPUT GETCAMERAX GETCAMERAY GETCAMERATARGET"
'' 28 instrucciones actualmente.
'
'Private scriptEngine As ScriptControl       ' Referencia al motor de scripts.
'Private Code() As String                    ' Lista de instrucciones del script con sus parametros.
'Private Vars As Collection                  ' Lista de variables del script.
'Private BookMarks As Collection             ' Lista de marcadores del script.
'Private CurrentCode As Long                 ' Instruccion actual.
'Private CurrentObject As Object             ' Entidad actual a la que llamara el script.
'Private CurrentObjectName As String         ' Clave de la entidad actual.
'Private Running As Boolean                  ' Indica si el script esta en ejecucion.
'Private WaitStatus As Boolean               ' Indica si se esta ejecutando una espera.
'Private Timer As New System.Timer           ' Cronometro para controlar las esperas.
'Private WaitDelay As Long                   ' Tiempo de espera del cronometro.
'
'' Carga un script en memoria:
'Public Sub Import(Filename As String)
'    On Error GoTo ErrOut
'
'    ' Leemos el contenido del archivo:
'    Call System.IO.ReadTextFile(Filename, Code())
'
'    ' Inicializamos el motor de script:
'    Set scriptEngine = New ScriptControl
'
'    scriptEngine.AllowUI = False
'    scriptEngine.UseSafeSubset = True
'
'    ' Agregamos el codigo de las funciones de interactuacion con el motor del juego:
'    'call scriptengine.AddCode("")
'
'    ' Inicializamos el cronometro para las esperas:
'    Set Timer = New System.Timer
'
'    Exit Sub
'
'ErrOut:
'   Stop
'End Sub
'
'' Agrega al motor de script las instrucciones de interactuacion con el motor del juego:
'Private Sub AddInstructions()
'    ' Y POR QUE NO MEJOR PROGRAMARLO EN EL PROPIO CODIGO DEL PROGRAMA COMO UNA CLASE Y AGREGARLA AL SCRIPT ENGINE?
'End Sub
'
'' Ejecuta el script:
'Public Sub Run()
'    If CurrentCode > Codes Then
'        Running = False
'        CurrentCode = 0
'    End If
'
'    ' Si el script esta activo se ejecutan las instrucciones:
'    If Running Then
'        ' Si no hay espera programada o en caso contrario la espera ha finalizado ejecutamos el codigo:
'        If (Not WaitStatus) Or (WaitStatus And Timer.value >= WaitDelay) Then
'            Dim i As Long
'            For i = CurrentCode To UBound(Code)
'                CurrentCode = i ' Se marca la linea de la instruccion actual a ejecutar.
'                ' Ejecutamos la instruccion correspondiente:
'                Call scriptEngine.Eval(Code(i))
'            Next
'        End If
'    End If
'End Sub
'
'' Detiene la ejecucion del script:
'Public Sub Terminate()
'    Set Timer = Nothing
'End Sub
