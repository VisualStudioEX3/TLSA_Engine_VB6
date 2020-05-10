VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ParticleEmitter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private varLifeTime As Long                 ' Tiempo que duraran las particulas (-1 = infinito)
Private varMaxParticles As Long             ' Numero maximo de particulas que generara el emisor.
Private varLoops As Long                    ' Numero de veces que generara particulas (-1 = indefinido)
Private varSize As Long                     ' Alto/Ancho de las particulas (seran cuadradas).
Private varWeight As Long                   ' Peso de las particulas.
Private varDirection As Single              ' Angulo que definira la direccion prioritaria hacia la que se emitiran las particulas (-1 ninguna, se generaran hacia cualquier direccion aleatoria).
Private varLocation As Core.POINT         ' Posicion del emisor.
Private varEnabled As Boolean               ' Indica si el emisor esta activado.

Private varFatherSim As Physics.Simulator   ' Referencia al simulador al que esta asociado el emisor.
Private varParticles As Collection          ' Lista interna para gestionar las instancias de las particulas dentro del simulador de fisica.

Public Sub Start()
'
End Sub

Friend Sub Update()
'
End Sub

Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
'
End Sub
