VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AvisoFinPeriodoDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Declaración de propiedades
Public AvisoFinPeriodoDemoPropTituloForm As String
Public AvisoFinPeriodoDemoPropNomAplicacion As String
Public AvisoFinPeriodoDemoPropSistemaAplicacion As String
Public AvisoFinPeriodoDemoPropVersionAplicacion As String
Public AvisoFinPeriodoDemoPropPeriodoTerminado As String
Public AvisoFinPeriodoDemoPropDerechos As String
Public AvisoFinPeriodoDemoPropExtension As Boolean

'Declaración de variables
Private mFrmAvisoFinPeriodoDemo As frmAvisoFinPeriodoDemo

Private Sub Class_Initialize()
    'objeto creado
    Set mFrmAvisoFinPeriodoDemo = New frmAvisoFinPeriodoDemo
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    
    Unload frmAvisoFinPeriodoDemo
    Set frmAvisoFinPeriodoDemo = Nothing
    'Cuando se termina, un objeto que controla un formulario, siempre debe descargar
    'el formulario y establecer su referencia a nothing para evitar ocupar recursos
    'con formularios huérfanos.
End Sub

Public Sub MostrarAvisoFinPeriodoDemo()
    'inicializo la propiedad del formulario que indica si se muestra el aviso
    'de extensión del período de evaluación.
    mFrmAvisoFinPeriodoDemo.propMuestroAvisoExtension = AvisoFinPeriodoDemoPropExtension
    
    'establesco el título del formulario
    mFrmAvisoFinPeriodoDemo.Caption = AvisoFinPeriodoDemoPropTituloForm
    
    'asigno el valor de las diferentes propiedades del objeto
    'a las etiquetas del formulario correspondientes.
    mFrmAvisoFinPeriodoDemo.lblNomAplicacion.Caption = AvisoFinPeriodoDemoPropNomAplicacion
    mFrmAvisoFinPeriodoDemo.lblSistemaAplicacion.Caption = AvisoFinPeriodoDemoPropSistemaAplicacion
    mFrmAvisoFinPeriodoDemo.lblVersionAplicación.Caption = AvisoFinPeriodoDemoPropVersionAplicacion
    mFrmAvisoFinPeriodoDemo.lblPeriodoTerminado.Caption = AvisoFinPeriodoDemoPropPeriodoTerminado
    mFrmAvisoFinPeriodoDemo.lblDerechos.Caption = AvisoFinPeriodoDemoPropDerechos
    'muestro formulario
    mFrmAvisoFinPeriodoDemo.Show 1
End Sub
