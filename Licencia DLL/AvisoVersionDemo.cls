VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AvisoVersionDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Declaracíon de propiedades
Public AvisoVersionDemoPropTituloForm As String
Public AvisoVersionDemoPropNomAplicacion As String
Public AvisoVersionDemoPropSistemaAplicacion As String
Public AvisoVersionDemoPropVersionAplicacion As String
Public AvisoVersionDemoPropPeriodoDeUso As String
Public AvisoVersionDemoPropDerechos As String
Public AvisoVersionDemoPropDiasDemos As String

Private mFrmAvisoVersionDemo As frmAvisoVersionDemo
Attribute mFrmAvisoVersionDemo.VB_VarHelpID = -1

Private Sub Class_Initialize()
    'Normalmente, el procedimiento Initialize contiene cualquier código
    'que es necesario ejecutar en el momento en que se crea el objeto.
    
    'objeto creado
    Set mFrmAvisoVersionDemo = New frmAvisoVersionDemo
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    'Siempre debe de controlar los errores del procedimiento de evento
    'Class_Terminate. Las aplicaciones cliente que utilizan su componente no pueden controlar
    'errores de Class_Terminate y,por lo tanto, serán fatales para la aplicación.
    'objeto terminado
    
    Unload frmAvisoVersionDemo
    Set frmAvisoVersionDemo = Nothing
    'Cuando se termina, un objeto que controla un formulario, siempre debe descargar
    'el formulario y establecer su referencia a nothing para evitar ocupar recursos
    'con formularios huérfanos.
End Sub

Public Sub MostrarAvisoVersionDemo()
    'establesco el valor de la propiedad del forumlario que indica
    'el total de días permitidos para la licencia demo.
    mFrmAvisoVersionDemo.propDiasVersionDemo = AvisoVersionDemoPropDiasDemos

    'asigno el valor de las diferentes propiedades del objeto
    'a las etiquetas del formulario correspondientes.
    mFrmAvisoVersionDemo.lblNomAplicacion.Caption = AvisoVersionDemoPropNomAplicacion
    mFrmAvisoVersionDemo.lblSistemaAplicacion.Caption = AvisoVersionDemoPropSistemaAplicacion
    mFrmAvisoVersionDemo.lblVersionAplicación.Caption = AvisoVersionDemoPropVersionAplicacion
    mFrmAvisoVersionDemo.lblPeriodoDeUso.Caption = AvisoVersionDemoPropPeriodoDeUso
    mFrmAvisoVersionDemo.lblDerechos.Caption = AvisoVersionDemoPropDerechos
    'cambio título del formulario
    mFrmAvisoVersionDemo.Caption = AvisoVersionDemoPropTituloForm
    'muestro formulario
    mFrmAvisoVersionDemo.Show 1
End Sub

