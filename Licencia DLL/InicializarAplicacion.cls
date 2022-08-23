VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InicializarAplicacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Declaración de constantes
Private Const cCantDias As Byte = 30    'Establece la cantidad de días de duración, del período de evaluación

'Declaración de variables
Private mfrmInicializarAplicacion As frmInicializarAplicacion

Public Function funInicializarAplicacion(tabla As Recordset, idAplicacion As Long, diaInstalacion As Date) As Integer
    'Este procedimiento crea un registro en la tabla SISTEMA_LICENCIA
    'ubicada en la base de datos de la aplicación que utiliza este procedimiento.
    'La clave del registro es el número de identificación que se encuentra en el archivo id.txt
    'perteneciente a la aplicación.
    'Este procedimiento se ejecutará siempre que no se encuentre un registro en la tabla anteriormente
    'mencionada, el cual tenga como clave el número de Id.
    'Esto ocurrirá la primera vez que se ejecuta una aplicación luego de instalada.
    
    '----- Parámetros --------
    'Entrada:   [tabla]             tabla SISTEMA_LICENCIA la cual se abre en la aplicación principal.
    '           [idAplicacion]      número que identifica a la aplicación. El mismo se obtiene del archvio
    '                               id.txt
    '           [diaInstalacion]    fecha actual del sistema. Es importante inicializar este valor,
    '                               ya que el mismo es utilizado por las rutinas que controlan
    '                               del período de evalución de las versiones demo.
    '
    'Salida:    Devuelve un código que establece si se pudo realizar la inicialización en forma correcta
    '           o en el caso contrario, si se produjo un error.
    '           Códigos devueltos:
    '
    '           0:  La inicialización se efectuó correctamente
    '           Cualquier error que pueda surgir de las sentencias de la función
    '-------------------------
    On Error GoTo error
    'muestro formulario de inicialización
    mfrmInicializarAplicacion.Show 1
    'agrego nuevo registro
    tabla.AddNew
        tabla("AplicacionId") = idAplicacion
        tabla("AplicacionSerieDisco") = "0"     'valor inicial establecido para la versión demo
        tabla("AplicacionNroLicencia") = "0"    'idem
        tabla("AplicacionDiaInstalacion") = diaInstalacion
        tabla("AplicacionVD") = True            'todas las versiones de la aplicación son demos, por defecto
        tabla("AplicacionVDUltimoDiaEjecutada") = diaInstalacion
        tabla("AplicacionVDCantDias") = cCantDias
    tabla.Update
error:
    funInicializarAplicacion = Err.Number
End Function

Private Sub Class_Initialize()
    'Normalmente, el procedimiento Initialize contiene cualquier código
    'que es necesario ejecutar en el momento en que se crea el objeto.

    'creo objeto
    Set mfrmInicializarAplicacion = New frmInicializarAplicacion
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    'Siempre debe de controlar los errores del procedimiento de evento
    'Class_Terminate. Las aplicaciones cliente que utilizan su componente no pueden controlar
    'errores de Class_Terminate y,por lo tanto, serán fatales para la aplicación.
    'objeto terminado

    Unload frmInicializarAplicacion
    Set frmInicializarAplicacion = Nothing
    'Cuando se termina, un objeto que controla un formulario, siempre debe descargar
    'el formulario y establecer su referencia a nothing para evitar ocupar recursos
    'con formularios huérfanos.
End Sub
