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

'Declaraci?n de constantes
Private Const cCantDias As Byte = 30    'Establece la cantidad de d?as de duraci?n, del per?odo de evaluaci?n

'Declaraci?n de variables
Private mfrmInicializarAplicacion As frmInicializarAplicacion

Public Function funInicializarAplicacion(tabla As Recordset, idAplicacion As Long, diaInstalacion As Date) As Integer
    'Este procedimiento crea un registro en la tabla SISTEMA_LICENCIA
    'ubicada en la base de datos de la aplicaci?n que utiliza este procedimiento.
    'La clave del registro es el n?mero de identificaci?n que se encuentra en el archivo id.txt
    'perteneciente a la aplicaci?n.
    'Este procedimiento se ejecutar? siempre que no se encuentre un registro en la tabla anteriormente
    'mencionada, el cual tenga como clave el n?mero de Id.
    'Esto ocurrir? la primera vez que se ejecuta una aplicaci?n luego de instalada.
    
    '----- Par?metros --------
    'Entrada:   [tabla]             tabla SISTEMA_LICENCIA la cual se abre en la aplicaci?n principal.
    '           [idAplicacion]      n?mero que identifica a la aplicaci?n. El mismo se obtiene del archvio
    '                               id.txt
    '           [diaInstalacion]    fecha actual del sistema. Es importante inicializar este valor,
    '                               ya que el mismo es utilizado por las rutinas que controlan
    '                               del per?odo de evaluci?n de las versiones demo.
    '
    'Salida:    Devuelve un c?digo que establece si se pudo realizar la inicializaci?n en forma correcta
    '           o en el caso contrario, si se produjo un error.
    '           C?digos devueltos:
    '
    '           0:  La inicializaci?n se efectu? correctamente
    '           Cualquier error que pueda surgir de las sentencias de la funci?n
    '-------------------------
    On Error GoTo error
    'muestro formulario de inicializaci?n
    mfrmInicializarAplicacion.Show 1
    'agrego nuevo registro
    tabla.AddNew
        tabla("AplicacionId") = idAplicacion
        tabla("AplicacionSerieDisco") = "0"     'valor inicial establecido para la versi?n demo
        tabla("AplicacionNroLicencia") = "0"    'idem
        tabla("AplicacionDiaInstalacion") = diaInstalacion
        tabla("AplicacionVD") = True            'todas las versiones de la aplicaci?n son demos, por defecto
        tabla("AplicacionVDUltimoDiaEjecutada") = diaInstalacion
        tabla("AplicacionVDCantDias") = cCantDias
    tabla.Update
error:
    funInicializarAplicacion = Err.Number
End Function

Private Sub Class_Initialize()
    'Normalmente, el procedimiento Initialize contiene cualquier c?digo
    'que es necesario ejecutar en el momento en que se crea el objeto.

    'creo objeto
    Set mfrmInicializarAplicacion = New frmInicializarAplicacion
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    'Siempre debe de controlar los errores del procedimiento de evento
    'Class_Terminate. Las aplicaciones cliente que utilizan su componente no pueden controlar
    'errores de Class_Terminate y,por lo tanto, ser?n fatales para la aplicaci?n.
    'objeto terminado

    Unload frmInicializarAplicacion
    Set frmInicializarAplicacion = Nothing
    'Cuando se termina, un objeto que controla un formulario, siempre debe descargar
    'el formulario y establecer su referencia a nothing para evitar ocupar recursos
    'con formularios hu?rfanos.
End Sub
