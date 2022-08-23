Attribute VB_Name = "VariablesGlobales"

Option Explicit
'Manejo de base de datos
Public bdAplicacion As Database, bdWK As Workspace

Public tbSISTEMA_BITACORA As Recordset
Public tbSISTEMA_BITACORAlistados As Recordset
Public tbSISTEMA_BITACORAparametros As Recordset
Public tbSISTEMA_PARAMETROS As Recordset
Public tbSISTEMA_USUARIOS As Recordset
Public tbSISTEMA_OPERACIONES As Recordset
Public tbSISTEMA_PERFILES As Recordset
Public tbSISTEMA_LICENCIA As Recordset

Public m_vardirBD As String     'ubicaci�n en disco de la base de datos
Public m_vardirRpt As String    'ubicaci�n en disco de los reportes

Public CaminoBaseDeDatos As String
Public Resolucion As String

'Determina la forma de trabajar del formulario
'de fechas
'1 pido fecha �nica
'2 pido ambas fechas
Public tipo_accion_fechas As Byte
'*

'Determina la forma de trabajar del formulario
'de selecci�n de listados
'1 = ejecutar
'2 = imprimir
'3 = eliminar
'4 = predeterminado
Public tipo_accion_selec As Byte
'*

Public rst_opr As Recordset    'Este recorset es utilizado para realizar todas las
                                'consultas de operaciones

Public m_UsuarioSisNom As String    'usuario que ingresa a la aplicaci�n


'Determina si se cancelo el ingreso de fechas al
'ejecutar un listado, si es as� no se continu� con la ejecuci�n
'del mismo.
Public CanceloIngresoFechas As Boolean

'Determina si estoy ejecutando una versi�n demo.
'La misma se inicializa en el m�dulo ControlDeLicenias
'y es utilizada en el formulario frmMain, cuando de produce el evento
'click de la opci�n del menu salir. Si la misma esta inicializa a true, se muestra el
'aviso de versi�n demo al salir de la aplicaci�n.
    Public gEsUnaVersionDemo As Boolean
'*

'Contiene el c�digo Id de la aplicaci�n, obtenido del archivo aplicaci�n.Id.txt
'Se utilza en el m�dulo de ControlDeLicencia
'Adem�s se utiliza para obtener informaci�n de la licencia de la aplicaci�n
'en el formulario de AcercaDe
    Public idApli As Long
'*
