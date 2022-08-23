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

Public m_vardirBD As String     'ubicación en disco de la base de datos
Public m_vardirRpt As String    'ubicación en disco de los reportes

Public CaminoBaseDeDatos As String
Public Resolucion As String

'Determina la forma de trabajar del formulario
'de fechas
'1 pido fecha única
'2 pido ambas fechas
Public tipo_accion_fechas As Byte
'*

'Determina la forma de trabajar del formulario
'de selección de listados
'1 = ejecutar
'2 = imprimir
'3 = eliminar
'4 = predeterminado
Public tipo_accion_selec As Byte
'*

Public rst_opr As Recordset    'Este recorset es utilizado para realizar todas las
                                'consultas de operaciones

Public m_UsuarioSisNom As String    'usuario que ingresa a la aplicación


'Determina si se cancelo el ingreso de fechas al
'ejecutar un listado, si es así no se continuá con la ejecución
'del mismo.
Public CanceloIngresoFechas As Boolean

'Determina si estoy ejecutando una versión demo.
'La misma se inicializa en el módulo ControlDeLicenias
'y es utilizada en el formulario frmMain, cuando de produce el evento
'click de la opción del menu salir. Si la misma esta inicializa a true, se muestra el
'aviso de versión demo al salir de la aplicación.
    Public gEsUnaVersionDemo As Boolean
'*

'Contiene el código Id de la aplicación, obtenido del archivo aplicación.Id.txt
'Se utilza en el módulo de ControlDeLicencia
'Además se utiliza para obtener información de la licencia de la aplicación
'en el formulario de AcercaDe
    Public idApli As Long
'*
