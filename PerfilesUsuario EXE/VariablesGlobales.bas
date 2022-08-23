Attribute VB_Name = "VariablesGlobales"
Option Explicit
'Manejo de base de datos
Public bdAplicacion As Database, bdWK As Workspace

Public tbSISTEMA_USUARIOS As Recordset
Public tbSISTEMA_PERFILES As Recordset
Public tbSISTEMA_OPERACIONES As Recordset
Public tbSISTEMA_PARAMETROS As Recordset
Public tbSISTEMA_LICENCIA As Recordset
Public tbSISTEMA_LISTADOS As Recordset
Public tbSISTEMA_MENSAJES As Recordset

Public m_vardirBD As String     'ubicación en disco de la base de datos
Public m_vardirRpt As String    'ubicación en disco de los reportes

Public CaminoBaseDeDatos As String
Public Resolucion As String

'Determina como se muestran los datos en el listview
Public m_TipoLista As Byte      '2= tipo lista
                                '3= tipo reporte


'Determina si estoy ejecutando una versión demo.
'La misma se inicializa en el módulo ControlDeLicenias
'y es utilizada en el formulario frmMain, cuando de produce el evento
'click de la opción del menu salir. Si la misma esta inicializa a true, se muestra el
'aviso de versión demo al salir de la aplicación.
    Public gEsUnaVersionDemo As Boolean
'*

'Contiene el código Id de la aplicación, obtenido del archivo aplicacion.Id.txt
'Es utilizado en el módulo de ControlDeLicencia
'Además se utiliza para obtener información de la licencia de la aplicación en el
'formulario de AcercaDe
    Public idApli As Long
'*
