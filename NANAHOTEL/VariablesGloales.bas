Attribute VB_Name = "VariablesGlobales"

'Declaración de variables de color configurables
    Public mSisColor_1DetalleDeGastos As OLE_COLOR
    Public mSisColor_2TotalDeGastosDiarios As OLE_COLOR
    Public mSisColor_3TotalDeGastosTitular As OLE_COLOR
    Public mSisColor_6SaldoMonedaNacional As OLE_COLOR
    Public mSisColor_7SaldoDolares As OLE_COLOR
    Public mSisColor_10CheckinSeleccionHab As OLE_COLOR
    Public mSisColor_11SeleccionHabLibre As OLE_COLOR
    Public mSisColor_12SeleccionHabOcupada As OLE_COLOR
    Public mSisColor_15FilaSeleccionada As OLE_COLOR
    Public mSisColor_18ControlesNoHabilitados As OLE_COLOR
    Public mSisColor_19FilaSeleccionadaTexto As OLE_COLOR
    
'Declaración de constantes de color fijas
    Public Const mConstSisColor_Blanco = &H80000005
    
'Declaración de variables de fuentes
    Public mSisFuente_1GeneralTipo As String
    Public msisFuente_1GeneralTam As Byte
'*

'Representa la fecha del sistema. Se carga desde el archivo parámetros y
'se incrementa en 1, cada vez que se ejecuta el cierre diario.
    Public m_FechaSis As Date

'*

'Representa al usuario actual en el sistema
    Public m_UsuarioSisNom As String
'*

'Conjunto de constantes utilizadas para representar los diferentes estados
'posibles de una habitación
'Éstos estados simpre serán los mismos, independientemente del hotel, por
'eso de implementa de esta manera
Public Enum estados
    libres = 0
    Ocupadas = 1
    reservada = 2
    Bloqueadas = 3
    Noasignadas = 4
    Total_estados = 5
End Enum
'*

'El vector se utiliza para guardar el nombre de los diferentes
'tipos de estados; los valores correspóndientes a cada elemento de la
'enumeracion corresponden con el índice de éste vector,
'en donde se encuntra la descripción correspondiente.
Public vec_estados(4) As String

'*
'Constantes usadas para determinar el color usado para trabajar con los diferentes
'tipos de estados de habitaciones
    Public Const const_color_ocupada As Long = &HF8A932           'azul
    Public Const const_color_reservada As Long = &HFF&            'rojo
    Public Const const_color_bloqueada As Long = &H8000&          'verde
    Public Const const_color_noAsignada As Long = &H80FF&            'marron
    
    'Determina el color de la barra de ocupación del hotel en consulta completa
    Public Const const_color_ocupacion As Long = &HFF0000            'azul
    'Determina el color de la barra de habitaciones limpias
    Public Const const_color_limpias As Long = &HC0FFC0      'verde claro
    'Determina el color de la barra de habitaciones sucias
    Public Const const_color_sucias As Long = &HC0FFFF       'amarillo claro
 '*
   
'Variables globales para usar como parametro entre formularios frmCargaReserva y frmMain
'La misma determina la función que desarrollará este formulario.
    Public tipo_accion_reserva As String
    Public nro_reserva As Long
'*
'Variable para determinar si la reserva seleccionada para consultar, esta anulada
    Public consulta_reserva_anulada As Byte
'*

'Variables para utilizar en los datas
'contiene la ubicación en disco de los reportes y de la base de datos
    Public vardir As String
    Public vardir2 As String
'*

    Public tipo_cuentas(4) As String

'Trabaja con el formulario frmIngHabitación, indicándole
'para que acción es que se pide llama a dicho formulario
'1 = para ingreso de extras
'2 = para consulta de cuentas
'3 = para tarifas
'4 = carga alojamiento manual
'5 = cambio titular
'6 = facturacion
'7 = pasajeros por habitacion
'8 = check-out
'9 = cambio de fecha de egreso
    Public tipo_accion_inghabitacion As Byte
'*

'Trabaja con el formulario frmIngHabitacion2
'1 = cambio de situación
'2 = consulta de situacion
'3 = bloqueo de habitaciones
    Public tipo_accion_inghabitacion2 As Byte
'*

'Trabaja con el formulario frmCheck_in
'1=Es un Checkin
'2=Es un Walkin
'3=Es un Walkin a una habitacion ocupada
    Public tipo_accion_checkin As Byte
'*

'Trabaja con el formulario de frmMant_General_1:
'paises             = 2
'nacionalidades     = 3
'punto de ventas    = 4
    Public tipo_mantenimiento As Byte
'*

'Trabaja con el formulario frmReservaSelehab
'1 = llamada para trabajar con las reservas
'2 = llamada para trabajar con el Checkin
'3 = llamada para trabajar con el WalkinL
'4 = llamada para trabajar con el walkinO
    Public tipo_accion_SeleccionHab As Byte
'*


'Sirve para trabajar con la ayuda der reservas
    Public tipo_accion_busqueda_reserva As Byte
'*

'Sirve para saber si cancelo el formulario de seleccion de habitaciones
    Public cancelo_seleccion_habitaciones As Boolean
'*

'Representa el largo máximo (todas las habitaciones),
'de cada barra de la gráfica del cuadro verdisponibilidad

    Public Const maximo_barr_graf = 6150
'*

'Sirve para determinar que cliente fue seleccionado en el
'opción ubicación formulario,
    Public cliente_a_ubicar As Long

'Sirve para determinar la función de frmIngEstadocuenta
'Luego de ingresar cliente y dependiendo del parámetro, llama a:
'1= frmEstadoCuentas, muestras el estado de cuenta del cliente
'2= frmConsultaCuentas, muestra los gastos pendientes de facturación de un cliente.
    Public tipo_accion_IngEstadoCuenta As Byte
'*

'Determina la funcionalidad de frmTipoDocumento
'1= selecciono facturas o boletas en ambas monedas
'e ingreso número de docuumento para anular o consular
'2=nueva devolución
'3= consulto o elimino devoluciones
'4= consulto o elimino recivo automático
'5= nuevo recivo automático.
'6= nuevo recivo manual
'7= consulto o elimino recivo manual

'NOTA:Cuando se realiza una factura nueva
'el tipo de documento se selecciona, dentro del formulario frmFacturación.
    Public tipo_accion_tipodocumento As Byte
'*

'Determina la funcionalidad del formulario de facturación
'1= nueva
'2= consulta
'3=anulación
    Public tipo_accion_facturas As Byte
'*

'Determina la funcionalidad del formulario frmConsultaCuentas
'1=muestra los gastos de una habitación ocupada actualmente (pido nro. habitación)
'2=muestra los gastos de un cliente (pido nro.cliente)
'Si el cliente está alojado actualmente en el hotel y es titular de alguna habitación,
'las dos consultas mostrarán los mismos datos.
    Public tipo_accion_ConsultaCuentas As Byte
'*

'Determina la funcionalidad de frmRecivos
'1= nuevo recivo automático (imprimir recivo)
'2= ingreso recivo manual
'3=consulto recivo automatico
'4=consulto recivo manual
'5=borro recivo automatico
'6=borro recivo manual
    Public tipo_accion_recivo As Byte
'*

'Determino la funcionalidad del formulario frmDevolucion
'1=nueva devolución
'2=anulo devolución
'3=consulta devolución
    Public tipo_accion_devo As Byte
'*

'Determina la funcionalidad del formulario frmCuadroHabInf
'1=Reservas
'2=Checkin
'3=Bloqueo
'4=No asignadas
    Public tipoAccionCuadroHabInf As Byte
'*

'Determina si estoy ejecutando una versión demo.
'La misma se inicializa en el módulo ControlDeLicenias
'y es utilizada en el formulario frmMain, cuando de produce el evento
'click de la opción del menu salir. Si la misma esta inicializa a true, se muestra el
'aviso de versión demo al salir de la aplicación.
'Tambien se utiliza en el formulario de AcercaDe y Main para determinar los datos
'a mostrar con respecto a la licenca de la aplicación.
    Public gEsUnaVersionDemo As Boolean
'*

'Determina el camino y la base de datos con la cual trabaja la aplicación
'La misma se inicializa desde el archivo de configuración de la aplicación: EasyHotel.txt
'Si el archivo no existe se crea mediante rutinas de Configuracion.DLL
'Esta variable se utiliza en frmMain_Load
    Public BaseDeDatosAplicacion As String
'*

'Contiene el código Id de la aplicación, obtenido del archivo
'aplicacion.Id.txt
'Esta variable se inicializa cuando se ejecuta la aplicación, dentro de la función
'mFunAplicacionValida que se encuentra en el módulo ControlDeLicencia.
'Se declara en este módulo ya que es utilizada por el formulario de AcercaDe...
'También es utilizada al momento de imprimir para obtener información acerca de la
'aplicación, esta información se imprime generalmente en el cabezal de los resportes.
    Public idApli As Long
'*

'Esta variable se inicializa a False en el procedimiento sub Main  (inicio aplicación)
'Luego en el evento frmMAIN_Load se puede establecer a True en caso de que la aplicación
'no se pueda ejecutar, por ejemplo:
'   copìa no registrada
'   versión demo fuera del período
'   el usuario cancela el cuadro de díalogo de ingreso de contraseña
'Debo de utilizar esta variable para poder ejecutar el evento frmMain_Load sin interrupciones,
'es decir evitando el uso de la sentencia Unload, ya que se origina un error si la utilizo
'dentro del evento. Por ese motivo luego de ejecutar la sentencia Load frmMain
'evalúo el contenido de ésta variable determinando si ejecuto la sentencia UnLoad frmMain(true)
'frmMain.Show (false)
    Public terminarEjecucion As Boolean
'*

'Estas variables contiene el signo que se utiliza para las dos monedas
'que permite el sistema moneda nacional y dólares.
'Las mismas se inicializan cuando se ejecuta la aplicación y son usadas durante todo el
'programa cada vez que se desee mostrar uno de dichos signos.
    Public gblSignoMonedaNacional As String
    Public gblSignoDolares As String
'*

