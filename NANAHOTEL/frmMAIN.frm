VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMAIN 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hotel"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   11910
   Icon            =   "frmMAIN.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   11910
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   1
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin Hotel_Nana.gaHOTELbarra gaHOTELbarra1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   7
      Top             =   7605
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   582
   End
   Begin MSFlexGridLib.MSFlexGrid gDerecha 
      Height          =   2415
      Left            =   5640
      TabIndex        =   6
      Top             =   3720
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   4260
      _Version        =   393216
      FixedCols       =   0
      BackColorBkg    =   -2147483643
      Redraw          =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   0
      SelectionMode   =   2
   End
   Begin VB.Frame Frame3 
      Caption         =   "para poner objetos"
      Height          =   1575
      Left            =   6240
      TabIndex        =   4
      Top             =   1200
      Visible         =   0   'False
      Width           =   2055
      Begin ComctlLib.ImageList ImageList2 
         Left            =   960
         Top             =   360
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   5
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMAIN.frx":030A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMAIN.frx":065C
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMAIN.frx":09AE
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMAIN.frx":0D00
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMAIN.frx":10C6
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin ComctlLib.ImageList ImageList1 
         Left            =   240
         Top             =   240
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         UseMaskColor    =   0   'False
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   7
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMAIN.frx":13E8
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMAIN.frx":173A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMAIN.frx":1A8C
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMAIN.frx":1DDE
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMAIN.frx":2130
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMAIN.frx":2482
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMAIN.frx":27D4
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame1"
      Height          =   7215
      Left            =   3480
      MousePointer    =   9  'Size W E
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   7215
      Left            =   4560
      MousePointer    =   9  'Size W E
      TabIndex        =   2
      Top             =   480
      Width           =   75
   End
   Begin ComctlLib.TreeView twMenu 
      Height          =   7155
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   12621
      _Version        =   327682
      Indentation     =   531
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
      MousePointer    =   4
   End
   Begin ComctlLib.ListView lwDerecha 
      Height          =   7155
      Left            =   4605
      TabIndex        =   5
      Top             =   480
      Width           =   7290
      _ExtentX        =   12859
      _ExtentY        =   12621
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      Icons           =   "ImageList2"
      SmallIcons      =   "ImageList2"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Menu mnuReservas 
      Caption         =   "&Reservas"
      Begin VB.Menu mnuReservasNueva 
         Caption         =   "Nueva"
      End
      Begin VB.Menu mnuReservasModificar 
         Caption         =   "Modificar"
      End
      Begin VB.Menu mnuReservasConsultar 
         Caption         =   "Consultar"
      End
      Begin VB.Menu mnuReservasAnular 
         Caption         =   "Anular"
      End
   End
   Begin VB.Menu mnuIngresoPasa 
      Caption         =   "&Ingreso pasajeros"
      Begin VB.Menu mnuIngresoPasaCheckin 
         Caption         =   "Checkin"
      End
      Begin VB.Menu mnuIngresoPasaWalkin 
         Caption         =   "Walkin"
      End
      Begin VB.Menu mnuIngresoPasaWalkinHabOcupada 
         Caption         =   "Walkin habitación ocupada"
      End
   End
   Begin VB.Menu menuGastos 
      Caption         =   "&Gastos"
      Begin VB.Menu menuGastosExtras 
         Caption         =   "Gastos extras"
      End
      Begin VB.Menu menuGastosAlojamiento 
         Caption         =   "Gastos alojamiento"
      End
      Begin VB.Menu menuGastosResumenHabitacion 
         Caption         =   "Resumen de cuenta habitación"
      End
      Begin VB.Menu menuGastosResumenClientes 
         Caption         =   "Resumen de cuenta clientes"
      End
   End
   Begin VB.Menu mnuFacturacion 
      Caption         =   "&Facturación"
      Begin VB.Menu mnuFacturacionFacturas 
         Caption         =   "Facturas"
         Begin VB.Menu mnuFacturacionFacturasEmitir 
            Caption         =   "Emitir factura"
         End
         Begin VB.Menu mnuFacturacionFacturasConsultar 
            Caption         =   "Consultar"
         End
         Begin VB.Menu mnuFacturacionFacturasAnular 
            Caption         =   "Anular"
         End
      End
      Begin VB.Menu mnuFacturacionDevoluciones 
         Caption         =   "Devoluciones"
         Begin VB.Menu mnuFacturacionDevolucionesEmitir 
            Caption         =   "Emitir devolución"
         End
         Begin VB.Menu mnuFacturacionDevolucionesConsultar 
            Caption         =   "Consultar"
         End
      End
      Begin VB.Menu mnuRecivos 
         Caption         =   "Recivos"
         Begin VB.Menu mnuRecivosIngresar 
            Caption         =   "Ingresar recivo"
         End
         Begin VB.Menu mnuRecivosConsultar 
            Caption         =   "Consultar"
         End
         Begin VB.Menu mnuRecivosAnular 
            Caption         =   "Anular"
         End
      End
   End
   Begin VB.Menu mnuCheckOut 
      Caption         =   "&CheckOut"
   End
   Begin VB.Menu mnuInformes 
      Caption         =   "I&nformes"
      Begin VB.Menu mnuInformesCuadroSituacion 
         Caption         =   "Cuadro de situación"
      End
      Begin VB.Menu mnuInformesDisponibilidad 
         Caption         =   "Cuadro de disponibilidad"
      End
      Begin VB.Menu mnuInformesSituacionActual 
         Caption         =   "Resumen de  situación actual"
      End
      Begin VB.Menu mnuInformesConsultaCompleta 
         Caption         =   "Consulta de habitaciones completa"
      End
      Begin VB.Menu mnuInformesIngresos 
         Caption         =   "Ingresos previstos"
      End
      Begin VB.Menu mnuInformesEgresos 
         Caption         =   "Egresos previstos"
      End
      Begin VB.Menu mnuInformesPasajerosHabitacion 
         Caption         =   "Pasajeros por habitación"
      End
      Begin VB.Menu mnuInformesPoblacionFlotante 
         Caption         =   "Población flotante"
      End
      Begin VB.Menu mnuInformesUbicacionPasajeros 
         Caption         =   "Ubicación de pasajeros"
      End
   End
   Begin VB.Menu mnuHabitacion 
      Caption         =   "&Habitaciones"
      Begin VB.Menu mnuHabitacionCambioTitular 
         Caption         =   "Cambio de titular"
      End
      Begin VB.Menu mnuHabitacionCambioFechaEgreso 
         Caption         =   "Cambio de fecha de egreso"
      End
      Begin VB.Menu mnuHabitacionCambioSituacion 
         Caption         =   "Cambio de situación"
      End
      Begin VB.Menu mnuHabitacionBloquear 
         Caption         =   "Bloquear "
      End
      Begin VB.Menu mnuHabitacionCambioHabitacion 
         Caption         =   "Cambio de habitación"
      End
      Begin VB.Menu mnuHabitacionCambioTarifa 
         Caption         =   "Cambio de tarifa"
      End
      Begin VB.Menu mnuHabitacionConsultaTitular 
         Caption         =   "Consulta de titular"
      End
   End
   Begin VB.Menu mnuCierreDiario 
      Caption         =   "Cierre &diario"
   End
   Begin VB.Menu mnuEstadoCuenta 
      Caption         =   "&Estados de cuenta"
   End
   Begin VB.Menu mnuSis 
      Caption         =   "&Sistema"
      Begin VB.Menu mnuSisCambioUsuario 
         Caption         =   "Cambio de usuario"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuSisStandBy 
         Caption         =   "Aplicación en Stand by"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuMante 
         Caption         =   "Mantenimiento"
         Begin VB.Menu mnuManteClientes 
            Caption         =   "Clientes"
         End
         Begin VB.Menu mnuManteEmpresas 
            Caption         =   "Empresas"
         End
         Begin VB.Menu mnuManteTarifas 
            Caption         =   "Tarifas de habitaciones"
         End
         Begin VB.Menu mnuMantArticulos 
            Caption         =   "Artículos"
         End
         Begin VB.Menu mnuManteNacionalidad 
            Caption         =   "Nacionalidades"
         End
         Begin VB.Menu mnuPaises 
            Caption         =   "Países"
         End
         Begin VB.Menu mnuMantePuntoVentas 
            Caption         =   "Punto de ventas"
         End
      End
      Begin VB.Menu mnuRecivosAuto 
         Caption         =   "Recivos automáticos"
         Begin VB.Menu mnuRecivosAutoImprimir 
            Caption         =   "Imprimir"
         End
         Begin VB.Menu mnuRecivosAutoConsultar 
            Caption         =   "Consultar"
         End
         Begin VB.Menu mnuRecivosAutoAnular 
            Caption         =   "Anular"
         End
      End
      Begin VB.Menu mnuSisCong 
         Caption         =   "Configuración del sistema"
      End
   End
   Begin VB.Menu mnuSalir 
      Caption         =   "Sa&lir"
   End
   Begin VB.Menu mnuAcerca 
      Caption         =   "&Acerca de ..."
   End
End
Attribute VB_Name = "frmMAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'El 1 de abril del 2002, remplazé este formulario que hasta el momento
'servía de menú principal de la aplicación, el cuál tenía las caracterísitcas de un árbol de
'nodos que se mostraban en una de las dos divisiones existentes en la pantalla (al estilo explorer)
'mientras que en la segunda divión se mostraba información de reservas, habitaciones, facturas, etc.
'Este formato se abandonó por conciderarlo demasiado complejo para el usuario final.
'Los procedimientos que realizaban parte del trabajo se encuentran en el módulo: MenuArbol
'el cual es imprecindible para el funcionamiento de este formulario.
'Tenemos entonces que este formulario fue remplazado por frmMAINver2, en el cual se implementa
'un nuevo formato de menu, basado en botones, el cual esperamos seá el definitivo y el correcto.

Option Explicit
Private xIni As Single  'trabaja para mover la barra divisoria
Private posX As Single  'utilizada para controlar los márgenes
Private AccesoPermitido As Boolean
Private PermitoDbleClik As Boolean

Private WithEvents PidoClave As UsuarioMuestro
Attribute PidoClave.VB_VarHelpID = -1
Private WithEvents ModoStandBy As UsuarioMuestro
Attribute ModoStandBy.VB_VarHelpID = -1

Private Sub Form_Activate()
    'Cuando el formulario activo es el main muestro barra de tareas
    Me.gaHOTELbarra1.Visible = True
End Sub

Private Sub Form_Deactivate()
    'Muestro la barra de tareas solo en el formulario activo
    Me.gaHOTELbarra1.Visible = False
End Sub

Private Sub Form_Load()
    'Abro base de datos
    mSubAbroBaseDeDatos
    
    'Cargo variables desde archivo parámetros
    mSubInicioAplicacion
    
    'pido autorización para entrar al programa
    subAutorizacion
    
    'Avilito la utilización del procedimiento de bitacora.dll
    Set ControlOperaciones = New GraboOperacion
    
    'Cargo nodos del menu
    If AccesoPermitido Then
        subCargoNodosMenu
    Else
        Unload Me
    End If
End Sub

Private Sub subAutorizacion()
    'Determino si el usuario puede ingresar a la aplicación
    
    AccesoPermitido = False
    'Valido acceso a la aplicación
    If tbPARAMETROS("SisAdminTF") = 0 Then
        'Nunca definí perfiles de usuario, por ese motivo
        'no pido contraseña ninguna.
        AccesoPermitido = True
        'culto opciones de usuario
        mnuSisCambioUsuario.Visible = False
        Me.mnuSisStandBy.Visible = False
        'tampoco permito cambiar el usuario con dblclick
        'sobre la barra de estado
        PermitoDbleClik = False
    Else
        'Tengo definido perfiles de usuarios por lo que
        'tengo que ingresar contraseña
        If PidoClave Is Nothing Then
            Set PidoClave = New UsuarioMuestro
        End If
        'Ejecuto dll para pedir contraseña
        PidoClave.MuestroUsuario tbSISTEMA_USUARIOS
        
        PermitoDbleClik = True
    End If
End Sub

Private Sub gaHOTELbarra1_DblClickSobreUsuario()
    'Hacer doble click sobre la barra de estado
    'equivale a ctrol+u
    
    'Esta bandera controla que cuando el sistema esta trabajando en modo
    'libre no trabaje tampoco el dblclick sobre la barra de estado
    If PermitoDbleClik Then
        mnuSisCambioUsuario_Click
    End If
End Sub

Private Sub mnuSisCambioUsuario_Click()
    'Cambio de usuario activo
    PidoClave.MuestroUsuario tbSISTEMA_USUARIOS
End Sub

Private Sub mnuSisStandBy_Click()
    'Muestra ventana de usuarios con la posibilidad
    'de salir del programa si no es un usuario
    If ModoStandBy Is Nothing Then
        Set ModoStandBy = New UsuarioMuestro
    End If
    ModoStandBy.MuestroUsuarioStandBy tbSISTEMA_USUARIOS
End Sub

Private Sub ModoStandBy_NotificoCliente(usuario As String, boton As Byte)
    'Este evento se ejecuta cuando hago clik
    'en algun boton del cuadro de dialogo de StandBy
    If boton = 1 Then
        'muestro usuario activo
        m_UsuarioSisNom = usuario
        Me.gaHOTELbarra1.InicializoUsuario
    Else
        'Termino con la ejecución del programa
        Unload Me
    End If
End Sub

Private Sub PidoClave_NotificoCliente(usuario As String, boton As Byte)
    'Este evento se ejecuta cuando hago click
    'en algun boton del cuadro de díalogo de contraseña de clientes
    If boton = 1 Then   'aceptar
        'muestro usuario activo
        m_UsuarioSisNom = usuario
        Me.gaHOTELbarra1.InicializoUsuario
        
        'muestro fecha del sistema
        Me.gaHOTELbarra1.InicializoFecha
        
        AccesoPermitido = True
    Else
        AccesoPermitido = False
    End If
End Sub

Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    xIni = Frame1.Left
    Frame2.Visible = True
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then  'boton del mouse apretado
        Frame2.Top = Frame1.Top
        If (xIni + X) > 600 And (xIni + X) < 9000 Then 'controlo márgenes
            posX = X
            Frame2.Left = xIni + posX
        End If
    End If
End Sub

Private Sub Frame1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Frame2.Visible = False
    'controlo márgenes nuevamente para corregir pequeña diferencia.
    If (xIni + posX) > 600 And (xIni + posX) < 9000 Then
        Frame1.Left = xIni + posX
        'Cambio ancho de ventana de árbol
        twMenu.Width = Frame1.Left + 15    'los 15 son para mejorar vista
        'Muevo ventana de grilla
        lwDerecha.Left = Frame1.Left + 40
        lwDerecha.Width = frmMAIN.Width - lwDerecha.Left 'y cambio tamaño
    End If
End Sub

Private Sub subCargoNodosMenu()
    'Principal
    twMenu.Nodes.Add , , "mnuMain", tbPARAMETROS("SisNombreHotelMenu"), 1, 2
    'Reservas
    twMenu.Nodes.Add "mnuMain", tvwChild, "mnuRes", "Reservas", 3
    twMenu.Nodes.Add "mnuRes", tvwChild, "mnuResNueva", "Nueva reserva", 4
    twMenu.Nodes.Add "mnuRes", tvwChild, "mnuResModificar", "Modificar reserva", 4
    twMenu.Nodes.Add "mnuRes", tvwChild, "mnuResConsultar", "Consultar reserva", 4
    twMenu.Nodes.Add "mnuRes", tvwChild, "mnuResAnular", "Anular reserva", 4
    'Ingreso pasajeros
    twMenu.Nodes.Add "mnuMain", tvwChild, "mnuIng", "Ingreso pasajeros", 3
    twMenu.Nodes.Add "mnuIng", tvwChild, "mnuIngCheckin", "Checkin", 4
    twMenu.Nodes.Add "mnuIng", tvwChild, "mnuIngWalkin", "Walkin libre", 4
    twMenu.Nodes.Add "mnuIng", tvwChild, "mnuIngWalkinO", "Walkin ocupada", 4
    'Gastos
    twMenu.Nodes.Add "mnuMain", tvwChild, "mnuGastos", "Gastos pasajeros", 1, 2
    twMenu.Nodes.Add "mnuGastos", tvwChild, "mnuGastosExtras", "Ingreso extras", 4
    twMenu.Nodes.Add "mnuGastos", tvwChild, "mnuGastosAloja", "Ingreso alojamiento", 4
    twMenu.Nodes.Add "mnuGastos", tvwChild, "mnuGastosResumenHab", "Resumen habitación", 4
    twMenu.Nodes.Add "mnuGastos", tvwChild, "mnuGastosResumenCli", "Resumen cliente", 4
    'Facturación
    twMenu.Nodes.Add "mnuMain", tvwChild, "mnuFact", "Facturación", 1, 2
        'facturas
    twMenu.Nodes.Add "mnuFact", tvwChild, "mnuFactFact", "Facturas", 3
    twMenu.Nodes.Add "mnuFactFact", tvwChild, "mnuFactFactEmitir", "Emitir", 4
    twMenu.Nodes.Add "mnuFactFact", tvwChild, "mnuFactFactConsultar", "Consultar", 4
    twMenu.Nodes.Add "mnuFactFact", tvwChild, "mnuFactFactAnular", "Anular", 4
        'devoluciones
    twMenu.Nodes.Add "mnuFact", tvwChild, "mnuFactDevo", "Devoluciones", 3
    twMenu.Nodes.Add "mnuFactDevo", tvwChild, "mnuFactDevoEmitir", "Emitir", 4
    twMenu.Nodes.Add "mnuFactDevo", tvwChild, "mnuFactDevoConsultar", "Consultar", 4
        'recivos
    twMenu.Nodes.Add "mnuFact", tvwChild, "mnuFactRes", "Recibos", 3
    twMenu.Nodes.Add "mnuFactRes", tvwChild, "mnuFactResNuevo", "Ingresar", 4
    twMenu.Nodes.Add "mnuFactRes", tvwChild, "mnuFactResConsultar", "Consultar", 4
    twMenu.Nodes.Add "mnuFactRes", tvwChild, "mnuFactResAnular", "Anular", 4
    'checkout
    twMenu.Nodes.Add "mnuMain", tvwChild, "mnuCheckout", "Checkout", 3
    'informes
    twMenu.Nodes.Add "mnuMain", tvwChild, "mnuInf", "Informes", 5
    twMenu.Nodes.Add "mnuInf", tvwChild, "mnuInfCuadroSitu", "Cuadro de situación", 4
    twMenu.Nodes.Add "mnuInf", tvwChild, "mnuInfCuadroDispo", "Cuadro de disponibilidad", 4
    twMenu.Nodes.Add "mnuInf", tvwChild, "mnuInfResumenActual", "Resumen de situación actual", 4
    twMenu.Nodes.Add "mnuInf", tvwChild, "mnuInfConsultaCompleta", "Consulta completa de habitaciones", 4
    twMenu.Nodes.Add "mnuInf", tvwChild, "mnuInfIngPre", "Ingresos previstos", 4
    twMenu.Nodes.Add "mnuInf", tvwChild, "mnuInfEgrPre", "Egresos previstos", 4
    twMenu.Nodes.Add "mnuInf", tvwChild, "mnuInfPasajerosHab", "Pasajeros por habitación", 4
    twMenu.Nodes.Add "mnuInf", tvwChild, "mnuInfPoblacionFlotante", "Población flotante", 4
    twMenu.Nodes.Add "mnuInf", tvwChild, "mnuInfUbicaconPasa", "Ubicación de pasajeros", 4
    'habitaciones
    twMenu.Nodes.Add "mnuMain", tvwChild, "mnuHab", "Habitaciones", 3
    twMenu.Nodes.Add "mnuHab", tvwChild, "mnuHabCambioTit", "Cambio de titular", 4
    twMenu.Nodes.Add "mnuHab", tvwChild, "mnuHabCambioFechaEgr", "Cambio fecha de egreso", 4
    twMenu.Nodes.Add "mnuHab", tvwChild, "mnuHabCambioSitu", "Cambio de situación", 4
    twMenu.Nodes.Add "mnuHab", tvwChild, "mnuHabBloq", "Bloquear", 4
    twMenu.Nodes.Add "mnuHab", tvwChild, "mnuHabCambioHab", "Cambio de habitación", 4
    twMenu.Nodes.Add "mnuHab", tvwChild, "mnuHabCambioTarifa", "Cambio de tarifa", 4
    twMenu.Nodes.Add "mnuHab", tvwChild, "mnuHabConsultaTit", "Consulta de titular", 4
    'cierre diario
    twMenu.Nodes.Add "mnuMain", tvwChild, "mnuCierre", "Cierre diario", 4
    'estados de cuenta
    twMenu.Nodes.Add "mnuMain", tvwChild, "mnuEstadoCuenta", "Estados de cuenta", 3
    'sistema
    twMenu.Nodes.Add "mnuMain", tvwChild, "mnuSis", "Sistema", 6
        'mantenimiento
    twMenu.Nodes.Add "mnuSis", tvwChild, "mnuSisMant", "Mantenimiento", 1, 2
    twMenu.Nodes.Add "mnuSisMant", tvwChild, "mnuSisMantCli", "Clientes", 4
    twMenu.Nodes.Add "mnuSisMant", tvwChild, "mnuSisMantEmp", "Empresas", 4
    twMenu.Nodes.Add "mnuSisMant", tvwChild, "mnuSisMantTarifas", "Tarifas", 4
    twMenu.Nodes.Add "mnuSisMant", tvwChild, "mnuSisMantArticulos", "Artículos", 4
    twMenu.Nodes.Add "mnuSisMant", tvwChild, "mnuSisMantNacionalidades", "Nacionalidades", 4
    twMenu.Nodes.Add "mnuSisMant", tvwChild, "mnuSisMantPaises", "Paises", 4
    twMenu.Nodes.Add "mnuSisMant", tvwChild, "mnuSisMantPuntoVenta", "Punto de venta", 4
        'recivos
    twMenu.Nodes.Add "mnuSis", tvwChild, "mnuSisReci", "Recivos", 3
    twMenu.Nodes.Add "mnuSisReci", tvwChild, "mnuSisReciImp", "Emitir", 4
    twMenu.Nodes.Add "mnuSisReci", tvwChild, "mnuSisReciConsultar", "Consultar", 4
    twMenu.Nodes.Add "mnuSisReci", tvwChild, "mnuSisReciAnular", "Anular", 4
        'configuracion del sistema
    twMenu.Nodes.Add "mnuSis", tvwChild, "mnuSisConf", "Configuración del sistema", 7
    'salir
    twMenu.Nodes.Add "mnuMain", tvwChild, "mnuSalir", "Salir", 1
    'acerca de
    twMenu.Nodes.Add "mnuMain", tvwChild, "mnuAcerca", "Acerca de...", 1
End Sub
Private Sub twMenu_KeyPress(KeyAscii As Integer)
    'Para ejecutar una opción tengo dos posibilidades ya sea
    'haciendo doble clik o presionando la tecja enter
    If KeyAscii = 13 Then 'enter
        twMenu_DblClick
    End If
End Sub

Private Sub twMenu_NodeClick(ByVal Node As ComctlLib.Node)
    'Cuando me posiciono sobre un nodo del tipo que muestran datos
    'muestro los datos correspondientes.
    Select Case Node.Key
        Case "mnuRes"
            subBorroListV
            subMuestroListView
            mSubMenuMuestroReservas
        Case "mnuIng"
            subBorroListV
            subMuestroListView
            mSubMenuMuestroIngresos
        Case "mnuFactFact"
            subBorroListV
            subMuestroListView
            mSubMenuMuestroFacturas
        Case "mnuFactRes"
            subBorroListV
            subMuestroListView
            mSubMenuMuestroRecivos
        Case "mnuFactDevo"
            subBorroListV
            subMuestroListView
            mSubMenuMuestroDevoluciones
        Case "mnuCheckout"
            subMuestroGrilla
            mSubMenuMuestroEgresos gDerecha
        Case "mnuHab"
            subMuestroGrilla
            mSubMenuMuestroHabitaciones gDerecha
    End Select
End Sub

Private Sub twMenu_DblClick()
    'Para ejecutar una opción tengo dos posibilidades ya sea
    'haciendo doble clik o presionando la tecla enter
    
    'Ejecuto opciones
    Select Case twMenu.SelectedItem.Key
        Case "mnuResNueva"
            mnuReservasNueva_Click
        Case "mnuResModificar"
            mnuReservasModificar_Click
        Case "mnuResConsultar"
            mnuReservasConsultar_Click
        Case "mnuAnular"
            mnuReservasAnular_Click
        Case "mnuIngCheckin"
            mnuIngresoPasaCheckin_Click
        Case "mnuIngWalkin"
            mnuIngresoPasaWalkin_Click
        Case "mnuIngWalkinO"
            mnuIngresoPasaWalkinHabOcupada_Click
        Case "mnuGastosExtras"
            menuGastosExtras_Click
        Case "mnuGastosAloja"
            menuGastosAlojamiento_Click
        Case "mnuGastosResumenHab"
            menuGastosResumenHabitacion_Click
        Case "mnuGastosResumenCli"
            menuGastosResumenClientes_Click
        Case "mnuFactFactEmitir"
            mnuFacturacionFacturasEmitir_Click
        Case "mnuFactFactConsultar"
            mnuFacturacionFacturasConsultar_Click
        Case "mnuFactFactAnular"
            mnuFacturacionFacturasAnular_Click
        Case "mnuFactDevoEmitir"
            mnuFacturacionDevolucionesEmitir_Click
        Case "mnuFactDevoConsultar"
            mnuFacturacionDevolucionesConsultar_Click
        Case "mnuFactResNuevo"
            mnuRecivosIngresar_Click
        Case "mnuFactResConsultar"
            mnuRecivosConsultar_Click
        Case "mnuFactResAnular"
            mnuRecivosAnular_Click
        Case "mnuCheckout"
            mnuCheckOut_Click
        Case "mnuInfCuadroSitu"
            mnuInformesCuadroSituacion_Click
        Case "mnuInfCuadroDispo"
            mnuInformesDisponibilidad_Click
        Case "mnuInfResumenActual"
            mnuInformesSituacionActual_Click
        Case "mnuInfConsultaCompleta"
            mnuInformesConsultaCompleta_Click
        Case "mnuInfIngPre"
            mnuInformesIngresos_Click
        Case "mnuInfEgrPre"
            mnuInformesEgresos_Click
        Case "mnuInfPasajerosHab"
            mnuInformesPasajerosHabitacion_Click
        Case "mnuInfPoblacionFlotante"
        
        Case "mnuInfUbicaconPasa"
            mnuInformesUbicacionPasajeros_Click
        Case "mnuHabCambioTit"
            mnuHabitacionCambioTitular_Click
        Case "mnuHabCambioFechaEgr"
            mnuHabitacionCambioFechaEgreso_Click
        Case "mnuHabCambioSitu"
            mnuHabitacionCambioSituacion_Click
        Case "mnuHabBloq"
            mnuHabitacionBloquear_Click
        Case "mnuHabCambioHab"
            mnuHabitacionCambioHabitacion_Click
        Case "mnuHabCambioTarifa"
            mnuHabitacionCambioTarifa_Click
        Case "mnuHabConsultaTit"
            mnuHabitacionConsultaTitular_Click
        Case "mnuCierre"
            mnuCierreDiario_Click
        Case "mnuEstadoCuenta"
            mnuEstadoCuenta_Click
        Case "mnuSisMantCli"
            mnuManteClientes_Click
        Case "mnuSisMantEmp"
            mnuManteEmpresas_Click
        Case "mnuSisMantTarifas"
            mnuManteTarifas_Click
        Case "mnuSisMantArticulos"
            mnuMantArticulos_Click
        Case "mnuSisMantNacionalidades"
            mnuManteNacionalidad_Click
        Case "mnuSisMantPaises"
            mnuPaises_Click
        Case "mnuSisMantPuntoVenta"
             mnuMantePuntoVentas_Click
        Case "mnuSisReciImp"
            mnuRecivosAutoImprimir_Click
        Case "mnuSisReciConsultar"
            mnuRecivosAutoConsultar_Click
        Case "mnuSisReciAnular"
            mnuRecivosAutoAnular_Click
        Case "mnuSisConf"
            mnuSisCong_Click
        Case "mnuSalir"
            mnuSalir_Click
    End Select
End Sub

Private Sub subSeleccionoOpcion(op As String)
    'Evaluo el nodo seleccionado y ejecuto la opcion correspondiente
    
End Sub

Private Sub subMuestroGrilla()
    Me.gDerecha.Height = 7155
    Me.gDerecha.Left = 4600
    Me.gDerecha.Top = 480
    Me.gDerecha.Width = 7295
    Me.gDerecha.Visible = True
    Me.lwDerecha.Visible = False
End Sub

Private Sub subMuestroListView()
    Me.gDerecha.Visible = False
    Me.lwDerecha.Visible = True
End Sub

Private Sub subBorroListV()
    lwDerecha.ColumnHeaders.Clear
    lwDerecha.ListItems.Clear
End Sub

Private Sub lwDerecha_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
    lwDerecha.SortKey = ColumnHeader.Index - 1
    ' Establece Verdadero en Sorted para ordenar la lista.
    lwDerecha.Sorted = True
End Sub

'******************************************************
'*
'*
'*  Click del menu flotante
'*
'*
'*
'******************************************************

Private Sub menuGastosAlojamiento_Click()
    'hora de inicio de la operación
    HoraIni = Time
    OprEjecutada = 9
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        tipo_accion_inghabitacion = 4
        frmIngHabitacion.Show 1
    End If
End Sub

Private Sub menuGastosExtras_Click()
    'hora de inicio de la operación
    HoraIni = Time
    OprEjecutada = 8
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        tipo_accion_inghabitacion = 1
        frmIngHabitacion.Show 1
    End If
End Sub

Private Sub menuGastosResumenClientes_Click()
    OprEjecutada = 11
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        tipo_accion_ConsultaCuentas = 2
        tipo_accion_IngEstadoCuenta = 2
        frmIngPaxEmp.Show 1
    End If
End Sub

Private Sub menuGastosResumenHabitacion_Click()
    OprEjecutada = 10
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        tipo_accion_ConsultaCuentas = 1
        tipo_accion_inghabitacion = 2
        frmIngHabitacion.Show 1
    End If
End Sub

Private Sub mnuCheckOut_Click()
    'hora de inicio de la operación
    HoraIni = Time
    OprEjecutada = 54
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        tipo_accion_inghabitacion = 8
        frmIngHabitacion.Show 1
    End If
End Sub

Private Sub mnuCierreDiario_Click()
    'hora de inicio de la operación
    HoraIni = Time

    OprEjecutada = 57
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        frmCierreDiario.Show 1
    End If
End Sub

Private Sub mnuEstadoCuenta_Click()
    OprEjecutada = 58
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        tipo_accion_IngEstadoCuenta = 1
        frmIngPaxEmp.Show 1
    End If
End Sub

Private Sub mnuFacturacionDevolucionesConsultar_Click()
    OprEjecutada = 16
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        tipo_accion_tipodocumento = 3
        tipo_accion_devo = 3
        frmTipoDocumento.Show 1
    End If
End Sub

Private Sub mnuFacturacionDevolucionesEmitir_Click()
    'hora de inicio de la operación
    HoraIni = Time
    OprEjecutada = 15
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        tipo_accion_tipodocumento = 2
        tipo_accion_devo = 1
        frmTipoDocumento.Show 1
    End If
End Sub

Private Sub mnuFacturacionFacturasAnular_Click()
    'hora de inicio de la operación
    HoraIni = Time
    OprEjecutada = 14
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        tipo_accion_facturas = 3
        tipo_accion_tipodocumento = 1
        frmTipoDocumento.Show 1
    End If
End Sub

Private Sub mnuFacturacionFacturasConsultar_Click()
    OprEjecutada = 13
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        tipo_accion_facturas = 2
        tipo_accion_tipodocumento = 1
        frmTipoDocumento.Show 1
    End If
End Sub

Private Sub mnuFacturacionFacturasEmitir_Click()
    'hora de inicio de la operación
    HoraIni = Time
    OprEjecutada = 12
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        tipo_accion_facturas = 1
        tipo_accion_inghabitacion = 6
        frmIngHabitacion.Show 1
    End If
End Sub

Private Sub mnuHabitacionBloquear_Click()
    'hora de inicio de la operación
    HoraIni = Time

    OprEjecutada = 33
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        tipo_accion_inghabitacion2 = 3
        frmIngHabitacion2.Show 1
    End If
End Sub

Private Sub mnuHabitacionCambioFechaEgreso_Click()
    'hora de inicio de la operación
    HoraIni = Time
    OprEjecutada = 31
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        tipo_accion_inghabitacion = 9
        frmIngHabitacion.Show 1
    End If
End Sub

Private Sub mnuHabitacionCambioHabitacion_Click()
    'hora de inicio de la operación
    HoraIni = Time
    OprEjecutada = 34
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        frmCambioHabitacion.Show 1
    End If
End Sub

Private Sub mnuHabitacionCambioSituacion_Click()
    'hora de inicio de la operación
    HoraIni = Time
    OprEjecutada = 32
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        tipo_accion_inghabitacion2 = 1
        frmIngHabitacion2.Show 1
    End If
End Sub

Private Sub mnuHabitacionCambioTarifa_Click()
    'hora de inicio de la operación
    HoraIni = Time
    OprEjecutada = 35
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        tipo_accion_inghabitacion = 3
        frmIngHabitacion.Show 1
    End If
End Sub

Private Sub mnuHabitacionCambioTitular_Click()
    'hora de inicio de la operación
    HoraIni = Time
    OprEjecutada = 30
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        tipo_accion_inghabitacion = 5
        frmIngHabitacion.Show 1
    End If
End Sub

Private Sub mnuHabitacionConsultaTitular_Click()
    OprEjecutada = 36
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        frmConsultaTitular.Show 1
    End If
End Sub

Private Sub mnuInformesConsultaCompleta_Click()
    OprEjecutada = 24
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        frmConsultaCompleta.Show 1
    End If
End Sub

Private Sub mnuInformesCuadroSituacion_Click()
    OprEjecutada = 21
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        frmCuadroHab.Show 1
    End If
End Sub

Private Sub mnuInformesDisponibilidad_Click()
    OprEjecutada = 22
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        frmVerDisponibilidad.Show 1
    End If
End Sub

Private Sub mnuInformesEgresos_Click()
    OprEjecutada = 26
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        tipo_accion_LisIngEgr = 2
        frmLisIngEgr.Show 1
    End If
End Sub

Private Sub mnuInformesIngresos_Click()
    OprEjecutada = 25
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        tipo_accion_LisIngEgr = 1
        frmLisIngEgr.Show 1
    End If
End Sub

Private Sub mnuInformesPasajerosHabitacion_Click()
    OprEjecutada = 27
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        tipo_accion_inghabitacion = 7
        frmIngHabitacion.Show 1
    End If
End Sub

Private Sub mnuInformesSituacionActual_Click()
    OprEjecutada = 23
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        frmEstadoActualHotel.Show 1
    End If
End Sub

Private Sub mnuInformesUbicacionPasajeros_Click()
    OprEjecutada = 29
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        'Busca los pasajeros que están hospedados en el hotel actualmente.
        Dim cli_aux As String
        
        cli_aux = mFunBusqueda(2)
        If Val(cli_aux) <> 0 Then
            cliente_a_ubicar = cli_aux
            frmConsultaPasajeros.Show 1
        End If
    End If
End Sub

Private Sub mnuIngresoPasaCheckin_Click()
    'hora de inicio de la operación
    HoraIni = Time
    OprEjecutada = 5
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        tipo_accion_reserva = "Check-in"
        frmModificacionReserva.Show 1
    End If
End Sub

Private Sub mnuIngresoPasaWalkin_Click()
    'hora de inicio de la operación
    HoraIni = Time
    OprEjecutada = 6
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        corr_reserva
        tipo_accion_checkin = 2
        frmCheck_in.Show 1
    End If
End Sub

Private Sub mnuIngresoPasaWalkinHabOcupada_Click()
    'hora de inicio de la operación
    HoraIni = Time
    OprEjecutada = 7
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        corr_reserva
        tipo_accion_checkin = 3
        frmCheck_in.Show 1
    End If
End Sub

Private Sub mnuMantArticulos_Click()
    'hora de inicio de la operación
    HoraIni = Time
    OprEjecutada = 42
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        frmMant_Articulos.Show 1
    End If
End Sub

Private Sub mnuManteClientes_Click()
    'hora de inicio de la operación
    HoraIni = Time
    OprEjecutada = 39
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        frmMant_Clientes.Show 1
    End If
End Sub

Private Sub mnuManteEmpresas_Click()
    'hora de inicio de la operación
    HoraIni = Time
    OprEjecutada = 40
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        frmMant_Empre.Show 1
    End If
End Sub

Private Sub mnuManteNacionalidad_Click()
    'hora de inicio de la operación
    HoraIni = Time
    OprEjecutada = 43
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        tipo_mantenimiento = 3
        frmMant_General_1.Show 1
    End If
End Sub

Private Sub mnuMantePuntoVentas_Click()
    OprEjecutada = 45
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        tipo_mantenimiento = 4
        frmMant_General_1.Show 1
    End If
End Sub

Private Sub mnuManteTarifas_Click()
    'hora de inicio de la operación
    HoraIni = Time
    OprEjecutada = 41
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        frmMant_Tarifas.Show 1
    End If
End Sub

Private Sub mnuPaises_Click()
    OprEjecutada = 44
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        tipo_mantenimiento = 2
        frmMant_General_1.Show 1
    End If
End Sub

Private Sub mnuRecivosAnular_Click()
    'hora de inicio de la operación
    HoraIni = Time
    OprEjecutada = 19
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        tipo_accion_recivo = 6
        tipo_accion_tipodocumento = 7
        frmTipoDocumento.Show 1
    End If
End Sub

Private Sub mnuRecivosAutoAnular_Click()
    OprEjecutada = 48
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        tipo_accion_tipodocumento = 4
        tipo_accion_recivo = 5
        frmTipoDocumento.Show 1
    End If
End Sub

Private Sub mnuRecivosAutoConsultar_Click()
    OprEjecutada = 47
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        tipo_accion_tipodocumento = 4
        tipo_accion_recivo = 3
        frmTipoDocumento.Show 1
    End If
End Sub

Private Sub mnuRecivosAutoImprimir_Click()
    'hora de inicio de la operación
    HoraIni = Time
    OprEjecutada = 46
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        tipo_accion_tipodocumento = 5
        tipo_accion_recivo = 1
        frmTipoDocumento.Show 1
    End If
End Sub

Private Sub mnuRecivosConsultar_Click()
    OprEjecutada = 18
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        tipo_accion_recivo = 4
        tipo_accion_tipodocumento = 7
        frmTipoDocumento.Show 1
    End If
End Sub

Private Sub mnuRecivosIngresar_Click()
    'hora de inicio de la operación
    HoraIni = Time
    OprEjecutada = 17
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        tipo_accion_recivo = 2
        tipo_accion_tipodocumento = 6
        frmTipoDocumento.Show 1
    End If
End Sub

Private Sub mnuReservasAnular_Click()
    'hora de inicio de la operación
    HoraIni = Time
    OprEjecutada = 4
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        tipo_accion_reserva = "ANULAR"
        frmModificacionReserva.Show 1
    End If
End Sub

Private Sub mnuReservasConsultar_Click()
    OprEjecutada = 3
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        tipo_accion_reserva = "CONSULTAR"
        frmModificacionReserva.Show 1
    End If
End Sub

Private Sub mnuReservasModificar_Click()
    'hora de inicio de la operación
    HoraIni = Time
    OprEjecutada = 2
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        tipo_accion_reserva = "MODIFICAR"
        frmModificacionReserva.Show 1
    End If
End Sub

Private Sub mnuReservasNueva_Click()
    'hora de inicio de la operación
    HoraIni = Time
    OprEjecutada = 1
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        tipo_accion_reserva = "ALTA"
        frmCargaReserva.Show 1
    End If
End Sub

Private Sub mnuSalir_Click()
    Unload Me
End Sub

Private Sub mnuSisCong_Click()
    'hora de inicio de la operación
    HoraIni = Time

    OprEjecutada = 49
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        frmSisConfig.Show 1
    End If
End Sub

