VERSION 5.00
Object = "{9A9C8E95-7C99-11D6-AE38-98046E05332B}#15.0#0"; "MantenimientoBD.ocx"
Begin VB.Form frmPrueba 
   BackColor       =   &H8000000A&
   Caption         =   "Form1"
   ClientHeight    =   5505
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   ScaleHeight     =   5505
   ScaleWidth      =   9480
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Empresas"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   2520
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   6960
      TabIndex        =   8
      Top             =   4320
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "OcultoBotones"
      Height          =   375
      Left            =   5400
      TabIndex        =   7
      Top             =   4320
      Width           =   1335
   End
   Begin VB.TextBox txtCli 
      Height          =   405
      Left            =   2280
      TabIndex        =   6
      Top             =   4560
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Motrar"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   4560
      Width           =   1575
   End
   Begin MantenimientoBD.MantBD MantBD1 
      Height          =   3615
      Left            =   1920
      TabIndex        =   4
      Top             =   600
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   6376
      Campo           =   ""
      Tabla           =   "ARTICULOS"
      SugerirProxLibre=   1
      ColorFondoDatos =   12697856
      ColorFondoGrilla=   16777088
      ColorCaracteresIngreso=   255
      ColorFondoCampoIngreso=   13553152
      AnchoCeldas     =   375
      LargoCeldas     =   2000
      FuenteNombreCampo=   12
      FuenteDatosIngresados=   10
      FuenteDatosAIngresar=   10
   End
   Begin VB.CommandButton botPaises 
      Caption         =   "Países"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton botArticulos 
      Caption         =   "Artículos"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton botTarifa 
      Caption         =   "Tarifa"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton botClientes 
      Caption         =   "Clientes"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1575
   End
   Begin VB.Menu mnuGuardar 
      Caption         =   "Guardar"
   End
   Begin VB.Menu mnuModificar 
      Caption         =   "Modificar"
   End
   Begin VB.Menu mnuBorrar 
      Caption         =   "Borrar"
   End
   Begin VB.Menu mnuLimpiar 
      Caption         =   "Limpiar"
   End
   Begin VB.Menu mnuSugerir 
      Caption         =   "SugerirPróxmo"
   End
End
Attribute VB_Name = "frmPrueba"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub botArticulos_Click()
    'Ejecuto mantenimiento de articulos
    'Asigno propiedades al control antes de utilizarlo
    'camino y nombre de base de datos
    Me.MantBD1.CaminoBaseDeDatos = "c:\NANAHOTEL\hotel.mdb"
    'tabla a trabajar en el control
    Me.MantBD1.tabla = "ARTICULOS"
    'indice del campo clave de la tabla
    Me.MantBD1.IndiceCampoClave = 0
    'tipo de clave de la tabla
    Me.MantBD1.TipoClave = 0 'determinada por usuario
    'propiedad campo
    'en esta propiedad se incluye los valores de un combo esto son:
    'Moneda Nacional
    'Dolares
    'Estos valores no se pueden modificar ya que todo el sistema, asume que 0 es para moneda nacional
    'y 1 para dólares, siendo estos valores extraídos del valor listindex del control listbox
    MantBD1.campo = "Descripción;0;1;0;1;255;0;;;;;;;;;@Tipo de I.V.A.;4;2;0;;;;;;;;;;IVA;1;0@Moneda;3;3;0;;;;;;;;;Moneda Nacional#Dolares#;;;@Punto de Venta;4;4;0;;;;;;;;;;PUNTO_VENTA;1;0@Precio final;1;5;0;;;;1;999999;1;0;;;;;@Código;1;0;0;;;;1;99999;0;0;;;;;@"
    'propiedad integridad
    MantBD1.integridad = "FAC_LINEAS;3;Facturas;No puedo eliminar un artículos que halla sido facturado@CUENTAS_EXTRA;7;Gastos extras;No se puede borrar un artículo que  halla sido consumido por un pasajero@"
    'propiedad sugerir próximo
    Me.MantBD1.SugerirProxLibre = 1
    'armo el control para realizar el mantenimiento
    Me.MantBD1.IniciarMantenimiento
End Sub

Private Sub botClientes_Click()
    'Ejecuto mantenimiento de clientes
    'camino y nombre de base de datos
    Me.MantBD1.CaminoBaseDeDatos = "c:\NANAHOTEL\hotel.mdb"
    
            Me.MantBD1.tabla = "CLIENTES"                   'tabla a trabajar en el control
            Me.MantBD1.TipoClave = 1                        'tipo correlativo
            Me.MantBD1.SugerirProxLibre = 0                 'propiedad sugerir próximo
            Me.MantBD1.TablaContador = "SISTEMA_PARAMETROS" 'tabla de correlativos
            Me.MantBD1.IndiceCampoClaveContador = 0         'indice del campo clave de la tabla
            Me.MantBD1.CampoCont = 3                        'posición del campo en la tabla sistema_parametros
                                                            'que almacena el próximo correlativo
            'propiedad campo
            Me.MantBD1.campo = "Primer nombre;0;2;1;1;255;0;;;;;;;;;@Segundo nombre;0;3;1;1;255;0;;;;;;;;;@Primer apellido;0;4;0;1;255;0;;;;;;;;;@Segundo apellido;0;5;1;1;255;0;;;;;;;;;@Dirección;0;6;1;1;255;0;;;;;;;;;@Localidad;0;7;1;1;255;0;;;;;;;;;@País de residencia;4;8;1;;;;;;;;;;PAISES;1;0@Código postal;0;9;1;1;255;0;;;;;;;;;@Teléfono;0;10;1;1;255;0;;;;;;;;;@Fax;0;11;1;1;255;0;;;;;;;;;@e-mail;0;20;1;1;255;0;;;;;;;;;@Otros tel/fax;0;12;1;1;255;0;;;;;;;;;@Sexo;4;13;1;;;;;;;;;;SEXO;1;0@Nacionalidad;4;14;1;;;;;;;;;;NACIONALIDADES;1;0@Fecha de nacimiento;2;15;1;;;;;;;;1;;;;@Estado civil;4;16;1;;;;;;;;;;ESTADO_CIVIL;1;0@Documento;0;17;1;1;255;0;;;;;;;;;@Ruc;0;18;1;1;255;0;;;;;;;;;@Observaciones;0;19;1;1;1000;1;;;;;;;;;@"
            'propiedad integridad
            Me.MantBD1.integridad = "CHECKIN;1;Alojamientos;No se puede eliminar un cliente que este actualmente alojado en el hotel.@CHECKOUT;2;Alojamientos históricos;No se puede eliminar un cliente que estuvo alojado en el hotel.@CUENTAS_ALOJA;1;Gastos de alojamiento;No se puede eliminar un cliente que tenga gastos de alojamiento.@CUENTAS_EXTRA;1;Gastos extras;No se puede eliminar un cliente que tenga gastos extras.@ESTADO_CUENTAS;2;Cuentas pendientes;No se puede eliminar un cliente que tenga deudas con el hotel.@FAC_CABEZAL;9;Facturas ;No se puede eliminar un cliente al que se le halla realizado una factura.@RECIVOS;3;Recivos;No se puede eliminar un cliente al que se la halla realizado un recibo.@"
    'armo el control para realizar el mantenimiento
    Me.MantBD1.IniciarMantenimiento
    
End Sub

Private Sub botPaises_Click()
    'Ejecuto mantenimiento de paises
    'Asigno propiedades al control antes de utilizarlo
    'camino y nombre de base de datos
    Me.MantBD1.CaminoBaseDeDatos = "c:\NANAHOTEL\hotel.mdb"
    'tabla a trabajar en el control
    Me.MantBD1.tabla = "PAISES"
    'indice del campo clave de la tabla
    Me.MantBD1.IndiceCampoClave = 0
    'tipo de clave de la tabla
    Me.MantBD1.TipoClave = 0 'determinada por usuario
    'propiedad campo
    Me.MantBD1.campo = "Código;1;0;0;;;;0;99999;0;0;;;;;@Descripción;0;1;0;1;255;0;;;;;;;;;@"
    'propiedad integridad
    Me.MantBD1.integridad = "CLIENTES;8;Clientes;No se puede eliminar un país que tenga un cliente asignado.@FAC_CABEZAL;8;Facturas;No puedo borrar un país que halla sido utilizado en una factura@"
    'propiedad sugerir próximo
    Me.MantBD1.SugerirProxLibre = 1
    'armo el control para realizar el mantenimiento
    Me.MantBD1.IniciarMantenimiento
End Sub

Private Sub botTarifa_Click()
    'Ejecuto mantenimiento de tarifas
    Me.MantBD1.CaminoBaseDeDatos = "c:\NANAHOTEL\hotel.mdb"
    Me.MantBD1.tabla = "TIPO_HABITACIONES"      'tabla a trabajar en el control
    Me.MantBD1.IndiceCampoClave = 0          'indice del campo clave de la t
    Me.MantBD1.SugerirProxLibre = 1         'propiedad sugerir próximo
    Me.MantBD1.TipoClave = 0                'tipo de clave de la tabla
                                            '(determinada por usuario)
    Me.MantBD1.campo = "Tipo de habitación;4;0;0;;;;;;;;;;TIPO_HABITACIONES;1;0@Tarifa;1;2;0;;;;1;9999999;1;0;;;;;@"
    'no se asigna propiedad integridad ya que no se permite eliminar registros
    Me.MantBD1.IniciarMantenimiento
End Sub

Private Sub Command3_Click()
    Me.MantBD1.CaminoBaseDeDatos = "c:\NANAHOTEL\hotel.mdb"
    Me.MantBD1.tabla = "EMPRESAS"                   'tabla a trabajar en el control
    Me.MantBD1.TipoClave = 1                        'tipo correlativo
    Me.MantBD1.SugerirProxLibre = 0                 'propiedad sugerir próximo
    Me.MantBD1.TablaContador = "SISTEMA_PARAMETROS" 'tabla de correlativos
    Me.MantBD1.IndiceCampoClaveContador = 0         'indice del campo clave de la tabla
    Me.MantBD1.CampoCont = 3                        'posición del campo en la tabla sistema_parametros
    
    'propiedad campo
    Me.MantBD1.campo = "Nombre;0;1;0;1;50;0;;;;;;;;;@RazónSocial;0;2;1;1;50;0;;;;;;;;;@Ruc;0;3;1;1;50;0;;;;;;;;;@Dirección;0;4;1;1;50;0;;;;;;;;;@Teléfono;0;5;1;1;50;0;;;;;;;;;@Fax;0;6;1;1;50;0;;;;;;;;;@Email;0;7;1;1;50;0;;;;;;;;;@Contacto;0;8;1;1;50;0;;;;;;;;;@"
    'propiedad integridad
    Me.MantBD1.integridad = "RESERVAS;9;Reservas;No se puede eliminar una empresa a la  que se le halla realizado una reserva.@ANULADAS;9;Reservas anuladas;No se puedo eliminar una empresa a la cual se le  halla anulado una reserva.@HABITACIONES;3;Habitaciones;No se puedo eliminar una empresa que sea titular (tipo único) de una habitación.@HABITACIONES;4;Habitaciones;No se puedo eliminar una empresa que sea titular (tipo alojamiento) de una habitación.@HABITACIONES;5;Habitaciones;No se puedo eliminar una empresa que sea titular (tipo extras) de una habitación.@"
    Me.MantBD1.IniciarMantenimiento
End Sub

Private Sub Command1_Click()
    Me.MantBD1.MostrarRegistro Me.txtCli
End Sub

Private Sub Command2_Click()
    Me.MantBD1.MostrarBoton Val(Text1.Text), False
End Sub

Private Sub Form_Load()
    Me.MantBD1.ContraseñaBaseDeDatos = ";PWD=manyacapo;"
End Sub

Private Sub MantBD1_ErrorEnIngreso(tipo As Byte, desc As String)
    MsgBox tipo & " " & desc
End Sub

Private Sub MantBD1_GotFocus()
    Me.MantBD1.MuestroSeñalDeFocus True
End Sub

Private Sub MantBD1_LostFocus()
    Me.MantBD1.MuestroSeñalDeFocus False
End Sub

Private Sub MantBD1_NoHayDatosSuficientes(archivo As String)
    'este evnto ocurre cuando se quiere trabajr con un combo que se inicializa
    'desde archivo y el mismo no tiene datos
    MsgBox "Debe de ingresar datos en el archivo " & archivo & " para continuar."
End Sub

Private Sub MantBD1_SeEliminoTabla(claveEliminada As Variant)
    MsgBox "se eliminó la tabla correctamente " & claveEliminada
End Sub

Private Sub MantBD1_SeGraboTabla(claveGrabada As Variant)
    'determino si estoy trabajando con la tablas de clientes
    'busco clientes
    
    MsgBox "Se grabo la tabla correctamente " & claveGrabada
End Sub

Private Sub MantBD1_SeModificoTabla(claveModifica As Variant)
    MsgBox "Se modificó la tabla correctamente " & claveModifica
End Sub

Private Sub mnuBorrar_Click()
    Me.MantBD1.BorrarRegistro
End Sub

Private Sub mnuGuardar_Click()
    Me.MantBD1.GragarRegistro
End Sub

Private Sub mnuLimpiar_Click()
    Me.MantBD1.LimpioRegistro
End Sub

Private Sub mnuModificar_Click()
    Me.MantBD1.ModificarRegistro
End Sub

Private Sub mnuSugerir_Click()
    Me.MantBD1.SugerirProximo
End Sub

