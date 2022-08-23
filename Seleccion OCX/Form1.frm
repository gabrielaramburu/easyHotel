VERSION 5.00
Object = "{08825A62-8182-11D6-AE38-FDECBDCC172B}#17.0#0"; "SeleccionRegistrosBD.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "dbMemo"
   ClientHeight    =   6510
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11610
   LinkTopic       =   "Form1"
   ScaleHeight     =   6510
   ScaleWidth      =   11610
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton botActualizar 
      Caption         =   "Actualizar"
      Height          =   375
      Left            =   360
      TabIndex        =   19
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton botHabOcu 
      Caption         =   "Hab. ocupadas"
      Height          =   375
      Left            =   1680
      TabIndex        =   18
      Top             =   4800
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "Selección complemetaria"
      Height          =   2655
      Left            =   6720
      TabIndex        =   9
      Top             =   3720
      Width           =   4815
      Begin VB.CommandButton botRecivos 
         Caption         =   "&Recivos"
         Height          =   375
         Left            =   1680
         TabIndex        =   17
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtFechaTrabajo 
         Height          =   360
         Left            =   120
         TabIndex        =   15
         Top             =   2160
         Width           =   1575
      End
      Begin VB.ComboBox Combo2 
         Height          =   360
         ItemData        =   "Form1.frx":0000
         Left            =   1680
         List            =   "Form1.frx":001F
         TabIndex        =   14
         Text            =   "Combo2"
         Top             =   1440
         Width           =   3015
      End
      Begin VB.CommandButton botReservas 
         Caption         =   "&Reservas"
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   1440
         Width           =   1335
      End
      Begin VB.ComboBox Combo1 
         Height          =   360
         ItemData        =   "Form1.frx":00CD
         Left            =   1680
         List            =   "Form1.frx":00F0
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   960
         Width           =   3015
      End
      Begin VB.CommandButton botSelCom 
         Caption         =   "Pasajeros"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton botDocSelCom 
         Caption         =   "Documentos "
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de trabajo para reservas"
         Height          =   240
         Left            =   120
         TabIndex        =   16
         Top             =   1920
         Width           =   2835
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cambio criterio"
      Height          =   375
      Left            =   4800
      TabIndex        =   8
      Top             =   3840
      Width           =   1695
   End
   Begin VB.CommandButton botTar 
      Caption         =   "Tarifas"
      Height          =   375
      Left            =   1560
      TabIndex        =   7
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton botNac 
      Caption         =   "Nacionalidades"
      Height          =   375
      Left            =   2880
      TabIndex        =   6
      Top             =   4320
      Width           =   1815
   End
   Begin VB.CommandButton botPa 
      Caption         =   "Países"
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton botPV 
      Caption         =   "Punto de Venta"
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   3840
      Width           =   1815
   End
   Begin VB.CommandButton botEmpresas 
      Caption         =   "Empresas"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   4800
      Width           =   1095
   End
   Begin SeleccionRegistrosBD.SeleccionBD SeleccionBD1 
      Height          =   3495
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   6165
      GrillaForeColor =   -2147483640
      GrillaBackColor =   -2147483643
      BeginProperty GrillaFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton botCli 
      Caption         =   "Clientes"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton botArt 
      Cancel          =   -1  'True
      Caption         =   "Artículos"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Menu mnuCambiarCri 
      Caption         =   "Cambiar criterio"
   End
   Begin VB.Menu mnuCambiarValor 
      Caption         =   "Cambiar valor criterio"
   End
   Begin VB.Menu mnuColumnas 
      Caption         =   "Columnas"
      Begin VB.Menu mnuColumnas1 
         Caption         =   "1era"
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu mnuColumnas2 
         Caption         =   "2da"
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu mnuColumnas3 
         Caption         =   "3era"
         Shortcut        =   ^{F3}
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub botActualizar_Click()
    Me.SeleccionBD1.ActualizarDatos
End Sub

Private Sub botArt_Click()
    Muestro "art"
End Sub

Private Sub botCli_Click()
    Muestro "cli"
End Sub

Private Sub botDocSelCom_Click()
    'primero asigno las propiedades que son común a todas las tablas
    Me.SeleccionBD1.TeclaSeleccion = 13         'la tecla de selección es el enter
    Me.SeleccionBD1.NroCampoInicial = 0         'por defecto ordeno por número de documento
    Me.SeleccionBD1.IndiceCampoRetorno = 0      'retorno el número del documento
    Me.SeleccionBD1.TablasRelacionadas = ""
    
    Me.SeleccionBD1.tabla = "FAC_CABEZAL"
    Me.SeleccionBD1.campos = "1;FAC_CABEZAL;NroDocumento;1500@2;FAC_CABEZAL;Fecha;1500@3;FAC_CABEZAL;Nombre;1500@4;FAC_CABEZAL;Dirección;1500@5;FAC_CABEZAL;Localidad;1500@10;FAC_CABEZAL;Total;1500@"
    
    'Selecciono solo los pasajeros hospedados
    Me.SeleccionBD1.SeleccionComplementaria = " tipo_docu = " & Me.Combo1.ItemData(Combo1.ListIndex)
    
    'una vez cargada las porpiedades del control lo muestro
    Me.SeleccionBD1.Mostrar
End Sub

Private Sub botEmpresas_Click()
    Muestro "emp"
End Sub

Private Sub botHabOcu_Click()
    Dim conscompleaux  As String
    Me.SeleccionBD1.TeclaSeleccion = 13         'la tecla de selección es el enter
    Me.SeleccionBD1.NroCampoInicial = 0         'por defecto ordeno por número de documento
    Me.SeleccionBD1.IndiceCampoRetorno = 0      'retorno el número del documento
    Me.SeleccionBD1.TablasRelacionadas = "1;TIPO_HABITACIONES;0@10;TIPO_ESTADO_HAB;1@"
    
    Me.SeleccionBD1.tabla = "HABITACIONES"
    Me.SeleccionBD1.campos = "0;HABITACIONES;Habitación;1500@1;TIPO_HABITACIONES;TipoHabitación;2500@2;TIPO_ESTADO_HAB;Situación;2500@9;HABITACIONES;Tarifa;1500@"
    
    
    
    'el join con el archivo TIPO_ESTADO_HAB debe der ser con los registros de tipo 2 (situaciones)
        conscompleaux = " TIPO_ESTADO_HAB.tipo_cod = 2 "
        'muestro solo las habitaciones que se encuentren en el archivo checkin (ocupadas).
        conscompleaux = conscompleaux & " and  HABITACIONES.nroHab IN " & _
        "(Select CHECKIN.nroHab " & _
        " from CHECKIN) "
    Me.SeleccionBD1.SeleccionComplementaria = conscompleaux
    
    'una vez cargada las porpiedades del control lo muestro
    Me.SeleccionBD1.Mostrar

End Sub

Private Sub botNac_Click()
    Muestro "nacio"
End Sub

Private Sub botPa_Click()
    Muestro "paises"
End Sub

Private Sub botPV_Click()
    Muestro "pv"
End Sub

Private Sub botRecivos_Click()
    Me.SeleccionBD1.TeclaSeleccion = 13         'la tecla de selección es el enter
    Me.SeleccionBD1.NroCampoInicial = 1         'por defecto ordeno por descripción
    Me.SeleccionBD1.IndiceCampoRetorno = 1      'retorno el código del artículo
    
    Me.SeleccionBD1.tabla = "RECIVOS"
    Me.SeleccionBD1.campos = "0;RECIVOS;TipoRecivo;1500@1;RECIVOS;NroRecivo;1500@2;RECIVOS;Fecha;1500@3;RECIVOS;RealizadoA;1500@5;RECIVOS;Importe;1500@1;MONEDAS;Moneda;1500@"
    Me.SeleccionBD1.TablasRelacionadas = "6;MONEDAS;0@"
    Me.SeleccionBD1.Mostrar
    
End Sub

Private Sub botReservas_Click()
    Dim consultaComple As String
    Me.SeleccionBD1.TeclaSeleccion = 13         'la tecla de selección es el enter
    Me.SeleccionBD1.NroCampoInicial = 1         'por defecto ordeno por primer nombre
    Me.SeleccionBD1.IndiceCampoRetorno = 0      'retorno el número de reserva
    Me.SeleccionBD1.TablasRelacionadas = ""
    
    Me.SeleccionBD1.tabla = "RESERVAS"
    Me.SeleccionBD1.campos = "0;RESERVAS;NroReserva;1500@1;RESERVAS;PrimerNombre;1500@2;RESERVAS;SegundoNombre;1500@3;RESERVAS;PrimerApellido;1500@4;RESERVAS;SegundoApellido;1500@5;RESERVAS;FechaIngreso;1500@6;RESERVAS;FechaEgreso;1500@"

    'dependiendo del valor del combo muestro, muestro uno u otro tipo de reservas
    Select Case Me.Combo2.ItemData(Combo2.ListIndex)
        Case 1  'Todas
            consultaComple = "" 'no realizo selección complementaria
            
        Case 2  'Vigentes sin ocupar (futuras)
            'selecciono las reservas cuya fecha de ingreso sea mayor a la fecha de hoy
            consultaComple = "RESERVAS.fechaing > " & fechaSQL(CDate(Me.txtFechaTrabajo.Text))
            
        Case 3  'Vigentes ocupadas
            'Aparecen las que:  las que estan dentro del período de ocupación (ocupadas)
            '                   (incluyendo las que ingresaron hoy y estan ocupadas y se van hoy)
            ' No aparecen las no show, es decir las que estan dentro del período de ocupación
            'pero no ingresaron al hotel, ya que no estan en el archivo CHECKIN.
            
            consultaComple = "RESERVAS.fechaing <= " & fechaSQL(CDate(Me.txtFechaTrabajo.Text)) & _
            " and RESERVAS.fechaegr >= " & fechaSQL(CDate(Me.txtFechaTrabajo.Text)) & _
            " and RESERVAS.nroreserva IN " & _
            "(Select nroreserva From CHECKIN) "
            
        Case 4  'No vigentes
            'Incluye también las no show
            consultaComple = "RESERVAS.fechaegr < " & fechaSQL(CDate(Me.txtFechaTrabajo.Text))
            
        Case 5  'Anuladas
            Me.SeleccionBD1.tabla = "ANULADAS"
            Me.SeleccionBD1.campos = "0;ANULADAS;NroReserva;1500@1;ANULADAS;PrimerNombre;1500@2;ANULADAS;SegundoNombre;1500@3;ANULADAS;PrimerApellido;1500@4;ANULADAS;SegundoApellido;1500@5;ANULADAS;FechaIngreso;1500@6;ANULADAS;FechaEgreso;1500@"
            
        Case 6  'No show
            'NOTA: este tipo de reservas no se muestra más ya que no es útil para nada.
        
        Case 7  'Ingresan hoy o ya ingresaron
            consultaComple = "RESERVAS.fechaing = " & fechaSQL(CDate(Me.txtFechaTrabajo.Text))
            
        Case 8  'Ambas (caso2 y caso 9)
            consultaComple = "RESERVAS.fechaing >= " & fechaSQL(CDate(Me.txtFechaTrabajo.Text)) & _
            " and RESERVAS.nroreserva NOT IN " & _
            "(Select nroreserva From CHECKIN Where CHECKIN.fcheckdes = " & fechaSQL(CDate(Me.txtFechaTrabajo.Text)) & ")"
            
        Case 9  'Ingresan hoy pero todavía no lo hicieron
            consultaComple = "RESERVAS.fechaing = " & fechaSQL(CDate(Me.txtFechaTrabajo.Text)) & _
            " and RESERVAS.nroreserva NOT IN " & _
            "(Select nroreserva From CHECKIN Where CHECKIN.fcheckdes = " & fechaSQL(CDate(Me.txtFechaTrabajo.Text)) & ")"
    End Select
    
    Me.SeleccionBD1.SeleccionComplementaria = consultaComple
    
    'una vez cargada las porpiedades del control lo muestro
    Me.SeleccionBD1.Mostrar
End Sub

Private Sub botSelCom_Click()

    Me.SeleccionBD1.TeclaSeleccion = 13         'la tecla de selección es el enter
    Me.SeleccionBD1.NroCampoInicial = 1         'por defecto ordeno por descripción
    Me.SeleccionBD1.IndiceCampoRetorno = 0      'retorno el código del artículo
    Me.SeleccionBD1.TablasRelacionadas = ""
    
    Me.SeleccionBD1.tabla = "CLIENTES"
    Me.SeleccionBD1.campos = "1;CLIENTES;NombreCompleto;1500@2;CLIENTES;PrimerNombre;1500@3;CLIENTES;SegundoNombre;1500@4;CLIENTES;PrimerApellido;1500@5;CLIENTES;SegundoApellido;1500@17;CLIENTES;Documento;1500@6;CLIENTES;Dirección;1500@7;CLIENTES;Localidad;1500@1;PAISES;PaísResidencia;1500@9;CLIENTES;CódigoPostal;1500@10;CLIENTES;Teléfono;1500@11;CLIENTES;Fax;1500@20;CLIENTES;Email;1500@12;CLIENTES;OtrosTelFax;1500@1;SEXO;Sexo;1500@1;NACIONALIDADES;Nacionalidad;1500@15;CLIENTES;FechaNacimiento;1500@1;ESTADO_CIVIL;EstadoCivil;1500@18;CLIENTES;Ruc;1500@19;CLIENTES;Observaciones;1500@0;CLIENTES;Código;1500@"
    Me.SeleccionBD1.TablasRelacionadas = "8;PAISES;0@13;SEXO;0@14;NACIONALIDADES;0@16;ESTADO_CIVIL;0@" & "0;CHECKIN;1@"
    
    'Selecciono solo los pasajeros hospedados
    Me.SeleccionBD1.SeleccionComplementaria = " CLIENTES.nrocorr = CHECKIN.nrocorrcli "
    
    'una vez cargada las porpiedades del control lo muestro
    Me.SeleccionBD1.Mostrar
End Sub

Private Sub botTar_Click()
    Muestro "tarifas"
End Sub

Private Sub Command1_Click()
    Me.SeleccionBD1.CambiarValorCriterio
End Sub

Private Sub Form_Load()
    Me.SeleccionBD1.BaseDeDatos = "c:\nanahotel\hotel.mdb"
    Me.SeleccionBD1.ContraseñaBaseDeDatos = ";PWD=manyacapo;"
    Me.Combo1.ListIndex = 0
    Me.Combo2.ListIndex = 0
    Me.txtFechaTrabajo.Text = "01/09/01"
End Sub

Private Sub Muestro(nodo As String)
    'Dependiendo del nodo seleccionado preparo el control control de selección
    Dim consulta As String
    'primero asigno las propiedades que son común a todas las tablas
    Me.SeleccionBD1.TeclaSeleccion = 13         'la tecla de selección es el enter
    Me.SeleccionBD1.NroCampoInicial = 1         'por defecto ordeno por descripción
    Me.SeleccionBD1.IndiceCampoRetorno = 0      'retorno el código del artículo
    Me.SeleccionBD1.TablasRelacionadas = ""
    Me.SeleccionBD1.SeleccionComplementaria = ""
    Select Case nodo
        Case "cli"
            Me.SeleccionBD1.tabla = "CLIENTES"
            Me.SeleccionBD1.campos = "1;CLIENTES;NombreCompleto;1500@2;CLIENTES;PrimerNombre;1500@3;CLIENTES;SegundoNombre;1500@4;CLIENTES;PrimerApellido;1500@5;CLIENTES;SegundoApellido;1500@17;CLIENTES;Documento;1500@6;CLIENTES;Dirección;1500@7;CLIENTES;Localidad;1500@1;PAISES;PaísResidencia;1500@9;CLIENTES;CódigoPostal;1500@10;CLIENTES;Teléfono;1500@11;CLIENTES;Fax;1500@20;CLIENTES;Email;1500@12;CLIENTES;OtrosTelFax;1500@1;SEXO;Sexo;1500@1;NACIONALIDADES;Nacionalidad;1500@15;CLIENTES;FechaNacimiento;1500@1;ESTADO_CIVIL;EstadoCivil;1500@18;CLIENTES;Ruc;1500@19;CLIENTES;Observaciones;1500@0;CLIENTES;Código;1500@"
            Me.SeleccionBD1.TablasRelacionadas = "8;PAISES;0@13;SEXO;0@14;NACIONALIDADES;0@16;ESTADO_CIVIL;0@"
            
        Case "emp"
            Me.SeleccionBD1.tabla = "EMPRESAS"
            Me.SeleccionBD1.campos = "0;EMPRESAS;Código;1500@1;EMPRESAS;Nombre;1500@2;EMPRESAS;RazónSocial;1500@3;EMPRESAS;Ruc;1500@4;EMPRESAS;Dirección;1500@5;EMPRESAS;Teléfono;1500@6;EMPRESAS;Fax;1500@7;EMPRESAS;Email;1500@8;EMPRESAS;Contacto;1500@"
            
        Case "art"
'            consulta = "Select ARTICULOS.nroArticulo as Código, " & _
'                               "ARTICULOS.descriArticulo as Descripción, " & _
'                               "PUNTO_VENTA.descripcionPv as PuntoDeVenta, " & _
'                               "MONEDAS.descMoneda as Moneda, " & _
'                               "IVA.descIva as TipoIva, " & _
'                               "ARTICULOS.precioArticulo as PrecioSinIva " & _
'                        "From ARTICULOS,PUNTO_VENTA,MONEDAS,IVA " '& _
'                        "Where ARTICULOS.puntoVentaArticulo = PUNTO_VENTA.nroPv and " & _
'                              "ARTICULOS.monedaArticulo = MONEDAS.codMoneda and " & _
'                              "ARTICULOS.CodIvaArticulo = IVA.codIva "
'
'            Me.SeleccionBD1.MostrarRapido consulta
            Me.SeleccionBD1.tabla = "ARTICULOS"
            Me.SeleccionBD1.campos = "0;ARTICULOS;Código;750@1;ARTICULOS;Descripción;3500@1;PUNTO_VENTA;PuntoDeVenta;3000@1;MONEDAS;Moneda;1500@1;IVA;TipoIVA;1000@5;ARTICULOS;PrecioSinIVA;1100@"
            Me.SeleccionBD1.TablasRelacionadas = "4;PUNTO_VENTA;0@3;MONEDAS;0@2;IVA;0@"

        Case "pv"
            Me.SeleccionBD1.tabla = "PUNTO_VENTA"
            Me.SeleccionBD1.campos = "0;PUNTO_VENTA;Código;750@1;PUNTO_VENTA;Descripción;3000@"
            
        Case "paises"
            Me.SeleccionBD1.tabla = "PAISES"
            Me.SeleccionBD1.campos = "0;PAISES;Codigo;1000@1;PAISES;Descripcion;3000@"
            
        Case "nacio"
            Me.SeleccionBD1.tabla = "NACIONALIDADES"
            Me.SeleccionBD1.campos = "0;NACIONALIDADES;Código;1500@1;NACIONALIDADES;Descripción;3000@"
            
        Case "tarifas"
            Me.SeleccionBD1.tabla = "TIPO_HABITACIONES"
            Me.SeleccionBD1.campos = "1;TIPO_HABITACIONES;TipoHabitación;3000@2;TIPO_HABITACIONES;Tarifa;3000@3;TIPO_HABITACIONES;TotaldeHabitaciones;500@"
    
    End Select
    'una vez cargada las porpiedades del control lo muestro
    Me.SeleccionBD1.Mostrar
End Sub

Private Sub mnuCambiarCri_Click()
    Me.SeleccionBD1.CambiarCriterios
End Sub

Private Sub mnuCambiarValor_Click()
    Me.SeleccionBD1.CambiarValorCriterio
End Sub

Private Sub mnuColumnas1_Click()
    Me.SeleccionBD1.ClickColumna 0
End Sub

Private Sub mnuColumnas2_Click()
    Me.SeleccionBD1.ClickColumna 1
End Sub

Private Sub mnuColumnas3_Click()
    Me.SeleccionBD1.ClickColumna 6
End Sub

Private Sub SeleccionBD1_Seleccionar(ValorClaveTablaPrincipal As Variant)
    MsgBox ValorClaveTablaPrincipal
End Sub

Private Function fechaSQL(f As Variant)
    'Devuelve una cadena de caracters de formato: #MM/DD/AA#
    'con el objetivo de poder usarla en una consulta SQL
    
    Dim aux As String
    If IsDate(f) Then
        aux = Format(f, "mm/dd/yy")
        fechaSQL = "#" & aux & "#"
    End If
End Function


