VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{9A9C8E95-7C99-11D6-AE38-98046E05332B}#20.0#0"; "MantenimientoBD.ocx"
Object = "{08825A62-8182-11D6-AE38-FDECBDCC172B}#17.0#0"; "SeleccionRegistrosBD.ocx"
Begin VB.Form frmMantenimientos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento de base de datos"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   11910
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox lblDescMant 
      Height          =   1335
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "frmMantenimientos.frx":0000
      Top             =   2520
      Width           =   2895
   End
   Begin SeleccionRegistrosBD.SeleccionBD SeleccionBD1 
      Height          =   3675
      Left            =   120
      TabIndex        =   4
      Top             =   3840
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   6482
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
   Begin MantenimientoBD.MantBD MantBD1 
      Height          =   3735
      Left            =   3240
      TabIndex        =   3
      Top             =   120
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   6588
      Campo           =   ""
      ColorFondoDatos =   16744576
      ColorFondoGrilla=   16762705
      ColorCaracteresIngreso=   65535
      ColorFondoCampoIngreso=   16771513
      AnchoCeldas     =   315
      LargoCeldas     =   3000
      FuenteNombreCampo=   10
      FuenteDatosIngresados=   10
      FuenteDatosAIngresar=   10
   End
   Begin Hotel_Nana.gaHOTELbarra gaHOTELbarra1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   7605
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   582
   End
   Begin ComctlLib.TreeView twMantenimientos 
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   3625
      _Version        =   327682
      Indentation     =   441
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin VB.Label lblMantSel 
      AutoSize        =   -1  'True
      Caption         =   "lblMantSel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   900
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   2640
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMantenimientos.frx":0006
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMantenimientos.frx":0320
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuSeleccionar 
      Caption         =   "S&eleccionar mantenimiento"
      Begin VB.Menu mnuSeleccionarClientes 
         Caption         =   "Clientes"
         Shortcut        =   +{F1}
      End
      Begin VB.Menu mnuSeleccionarEmpresas 
         Caption         =   "Empresas"
         Shortcut        =   +{F2}
      End
      Begin VB.Menu mnuSeleccionarArticulos 
         Caption         =   "Artículos"
         Shortcut        =   +{F3}
      End
      Begin VB.Menu mnuSeleccionarPuntoDeVenta 
         Caption         =   "Puntos de venta"
         Shortcut        =   +{F4}
      End
      Begin VB.Menu mnuSeleccionarPaises 
         Caption         =   "Países"
         Shortcut        =   +{F5}
      End
      Begin VB.Menu mnuSeleccionarNacionalidades 
         Caption         =   "Nacionalidades"
         Shortcut        =   +{F6}
      End
      Begin VB.Menu mnuSeleccionarTarifas 
         Caption         =   "Tarifas"
         Shortcut        =   +{F7}
      End
   End
   Begin VB.Menu mnuOperaciones 
      Caption         =   "&Operaciones"
      Begin VB.Menu mnuOperacionesGuardar 
         Caption         =   "Guardar"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuOperacionesModificar 
         Caption         =   "Modificar"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuOperacionesEliminar 
         Caption         =   "Eliminar"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuOperacionesLimpiar 
         Caption         =   "Limpiar datos para nuevo ingreso"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuOperacionesProximo 
         Caption         =   "Próximo disponible"
         Shortcut        =   ^P
      End
      Begin VB.Menu div 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOperacionesCambiarCriterio 
         Caption         =   "Cambiar criterio"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuOperacionesIngresarCriterio 
         Caption         =   "Ir a ingresar criterio"
         Shortcut        =   ^I
      End
      Begin VB.Menu div2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOperacionesIrAMant 
         Caption         =   "Ir a mantenimiento"
         Shortcut        =   ^R
      End
   End
   Begin VB.Menu mnuSalir 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "frmMantenimientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private NodoSeleccionadoActualmente As String       'utilizada para no repetir tareas de inicialización
                                                    'si se da click sobre un nodo ya seleccinado

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        mnuSalir_Click
    End If
End Sub

Private Sub Form_Load()
    'Cargo árbol de opciones
    subCargoArbolOpciones
    'Inicializo propiedades generales del los controles
    subInicializoControles
    'inicializo etiquetas de informacion.
    Me.lblDescMant.Text = "Seleccione un archivo de la lista"
    Me.lblMantSel.Caption = "No hay archivo seleccionado"
    mSubBloqueoControlFormulario Me.lblDescMant, True
    'inicializo menu de opciones
    subInicializoOpcionesMenu False
End Sub

Private Sub subInicializoControles()
    'Cargo las propiedades de los controles que son genéricas, es decir
    'son independientes de la tabla con la cual se trabaja.
    Me.MantBD1.CaminoBaseDeDatos = vardir
    Me.SeleccionBD1.BaseDeDatos = vardir
    Me.MantBD1.ContraseñaBaseDeDatos = cContraseñaBD
    Me.SeleccionBD1.ContraseñaBaseDeDatos = cContraseñaBD
End Sub

Private Sub subCargoArbolOpciones()
    'Crea un nodo para cada tabla con la que se trabaja.
    'Dependiendo del permiso del usuario que ingresa a el formulario son los nodos que se
    'muestran.
    'También se configura el menu de opciones.
    
    Me.twMantenimientos.Nodes.Add , , "mant", "Mantenimientos", 2
    
    Me.mnuSeleccionarClientes.Visible = False
    OprEjecutada = 39
    If funUsuarioAutorizoSinMensaje(m_UsuarioSisNom, OprEjecutada) Then
        Me.twMantenimientos.Nodes.Add "mant", 4, "cli", "Clientes", 1
        Me.mnuSeleccionarClientes.Visible = True
    End If
    
    Me.mnuSeleccionarEmpresas.Visible = False
    OprEjecutada = 40
    If funUsuarioAutorizoSinMensaje(m_UsuarioSisNom, OprEjecutada) Then
        Me.twMantenimientos.Nodes.Add "mant", 4, "emp", "Empresas", 1
        Me.mnuSeleccionarEmpresas.Visible = True
    End If
    
    Me.mnuSeleccionarArticulos.Visible = False
    OprEjecutada = 42
    If funUsuarioAutorizoSinMensaje(m_UsuarioSisNom, OprEjecutada) Then
        Me.twMantenimientos.Nodes.Add "mant", 4, "art", "Artículos", 1
        Me.mnuSeleccionarArticulos.Visible = True
    End If
    
    Me.mnuSeleccionarPuntoDeVenta.Visible = False
    OprEjecutada = 45
    If funUsuarioAutorizoSinMensaje(m_UsuarioSisNom, OprEjecutada) Then
        Me.twMantenimientos.Nodes.Add "mant", 4, "pv", "Puntos de ventas", 1
        Me.mnuSeleccionarPuntoDeVenta.Visible = True
    End If
    
    Me.mnuSeleccionarPaises.Visible = False
    OprEjecutada = 44
    If funUsuarioAutorizoSinMensaje(m_UsuarioSisNom, OprEjecutada) Then
        Me.twMantenimientos.Nodes.Add "mant", 4, "paises", "Países", 1
        Me.mnuSeleccionarPaises.Visible = True
    End If
    
    Me.mnuSeleccionarNacionalidades.Visible = False
    OprEjecutada = 43
    If funUsuarioAutorizoSinMensaje(m_UsuarioSisNom, OprEjecutada) Then
        Me.twMantenimientos.Nodes.Add "mant", 4, "nacio", "Nacionalidades", 1
        Me.mnuSeleccionarNacionalidades.Visible = True
    End If
    
    Me.mnuSeleccionarTarifas.Visible = False
    OprEjecutada = 41
    If funUsuarioAutorizoSinMensaje(m_UsuarioSisNom, OprEjecutada) Then
        Me.twMantenimientos.Nodes.Add "mant", 4, "tarifas", "Tarifas", 1
        Me.mnuSeleccionarTarifas.Visible = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Descargo formulario de memoria
    Set frmMantenimientos = Nothing
End Sub

Private Sub MantBD1_CambioOperacion(operacionActual As Byte)
    'Devuelve la operación actual del control, es decir, el boton que está
    'activo.
    
    Me.mnuOperacionesEliminar.Enabled = False
    Me.mnuOperacionesGuardar.Enabled = False
    Me.mnuOperacionesModificar.Enabled = False
    
    Select Case operacionActual
        Case 0  'no esta en uso
            'no hago nada ya que inicialmente inicializo las opciones a enabled=false
        Case 1  'se esta ingresado (boton guardar)
            Me.mnuOperacionesGuardar.Enabled = True
        Case 2  'se esta modificando
            Me.mnuOperacionesModificar.Enabled = True
    End Select
End Sub

Private Sub MantBD1_ErrorEnIngreso(tipo As Byte, desc As String)
    'Se produjo un error en el ingreso de datos
    Select Case tipo
        Case 1
            'La cantidad de caracteres ingresado es menor que el mínimo permitido
            mSubMensaje 4, 85
        Case 2
            'El valor ingresado es menor al mínimo permitido
            mSubMensaje 4, 86
        Case 3
            'El valor ingresado es mayor al máximo permitido.
            mSubMensaje 4, 87
        Case 4
            'No se permiten valores nulos.
            mSubMensaje 4, 88
        Case 5
            'El formato de la fecha no es correcto
            mSubMensaje 3, 1
        Case 6
            'La fecha ingresada debe de ser menor igual a la fecha de hoy
            mSubMensaje 3, 5
        Case 7
            'la fecha ingresada no puede ser menor a la fecha de hoy.
            mSubMensaje 3, 2
        Case 8
            'El formato del número ingresado no es el correcto
            mSubMensaje 4, 89
    End Select
End Sub

Private Sub MantBD1_GotFocus()
    'Para mejorar la interface con el usuario, muestro ícono que indica que
    'le estoy dando el focus al control de mantenimiento
    Me.MantBD1.MuestroSeñalDeFocus True
    'asistencia a usuarios
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 194
End Sub

Private Sub MantBD1_LostFocus()
    'Al perder el focus oculto ícono
    Me.MantBD1.MuestroSeñalDeFocus False
    'asistencia a usuarios
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub MantBD1_SeEliminoTabla(claveEliminada As Variant)
    'actualizo grlla
    Me.SeleccionBD1.ActualizarDatos
    'Este evento se ejecuta al eliminar un registro de la tabla
    GraboBitacora "Eliminación Art. " & claveEliminada
End Sub

Private Sub MantBD1_SeGraboTabla(claveGrabada As Variant)
    'Este evento se ejecua al grabar un nuevo registro en la tabla
    
    'si modifique la tabla de clientes
    If Me.twMantenimientos.SelectedItem.Key = "cli" Then
        'primero modifico nombres del cliente
        subModificoNombre CLng(claveGrabada)

        'grabo campo nombre Completo
        subGraboCampoNombreCompleto CLng(claveGrabada)
    End If
    'actualizo grlla
    Me.SeleccionBD1.ActualizarDatos
    'grabo operación en bitácora
    GraboBitacora "Nuevo " & Me.twMantenimientos.SelectedItem.Key & " " & CStr(claveGrabada)
End Sub

Private Sub MantBD1_SeModificoTabla(claveModifica As Variant)
    'Este evento se ejecuta al modificar un nuevo registro en la tabla
    
    'si modifique la tabla de clientes
    If Me.twMantenimientos.SelectedItem.Key = "cli" Then
        'cambio tipo de letra a los nombres
        subModificoNombre CLng(claveModifica)
    End If
    'actualizo grlla
    Me.SeleccionBD1.ActualizarDatos
    'grabo operación en bitácora
    GraboBitacora "Modificación " & Me.twMantenimientos.SelectedItem.Key & " " & CStr(claveModifica)
End Sub

Private Sub subModificoNombre(cli As Long)
    'Modifico la primer letra del nombre a mayúscula y todo el apellido a mayúscula.
    Dim nomAux As String

    If busco_clienteTF(cli) Then
        tbCLIENTES.Edit
            nomAux = StrConv(tbCLIENTES("primer_nom_titular"), 2)   'convierte a minúsculas
            tbCLIENTES("primer_nom_titular") = StrConv(nomAux, 3)   'convierte la primera letra a mayúsculas
            nomAux = StrConv(tbCLIENTES("segundo_nom_titular"), 2)
            tbCLIENTES("segundo_nom_titular") = StrConv(nomAux, 3)
            tbCLIENTES("primer_ape_titular") = _
                        StrConv(tbCLIENTES("primer_ape_titular"), 1) 'convierte a mayúsculas
            tbCLIENTES("segundo_ape_titular") = _
                        StrConv(tbCLIENTES("segundo_ape_titular"), 1)
        tbCLIENTES.Update
    End If
End Sub

Private Sub subGraboCampoNombreCompleto(cli As Long)
    'Este campo es la concatenación de cuatro campos de la tabla de clientes
    'y como hay que obtener su valor, es necesario implementar este procedimiento.
    
    'busco cliente
    If busco_clienteTF(cli) Then
        tbCLIENTES.Edit
        'armo campo nombre completo
            tbCLIENTES("nombre_completo_titular") = _
                tbCLIENTES("primer_nom_titular") & " " & _
                tbCLIENTES("segundo_nom_titular") & " " & _
                tbCLIENTES("primer_ape_titular") & " " & _
                tbCLIENTES("segundo_ape_titular")
        tbCLIENTES.Update
    End If
End Sub

Private Sub SeleccionBD1_Seleccionar(ValorClaveTablaPrincipal As Variant)
    'Cuando se selecciona un elemento desde el control de selección,
    'muestro los datos en el control de mantenimiento
    Me.MantBD1.MostrarRegistro ValorClaveTablaPrincipal
    'le doy el focus al control de mantenimiento
    Me.MantBD1.SetFocus
End Sub

Private Sub subSeleccionoMante(nodo As String)
    'Muestro informacíon en etiquta de descripción
    subMuestroDescMant nodo
    'Preparo control de mantenimiento
    subPreparoControlMantenimiento nodo
    'Preparo control de selección
    subPreparoControlSeleccion nodo
    'inicializo menu de opciones
    subInicializoOpcionesMenu True
    'modifico acceso a opción próximo disponible
    Me.mnuOperacionesProximo.Enabled = CBool(Me.MantBD1.SugerirProxLibre)
    'después de seleccionar un mantenimiento le doy el focus a criterio
    mnuOperacionesIngresarCriterio_Click
End Sub

Private Sub twMantenimientos_NodeClick(ByVal Node As ComctlLib.Node)
    'Se seleccionó un nodo en el control tree
    If Node.Key <> NodoSeleccionadoActualmente Then
        subSeleccionoMante Node.Key
        NodoSeleccionadoActualmente = Node.Key
    End If
End Sub

Private Sub subPreparoControlMantenimiento(nodo As String)
    'Dependiendo del nodo seleccionado, preparo el control de mantenimiento.
    Dim muestro As Boolean
    muestro = True
    Me.MantBD1.MostrarBoton 1, True
    Me.MantBD1.MostrarBoton 3, True
    Select Case nodo
        Case "cli"
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
            
        Case "emp"
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
            
        Case "art"
            Me.MantBD1.tabla = "ARTICULOS"          'tabla a trabajar en el control
            Me.MantBD1.IndiceCampoClave = 0         'indice del campo clave de la tabla
            Me.MantBD1.SugerirProxLibre = 1         'propiedad sugerir próximo
            Me.MantBD1.TipoClave = 0                'tipo de clave de la tabla
                                                    '(determinada por usuario)
            'propiedad campo
            Me.MantBD1.campo = "Código;1;0;0;;;;1;9999999;0;0;;;;;@Descripción;0;1;0;1;255;0;;;;;;;;;@Tipo de I.V.A.;4;2;0;;;;;;;;;;IVA;1;0@Moneda;4;3;0;;;;;;;;;;MONEDAS;1;0@Punto de venta;4;4;0;;;;;;;;;;PUNTO_VENTA;1;0@Precio final;1;5;0;;;;1;999999;1;0;;;;;@"
            'propiedad integridad
            Me.MantBD1.integridad = "FAC_LINEAS;3;Facturas;No puedo eliminar un artículos que halla sido facturado@CUENTAS_EXTRA;7;Gastos extras;No se puede borrar un artículo que  halla sido consumido por un pasajero@"
        
        Case "pv"
            Me.MantBD1.tabla = "PUNTO_VENTA"        'tabla a trabajar en el control
            Me.MantBD1.IndiceCampoClave = 0         'indice del campo clave de la tabla
            Me.MantBD1.SugerirProxLibre = 1         'propiedad sugerir próximo
            Me.MantBD1.TipoClave = 0                'tipo de clave de la tabla
                                                    '(determinada por usuario)
            'propiedad campo
            Me.MantBD1.campo = "Código;1;0;0;;;;1;99999;0;0;;;;;@Descripción;0;1;0;1;255;0;;;;;;;;;@"
            'propiedad integridad
            Me.MantBD1.integridad = "ARTICULOS;4;Artículos;No se puede eliminar un punto de venta que tenga un artículo asignado.@"
        
        Case "paises"
            Me.MantBD1.tabla = "PAISES"             'tabla a trabajar en el control
            Me.MantBD1.IndiceCampoClave = 0         'indice del campo clave de la tabla
            Me.MantBD1.SugerirProxLibre = 1         'propiedad sugerir próximo
            Me.MantBD1.TipoClave = 0                'tipo de clave de la tabla
                                                    '(determinada por usuario)
            'propiedad campo
            Me.MantBD1.campo = "Código;1;0;0;;;;0;99999;0;0;;;;;@Descripción;0;1;0;1;255;0;;;;;;;;;@"
            'propiedad integridad
            Me.MantBD1.integridad = "CLIENTES;8;Clientes;No se puede eliminar un país que tenga un cliente asignado.@FAC_CABEZAL;8;Facturas;No puedo borrar un país que halla sido utilizado en una factura@"
            
        Case "nacio"
            Me.MantBD1.tabla = "NACIONALIDADES"      'tabla a trabajar en el control
            Me.MantBD1.IndiceCampoClave = 0          'indice del campo clave de la tabla
            Me.MantBD1.SugerirProxLibre = 1          'propiedad sugerir próximo
            Me.MantBD1.TipoClave = 0                 'tipo de clave de la tabla
                                                     '(determinada por usuario)
            'propiedad campo
            Me.MantBD1.campo = "Código;1;0;0;;;;1;99999;0;0;;;;;@Descripción;0;1;0;1;255;0;;;;;;;;;@"
            'propiedad integridad
            Me.MantBD1.integridad = "CLIENTES;14;Clientes;No se puede eliminar una nacionalidad asignada a un cliente@"
            
        Case "tarifas"
            Me.MantBD1.tabla = "TIPO_HABITACIONES"      'tabla a trabajar en el control
            Me.MantBD1.IndiceCampoClave = 0             'indice del campo clave de la t
            Me.MantBD1.SugerirProxLibre = 0            'propiedad sugerir próximo
            Me.MantBD1.TipoClave = 0                    'tipo de clave de la tabla
                                                        '(determinada por usuario)
            Me.MantBD1.campo = "Tipo de habitación;4;0;0;;;;;;;;;;TIPO_HABITACIONES;1;0@Tarifa;1;2;0;;;;1;9999999;1;0;;;;;@"
            'no se asigna propiedad integridad ya que no se permite eliminar registros
            Me.MantBD1.MostrarBoton 1, False    'para esta opcióm no permito agregar ni eliminar botones
            Me.MantBD1.MostrarBoton 3, False
        Case Else
            muestro = False
    End Select
    If muestro Then
        'armo el control para realizar el mantenimiento
        Me.MantBD1.IniciarMantenimiento
    End If
End Sub

Private Sub subPreparoControlSeleccion(nodo As String)
    'Dependiendo del nodo seleccionado preparo el control control de selección
    Dim muestro As Boolean
    
    muestro = True
    'primero asigno las propiedades que son común a todas las tablas
    Me.SeleccionBD1.TeclaSeleccion = 13         'la tecla de selección es el enter
    Me.SeleccionBD1.IndiceCampoRetorno = 0      'retorno la clave del archivo
    Me.SeleccionBD1.TablasRelacionadas = ""
    Select Case nodo
        Case "cli"
            Me.SeleccionBD1.IndiceCampoRetorno = 20      'retorno la clave del archivo
            Me.SeleccionBD1.tabla = "CLIENTES"
            Me.SeleccionBD1.campos = "1;CLIENTES;NombreCompleto;4500@2;CLIENTES;PrimerNombre;1500@3;CLIENTES;SegundoNombre;1500@4;CLIENTES;PrimerApellido;1500@5;CLIENTES;SegundoApellido;1500@17;CLIENTES;Documento;1500@6;CLIENTES;Dirección;3000@7;CLIENTES;Localidad;1500@1;PAISES;PaísResidencia;1500@9;CLIENTES;CódigoPostal;1500@10;CLIENTES;Teléfono;1500@11;CLIENTES;Fax;1500@20;CLIENTES;Email;1500@12;CLIENTES;OtrosTelFax;1500@1;SEXO;Sexo;1000@1;NACIONALIDADES;Nacionalidad;1500@15;CLIENTES;FechaNacimiento;1500@1;ESTADO_CIVIL;EstadoCivil;1000@18;CLIENTES;Ruc;1500@19;CLIENTES;Observaciones;5500@0;CLIENTES;Código;750@"
            Me.SeleccionBD1.TablasRelacionadas = "8;PAISES;0@13;SEXO;0@14;NACIONALIDADES;0@16;ESTADO_CIVIL;0@"
            Me.SeleccionBD1.NroCampoInicial = 3         'por defecto ordeno por primer apellido
            
        Case "emp"
            Me.SeleccionBD1.IndiceCampoRetorno = 8      'retorno la clave del archivo
            Me.SeleccionBD1.tabla = "EMPRESAS"
            Me.SeleccionBD1.campos = "1;EMPRESAS;Nombre;2500@2;EMPRESAS;RazónSocial;2500@3;EMPRESAS;Ruc;1500@4;EMPRESAS;Dirección;2500@5;EMPRESAS;Teléfono;1500@6;EMPRESAS;Fax;1500@7;EMPRESAS;Email;1500@8;EMPRESAS;Contacto;2000@0;EMPRESAS;Código;750@"
            Me.SeleccionBD1.NroCampoInicial = 0         'por defecto ordeno por nombre
            
        Case "art"
            Me.SeleccionBD1.tabla = "ARTICULOS"
            Me.SeleccionBD1.campos = "0;ARTICULOS;Código;750@1;ARTICULOS;Descripción;3500@1;PUNTO_VENTA;PuntoDeVenta;3000@1;MONEDAS;Moneda;1500@1;IVA;TipoIVA;1000@5;ARTICULOS;PrecioSinIVA;1100@"
            Me.SeleccionBD1.TablasRelacionadas = "4;PUNTO_VENTA;0@3;MONEDAS;0@2;IVA;0@"
            Me.SeleccionBD1.NroCampoInicial = 1         'por defecto ordeno por descripción
            
        Case "pv"
            Me.SeleccionBD1.tabla = "PUNTO_VENTA"
            Me.SeleccionBD1.campos = "0;PUNTO_VENTA;Código;750@1;PUNTO_VENTA;Descripción;10000@"
            Me.SeleccionBD1.NroCampoInicial = 1         'por defecto ordeno por descripción
            
        Case "paises"
            Me.SeleccionBD1.tabla = "PAISES"
            Me.SeleccionBD1.campos = "0;PAISES;Codigo;750@1;PAISES;Descripcion;10000@"
            Me.SeleccionBD1.NroCampoInicial = 1         'por defecto ordeno por descripción
            
        Case "nacio"
            Me.SeleccionBD1.tabla = "NACIONALIDADES"
            Me.SeleccionBD1.campos = "0;NACIONALIDADES;Código;750@1;NACIONALIDADES;Descripción;10000@"
            Me.SeleccionBD1.NroCampoInicial = 1         'por defecto ordeno por descripción
            
        Case "tarifas"
            Me.SeleccionBD1.IndiceCampoRetorno = 3      'retorno la clave del archivo
            Me.SeleccionBD1.tabla = "TIPO_HABITACIONES"
            Me.SeleccionBD1.campos = "1;TIPO_HABITACIONES;TipoHabitación;3000@2;TIPO_HABITACIONES;Tarifa;3000@3;TIPO_HABITACIONES;TotaldeHabitaciones;3000@0;TIPO_HABITACIONES;Código;750@"
            Me.SeleccionBD1.NroCampoInicial = 0         'por defecto ordeno por tipo de tarifa
            
        Case Else
            muestro = False
    End Select
    If muestro Then
        'una vez cargada las porpiedades del control lo muestro
        Me.SeleccionBD1.Mostrar
    End If
End Sub

Private Sub subMuestroDescMant(nodo As String)
    'Muestro información del mantenimiento seleccionado
    Select Case nodo
        Case "cli"
            lblDescMant.Text = "Permite trabajar con el archivo de clientes, " & _
                                    "pudiendo ingresar nuevos clientes, modificar datos o eliminar clientes ya existentes, " & _
                                    "o simplemente consultar información útil."
            'muestro mantenimiento seleccionado
            lblMantSel.Caption = "Clientes"

        Case "emp"
            lblDescMant.Text = "Permite trabajar con el archivo de empresas, " & _
                                    "pudiendo ingresar nuevas empresas, modificar datos o eliminar empresas ya existentes, " & _
                                    "o simplemente consultar información útil."
            'muestro mantenimiento seleccionado
            lblMantSel.Caption = "Empresas"

        Case "art"
            lblDescMant.Text = "Permite trabajar con el archivo de artículos, " & _
                                    "pudiendo ingresar nuevos artículos, modificar datos o eliminar artículos ya existentes, " & _
                                    "o simplemente consultar información útil."
            'muestro mantenimiento seleccionado
            lblMantSel.Caption = "Artículos"

        Case "pv"
            lblDescMant.Text = "Permite trabajar con el archivo de puntos de venta, " & _
                                    "pudiendo ingresar nuevos puntos de venta, modificar datos o eliminar punto de ventas ya existentes, " & _
                                    "o simplemente consultar información útil."
            'muestro mantenimiento seleccionado
            lblMantSel.Caption = "Punto de venta"

        Case "paises"
            lblDescMant.Text = "Permite trabajar con el archivo de países, " & _
                                    "pudiendo ingresar nuevos países, modificar datos o eliminar países ya existentes, " & _
                                    "o simplemente consultar información útil."
            'muestro mantenimiento seleccionado
            lblMantSel.Caption = "Paises"

        Case "nacio"
            lblDescMant.Text = "Permite trabajar con el archivo de nacionalidades, " & _
                                    "pudiendo ingresar nuevas nacionalidades, modificar datos o eliminar nacionalidades ya existentes, " & _
                                    "o simplemente consultar información útil."
            'muestro mantenimiento seleccionado
            lblMantSel.Caption = "Nacionalidades"

        Case "tarifas"
            lblDescMant.Text = "Permite trabajar con el archivo de tarifas, " & _
                                    "pudiendo modificar los importe de las tarifas correspondientes a cada tipo de habitación del hotel, " & _
                                    "o simplemente consultar el importe actual de las mismas."
            'muestro mantenimiento seleccionado
            lblMantSel.Caption = "Tarifas"
    End Select
End Sub

Private Sub mnuSeleccionarArticulos_Click()
    'Selecciono nodo desde menu
    subSeleccionoMante "art"
End Sub

Private Sub mnuSeleccionarClientes_Click()
    'Selecciono nodo desde menu
    subSeleccionoMante "cli"
End Sub

Private Sub mnuSeleccionarEmpresas_Click()
    'Selecciono nodo desde menu
    subSeleccionoMante "emp"
End Sub

Private Sub mnuSeleccionarNacionalidades_Click()
    'Selecciono nodo desde menu
    subSeleccionoMante "nacio"
End Sub

Private Sub mnuSeleccionarPaises_Click()
    'Selecciono nodo desde menu
    subSeleccionoMante "paises"
End Sub

Private Sub mnuSeleccionarPuntoDeVenta_Click()
    'Selecciono nodo desde menu
    subSeleccionoMante "pv"
End Sub

Private Sub mnuSeleccionarTarifas_Click()
    'Selecciono nodo desde menu
    subSeleccionoMante "tarifas"
End Sub

Private Sub subSeleccionoNodoTeclado(nodo As Byte)
    'Selecciono el elemento correspondiente
    Me.twMantenimientos.Nodes.Item(nodo).Selected = True
    'Ejecuto el evento click para realizar mantenimiento
    twMantenimientos_NodeClick twMantenimientos.SelectedItem
End Sub

Private Sub mnuOperacionesCambiarCriterio_Click()
    'Cambio el criterio desde la lista de criterios
    Me.SeleccionBD1.CambiarCriterios
End Sub

Private Sub mnuOperacionesIngresarCriterio_Click()
    'Cambio el valor de la cadena de caracteres del criterio
    Me.SeleccionBD1.CambiarValorCriterio
End Sub

Private Sub mnuOperacionesGuardar_Click()
    'Implemento tecla de función para boton guardar
    'Equivale a hacer click sobre el boton guardar
    Me.MantBD1.GragarRegistro
End Sub

Private Sub mnuOperacionesEliminar_Click()
    'Implemento tecla de función para boton eliminar
    'Equivale a hacer click sobre el boton eliminar
    Me.MantBD1.BorrarRegistro
End Sub

Private Sub mnuOperacionesLimpiar_Click()
    'Implemento tecla de función para boton limpiar
    'Equivale a hacer click sobre el boton limpiar
    Me.MantBD1.LimpioRegistro
End Sub

Private Sub mnuOperacionesModificar_Click()
    'Implemento tecla de función para boton modificar
    'Equivale a hacer click sobre el boton modificar
    Me.MantBD1.ModificarRegistro
End Sub

Private Sub mnuOperacionesProximo_Click()
    'Implemento tecla de función para boton próximo libre
    'Equivale a hacer click sobre el boton próximo
    Me.MantBD1.SugerirProximo
End Sub

Private Sub subInicializoOpcionesMenu(x As Boolean)
    'Establesco la propedad enabled de las diferentes opciones del menú
    Me.mnuOperaciones.Enabled = x
End Sub

Private Sub mnuOperacionesIrAMant_Click()
    'Le doy el focus al control de mantenimiento por medio de una tecla de función.
    MantBD1.SetFocus
End Sub

Private Sub mnuSalir_Click()
    Unload Me
End Sub

'*******************************************
'*
'*  Asistencia a usuarios
'*
'*******************************************

Private Sub twMantenimientos_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 192
End Sub

Private Sub SeleccionBD1_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 193
End Sub

Private Sub SeleccionBD1_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub twMantenimientos_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

