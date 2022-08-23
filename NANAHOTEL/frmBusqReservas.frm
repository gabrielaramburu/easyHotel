VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{08825A62-8182-11D6-AE38-FDECBDCC172B}#16.0#0"; "SeleccionRegistrosBD.ocx"
Begin VB.Form frmBusqReservas 
   Caption         =   "Búsqueda de reservas"
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   10980
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   10980
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   10320
      Top             =   5400
   End
   Begin SeleccionRegistrosBD.SeleccionBD SeleccionBD1 
      Height          =   4695
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   8281
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
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   9551
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   1
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFormualario 
      Caption         =   "&Formulario"
      Begin VB.Menu mnuFormularioSeleccion 
         Caption         =   "Seleccionar"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuFormularioCancelar 
         Caption         =   "Cancelar"
      End
      Begin VB.Menu mnuDiv1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFormularioCambiarCriterio 
         Caption         =   "Cambiar criterio"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuFormularioIrAIngreso 
         Caption         =   "Ir a ingreso de criterio"
         Shortcut        =   ^I
      End
   End
End
Attribute VB_Name = "frmBusqReservas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Declaración de propiedades
Public propTipoAccion As Byte               'Determina los tabs que se muestran en el formulario
Public propRetorno As String                'En esta propiedad se carga el valor devuelto por el control para
                                            'que la misma sea consultada desde el formulario que utiliza la ayuda

Private vecTitulos(9, 1) As String

Private Sub Form_Activate()
    'Como esta propiedad es de retorno
    'la inicializo cada vez que se activa el formulario.
    Me.propRetorno = Empty
End Sub

Private Sub Form_Load()
    'Cargo un vector con los títulos de los diferentes tab
    subCargoVectorTitulos
    
    'Crea las diferentes fichas que se necesitam dependiendo de la función para la que se
    'llama al formulario.
    subMuestroTabs
    
    'inicializo propiedades del control genéricas
    subInicializo
    
    'desencadeno evento
    Timer1.Enabled = True
End Sub

Private Sub Form_Resize()
    'Al modificar el tamaño del formulario también modifico el tamaño del
    'control de selección y el tamaño del tabs
    Me.TabStrip1.Width = Me.Width - 400
    Me.TabStrip1.Height = Me.Height - 1000
    
    Me.SeleccionBD1.Width = Me.TabStrip1.Width - 400
    Me.SeleccionBD1.Height = Me.TabStrip1.Height - 800
    Me.Refresh
End Sub

Private Sub subInicializo()
    'Inicializo las propiedades del control de selección que son comunes a todas las consultas
    'que se realizan con el mismo.
    Me.SeleccionBD1.TeclaSeleccion = 13         'la tecla de selección es el enter
    Me.SeleccionBD1.NroCampoInicial = 1         'por defecto ordeno por primer nombre
    Me.SeleccionBD1.IndiceCampoRetorno = 0      'retorno el número de reserva
    Me.SeleccionBD1.TablasRelacionadas = ""
    Me.SeleccionBD1.BaseDeDatos = vardir
End Sub
    
Private Sub subMuestroTabs()
    'Para cambiar orden de los tabs, cambiar aqui.
        
    Select Case propTipoAccion
        Case 1  'todas los tabs
            'cargo tag y descpción
            
            TabStrip1.Tabs.Add , , vecTitulos(8, 0)   'ingresan hoy
            TabStrip1.Tabs.Item(TabStrip1.Tabs.Count).Tag = vecTitulos(8, 1)
            
            TabStrip1.Tabs.Add , , vecTitulos(1, 0)   'vigente sin ocupar
            TabStrip1.Tabs.Item(TabStrip1.Tabs.Count).Tag = vecTitulos(1, 1)
            
            TabStrip1.Tabs.Add , , vecTitulos(2, 0)   'vigente ocupadas
            TabStrip1.Tabs.Item(TabStrip1.Tabs.Count).Tag = vecTitulos(2, 1)
            
            TabStrip1.Tabs.Add , , vecTitulos(3, 0)   'no vigentes
            TabStrip1.Tabs.Item(TabStrip1.Tabs.Count).Tag = vecTitulos(3, 1)
            
            TabStrip1.Tabs.Add , , vecTitulos(4, 0)   'anuladas
            TabStrip1.Tabs.Item(TabStrip1.Tabs.Count).Tag = vecTitulos(4, 1)
            
            TabStrip1.Tabs.Add , , vecTitulos(0, 0)   'todas
            TabStrip1.Tabs.Item(TabStrip1.Tabs.Count).Tag = vecTitulos(0, 1)
            
        Case 2
            'Este caso se utiliza para mostrar las reservas que se pueden modificar o anular.
            'Dichas reservas estan activas.
            'Para el caso de las reservas que ingresan hoy, se puede dar el caso de que la misma
            'ya halla ingresado al hotel, por lo que no se puede (anular ni modificar).
            'Puede ocurrir que una reserva contenga más de una habitación. Si es así,
            'y una o de ellas ya ingresó al hotel, entonces tampoco se podrá modificar ni anular la reserva.
            
            TabStrip1.Tabs.Add , , vecTitulos(7, 0)   'ambas
            TabStrip1.Tabs.Item(TabStrip1.Tabs.Count).Tag = vecTitulos(7, 1)
            
            TabStrip1.Tabs.Add , , vecTitulos(1, 0)   'vigente sin ocupar
            TabStrip1.Tabs.Item(TabStrip1.Tabs.Count).Tag = vecTitulos(1, 1)
            
            TabStrip1.Tabs.Add , , vecTitulos(8, 0)   'ingresan hoy (sin ocupar)
            TabStrip1.Tabs.Item(TabStrip1.Tabs.Count).Tag = vecTitulos(8, 1)

        Case 3
            'Este caso se utiliza al realizar un checkin.
            'Simpre trabajo con las reservas que ingresan hoy, pero tengo que separarlas
            'en dos tipos:
            'a) las que ingresan hoy pero todavía no ha ingresado al hotel ningún pasajero
            'b) las que ingresan hoy y si se ha presentado algún pasajero al hotel.
            'Al mostrarlas en la ayuda de reservas, es necesario diferenciarlas, ya que
            'la opción que más voy a utilizar, es la a)
            'Sin embrago la opción b) puede ser util cuando los pasajeros no ingresan
            'al mismo tiempo al hotel.
            
            TabStrip1.Tabs.Add , , vecTitulos(8, 0)   'ingresan hoy (sin ocupar)
            TabStrip1.Tabs.Item(TabStrip1.Tabs.Count).Tag = vecTitulos(8, 1)
            
            TabStrip1.Tabs.Add , , vecTitulos(6, 0)   'ingresan hoy (ya ingresaron)
            TabStrip1.Tabs.Item(TabStrip1.Tabs.Count).Tag = vecTitulos(6, 1)
    End Select
    
    'borro el tabs que la máquina crea por defecto
    TabStrip1.Tabs.Remove 1
End Sub

Private Sub subCargoVectorTitulos()
    'Establece los títulos de los diferentes tabs.
    'Para cambiar descripción de títulos cambiar aqui.
    
    vecTitulos(0, 0) = "&Todas"
    vecTitulos(0, 1) = 1
    vecTitulos(1, 0) = "&Futuras"
    vecTitulos(1, 1) = 2
    vecTitulos(2, 0) = "&Ocupadas"
    vecTitulos(2, 1) = 3
    vecTitulos(3, 0) = "&Cumplidas"
    vecTitulos(3, 1) = 4
    vecTitulos(4, 0) = "&Anuladas"
    vecTitulos(4, 1) = 5
    'Esta posición(5) no se usa.
    vecTitulos(5, 0) = ""
    vecTitulos(5, 1) = 6
    vecTitulos(6, 0) = "&Ingresan hoy (ingresaron)"
    vecTitulos(6, 1) = 7
    vecTitulos(7, 0) = "A&mbas"
    vecTitulos(7, 1) = "8"
    vecTitulos(8, 0) = "I&ngresan hoy (no ingresaron aún)"      'solo las no ocupadas
    vecTitulos(8, 1) = "9"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Descargo formulario de memoria
    Set frmBusqReservas = Nothing
End Sub

Private Sub SeleccionBD1_Seleccionar(ValorClaveTablaPrincipal As Variant)
    'Se selecciono una fila
    Me.propRetorno = ValorClaveTablaPrincipal
    'oculto el formulario para poder leer las propiedades del mismo
    Me.Visible = False
End Sub

Private Sub TabStrip1_Click()
    'Este evento se ejecuta cuando cambio los tabs activos.
    'Por lo tanto,en este momento tengo que cambiar las propiedades del control de selección
    
    Dim consultaComple As String
    
    'dependiendo de la ficha seleccionada, muestro uno u otro tipo de reservas
    Select Case TabStrip1.SelectedItem.Tag
        Case 1  'Todas
            Me.SeleccionBD1.tabla = "RESERVAS"
            Me.SeleccionBD1.campos = "0;RESERVAS;NroReserva;1500@1;RESERVAS;PrimerNombre;1500@2;RESERVAS;SegundoNombre;1500@3;RESERVAS;PrimerApellido;1500@4;RESERVAS;SegundoApellido;1500@5;RESERVAS;FechaIngreso;1500@6;RESERVAS;FechaEgreso;1500@"
            consultaComple = "" 'no realizo selección complementaria
            
        Case 2  'Vigentes sin ocupar (futuras)
            'selecciono las reservas cuya fecha de ingreso sea mayor a la fecha del sistema
            Me.SeleccionBD1.tabla = "RESERVAS"
            Me.SeleccionBD1.campos = "0;RESERVAS;NroReserva;1500@1;RESERVAS;PrimerNombre;1500@2;RESERVAS;SegundoNombre;1500@3;RESERVAS;PrimerApellido;1500@4;RESERVAS;SegundoApellido;1500@5;RESERVAS;FechaIngreso;1500@6;RESERVAS;FechaEgreso;1500@"
            consultaComple = "RESERVAS.fechaing > " & fechaSQL(m_FechaSis)
            
        Case 3  'Vigentes ocupadas
            'Aparecen las que:  las que estan dentro del período de ocupación (ocupadas)
            '                   (incluyendo las que ingresaron hoy y estan ocupadas y se van hoy)
            ' No aparecen las no show, es decir las que estan dentro del período de ocupación
            'pero no ingresaron al hotel, ya que no estan en el archivo CHECKIN.
            Me.SeleccionBD1.tabla = "RESERVAS"
            Me.SeleccionBD1.campos = "0;RESERVAS;NroReserva;1500@1;RESERVAS;PrimerNombre;1500@2;RESERVAS;SegundoNombre;1500@3;RESERVAS;PrimerApellido;1500@4;RESERVAS;SegundoApellido;1500@5;RESERVAS;FechaIngreso;1500@6;RESERVAS;FechaEgreso;1500@"
            consultaComple = "RESERVAS.fechaing <= " & fechaSQL(m_FechaSis) & _
            " and RESERVAS.fechaegr >= " & fechaSQL(m_FechaSis) & _
            " and RESERVAS.nroreserva IN " & _
            "(Select nroreserva From CHECKIN) "
            
        Case 4  'No vigentes
            'Incluye también las no show
            Me.SeleccionBD1.tabla = "RESERVAS"
            Me.SeleccionBD1.campos = "0;RESERVAS;NroReserva;1500@1;RESERVAS;PrimerNombre;1500@2;RESERVAS;SegundoNombre;1500@3;RESERVAS;PrimerApellido;1500@4;RESERVAS;SegundoApellido;1500@5;RESERVAS;FechaIngreso;1500@6;RESERVAS;FechaEgreso;1500@"
            consultaComple = "RESERVAS.fechaegr < " & fechaSQL(m_FechaSis)
            
        Case 5  'Anuladas
            Me.SeleccionBD1.tabla = "ANULADAS"
            Me.SeleccionBD1.campos = "0;ANULADAS;NroReserva;1500@1;ANULADAS;PrimerNombre;1500@2;ANULADAS;SegundoNombre;1500@3;ANULADAS;PrimerApellido;1500@4;ANULADAS;SegundoApellido;1500@5;ANULADAS;FechaIngreso;1500@6;ANULADAS;FechaEgreso;1500@"
            
        Case 7  'Ingresan hoy y ya ingresaron
            Me.SeleccionBD1.tabla = "RESERVAS"
            Me.SeleccionBD1.campos = "0;RESERVAS;NroReserva;1500@1;RESERVAS;PrimerNombre;1500@2;RESERVAS;SegundoNombre;1500@3;RESERVAS;PrimerApellido;1500@4;RESERVAS;SegundoApellido;1500@5;RESERVAS;FechaIngreso;1500@6;RESERVAS;FechaEgreso;1500@"
            consultaComple = "RESERVAS.fechaing = " & fechaSQL(m_FechaSis) & _
            " and RESERVAS.nroreserva IN " & _
            "(Select nroreserva From CHECKIN Where CHECKIN.fcheckdes = " & fechaSQL(m_FechaSis) & ")"
            
        Case 8  'Ambas (caso2 y caso 9)
            Me.SeleccionBD1.tabla = "RESERVAS"
            Me.SeleccionBD1.campos = "0;RESERVAS;NroReserva;1500@1;RESERVAS;PrimerNombre;1500@2;RESERVAS;SegundoNombre;1500@3;RESERVAS;PrimerApellido;1500@4;RESERVAS;SegundoApellido;1500@5;RESERVAS;FechaIngreso;1500@6;RESERVAS;FechaEgreso;1500@"
            consultaComple = "RESERVAS.fechaing >= " & fechaSQL(m_FechaSis) & _
            " and RESERVAS.nroreserva NOT IN " & _
            "(Select nroreserva From CHECKIN Where CHECKIN.fcheckdes = " & fechaSQL(m_FechaSis) & ")"
            
        Case 9  'Ingresan hoy pero todavía no lo hicieron
            Me.SeleccionBD1.tabla = "RESERVAS"
            Me.SeleccionBD1.campos = "0;RESERVAS;NroReserva;1500@1;RESERVAS;PrimerNombre;1500@2;RESERVAS;SegundoNombre;1500@3;RESERVAS;PrimerApellido;1500@4;RESERVAS;SegundoApellido;1500@5;RESERVAS;FechaIngreso;1500@6;RESERVAS;FechaEgreso;1500@"
            consultaComple = "RESERVAS.fechaing = " & fechaSQL(m_FechaSis) & _
            " and RESERVAS.nroreserva NOT IN " & _
            "(Select nroreserva From CHECKIN Where CHECKIN.fcheckdes = " & fechaSQL(m_FechaSis) & ")"
    End Select
    
    Me.SeleccionBD1.SeleccionComplementaria = consultaComple
    
    'una vez cargada las porpiedades del control lo muestro
    Me.SeleccionBD1.Mostrar
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub mnuFormularioSeleccion_Click()
    'Al digitar esta opción del menu, simulo que digite la tecla de selección del control
    'seleccionando de esta manera el control.
    Dim tecla As Long
    tecla = Me.SeleccionBD1.TeclaSeleccion
    SendKeys (Chr(tecla))
End Sub

Private Sub mnuFormularioCambiarCriterio_Click()
    'Equivale a presionar el boton de cambiar criterios
    Me.SeleccionBD1.CambiarCriterios
End Sub

Private Sub mnuFormularioIrAIngreso_Click()
    'Le doy el focus al control que ingresa el criterio
    Me.SeleccionBD1.CambiarValorCriterio
End Sub

Private Sub mnuFormularioCancelar_Click()
    Unload Me
End Sub

Private Sub Timer1_Timer()
    'Este evento se crea para poder visualizar la barra de progreso al momento
    'de abrir el formulario. Es decir, si no fuera por este evento, el formulario
    'recién se mostraría después de haber realizado la consulta, por lo que la interface
    'con el usuario se ve perjudicada, ya que se genera un pequeño tiempo de espera.
    
    'muestro primer consulta
    TabStrip1_Click
    Timer1.Enabled = False
End Sub
