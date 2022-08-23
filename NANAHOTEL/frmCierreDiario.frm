VERSION 5.00
Object = "{EB8C3860-8603-11D0-A35D-7062E4000000}#1.1#0"; "VCBND.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmCierreDiario 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CierreDiario"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   9510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5895
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   10398
      _Version        =   327680
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Realizar"
      TabPicture(0)   =   "frmCierreDiario.frx":0000
      Tab(0).ControlCount=   2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      TabCaption(1)   =   "Consultar cierres anteriores"
      TabPicture(1)   =   "frmCierreDiario.frx":001C
      Tab(1).ControlCount=   1
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(0).Enabled=   0   'False
      Begin VB.Frame Frame3 
         Caption         =   "Consulta de cierres diarios"
         Height          =   5175
         Left            =   -74880
         TabIndex        =   13
         Top             =   480
         Width           =   9015
         Begin VB.CommandButton botConsultarCierreDiario 
            Height          =   375
            Left            =   7560
            Picture         =   "frmCierreDiario.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   16
            Tag             =   "Imprimir"
            Top             =   840
            Width           =   1215
         End
         Begin VcBndCtl.VcCalCombo fechaCierreConsulta 
            Height          =   375
            Left            =   240
            TabIndex        =   15
            Top             =   720
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _0              =   $"frmCierreDiario.frx":097A
            _1              =   $"frmCierreDiario.frx":0D83
            _2              =   $"frmCierreDiario.frx":118C
            _3              =   ")@f-@@@E@@@A@@@@@@@@@'@@@E@@@A@@@@@@@@@)h@@@E@@@A@@@@@@@@@,467D"
            _count          =   4
            _ver            =   2
         End
         Begin VB.Image Image1 
            Height          =   105
            Left            =   240
            Picture         =   "frmCierreDiario.frx":1595
            Stretch         =   -1  'True
            Top             =   1440
            Width           =   8490
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "&Ingrese fecha del cierre diario"
            Height          =   240
            Left            =   240
            TabIndex        =   14
            Top             =   360
            Width           =   2655
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Cierre diario"
         Height          =   3375
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   9015
         Begin VB.Data Data1 
            Caption         =   "Data1"
            Connect         =   "Access"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   405
            Left            =   4800
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   2  'Snapshot
            RecordSource    =   "select * from cierre_diario"
            Top             =   120
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.CommandButton botProcesar 
            Caption         =   "&Procesar"
            Height          =   375
            Left            =   7440
            TabIndex        =   8
            Top             =   480
            Width           =   1275
         End
         Begin MSFlexGridLib.MSFlexGrid gProcesos 
            Height          =   2295
            Left            =   240
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   960
            Width           =   8535
            _ExtentX        =   15055
            _ExtentY        =   4048
            _Version        =   393216
            Rows            =   11
            FixedRows       =   0
            FixedCols       =   0
            GridLines       =   0
            ScrollBars      =   2
         End
         Begin VB.Label lblFechaCierre 
            AutoSize        =   -1  'True
            Caption         =   "lblFechaCierre"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   3000
            TabIndex        =   12
            Top             =   360
            Width           =   1320
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Se realizará el cierre del día:"
            Height          =   240
            Left            =   240
            TabIndex        =   11
            Top             =   360
            Width           =   2550
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "&Procesos realizados"
            Height          =   240
            Left            =   240
            TabIndex        =   10
            Top             =   720
            Width           =   1860
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Nueva fecha del sistema "
         Height          =   1695
         Left            =   120
         TabIndex        =   3
         Top             =   3960
         Width           =   9015
         Begin VB.CommandButton botImprimir 
            Height          =   375
            Left            =   7560
            Picture         =   "frmCierreDiario.frx":1928
            Style           =   1  'Graphical
            TabIndex        =   4
            Tag             =   "Imprimir"
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "La fecha del sistema ahora es:"
            Height          =   240
            Left            =   240
            TabIndex        =   6
            Top             =   600
            Visible         =   0   'False
            Width           =   2730
         End
         Begin VB.Label lblNuevaFecha 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "lblNuevaFecha"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   3840
            TabIndex        =   5
            Top             =   600
            Visible         =   0   'False
            Width           =   1395
         End
      End
   End
   Begin VB.CommandButton botSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   8040
      TabIndex        =   0
      Top             =   6120
      Width           =   1275
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Bindings        =   "frmCierreDiario.frx":226A
      Left            =   1560
      Top             =   6000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      BoundReportHeading=   "rptCierreErrores"
   End
   Begin Hotel_Nana.gaHOTELbarra gaHOTELbarra1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   6495
      Width           =   9510
      _ExtentX        =   16775
      _ExtentY        =   582
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   360
      Top             =   5880
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
            Picture         =   "frmCierreDiario.frx":227A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCierreDiario.frx":25CC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFormulario 
      Caption         =   "&Formulario"
      Begin VB.Menu mnuFormularioProcesar 
         Caption         =   "Procesar"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuFormularioSalir 
         Caption         =   "Salir"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnudiv1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFormularioImprimir 
         Caption         =   "Imprimir nuevo cierre"
         Shortcut        =   ^I
      End
   End
   Begin VB.Menu mnuVer 
      Caption         =   "&Ver información de ..."
      Begin VB.Menu mnuVerRealizarCierre 
         Caption         =   "Realizar cierre"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuVerConsultarCierre 
         Caption         =   "Consultar cierres anteriores"
         Shortcut        =   {F6}
      End
   End
End
Attribute VB_Name = "frmCierreDiario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim m_ultimo_cierre As Date

'NOTA: es necesario modificar la interfaz del formulario para que
'muestre el listado de población flotante.

Private Sub botImprimir_Click()
    'Imprimo información
    muestro_listado m_ultimo_cierre
End Sub

Private Sub Form_Load()
    'inicializo control data
    subInicializoControlData Me.Data1
    
    'Apariencia formulario
    mSubConfiguroFuentesControlesSistema Me
    
    'configuro estado inicial del formulario
    subAparienciaIncialFormulario
    
    'El ultimo cierre realizado corresponde a la fecha actual del programa, antes de
    'ejecutado este proceso.
    'Comparada con la fecha real, devería ser un día menor.
    m_ultimo_cierre = m_FechaSis
    
    configuracion_apariencia
    
    Me.lblFechaCierre = Format(m_ultimo_cierre, "dddd dd mmmm yyyy")
End Sub

Private Sub subAparienciaIncialFormulario()
    'Configura la apariencia de los controles del formulario, al iniciar el proceso
    
    'configuro menú
    Me.mnuFormularioImprimir.Enabled = False
    Me.botImprimir.Enabled = False
    'configuro grilla de tareas
    gProcesos.FormatString = "      |                                                       "
    'configuro ancho de celdas
    gProcesos.ColWidth(0) = 500
    gProcesos.ColWidth(1) = gProcesos.Width - gProcesos.ColWidth(0)
End Sub

Private Sub subAparienciaTrabajoFormulario()
    'Congiguro la apariencia de los controles del formulario, después de
    'procesar.
    botProcesar.Enabled = False
    Me.mnuFormularioProcesar.Enabled = False
End Sub

Private Sub botProcesar_Click()
    If confirma Then
        'configuro apariencia de trabajo
        subAparienciaTrabajoFormulario
    
        'determino reservas noShow
        proceso_reservas_noshow
        'determino habitaciones ocupadas fuera del período
        proceso_checkout
        'cargo gastos de alojamientos a todas las habitaciones ocupadas
        '(dentro del período especificado)
        proceso_alojamientos
        'cambio situación a sucia de
        'las habitaciones ocupadas '(dentro del período especificado)
        subCambioSituacion
        'obtengo listado de población flotante
        subPoblacionFlotante
        
        'actualizo fecha de último cierre
        tbPARAMETROS.Edit
            tbPARAMETROS("fecha_ultimo_cierre_realizado") = tbPARAMETROS("fecha_ultimo_cierre_realizado") + 1
        tbPARAMETROS.Update
        
        'grabo bitácora
        'Tengo que grabar la bitácora antes de actualizar la fecha del sistema
        GraboBitacora "Cierre día " & m_FechaSis
        
        'muestro información en grilla de procesos
        subMuestroLineaGrillaProceso 9
        'Actualizo nueva fecha del sistema
        mSub_Cargo_Fecha_Sistema
        'Como la fecha en la barra de estado se actualiza cada vez que se crea una nueva
        'instancia del control es necesario llamar a este método del control para
        'que la fecha se actualize.
        Me.gaHOTELbarra1.InicializoFecha
        'muestro información en grilla de procesos
        subMuestroLineaGrillaProceso 10
    
        subMuestroFormularioAlFinalizarProceso
        
        'aviso de finalización de cierre diario
        mSubMensaje 4, 45
    End If
End Sub

Private Sub subPoblacionFlotante()
    'Luego de finalizado el día de trabajo se está en condiciones de generar
    'la información del listado de población flotante.
    'Este procedimiento obtiene esta información y la graba en el archivo de población
    'flotante.
    
    'obtengo ingresos
    subObtengoIngresos m_ultimo_cierre
    'obtengo egresos
    subObtengoEgresos m_ultimo_cierre
End Sub

Private Sub subObtengoIngresos(fechaing As Date)
    'Recorro el archivo CheckIn y obtengo los registros (personas) que ingresaron el día
    'que estoy cerrando.
    Dim tablaIngresos As Recordset
    Set tablaIngresos = tbCHECKIN
    
    tablaIngresos.Index = "i_fechaIng"
    tablaIngresos.Seek "=", fechaing
    If Not tablaIngresos.NoMatch Then
        Do While Not tablaIngresos.EOF
            'recorro todos los pasajeros que ingresaron el día a cerrar
            If tablaIngresos("fIngHab") = fechaing Then
                subGraboInfListadoPoblacionFlotante fechaing, _
                                                    tablaIngresos("nroHab"), _
                                                    tablaIngresos("nroCorrCli"), _
                                                    0, _
                                                    tablaIngresos("nroreserva"), _
                                                    tablaIngresos("fCheckDes"), _
                                                    tablaIngresos("fCheckHas"), _
                                                    tablaIngresos("HoraIngHab"), _
                                                    tablaIngresos("fIngHab")

                tablaIngresos.MoveNext
            Else
                Exit Do
            End If
        Loop
    End If
    
    Set tablaIngresos = Nothing
End Sub

Private Sub subObtengoEgresos(fechaegr As Date)
    'Recorro el archivo CheckOut y obtengo los registros (personas) que egresaron el día
    'que estoy cerrando.
    Dim tablaEgresos As Recordset
    Set tablaEgresos = tbCHECKOUT
    
    tablaEgresos.Index = "i_fegr"
    tablaEgresos.Seek "=", fechaegr
    If Not tablaEgresos.NoMatch Then
        Do While Not tablaEgresos.EOF
            'recorro todos los pasajeros que ingresaron el día a cerrar
            If tablaEgresos("fEgrHab") = fechaegr Then
                subGraboInfListadoPoblacionFlotante fechaegr, _
                                                    tablaEgresos("nroHab"), _
                                                    tablaEgresos("nroCorrCli"), _
                                                    1, _
                                                    , _
                                                    , _
                                                    , _
                                                    , _
                                                    , _
                                                    tablaEgresos("nroreserva"), _
                                                    tablaEgresos("fDes"), _
                                                    tablaEgresos("fHas"), _
                                                    tablaEgresos("HoraIngHab"), _
                                                    tablaEgresos("FIngHab"), _
                                                    tablaEgresos("HoraEgrHab"), _
                                                    tablaEgresos("FegrHab")
                tablaEgresos.MoveNext
            Else
                Exit Do
            End If
        Loop
    End If
    
    Set tablaEgresos = Nothing
End Sub

Private Sub subGraboInfListadoPoblacionFlotante(fechaListado As Date, _
                                                nrohab As Long, _
                                                nroCorrCli As Long, _
                                                tipoReg As Byte, _
                                                Optional IngNroReserva As Long, _
                                                Optional IngFCheckDes As Date, _
                                                Optional IngFcheckHas As Date, _
                                                Optional IngHoraIngHab As Date, _
                                                Optional IngFIngHab As Date, _
                                                Optional EgrNroReserva As Long, _
                                                Optional EgrFDes As Date, _
                                                Optional EgrFHas As Date, _
                                                Optional EgrHoraIngHab As Date, _
                                                Optional EgrFIngHab As Date, _
                                                Optional EgrHoraEgrHab As Date, _
                                                Optional EgrFEgrHab As Date)
                                                
    '--------------------------------------------------------------------------------------
    'Parámetros.
    '   [tipoReg]:  0 = grabo ingresos
    '               1 = grabo egresos
    '
    '---------------------------------------------------------------------------------------
    Dim proximaLinea As Integer
    Dim tablaListado As Recordset
    
    Dim fechanac As String
    
    Set tablaListado = tbPOBLACION_FLOTANTE
    'obtengo proxima linea
    proximaLinea = mFunObtengoCorrListadoPoblacionFlotante(fechaListado)
    'grabo registro
    tablaListado.AddNew
        tablaListado("fechaListado") = fechaListado
        tablaListado("nroLineaListado") = proximaLinea
        tablaListado("nroHab") = nrohab
        tablaListado("nroCorrCli") = nroCorrCli
        tablaListado("nombreCompletoCli") = obtengo_nombre_pasajero(nroCorrCli)
        'obtengo fecha de nacimiento
        fechanac = mfunObtengoDatosCli(1, nroCorrCli)
        If IsDate(fechanac) Then tablaListado("fechanac") = fechanac
        tablaListado("documento") = mfunObtengoDatosCli(2, nroCorrCli)
        tablaListado("estadocivil") = mfunObtengoDatosCli(3, nroCorrCli)
        tablaListado("nacionalidad") = mfunObtengoDatosCli(4, nroCorrCli)
        tablaListado("tipoDocu") = mfunObtengoDatosCli(5, nroCorrCli)
        
        tablaListado("tipoReg") = tipoReg
        tablaListado("IngNroReserva") = IngNroReserva
        tablaListado("IngFCheckDes") = IngFCheckDes
        tablaListado("IngFcheckHas") = IngFcheckHas
        tablaListado("IngHoraIngHab") = IngHoraIngHab
        tablaListado("IngFIngHab") = IngFIngHab
        tablaListado("EgrNroReserva") = EgrNroReserva
        tablaListado("EgrFDes") = EgrFDes
        tablaListado("EgrFHas") = EgrFHas
        tablaListado("EgrHoraIngHab") = EgrHoraIngHab
        tablaListado("EgrFIngHab") = EgrFIngHab
        tablaListado("EgrHoraEgrHab") = EgrHoraEgrHab
        tablaListado("EgrFEgrHab") = EgrFEgrHab
    tablaListado.Update
    
    Set tablaListado = Nothing
End Sub

Private Sub proceso_reservas_noshow()
    'Recorro todas las habitaciones, que pertenescan a las reservas que ingresan al
    'hotel, el día del cierre, si no ingresaron estan noshow.
    'Una reserva puede estar compuesta de varias habitaciones, por ese motivo,
    'es necesario recorrer este archivo y no tbRESERVAS, ya que incluso, dentro de una
    'reserva puede haber habitaciones que ingresaron y otras que quedan noshow.
    
    Dim existenReservasNoShow As Boolean      'determina si existen resevas noShow
                                        
    existenReservasNoShow = False             'por defecto asumo que no hay reservas noShow
    
    'muestro información en grilla de procesos
    subMuestroLineaGrillaProceso 1
        
    tbHAB_RESERVAS.Index = "ihab_reserva_fechai"
    tbHAB_RESERVAS.Seek "=", m_ultimo_cierre
    If Not tbHAB_RESERVAS.NoMatch Then
        Do While Not tbHAB_RESERVAS.EOF
            If tbHAB_RESERVAS("fechaing") = m_ultimo_cierre Then
                'busco si está alojada
                If Not busco_habita_checkin(tbHAB_RESERVAS("nrohabitacion")) Then
                    'si no esta alojada es una reserva no show
                    subRegistroReservaNoShow False
                    'indico que la habitación de la reserva quedó no show
                    tbHAB_RESERVAS.Edit
                        tbHAB_RESERVAS("noshow") = 1
                        'modifico el texto de que se muestra en la columna estado
                        'de la grilla de habitaciones en el formulario de reservas,
                        'con la idea de que si se consulta la reserva se sepa que fue no show.
                        tbHAB_RESERVAS("descri_Estado") = "       NoShow"
                    tbHAB_RESERVAS.Update
                    existenReservasNoShow = True    'existe por lo menos 1 reserva noShow
                End If
            Else
                Exit Do
            End If
            tbHAB_RESERVAS.MoveNext
        Loop
    End If
    'verifico si existen reservas noShow
    If Not existenReservasNoShow Then
        'grabo información en cierre diario
        subRegistroReservaNoShow True
    End If
    'muestro información en grilla de procesos
    subMuestroLineaGrillaProceso 2
End Sub

Private Sub proceso_checkout()
    'Recorro todos los pasajeros que se deberían de haber dejado el hotel y todavía
    'estan alojados.
    Dim existenAlojamientosOut As Boolean   'determina si existen alojamientos fuera del
                                            'período establecido
    
    existenAlojamientosOut = False          'por defecto asumo que no existen
    
    'muestro información en grilla de procesos
    subMuestroLineaGrillaProceso 3
    tbCHECKIN.Index = "i_habitacion"
    If tbCHECKIN.RecordCount > 0 Then
        tbCHECKIN.MoveFirst
        If Not tbCHECKIN.NoMatch Then
            Do While Not tbCHECKIN.EOF
                If tbCHECKIN("fcheckhas") <= m_ultimo_cierre Then
                    subRegistroCheckin False
                    existenAlojamientosOut = True 'existe por lo menos 1 alojamiento fuera
                                                  'del período establecido.
                End If
                tbCHECKIN.MoveNext
            Loop
        End If
    End If
    'determino si existen alojamientos fuera de período
    If Not existenAlojamientosOut Then
        subRegistroCheckin True
    End If
    'muestro información en grilla de procesos
    subMuestroLineaGrillaProceso 4
End Sub

Private Sub subRegistroReservaNoShow(lineaVacia As Boolean)
    'Información para listado de cierre diario.
    'Grabo cada habitación de las reservas que quedaron noshow.
    Dim desc As String
    Dim prox_corr As Integer
    prox_corr = obtengo_ultimo_corr_cierre(m_ultimo_cierre)
    tbCIERRE_DIARIO.AddNew
        tbCIERRE_DIARIO("fecha_cierre") = m_ultimo_cierre
        tbCIERRE_DIARIO("nrocorr_cierre") = prox_corr
        tbCIERRE_DIARIO("tipo_detalle_cierre") = 1   'reserva
        If lineaVacia Then
            'no existe ninguna reserva noShow
            tbCIERRE_DIARIO("desc_detalle_cierre") = "No se produjeron reservas noShow."
            tbCIERRE_DIARIO("hab_cierre") = 0
        Else
            'grabo información de la reserva noShow detectada
            desc = obtengo_desc_reserva
            tbCIERRE_DIARIO("desc_detalle_cierre") = desc
            tbCIERRE_DIARIO("hab_cierre") = tbHAB_RESERVAS("nrohabitacion")
        End If
    tbCIERRE_DIARIO.Update
End Sub

Private Function obtengo_desc_reserva()
    'Devuleve una cadena de caracteres con información respectiva a las habitaciones
    'reservadas que quedaron noshow.

    obtengo_desc_reserva = _
        tbHAB_RESERVAS("nroreserva") & " " & _
        mFunNombreTitularReserva(tbHAB_RESERVAS("nroreserva")) & " " & _
        tbHAB_RESERVAS("fechaing") & " " & _
        tbHAB_RESERVAS("fechaegr") & " " & _
        tbHAB_RESERVAS("nrohabitacion") & " " & _
        tbHAB_RESERVAS("nomtipohabitacion") & " " & _
        tbHAB_RESERVAS("tarifa")
End Function

Private Sub subRegistroCheckin(lineaVacia As Boolean)
    'Información para listado de cierre diario.
    'Grabo los pasajeros de las habitaciones que deberían haber ido
    'y no lo hicieron
    Dim prox_corr As Integer
    prox_corr = obtengo_ultimo_corr_cierre(m_ultimo_cierre)
    tbCIERRE_DIARIO.AddNew
        tbCIERRE_DIARIO("fecha_cierre") = m_ultimo_cierre
        tbCIERRE_DIARIO("nrocorr_cierre") = prox_corr
        tbCIERRE_DIARIO("tipo_detalle_cierre") = 2   'checkin
        If lineaVacia Then
            'no existen alojamientos fuera del período
            tbCIERRE_DIARIO("desc_detalle_cierre") = "No existen alojamientos fuera del período establecido."
            tbCIERRE_DIARIO("hab_cierre") = 0
        Else
            'grabo información del alojamiento fuera del período
            tbCIERRE_DIARIO("desc_detalle_cierre") = obtengo_desc_checkin
            tbCIERRE_DIARIO("hab_cierre") = tbCHECKIN("nrohab")
        End If
    tbCIERRE_DIARIO.Update
End Sub

Private Function obtengo_desc_checkin()
    'Devuleve información resepcto a los pasajeros de las habitacione que deverían de haber
    'dejaron el hotel y ahun no lo han hecho.
    obtengo_desc_checkin = _
        tbCHECKIN("nrohab") & " " & _
        obtengo_nombre_pasajero(tbCHECKIN("nrocorrcli")) & " " & _
        tbCHECKIN("fcheckdes") & " " & _
        tbCHECKIN("fcheckhas")
End Function

Private Sub proceso_alojamientos()
    'Cargo de forma automático los gastos de alojamientos de las habitaciones
    'que estan alojadas en el hotel, a la cuenta del titular de la habitación.
    'No se toman en cuanta aquellas que deberían de haber dejado el mismo y no lo hicieron.
    
    Dim existenAlojamientos As Boolean  'determino si se cargaron los gastos a por lo menos 1
                                        'habitación.
    
    existenAlojamientos = False         'por defecto asumo que no hay habitaciones ocupadas
    
    'Recorro todas las habitaciones del hotel y verifico si está alojada.
    'muestro información en grilla de procesos
    subMuestroLineaGrillaProceso 5
    tbHABITACIONES.Index = "inrohab"
    tbHABITACIONES.MoveFirst
    If Not tbHABITACIONES.NoMatch Then  'existe
        Do While Not tbHABITACIONES.EOF
            'busco si esta hospedada
            If busco_habita_checkin(tbHABITACIONES("nrohab")) Then
                'si esta hospedada, determino que este habilitada.
                'una habitación esta habilitada si esta dentro del rango de alojamiento
                'pactado, es decir si esta cumplendo con su alojamiento dentro de las fechas desde y hasta.
                'si tbcheckin(fcheckhas") <= ya devería de haber quedado libre.
                If tbCHECKIN("fcheckhas") > m_ultimo_cierre Then
                    grabo_costos_alojamiento
                    subRegistroCostosAlojamiento False
                    existenAlojamientos = True  'existe por lo menos 1 habitación ocupada
                End If
            End If
            tbHABITACIONES.MoveNext
        Loop
    End If
    'determino si se cargó por lo menos 1 gasto de alojamiento
    If Not existenAlojamientos Then
        'es necesario para mejorar la presentación del reporte y la información
        'histórica del cierre diario, crear un registro indicando que no existen
        'alojaminetos en esta fecha. Si no se indica esto, no aparece nada en el
        'listado de cierre diario.
        subRegistroCostosAlojamiento True
    End If
    'muestro información en grilla de procesos
    subMuestroLineaGrillaProceso 6
End Sub

Private Sub subCambioSituacion()
    'Cambia la situación de las habitaciones ocupadas a sucias.
    
    'Se supone que en los hoteles el servicio a las habitaciones es diario
    'y la forma de saber que habitaciones estan sucias es despues del cierre diario,
    'despues las mucamas o la gobernanta se encarga de ver que habitaciones han sido limpiadas
    'para cambiar la situación a limpa.
    subMuestroLineaGrillaProceso 7
    tbHABITACIONES.Index = "inrohab"
    tbHABITACIONES.MoveFirst
    If Not tbHABITACIONES.NoMatch Then  'existe
        Do While Not tbHABITACIONES.EOF
            'busco si esta hospedada
            If busco_habita_checkin(tbHABITACIONES("nrohab")) Then
                'si esta hospedada, determino que este habilitada.
                'una habitación esta habilitada si esta dentro del rango de alojamiento
                'pactado, es decir si esta cumplendo con su alojamiento dentro de las fechas desde y hasta.
                'si tbcheckin(fcheckhas") <= ya devería de haber quedado libre.
                If tbCHECKIN("fcheckhas") > m_ultimo_cierre Then
                    'cambio situación
                    tbHABITACIONES.Edit
                        tbHABITACIONES("situacionhab") = 2  'sucia
                        tbHABITACIONES("fechasituacionhab") = m_FechaSis
                    tbHABITACIONES.Update
                End If
            End If
            tbHABITACIONES.MoveNext
        Loop
    End If
    'muestro información en grilla de procesos
    subMuestroLineaGrillaProceso 8
End Sub

Private Sub grabo_costos_alojamiento()
    'Grabo los gastos de alojamineto a la cuanta
    'del titular de alojamiento correspondiente a la habitación que estoy procesando.
    Dim prox_gasto As Integer
    
    prox_gasto = obtengo_ultimo_corr_aloja(m_ultimo_cierre)
    tbCUENTAS_ALOJA.AddNew
        tbCUENTAS_ALOJA("fecha") = m_ultimo_cierre
        tbCUENTAS_ALOJA("nrocorr_cuenta_aloja") = prox_gasto
        
        tbCUENTAS_ALOJA("habitacion_cuenta_aloja") = tbHABITACIONES("nrohab")
        tbCUENTAS_ALOJA("titular_aloja") = _
                        busco_titular_hab2SinCambiarPunteroHab(tbHABITACIONES("nrohab"), "aloja")
        
        tbCUENTAS_ALOJA("tarifa") = tbHABITACIONES("tarifa")
        'alojamiento automático
        'este valor coincide con el establecido en el archivo de constantes (SISTEMA_CONSTANTES)
        'Indica que el alojamiento fue cargado de forma automática.
        'Dichos alojamientos se agrupan en una sola línea (total) conjuntamente
        'con los de tipo = 4,corrección de tarifa al momento de facturar.
        tbCUENTAS_ALOJA("tipoAloja") = 1
        'como observación muestro el día que se está cobrando.
        tbCUENTAS_ALOJA("obsAloja") = Format(m_ultimo_cierre, "dddd, d mmm")
    tbCUENTAS_ALOJA.Update
End Sub

Private Sub subRegistroCostosAlojamiento(lineaVacia As Boolean)
    'Información para listado de cierre diario.
    'Grabo las habitaciones a las cuales se le cargaron gastos de alojamiento.
    Dim prox_corr As Integer
    prox_corr = obtengo_ultimo_corr_cierre(m_ultimo_cierre)
    tbCIERRE_DIARIO.AddNew
        tbCIERRE_DIARIO("fecha_cierre") = m_ultimo_cierre
        tbCIERRE_DIARIO("nrocorr_cierre") = prox_corr
        tbCIERRE_DIARIO("tipo_detalle_cierre") = 3   'gastos alojamiento
        If lineaVacia Then
            'este caso ocurre cuando no existen habitaciones alojadas
            'siendo necesareio crear una línea que indique justamente esto.
            tbCIERRE_DIARIO("desc_detalle_cierre") = "No se cargaron gastos. No existen habitaciones ocupadas."
            tbCIERRE_DIARIO("hab_cierre") = 0
        Else
            tbCIERRE_DIARIO("desc_detalle_cierre") = obtengo_desc_gastos_aloja
            tbCIERRE_DIARIO("hab_cierre") = tbHABITACIONES("nrohab")
        End If
    tbCIERRE_DIARIO.Update
End Sub

Private Function obtengo_desc_gastos_aloja()
    'Devuelve información con respecto a los gastos cargados a las habitaciones alojadas.
    obtengo_desc_gastos_aloja = _
        tbHABITACIONES("nrohab") & " " & _
        busco_tipo_hab_descri(tbHABITACIONES("nrohab")) & " " & _
        tbHABITACIONES("tarifa")
End Function

Private Function confirma()
    confirma = False
    'aviso de confirmación de operación
    If mFunMensaje(4, 44) Then
        confirma = True
    End If
End Function

Private Sub muestro_listado(fechaCierre As Date)
    'esto es una cagada de primera tiene que haber otra manera más confortable de hacer
    'esta porquería en cristal
    Me.Data1.RecordSource = _
        "select * from cierre_diario where fecha_cierre = " & fechaSQL(Str(fechaCierre))
    Me.Data1.Refresh
    Me.CrystalReport1.ReportFileName = vardir2 + "rptcierreerrores.rpt"
    Me.CrystalReport1.Formulas(2) = fechaCierre
    Me.CrystalReport1.WindowState = crptMaximized
    Me.CrystalReport1.Action = 1
End Sub

Private Sub subMuestroFormularioAlFinalizarProceso()
    Me.lblNuevaFecha = Format(tbPARAMETROS("fecha_ultimo_cierre_realizado"), "dddd dd mmmm yyyy")
    Me.lblNuevaFecha.Visible = True
    Label4.Visible = True
    'permito imprimir formulario
    Me.botImprimir.Enabled = True
    Me.mnuFormularioImprimir.Enabled = True
    'muestro información en grilla de procesos
    subMuestroLineaGrillaProceso 11
    'oculto barra de proceso
    Me.gaHOTELbarra1.ProgresoFin
End Sub

Private Sub botConsultarCierreDiario_Click()
    'valido que la fecha sea válida
    If IsDate(Me.fechaCierreConsulta.Value) Then
        'busco cierre diario de la fecha ingresada
        tbCIERRE_DIARIO.Index = "pk_cierre"
        tbCIERRE_DIARIO.Seek ">=", Me.fechaCierreConsulta.Value, 0
        If Not tbCIERRE_DIARIO.NoMatch Then
            If tbCIERRE_DIARIO("fecha_cierre") = Me.fechaCierreConsulta.Value Then
                'listo cierre diario
                muestro_listado Me.fechaCierreConsulta.Value
            Else
                'aviso de que no existe cierre diario
                mSubMensaje 4, 127
            End If
        End If
    Else
        'aviso de fecha incorrecta
        mSubMensaje 3, 1
        Me.fechaCierreConsulta.SetFocus
    End If
End Sub

Private Sub botCerrar_Click()
    Unload Me
End Sub

Private Sub configuracion_apariencia()
    'Determina la apariencia del los elemento configurables del formulario
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Descargo formulario de memoria
    Set frmCierreDiario = Nothing
End Sub

Private Sub mnuFormularioImprimir_Click()
    'Equivale a presionar el boton de imprimir o la tecla Ctrol + I
    botImprimir_Click
End Sub

Private Sub mnuFormularioProcesar_Click()
    'Equivale a presionar el boton de procesar o la tecla F9
    botProcesar_Click
End Sub

Private Sub mnuFormularioSalir_Click()
    'Equivale a presionar el boton de salir o la tecla F12
    botSalir_Click
End Sub

Private Sub mnuVerConsultarCierre_Click()
    'Equivale a presionar la tecla F6
    Me.ssTab1.Tab = 1
End Sub

Private Sub mnuVerRealizarCierre_Click()
    'Equivale a presionar la tecla F5
    Me.ssTab1.Tab = 0
End Sub

Private Sub botSalir_Click()
    Unload Me
End Sub

Private Sub subMuestroLineaGrillaProceso(linea As Byte)
    'Muestro una línea determinada en la grilla de procesos
    'Cada línea tiene asociado un ícono específico
    Dim descLinea As String
    Dim iconoLinea As Byte
    Dim posLinea As Byte
    Select Case linea
        Case 1
            descLinea = "Buscando reservas noShow."
            iconoLinea = 1
            posLinea = 0
        Case 2
            descLinea = "Se identificaron las reservas no show."
            iconoLinea = 2
            posLinea = 0
        Case 3
            descLinea = "Buscando pasajeros que aún no dejaron el hotel."
            iconoLinea = 1
            posLinea = 1
        Case 4
            descLinea = "Se identificaron los pasajeros que no dejaron el hotel."
            iconoLinea = 2
            posLinea = 1
        Case 5
            descLinea = "Cargando gastos de alojamiento."
            iconoLinea = 1
            posLinea = 2
        Case 6
            descLinea = "Se cargaron gastos de alojamiento de las habitaciones ocupadas."
            iconoLinea = 2
            posLinea = 2
        Case 7
            descLinea = "Estableciendo situación de las habitaciones ocupadas a sucia."
            iconoLinea = 1
            posLinea = 3
        Case 8
            descLinea = "Se establecío a sucia la situación de las habitaciones ocupadas."
            iconoLinea = 2
            posLinea = 3
        Case 9
            descLinea = "Actualizando fecha del sistema."
            iconoLinea = 1
            posLinea = 4
        Case 10
            descLinea = "Se actualizó la fecha del sistema."
            iconoLinea = 2
            posLinea = 4
        Case 11
            descLinea = "El cierre diario finalizó de forma correcta."
            iconoLinea = 2
            posLinea = 5
    End Select
    'Creo línea en grilla
    'me posiciono en la fila correspondiente
    gProcesos.row = posLinea
    'muestro descripción
    gProcesos.TextMatrix(gProcesos.row, 1) = descLinea
    'muestro íconos
    Set gProcesos.CellPicture = ImageList1.ListImages(iconoLinea).Picture
    'muetro barra de progreso
    Me.gaHOTELbarra1.Progreso 0, 11, CLng(linea)
End Sub

'******************************************************************************
'*
'*  Asistencia al usuario
'*
'******************************************************************************

Private Sub botProcesar_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 18
End Sub

Private Sub botImprimir_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 19
End Sub

Private Sub botSalir_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 1, 3
End Sub

Private Sub fechaCierreConsulta_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 205
End Sub

Private Sub botConsultarCierreDiario_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 206
End Sub

Private Sub botConsultarCierreDiario_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub fechaCierreConsulta_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub botProcesar_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub botImprimir_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub botSalir_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

