VERSION 5.00
Object = "{EB8C3860-8603-11D0-A35D-7062E4000000}#1.1#0"; "VCBND.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmListadoIngresos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Listado de ingresos"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   11910
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin Hotel_Nana.gaHOTELbarra gaHOTELbarra1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   7605
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   582
   End
   Begin VB.Frame Frame1 
      Height          =   7455
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   11655
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   1080
         Width           =   1815
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "&Ingresos"
            Height          =   240
            Left            =   0
            TabIndex        =   7
            Top             =   0
            Width           =   780
         End
      End
      Begin MSFlexGridLib.MSFlexGrid dbIE 
         Bindings        =   "frmListadoIngresos.frx":0000
         Height          =   5415
         Left            =   360
         TabIndex        =   8
         Top             =   1320
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   9551
         _Version        =   393216
         FocusRect       =   2
         HighLight       =   0
      End
      Begin VB.ComboBox cboTipo_habitacion 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7200
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   480
         Width           =   2415
      End
      Begin VcBndCtl.VcCalCombo Fcons 
         Height          =   375
         Left            =   5760
         TabIndex        =   1
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _0              =   $"frmListadoIngresos.frx":0010
         _1              =   $"frmListadoIngresos.frx":0419
         _2              =   $"frmListadoIngresos.frx":0822
         _3              =   "-@@E@@@A@@@@@@@@@'@@@E@@@A@@@@@@@@@)h@@@E@@@A@@@@@@@@@,4E7F"
         _count          =   4
         _ver            =   2
      End
      Begin VB.CommandButton botImprimir 
         Height          =   375
         Left            =   8880
         Picture         =   "frmListadoIngresos.frx":0C2B
         Style           =   1  'Graphical
         TabIndex        =   5
         Tag             =   "Imprimir"
         Top             =   6960
         Width           =   1215
      End
      Begin VB.CommandButton botSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   375
         Left            =   10200
         TabIndex        =   6
         Top             =   6960
         Width           =   1215
      End
      Begin VB.CommandButton botProcesar 
         Caption         =   "&Procesar"
         Height          =   375
         Left            =   10200
         TabIndex        =   4
         Tag             =   "Procesar"
         Top             =   480
         Width           =   1215
      End
      Begin ComctlLib.TabStrip TabStrip1 
         Height          =   6135
         Left            =   240
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   720
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   10821
         _Version        =   327682
         BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
            NumTabs         =   3
            BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Todos"
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Estan por ingresar"
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Ya ingresados"
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "F&echa:"
         Height          =   240
         Left            =   5760
         TabIndex        =   0
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "&Tipo habitación:"
         Height          =   240
         Left            =   7200
         TabIndex        =   2
         Top             =   240
         Width           =   1440
      End
   End
   Begin VB.Menu mnuFormulario 
      Caption         =   "&Formulario"
      Begin VB.Menu mnuFormularioAceptar 
         Caption         =   "Aceptar"
         Shortcut        =   {F12}
      End
      Begin VB.Menu div1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFormularioProcesar 
         Caption         =   "Procesar"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuImprimirConsulta 
         Caption         =   "Imprimir"
         Shortcut        =   ^I
      End
   End
   Begin VB.Menu mnuSeleccion 
      Caption         =   "&Ver información de..."
      Begin VB.Menu mnuSeleccionTodos 
         Caption         =   "Todos"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuSeleccionEstanPorIngresar 
         Caption         =   "Estan por ingresar"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuSeleccionYaIngresados 
         Caption         =   "Ya ingresados"
         Shortcut        =   {F7}
      End
   End
   Begin VB.Menu mnuOrden 
      Caption         =   "&Ordenado por ..."
      Begin VB.Menu mnuOrdenReserva 
         Caption         =   "Número de reserva"
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu mnuOrdenTitular 
         Caption         =   "Titular"
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu mnuOrdenNroHab 
         Caption         =   "Número de habitación"
         Shortcut        =   ^{F3}
      End
      Begin VB.Menu mnuOrdenTipoHab 
         Caption         =   "Tipo de habitación"
         Shortcut        =   ^{F4}
      End
      Begin VB.Menu mnuOrdenCantidad 
         Caption         =   "Cantidad de pasajeros"
         Shortcut        =   ^{F5}
      End
      Begin VB.Menu mnuOrdenHoraIng 
         Caption         =   "Hora de ingreso"
         Shortcut        =   ^{F6}
      End
      Begin VB.Menu mnuOrdenAgencia 
         Caption         =   "Agencia o empresa"
         Shortcut        =   ^{F7}
      End
   End
   Begin VB.Menu mnuIr 
      Caption         =   "Ir a .."
      Begin VB.Menu mnuIrCheckin 
         Caption         =   "Realizar Checkin"
         Shortcut        =   ^C
      End
   End
End
Attribute VB_Name = "frmListadoIngresos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rst_principal As Recordset
Dim qdf As QueryDef

Dim tipoCriteroOrdenListadoImp As Byte    'tipo de criterio por el cual se estan ordenando los datos
                                          'en la grilla; se utiza para ordenar el listado impreso.

Private marco_color_seleccionadas As Boolean

Private Sub Form_Load()
    'Apariencia formulario
    mSubConfiguroFuentesControlesSistema Me
    
    'cargo tipo habitación
    carga_tipo_hab Me.cboTipo_habitacion
    cboTipo_habitacion.AddItem ("(Todas)")
    cboTipo_habitacion.Text = "(Todas)"
    
    'inicializo fecha de consulta
    Fcons.Value = m_FechaSis
    
    mnuSeleccionTodos_Click       'por defecto selecciono todos
End Sub

Private Sub botProcesar_Click()
    If valido_fecha Then
       subProcesoIngresos
       subMuestroIngresos
       mnuOrdenReserva_Click
    End If
End Sub

Private Sub subProcesoIngresos()
    'Se encarga de crear un recordset con los datos a mostrar.
    Dim qdf As QueryDef
    Dim consulta As String
    Dim criterio_ordenacion As String
    
    consulta = SQLIngresosPrevistos(Fcons.Text)
                
    If cboTipo_habitacion.Text <> "(Todas)" Then
        consulta = consulta & " and hab_reserva.tipohabitacion = " & _
                    cboTipo_habitacion.ItemData(cboTipo_habitacion.ListIndex)
    End If
    
    'originalmente ordeno por número de reserva
    criterio_ordenacion = " ORDER BY reservas.nroreserva"
    
    'ejecuto consulta
    Set qdf = bdHOTEL.CreateQueryDef("")
    qdf.SQL = consulta & criterio_ordenacion
    Set rst_principal = qdf.OpenRecordset(dbOpenSnapshot)
End Sub

Private Sub subMuestroIngresos()
    'Recorro el recordset generado y muestro en la grilla
    limpio_grilla dbIE
    dibujo_grilla_ing
    
    Dim linea As String
    
    If rst_principal.RecordCount > 0 Then
        rst_principal.MoveFirst
        Do While Not rst_principal.EOF
            If cumple_seleccion Then
                linea = Chr(9) & rst_principal!NroReserva & _
                        Chr(9) & rst_principal!primer_ape_titular & _
                        Chr(9) & rst_principal!nrohabitacion & _
                        Chr(9) & rst_principal!descripcion & _
                        Chr(9) & rst_principal!pasajeros & _
                        Chr(9) & _
                        Chr(9) & mFunBuscoNombreEmpresa(rst_principal!nroAgenciaEmpresa)
                dbIE.AddItem linea
                If marco_color_seleccionadas Then
                    marco_linea_grilla
                    cargo_hora_ingreso
                End If
            End If
            rst_principal.MoveNext
        Loop
    End If
End Sub

Private Function cumple_seleccion()
    'Para cada registro del recordset obtenido,
    'valido que cumpla con la condición de selección
    
    marco_color_seleccionadas = False
    
    If Me.mnuSeleccionTodos.Checked Then 'todos
        'si la habitacion fue ocupada cargo hora de ingreso
        If busco_ReservaHabita_checkin(rst_principal!NroReserva, rst_principal!nrohabitacion) Then
            marco_color_seleccionadas = True
        End If
        cumple_seleccion = True
    End If
        
    If Me.mnuSeleccionEstanPorIngresar.Checked Then 'por INGRESAR
        'si la habitación ya fue ocupada no muestro
        If busco_ReservaHabita_checkin(rst_principal!NroReserva, rst_principal!nrohabitacion) Then
            cumple_seleccion = False
        Else
            cumple_seleccion = True
        End If
    End If
    
    If Me.mnuSeleccionYaIngresados.Checked Then  'los INGRESADOS
        'si la habitacion ya fue ingresada muestro
        If busco_ReservaHabita_checkin(rst_principal!NroReserva, rst_principal!nrohabitacion) Then
            cumple_seleccion = True
            marco_color_seleccionadas = True
        Else
            cumple_seleccion = False
        End If
    End If
End Function

Private Sub cargo_hora_ingreso()
    'La hora de ingreso no la tengo cargada en el recordset,
    'por ese motivo tengo que obtenerla del archivo checkin
    dbIE.row = dbIE.row
    dbIE.col = dbIE.Cols - 2            'la última columna es para el nombre de la agencia
    dbIE.Text = tbCHECKIN("horainghab")
End Sub

Private Sub subOrdenoIngresos(criterio As Byte)
    'Cambia el orden de la consulta dependiendo del usuario
    Select Case criterio
        Case 1  'ordenado por reserva
            dbIE.col = 1
        Case 2  'ordenado por titular
            dbIE.col = 2
        Case 3  'ordeno por número de habitación
            dbIE.col = 3
        Case 4  'ordenado por tipo habitacion
            dbIE.col = 4
        Case 5  'ordeno por cantidad de pasajeros
            dbIE.col = 5
        Case 6  'ordeno por hora de ingreso
            dbIE.col = 6
        Case 7  'ordeno por agencia
            dbIE.col = 7
    End Select
    dbIE.Sort = 5
    'ordeno el listado impreso por el mismo criterio que la grilla
    tipoCriteroOrdenListadoImp = criterio
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Descargo formulario de memoria
    Set frmListadoIngresos = Nothing
End Sub

Private Sub dbIE_DblClick()
    'Al realizar doble click puedo ingresar la reserva
    mnuIrCheckin_Click
End Sub

Private Sub mnuIrCheckin_Click()
    'Permite realizar el Checkin, de una reserva que esta por ingresar
    Dim nrores_aux As String
    HoraIni = Time
    OprEjecutada = 5
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        'obtengo número de reserva
        nrores_aux = dbIE.TextMatrix(dbIE.row, 1)
        'valido que se este posicionado sobre un número de reserva
        If (Trim(nrores_aux) <> Empty) And dbIE.row > 1 Then
            tipo_accion_reserva = "Check-in"
            Load frmModificacionReserva
            'cargo el número de reserva
            frmModificacionReserva.txtNroReservaAnio.Text = Mid(nrores_aux, 1, 4)
            frmModificacionReserva.txtNroReserva.Text = Val(Mid(nrores_aux, 5, 8))
            'muestro formulario
            frmModificacionReserva.Show 1
            'después de realizar el ingreso ejecuto nuevamente la
            'consulta para que se actualize el listado de ingreso
            mnuSeleccionTodos_Click
        Else
            'aviso: debe de seleccionar reserva
            mSubMensaje 4, 119
        End If
    End If
End Sub

Private Sub mnuOrdenReserva_Click()
    'ordeno por reserva
    subDesmarcoOrden
    mnuOrdenReserva.Checked = True
    mSubMuestroIcono dbIE, 1
    subOrdenoIngresos 1
End Sub

Private Sub mnuOrdenTitular_Click()
    'ordeno por titular
    subDesmarcoOrden
    mnuOrdenTitular.Checked = True
    mSubMuestroIcono dbIE, 2
    subOrdenoIngresos 2
End Sub

Private Sub mnuOrdenNroHab_Click()
    'ordeno por número de habitación
    subDesmarcoOrden
    mnuOrdenNroHab.Checked = True
    mSubMuestroIcono dbIE, 3
    subOrdenoIngresos 3
End Sub

Private Sub mnuOrdenTipoHab_Click()
    'ordeno por tipo de habitación
    subDesmarcoOrden
    mnuOrdenTipoHab.Checked = True
    mSubMuestroIcono dbIE, 4
    subOrdenoIngresos 4
End Sub

Private Sub mnuOrdenCantidad_Click()
    'ordeno por cantidad de pasajeros
    subDesmarcoOrden
    mnuOrdenCantidad.Checked = True
    mSubMuestroIcono dbIE, 5
    subOrdenoIngresos 5
End Sub

Private Sub mnuOrdenHoraIng_Click()
    'ordeno por hora de ingreso
    subDesmarcoOrden
    mnuOrdenHoraIng.Checked = True
    mSubMuestroIcono dbIE, 6
    subOrdenoIngresos 6
End Sub

Private Sub mnuOrdenAgencia_Click()
    'ordeno por agencia
    subDesmarcoOrden
    mnuOrdenAgencia.Checked = True
    mSubMuestroIcono dbIE, 7
    subOrdenoIngresos 7
End Sub

Private Sub mnuSeleccionEstanPorIngresar_Click()
    'selecciono solo los que estan por ingresar
    TabStrip1.Tabs(2).Selected = True
End Sub

Private Sub mnuSeleccionTodos_Click()
    'selecciono todos los ingresos
    TabStrip1.Tabs(1).Selected = True
End Sub

Private Sub mnuSeleccionYaIngresados_Click()
    'selecciono solo los ya ingresados
    TabStrip1.Tabs(3).Selected = True
End Sub

Private Sub subDesmarcoSeleccion()
    mnuSeleccionTodos.Checked = False
    mnuSeleccionYaIngresados.Checked = False
    mnuSeleccionEstanPorIngresar.Checked = False
End Sub

Private Sub subDesmarcoOrden()
    mnuOrdenTitular.Checked = False
    mnuOrdenTipoHab.Checked = False
    mnuOrdenReserva.Checked = False
    mnuOrdenNroHab.Checked = False
    mnuOrdenHoraIng.Checked = False
    mnuOrdenCantidad.Checked = False
    mnuOrdenAgencia.Checked = False
End Sub

'****************************************************************************
'* Procedimientos para dibujar grilla
'****************************************************************************

Private Sub dibujo_grilla_ing()
    Dim i As Byte
    cabezal_grilla_ing
End Sub

Private Sub cabezal_grilla_ing()
    dbIE.FormatString = " | Reserva        | Titular                                                | Hab.      | Tipo           | Pax      | Hora Ingreso    | Agencia                                "
End Sub

Private Sub marco_linea_grilla()
    marco_celdas_grilla dbIE, 1, dbIE.Cols - 1, dbIE.Rows - 1, dbIE.Rows - 1
    dbIE.CellBackColor = &HFFFF80
End Sub

Private Function valido_fecha()
    valido_fecha = True
    If Not IsDate(Fcons.Text) Then
        'no es un formato de fecha válido
        mSubMensaje 3, 1
        Fcons.SetFocus
        valido_fecha = False
        Exit Function
    End If
    
    If Fcons.Value < m_FechaSis Then
        'fecha menor a la del día de hoy
        mSubMensaje 3, 2
        Fcons.SetFocus
        valido_fecha = False
        Exit Function
    End If
End Function

Private Sub TabStrip1_Click()
    'Cuando pincho un tab efectúo la mismas operaciones que cuando
    'selecciono las opciones desde el menú, dando mayor funcionalidad al programa.
    If TabStrip1.SelectedItem.Index = 1 Then  'ejecuto opcion todas
        subDesmarcoSeleccion
        mnuSeleccionTodos.Checked = True
        'ejecuto consulta
        subProcesoIngresos
        subMuestroIngresos
        
        mnuOrdenReserva_Click         'ordeno originalmente por reserva
    End If
    
    If TabStrip1.SelectedItem.Index = 2 Then
        subDesmarcoSeleccion
        mnuSeleccionEstanPorIngresar.Checked = True
        'ejecuto consulta
        subProcesoIngresos
        subMuestroIngresos
        
        mnuOrdenReserva_Click         'ordeno originalmente por reserva
    End If
    
    If TabStrip1.SelectedItem.Index = 3 Then
        subDesmarcoSeleccion
        mnuSeleccionYaIngresados.Checked = True
        'ejecuto consulta
        subProcesoIngresos
        subMuestroIngresos
        
        mnuOrdenReserva_Click         'ordeno originalmente por reserva
    End If
End Sub

'**********************************************************
'*
'   Impresión de reporte
'*
'**********************************************************

Private Sub botImprimir_Click()
    'Imprimo información
    If mfunAplicoConfImp(2, 13) = 1 Then
        'verifico que tipo de filtro estoy aplicando
        If Me.TabStrip1.Tabs(1).Selected = True Then
            'muestro todas
            subArmoReporteIngresosPrevistos 1, Fcons.Text
        Else
            If Me.TabStrip1.Tabs(2).Selected = True Then
                'las que estan por ingresar
                subArmoReporteIngresosPrevistos 2, Fcons.Text
            Else
                If Me.TabStrip1.Tabs(3).Selected = True Then
                    'las que ya ingresaron
                    subArmoReporteIngresosPrevistos 3, Fcons.Text
                End If
            End If
        End If
    End If
End Sub

Private Sub subArmoReporteIngresosPrevistos(tipoFiltro As Byte, fechaing As Date)
    '----------------------------------------------------------------------------------
    'Configuro el control Crystal genérico,obtengo datos y emite el listado
    'ingresos previstos
    '-------------------------------------------------------------------------------
    'Parámetros.
    '   Entrada:    [tipoFiltro] determina que tipo de información muestro.
    '               1 = todos los ingresos
    '               2 = solo los que estan por ingresar
    '               3 = solo los que ya ingresaron
    '               [fechaIng] fecha de la cual se quiere saber los ingresos previstos
    '-------------------------------------------------------------------------------
    
    Dim descTipoFiltro As String
    Dim descOrden As String
    
    'inicializo control data
     subInicializoControlData frmMAIN.Data1CrystalReport
    
    'NOTA: esta consulta debe de ser la misma que se utilizo para crear el reporte,
    'por ese motivo no debe de ser modificada en aspectos tales como campos a mostrar y
    'tablas utilizadas.

    'realizo un LEFT JOIN para incluir reservas-hab que no han ingresado al hotel todavía
     frmMAIN.Data1CrystalReport.RecordSource = _
     "select * from reservas INNER JOIN " & _
     "(hab_reserva LEFT JOIN checkin ON checkin.nroreserva = hab_reserva.nroreserva and checkin.nrohab = hab_reserva.nrohabitacion )" & _
     " ON reservas.nroreserva = hab_reserva.nroreserva " & _
     " where reservas.fechaing = " & fechaSQL(fechaing)
     
     'verifico si tengo que filtrar por tipo de habitación
     If cboTipo_habitacion.Text <> "(Todas)" Then
        frmMAIN.Data1CrystalReport.RecordSource = _
        frmMAIN.Data1CrystalReport.RecordSource & " and hab_reserva.tipohabitacion = " & _
                    cboTipo_habitacion.ItemData(cboTipo_habitacion.ListIndex)
     End If
      
     Select Case tipoFiltro
        Case 1
            'todos los ingresos
            descTipoFiltro = "Todos"
        Case 2
            'se leccionan las reservas que no han ingresado
            frmMAIN.Data1CrystalReport.RecordSource = frmMAIN.Data1CrystalReport.RecordSource & _
            " and checkin.horaIngHab = null"
            descTipoFiltro = "Estan por ingresar"
        Case 3
            'selecciono las reservas que ya ingresaron, es decir que ya estan en el
            'archivo checkin
            frmMAIN.Data1CrystalReport.RecordSource = frmMAIN.Data1CrystalReport.RecordSource & _
            " and checkin.horaIngHab <> ''"
            descTipoFiltro = "Ya ingresadas"
    End Select
    'establesco orden del listado
    descOrden = funEstablescoOrdenListado
    
    'ejecuto consulta control data
    
    frmMAIN.Data1CrystalReport.Refresh
    
    'como de aquí en más voy a trabajar directamente con el control data
    'me aseguro de que se allan encontrado reservas
    If frmMAIN.Data1CrystalReport.Recordset.RecordCount > 0 Then
        'determino el nombre del reporte a utilizar
        frmMAIN.CrystalReport1.ReportFileName = vardir2 + "rptIngP.rpt"

        'inicializo fórmulas
        With frmMAIN.CrystalReport1
            .Formulas(0) = "cabVersion = '" & mFunImpVersionAplicacion & "'"        'nombre y versión aplicación
            .Formulas(1) = "cabRegistro = '" & mFunImpRegistro & "'"                'hotel propietario de la aplicación
            .Formulas(2) = "cabHoraImp = ' " & CStr(Time()) & "'"                   'hora actual
            .Formulas(3) = "cabTipoFiltro = '" & descTipoFiltro & "'"
            .Formulas(4) = "parte1Fecha = '" & Format(fechaing, "dddd dd mmmm yyyy") & "'"
            .Formulas(5) = "parte1TipoHab = '" & Me.cboTipo_habitacion.Text & "'"
            .Formulas(6) = "parte1Orden = '" & descOrden & "'"
        End With
        'genero listado
        frmMAIN.CrystalReport1.Action = 1
        'aviso de confirmación de impresión de reporte
        mSubMensaje 4, 137  'se imprimieron los ingresos previstos
        'inicializo fórmulas
        mSubInicializoFormulas 6
    Else
        'aviso de que no hay datos para imprimir
        mSubMensaje 3, 9
    End If
End Sub

Private Function funEstablescoOrdenListado() As String
    '-------------------------------------------------------------------
    'Establesco orden del listado impreso según lo establecido para la
    'grilla.
    '-------------------------------------------------------------------
    'Parámetros.
    '   Salida: devuelve un string con información acerca del criterio de
    '           ordenación utilizado, el mismo se muestra en el listado.
    'NOTA:
    'tipoCriteroOrdenListadoImp se inicializa cada vez que se cambia
    'el criterio de otrdenación de la grilla.
    '-------------------------------------------------------------------
    Dim criterio As String
    Select Case tipoCriteroOrdenListadoImp
        Case 1  'por reserva
            criterio = "reservas.nroreserva"
            funEstablescoOrdenListado = "Ordenado por número de reserva"
        Case 2  'titular
            funEstablescoOrdenListado = "Ordenado por titular"
            criterio = "reservas.primer_ape_titular,reservas.segundo_ape_titular,primer_nom_titular,segundo_nom_titular"
        Case 3  'número de habitación
            funEstablescoOrdenListado = "Ordenado por número de habitación"
            criterio = "hab_reserva.nrohabitacion"
        Case 4  'tipo habitación
            funEstablescoOrdenListado = "Ordenado por tipo de habitación"
            criterio = "hab_reserva.tipohabitacion"
        Case 5  'cantidad pax
            funEstablescoOrdenListado = "Ordenado por cantidad de pax"
            criterio = "hab_reserva.pasajeros"
        Case 6  'hora ingreso
            funEstablescoOrdenListado = "Ordenado por hora de ingreso"
            criterio = "checkin.horaIngHab"
        Case 7  'agencia
            funEstablescoOrdenListado = "Ordenado por agencia o empresa"
            criterio = "reservas.nroAgenciaEmpresa"
    End Select
    frmMAIN.Data1CrystalReport.RecordSource = frmMAIN.Data1CrystalReport.RecordSource & _
    " order by " & criterio
End Function

'******************************************************************************
'* Fin de impresión
'******************************************************************************

Private Sub botSalir_Click()
    Unload Me
End Sub

Private Sub mnuFormularioAceptar_Click()
    'Equivale a digitar F12 o el boton aceptar
    botSalir_Click
End Sub

Private Sub mnuFormularioProcesar_Click()
    'Equivale a presionar el boton procesar o F9
    botProcesar_Click
End Sub

Private Sub mnuImprimirConsulta_Click()
    'Equivale a presionar el boton de imprimir o Ctrol+I
    botImprimir_Click
End Sub

'*********************************************
'*
'*  Asistencia a usuarios
'*
'*********************************************

Private Sub Fcons_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 49
End Sub

Private Sub cboTipo_habitacion_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 47
End Sub

Private Sub botProcesar_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 28
End Sub

Private Sub botImprimir_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 50
End Sub

Private Sub botSalir_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 1, 1
End Sub

Private Sub cboTipo_habitacion_LostFocus()
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

Private Sub Fcons_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub


