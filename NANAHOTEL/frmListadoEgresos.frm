VERSION 5.00
Object = "{EB8C3860-8603-11D0-A35D-7062E4000000}#1.1#0"; "VCBND.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmListadoEgresos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Listado de egresos"
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
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   7605
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   582
   End
   Begin VB.Frame Frame1 
      Height          =   7455
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Width           =   11655
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   1080
         Width           =   975
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "&Egresos"
            Height          =   240
            Left            =   0
            TabIndex        =   12
            Top             =   0
            Width           =   765
         End
      End
      Begin VB.CommandButton botImprimir 
         Height          =   375
         Left            =   8880
         Picture         =   "frmListadoEgresos.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Tag             =   "Imprimir"
         Top             =   6960
         Width           =   1215
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
         _0              =   $"frmListadoEgresos.frx":0942
         _1              =   $"frmListadoEgresos.frx":0D4B
         _2              =   $"frmListadoEgresos.frx":1154
         _3              =   "-E@@@A@@@@@@@@@'@@@E@@@A@@@@@@@@@)h@@@E@@@A@@@@@@@@@,4E7F"
         _count          =   4
         _ver            =   2
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
      Begin MSFlexGridLib.MSFlexGrid dbIE 
         Height          =   5415
         Left            =   360
         TabIndex        =   7
         Top             =   1320
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   9551
         _Version        =   393216
         FocusRect       =   2
         HighLight       =   0
      End
      Begin ComctlLib.TabStrip TabStrip1 
         Height          =   6135
         Left            =   240
         TabIndex        =   10
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
               Caption         =   "Aun no dejaron el hotel"
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Ya egresaron"
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "F&echa:"
         Height          =   255
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
         Caption         =   "Salir"
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
      Begin VB.Menu mnuSeleccionarTodos 
         Caption         =   "Todos"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuSeleccionNoDejaron 
         Caption         =   "Aun no dejaron el hotel"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuSeleccionYaDejaron 
         Caption         =   "Ya egresados"
         Shortcut        =   {F7}
      End
   End
   Begin VB.Menu mnuOrden 
      Caption         =   "&Ordenado por ..."
      Begin VB.Menu mnuOrden1 
         Caption         =   "1criterio"
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu mnuOrden2 
         Caption         =   "2criterio"
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu mnuOrden3 
         Caption         =   "3criterio"
         Shortcut        =   ^{F3}
      End
      Begin VB.Menu mnuOrden4 
         Caption         =   "4criterio"
         Shortcut        =   ^{F4}
      End
      Begin VB.Menu mnuOrden5 
         Caption         =   "5criterio"
         Shortcut        =   ^{F5}
      End
      Begin VB.Menu mnuOrden6 
         Caption         =   "6criterio"
         Shortcut        =   ^{F6}
      End
   End
   Begin VB.Menu mnuMostrar 
      Caption         =   "&Mostrar informacion por ..."
      Visible         =   0   'False
      Begin VB.Menu mnuMostrarPasajeros 
         Caption         =   "Pasajeros"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuMostrarHab 
         Caption         =   "Habitación"
         Shortcut        =   ^H
      End
   End
End
Attribute VB_Name = "frmListadoEgresos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' 09/04/03 Mostrar por habitación
'Oculto la opción de mostrar por habitación ya que en este momento del
'ciclo de desarrollo de la aplicación, concidero que no es necesario mostrar esa
'información. De todas maneras no se altera ningúna línea de código,con la idea de
'en un futuro, si es necesario, reactivar esta opción.
'----------------------------------------------------------------------------------------

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
    
    'inicialmente muestro todos
    mnuSeleccionarTodos.Checked = True
    mnuMostrarPasajeros.Checked = True
    
    subConfiguroMenuOrden
    subEjecutar
End Sub

Private Sub botProcesar_Click()
    subEjecutar
End Sub

Private Sub subEjecutar()
    'Cada vez que cambio una opción de ver o de mostrar
    'en el menu de opciones ejecuto este procedimiento
    'como así también cada vez que se pulsa el boton procesar.
    
    'valido fecha de egreso
    If mFunValidoFecha Then
        proceso_egresos
        If mnuMostrarHab.Checked = True Then
            'muestro habitaciones
            dibujo_grilla_egr_h
            muestro_egreso_habitacion
        Else
            'muestro pasajeros
            dibujo_grilla_egr_p
            muestro_egreso_pasajero
        End If
        'ordeno por primer campo luego de mostrar
        mnuOrden1_Click
    End If
End Sub

Private Sub subConfiguroMenuOrden()
    'Configuro las opciones del menu orden dependiendo de la información
    'que muestro en la grilla; pasajeros o habitación
    
    If Me.mnuMostrarHab.Checked Then
        Me.mnuOrden1.Caption = "Número habitación"
        Me.mnuOrden2.Caption = "Tipo de habitación"
        Me.mnuOrden3.Caption = "Titular única"
        Me.mnuOrden4.Caption = "Titular alojamiento"
        Me.mnuOrden5.Caption = "Titular extras"
        Me.mnuOrden6.Caption = "Hora de egreso"
        Me.mnuOrden5.Visible = True
        Me.mnuOrden6.Visible = True
    End If
    
    If Me.mnuMostrarPasajeros.Checked Then
        Me.mnuOrden1.Caption = "Pasajero"
        Me.mnuOrden2.Caption = "Número de habitación"
        Me.mnuOrden3.Caption = "Tipo de habitación"
        Me.mnuOrden4.Caption = "Hora de egreso"
        Me.mnuOrden5.Caption = "Agencia o empresa"
        Me.mnuOrden6.Visible = False
    End If
End Sub

Private Sub proceso_egresos()
    Dim consultaCheckIn As String
    Dim consultaCheckOut As String
    Dim consulta As String
    
    consultaCheckIn = SQLEgresosPorRealizar(Fcons.Text)
    
    'evalúo tipo de habitacion
    If cboTipo_habitacion.Text <> "(Todas)" Then
        consultaCheckIn = consultaCheckIn & " and habitaciones.tipohab = " & _
        cboTipo_habitacion.ItemData(cboTipo_habitacion.ListIndex)
    End If

    consultaCheckOut = SQLEgresosYaRealizados(Fcons.Text)
    
    'evalúo tipo de habitacion
    If cboTipo_habitacion.Text <> "(Todas)" Then
        consultaCheckOut = consultaCheckOut & " and habitaciones.tipohab = " & _
        cboTipo_habitacion.ItemData(cboTipo_habitacion.ListIndex)
    End If

    'Evalúo fecha
    If IsDate(Fcons.Text) Then
        If Fcons.Value = m_FechaSis Then  'trabajo con checkin y checkout
            consulta = consultaCheckIn & " UNION " & consultaCheckOut
        Else                        'solo me interesa trabajar con Checkin
            consulta = consultaCheckIn
        End If
    End If

    'Ejecuto consulta
    Set qdf = bdHOTEL.CreateQueryDef("")
    qdf.SQL = consulta
    
    Set rst_principal = qdf.OpenRecordset(dbOpenSnapshot)
End Sub

Private Sub muestro_egreso_pasajero()
    'Recorro el recordset principal y listo discriminado por pasajeros
    Dim linea As String
    Dim descri_tipohab As String
    limpio_grilla dbIE
    dibujo_grilla_egr_p
    
    If rst_principal.RecordCount > 0 Then
        rst_principal.MoveFirst
        Do While Not rst_principal.EOF
            If cumple_seleccion_egr Then
            
                If busco_tipo_habTF(rst_principal!tipoHab) Then
                    descri_tipohab = tbTIPO_HABITACIONES("descripcion")
                End If
                linea = Chr(9) & Trim(rst_principal!nombre_completo_titular) & _
                        Chr(9) & rst_principal!nrohab & _
                        Chr(9) & descri_tipohab & _
                        Chr(9) & rst_principal(5) & _
                        Chr(9) & funObtengoNombreEmpresa(rst_principal("nroReserva"))
                dbIE.AddItem linea
                If marco_color_seleccionadas Then
                    marco_linea_grilla
                End If
            End If
            rst_principal.MoveNext
        Loop
    End If
End Sub

Private Function funObtengoNombreEmpresa(NroReserva As Long) As String
    '-------------------------------------------------------------------------------
    'Obtiene el nombre de la agencia que realizó la reserva de la habitación que
    'dejo libre el hotel (archivo Checkout) o lo esta por hacer (archivo Checkin).
    '-------------------------------------------------------------------------------
    'Parámetros.
    '   Entrada:    [nroReserva] número de reserva almacenado en el archivo CheckIn
    '                            o CheckOut, dependiendo si estoy listando
    '                            egresos por realizar o ya realizados.
    '               Si existe la reserva (pasajero ingresó al hotel por medio de un
    '               Checkin y si la reserva fue realizada por medio de una agencia,
    '               devuelve el nombre de la agencia.
    'NOTA:
    'Es importante señalar que si el ingreso se hizo mediante un Walkin, el
    'número de reserva no se encontraá en el archivo de reserva (Reservas).
    '--------------------------------------------------------------------------------
    'declaro variables para trabajar con tablas
    Dim tbRes As Recordset
    Set tbRes = tbRESERVAS
    'por defecto asumo que no existe agencia o que la reserva fue realizada por
    'un particular.
    funObtengoNombreEmpresa = Empty
    tbRes.Index = "i_reservas"
    tbRes.Seek "=", NroReserva
    If Not tbRes.NoMatch Then
        'existe la reserva
        'verifico si fue realizada por una agencia
        If tbRes("agenciaEmpresa") = 1 Then
            funObtengoNombreEmpresa = mFunBuscoNombreEmpresa(tbRes("nroAgenciaEmpresa"))
        End If
    End If
End Function

Private Sub muestro_egreso_habitacion()
    Dim linea As String
    Dim descri_tipohab As String
    Dim titUnica As String
    Dim titAloja As String
    Dim titExtra As String
    
    limpio_grilla dbIE
    dibujo_grilla_egr_h
    
    If rst_principal.RecordCount > 0 Then
        rst_principal.MoveFirst
        Do While Not rst_principal.EOF
            If cumple_seleccion_egr Then
            
                If busco_tipo_habTF(rst_principal!tipoHab) Then
                    descri_tipohab = tbTIPO_HABITACIONES("descripcion")
                End If
                'obtengo nombre de los titulares de las habitaciones
                titUnica = busco_titular_hab(rst_principal!nrohab, "unica")
                If titUnica = "" Then
                    'si el titular no es única busco los otro tipos de titulares
                    titAloja = busco_titular_hab(rst_principal!nrohab, "aloja")
                    titExtra = busco_titular_hab(rst_principal!nrohab, "extra")
                End If
                linea = Chr(9) & rst_principal!nrohab & _
                        Chr(9) & descri_tipohab & _
                        Chr(9) & titUnica & _
                        Chr(9) & titAloja & _
                        Chr(9) & titExtra & _
                        Chr(9) & rst_principal(5)
                dbIE.AddItem linea
                If marco_color_seleccionadas Then
                    marco_linea_grilla
                End If
            End If
            rst_principal.MoveNext
        Loop
    End If
End Sub

Private Function cumple_seleccion_egr()
    'Para cada registro del recordset obtenido,
    'valido que cumpla con la condición de selección determinada
    'en el menú
    
    marco_color_seleccionadas = False
    
    If mnuSeleccionarTodos.Checked Then 'todos
        'si tiene hora de egreso muestro con otro color
        'If rst_principal!expr1004 <> "" Then
        If rst_principal(5) <> "" Then
            marco_color_seleccionadas = True
        End If
        cumple_seleccion_egr = True
    End If
    
    If mnuSeleccionYaDejaron.Checked Then 'ya EGRESADOS
        If rst_principal(5) = "" Then   'esta vacio porque no se fue
            cumple_seleccion_egr = False
        Else
            cumple_seleccion_egr = True
            marco_color_seleccionadas = True
        End If
    End If
    
    If mnuSeleccionNoDejaron.Checked Then 'por EGRESAR
        If rst_principal(5) = "" Then   'esta vacio porque no se fue
            cumple_seleccion_egr = True
        Else
            cumple_seleccion_egr = False
        End If
    End If
End Function

Private Sub subOrdenoGrilla(criterio As Byte)
    'Ordena la grilla segun el criterio
    Select Case criterio
        Case 1
            dbIE.col = 1
        Case 2
            dbIE.col = 2
        Case 3
            dbIE.col = 3
        Case 4
            dbIE.col = 4
        Case 5
            dbIE.col = 5
        Case 6
            dbIE.col = 6
    End Select
    dbIE.Sort = 5
    'almaceno el criterio actualmente utilizado para realizar el listado impreso
    tipoCriteroOrdenListadoImp = criterio
End Sub

'****************************************************************************
'* Procedimientos para dibujar grilla
'****************************************************************************

Private Sub dibujo_grilla_egr_p()
    Dim i As Byte
    cabezal_grilla_egr_p
End Sub

Private Sub dibujo_grilla_egr_h()
    Dim i As Byte
    cabezal_grilla_egr_h
End Sub

Private Sub cabezal_grilla_egr_p()
    dbIE.FormatString = " | Pasajero                                                                                    | Hab.    | Tipo            | Hora Egreso       | Agencia                            "
End Sub

Private Sub cabezal_grilla_egr_h()
    dbIE.FormatString = " | Hab.        | Tipo hab.         | Titular única                     | Titular alojamiento             | Titular extas                       |Hora Eng.         "
End Sub

Private Sub marco_linea_grilla()
    marco_celdas_grilla dbIE, 1, dbIE.Cols - 1, dbIE.Rows - 1, dbIE.Rows - 1
    dbIE.CellBackColor = &HFFFF80
End Sub

Private Function mFunValidoFecha() As Boolean
    mFunValidoFecha = True
    If Not IsDate(Fcons.Text) Then
        'formato de fecha no válido
        mSubMensaje 3, 1
        Fcons.SetFocus
        mFunValidoFecha = False
        Exit Function
    End If
    
    If Fcons.Value < m_FechaSis Then
        'fecha menor a la del día de hoy
        mSubMensaje 3, 2
        Fcons.SetFocus
        mFunValidoFecha = False
        Exit Function
    End If
End Function

Private Sub Form_Unload(Cancel As Integer)
    'Descargo formulario de memoria
    Set frmListadoEgresos = Nothing
End Sub

'*******************************************************
'*
'*   Imprimo listado
'*
'*******************************************************

Private Sub botImprimir_Click()
    'Imprimo información
    If mfunAplicoConfImp(2, 14) = 1 Then
        'verifico que tipo de filtro estoy aplicando
        If Me.TabStrip1.Tabs(1).Selected = True Then
            'muestro todas
            subArmoReporteEgresosPrevistos 1, Fcons.Text
        Else
            If Me.TabStrip1.Tabs(2).Selected = True Then
                'las que estan por egresar
                subArmoReporteEgresosPrevistos 2, Fcons.Text
            Else
                If Me.TabStrip1.Tabs(3).Selected = True Then
                    'las que ya egresaron
                    subArmoReporteEgresosPrevistos 3, Fcons.Text
                End If
            End If
        End If
    End If
End Sub

Private Sub subArmoReporteEgresosPrevistos(tipoFiltro As Byte, fechaegr As Date)
    '----------------------------------------------------------------------------------
    'Configuro el control Crystal genérico,obtiene datos y emite el listado
    'egresos previstos
    '-------------------------------------------------------------------------------
    'Parámetros.
    '   Entrada:    [tipoFiltro] determina que tipo de información muestro.
    '               1 = todos los egresos
    '               2 = solo los que estan por egresar
    '               3 = solo los que ya egresaron
    '               [fechaEgr] fecha de la cual se quiere saber los egresos previstos
    '-------------------------------------------------------------------------------
    
    Dim descTipoFiltro As String
    Dim descOrden As String
    
    Dim consultaCheckIn As String
    Dim consultaCheckOut As String
    Dim consulta As String
        
    'inicializo control data
     subInicializoControlData frmMAIN.Data1CrystalReport
    
    'NOTA: esta consulta debe de ser la misma que se utilizo para crear el reporte,
    'por ese motivo no debe de ser modificada en aspectos tales como campos a mostrar y
    'tablas utilizadas.
    
    'selecciono todos los ingresos no realizados desde checkin y los junto
    'con los ingresos ya realizados de CheckOut
    
    'SELECT checkin.nrohab,checkin.nrocorrcli,checkin.nroReserva,tipo_habitaciones.descripcion,empresas.NomEmp,'' as HoraEgr from tipo_habitaciones INNER JOIN (habitaciones INNER JOIN (checkin LEFT JOIN (reservas LEFT JOIN empresas
    'ON reservas.nroAgenciaEmpresa = empresas.NroCorrEmp) ON checkin.nroReserva = reservas.nroReserva) ON habitaciones.nroHab = checkin.nroHab) ON tipo_habitaciones.tipoHab = habitaciones.tipoHab UNION ALL
    'SELECT checkout.nrohab,checkout.nrocorrcli,checkout.nroReserva,tipo_habitaciones.descripcion,empresas.NomEmp,checkout.horaEgrHab as HoraEgr
    'from tipo_habitaciones INNER JOIN (habitaciones INNER JOIN (checkout LEFT JOIN (reservas LEFT JOIN empresas
    'ON reservas.nroAgenciaEmpresa = empresas.NroCorrEmp) ON checkOut.nroReserva = reservas.nroReserva) ON habitaciones.nroHab = checkout.nroHab) ON tipo_habitaciones.tipoHab = habitaciones.tipoHab
    
    consultaCheckIn = _
    "SELECT checkin.nrohab," & _
    "checkin.nrocorrcli," & _
    "checkin.nroReserva," & _
    "tipo_habitaciones.descripcion," & _
    "empresas.NomEmp," & _
    "'' as 'HoraEgr' " & _
    "from tipo_habitaciones INNER JOIN (habitaciones INNER JOIN (checkin LEFT JOIN (reservas LEFT JOIN empresas " & _
    "ON reservas.nroAgenciaEmpresa = empresas.NroCorrEmp) " & _
    "ON checkin.nroReserva = reservas.nroReserva) " & _
    "ON habitaciones.nroHab = checkin.nroHab) " & _
    "ON tipo_habitaciones.tipoHab = habitaciones.tipoHab " & _
    "WHERE checkin.fCheckHas = " & fechaSQL(fechaegr) & _
    funAplicoFiltroTipoHab
    
    consultaCheckOut = _
    "SELECT checkout.nrohab," & _
    "checkout.nrocorrcli," & _
    "checkout.nroReserva," & _
    "tipo_habitaciones.descripcion," & _
    "empresas.NomEmp," & _
    "checkout.horaEgrHab as 'HoraEgr' " & _
    "from tipo_habitaciones INNER JOIN (habitaciones INNER JOIN (checkout LEFT JOIN (reservas LEFT JOIN empresas " & _
    "ON reservas.nroAgenciaEmpresa = empresas.NroCorrEmp) " & _
    "ON checkOut.nroReserva = reservas.nroReserva) " & _
    "ON habitaciones.nroHab = checkout.nroHab) " & _
    "ON tipo_habitaciones.tipoHab = habitaciones.tipoHab " & _
    "WHERE checkout.fhas = " & fechaSQL(fechaegr) & _
    funAplicoFiltroTipoHab
      
    Select Case tipoFiltro
        Case 1
            'todos los egresos
            descTipoFiltro = "Todos"
            consulta = consultaCheckIn & " UNION " & consultaCheckOut
        Case 2
            'se seleccionan los que estan por egresar
            descTipoFiltro = "Aun no dejaron el hotel"
            consulta = consultaCheckIn
        Case 3
            'selecciono los que ya egresaron
            descTipoFiltro = "Ya egresaron"
            consulta = consultaCheckOut
    End Select
        
    frmMAIN.Data1CrystalReport.RecordSource = consulta
    'ejecuto consulta control data
    frmMAIN.Data1CrystalReport.Refresh
    
    'como de aquí en más voy a trabajar directamente con el control data
    'me aseguro de que se allan encontrado reservas
    If frmMAIN.Data1CrystalReport.Recordset.RecordCount > 0 Then
        'determino el nombre del reporte a utilizar
        frmMAIN.CrystalReport1.ReportFileName = vardir2 + "rptEgrP.rpt"
        'establesco orden del listado
        descOrden = funEstablescoOrdenListado
    
        'inicializo fórmulas
        With frmMAIN.CrystalReport1
            .Formulas(0) = "cabVersion = '" & mFunImpVersionAplicacion & "'"        'nombre y versión aplicación
            .Formulas(1) = "cabRegistro = '" & mFunImpRegistro & "'"                'hotel propietario de la aplicación
            .Formulas(2) = "cabHoraImp = ' " & CStr(Time()) & "'"                   'hora actual
            .Formulas(3) = "cabTipoFiltro = '" & descTipoFiltro & "'"
            .Formulas(4) = "parte1Fecha = '" & Format(fechaegr, "dddd dd mmmm yyyy") & "'"
            .Formulas(5) = "parte1TipoHab = '" & Me.cboTipo_habitacion.Text & "'"
            .Formulas(6) = "parte1Orden = '" & descOrden & "'"
        End With
        'genero listado
        frmMAIN.CrystalReport1.Action = 1
        'aviso de confirmación de impresión de reporte
        mSubMensaje 4, 138  'se imprimieron los egresos previstos
        'inicializo fórmulas
        mSubInicializoFormulas 6
        'inicializo campos de ordenación del informe
        mSubInicializoCamposOrden 1
    Else
        'aviso de que no hay datos para imprimir
        mSubMensaje 3, 9
    End If
End Sub

Private Function funAplicoFiltroTipoHab() As String
    '----------------------------------------------------------------------
    'Determino si tengo que filtarar por tipo de habitación o no.
    'Cuando muestro todas las habitaciones no realizo filtro.
    '----------------------------------------------------------------------
    'Parámetros.
    '   Salida: string con la nueva condición a filtrar en la clausura Where
    '   de la sentencia SQL.
    '-----------------------------------------------------------------------
    If cboTipo_habitacion.Text <> "(Todas)" Then
        funAplicoFiltroTipoHab = " and habitaciones.tipohab = " & _
        cboTipo_habitacion.ItemData(cboTipo_habitacion.ListIndex)
    Else
        funAplicoFiltroTipoHab = Empty
    End If
End Function

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
    'el criterio de ordenación de la grilla.
    '-------------------------------------------------------------------
    Dim criterio As String
    Select Case tipoCriteroOrdenListadoImp
        Case 1  'por pasajero
            funEstablescoOrdenListado = "Ordenado por pasajero"
            frmMAIN.CrystalReport1.SortFields(0) = "+{CLIENTES.nombre_completo_titular}"
        Case 2  'habitación
            funEstablescoOrdenListado = "Ordenado por número de habitación"
            frmMAIN.CrystalReport1.SortFields(0) = "+{Bound Control.nrohab}"
        Case 3  'tipo de habitación
            funEstablescoOrdenListado = "Ordenado por tipo de habitación"
            frmMAIN.CrystalReport1.SortFields(0) = "+{Bound Control.descripcion}"
        Case 4  'hora egreso
            funEstablescoOrdenListado = "Ordenado por hora de egreso"
            frmMAIN.CrystalReport1.SortFields(0) = "+{Bound Control.'HoraEgr'}"
        Case 5  'agencia empresa
            funEstablescoOrdenListado = "Ordenado por agencia"
            frmMAIN.CrystalReport1.SortFields(0) = "+{Bound Control.NomEmp}"
    End Select
End Function

'***********************************
'*
'* Fin impresión
'*
'***********************************

Private Sub mnuMostrarHab_Click()
    'Muestro egresos discriminados por habitación
    mnuMostrarPasajeros.Checked = False
    mnuMostrarHab.Checked = True
    'cambio menu de ordenación
    subConfiguroMenuOrden
        
    subEjecutar
End Sub

Private Sub mnuMostrarPasajeros_Click()
    'Muestro egresos discriminados por pasajeros
    mnuMostrarPasajeros.Checked = True
    mnuMostrarHab.Checked = False
    'cambio menu de ordenación
    subConfiguroMenuOrden
    
    subEjecutar
End Sub

Private Sub mnuOrden1_Click()
    'ordeno primer campo de cualquiera de las grillas
    subDesmarcoOrdenes
    mnuOrden1.Checked = True
    subOrdenoGrilla 1
    mSubMuestroIcono dbIE, 1
End Sub

Private Sub mnuOrden2_Click()
    'ordeno por el segundo campo de cualquiera de las grillas
    subDesmarcoOrdenes
    mnuOrden2.Checked = True
    subOrdenoGrilla 2
    mSubMuestroIcono dbIE, 2
End Sub

Private Sub mnuOrden3_Click()
    'ordeno por el tercer campo de cualquiera de las grillas
    subDesmarcoOrdenes
    mnuOrden3.Checked = True
    subOrdenoGrilla 3
    mSubMuestroIcono dbIE, 3
End Sub

Private Sub mnuOrden4_Click()
    'ordeno por el cuarto campo de cualquiera de las grillas
    subDesmarcoOrdenes
    mnuOrden4.Checked = True
    subOrdenoGrilla 4
    mSubMuestroIcono dbIE, 4
End Sub

Private Sub mnuOrden5_Click()
    'ordeno por el quinto campo de cualquiera de las grillas
    subDesmarcoOrdenes
    mnuOrden5.Checked = True
    subOrdenoGrilla 5
    mSubMuestroIcono dbIE, 5
End Sub

Private Sub mnuOrden6_Click()
    'ordeno por el sexto campo de cualquiera de las grillas
    subDesmarcoOrdenes
    mnuOrden6.Checked = True
    subOrdenoGrilla 6
    mSubMuestroIcono dbIE, 6
End Sub

Private Sub subDesmarcoOrdenes()
    mnuOrden1.Checked = False
    mnuOrden2.Checked = False
    mnuOrden3.Checked = False
    mnuOrden4.Checked = False
    mnuOrden5.Checked = False
    mnuOrden6.Checked = False
End Sub

Private Sub subDesmarcoSeleccion()
    mnuSeleccionarTodos.Checked = False
    mnuSeleccionNoDejaron.Checked = False
    mnuSeleccionYaDejaron.Checked = False
End Sub
    
Private Sub mnuSeleccionarTodos_Click()
    'Muestro todos los egresos
    TabStrip1.Tabs(1).Selected = True
End Sub

Private Sub mnuSeleccionNoDejaron_Click()
    'muestro solo los que aun no dejaron el hotel
    TabStrip1.Tabs(2).Selected = True
End Sub

Private Sub mnuSeleccionYaDejaron_Click()
    'muestro los que ya dejaron el hotel
    TabStrip1.Tabs(3).Selected = True
End Sub

Private Sub TabStrip1_Click()
    'Cuando picho un tab efectúo la mismas operaciones que cuando
    'selecciono las opciones desde el menú, dando mayor funcionalidad al programa.
    If TabStrip1.SelectedItem.Index = 1 Then  'ejecuto opcion todas
        subDesmarcoSeleccion
        mnuSeleccionarTodos.Checked = True
        subEjecutar
    End If
    If TabStrip1.SelectedItem.Index = 2 Then  'ejecuto opción aun no dejaron
        subDesmarcoSeleccion
        mnuSeleccionNoDejaron.Checked = True
        subEjecutar
    End If
    If TabStrip1.SelectedItem.Index = 3 Then  'ejecuto opción ya dejaron el hotel
        subDesmarcoSeleccion
        mnuSeleccionYaDejaron.Checked = True
        subEjecutar
    End If
End Sub

Private Sub botSalir_Click()
    Unload Me
End Sub

Private Sub mnuFormularioAceptar_Click()
    'Equivale a presionar el boton aceptar o la tecla F12
    botSalir_Click
End Sub

Private Sub mnuFormularioProcesar_Click()
    'Equivale a presionar el boton procesar
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
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 46
End Sub

Private Sub cboTipo_habitacion_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 47
End Sub

Private Sub botProcesar_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 28
End Sub

Private Sub botImprimir_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 48
End Sub

Private Sub botSalir_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 1, 1
End Sub

Private Sub botSalir_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub Fcons_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub botImprimir_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub botProcesar_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub cboTipo_habitacion_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

