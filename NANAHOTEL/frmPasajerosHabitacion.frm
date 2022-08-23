VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPasajerosHabitacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pasajeros por habitación"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Hotel_Nana.gaHOTELtitular gaHOTELtitular1 
      Height          =   1335
      Left            =   120
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   2355
      BackColor       =   -2147483633
   End
   Begin VB.Frame Frame1 
      Caption         =   "Información habitación"
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   9255
      Begin VB.CommandButton botImprimirTodas 
         Caption         =   "Imprimir &Todas"
         Height          =   375
         Left            =   4800
         TabIndex        =   13
         Top             =   4440
         Width           =   1695
      End
      Begin VB.CommandButton botSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   375
         Left            =   7920
         TabIndex        =   3
         Top             =   4440
         Width           =   1215
      End
      Begin VB.CommandButton botImprimir 
         Height          =   375
         Left            =   6600
         Picture         =   "frmPasajerosHabitacion.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         Tag             =   "Imprimir"
         Top             =   4440
         Width           =   1215
      End
      Begin MSFlexGridLib.MSFlexGrid gPasajeros 
         Bindings        =   "frmPasajerosHabitacion.frx":0942
         Height          =   3015
         Left            =   120
         TabIndex        =   1
         Top             =   1320
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   5318
         _Version        =   393216
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   5880
         Picture         =   "frmPasajerosHabitacion.frx":0952
         Top             =   360
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label lblAtencion 
         Caption         =   "lblAtencion"
         Height          =   855
         Left            =   6480
         TabIndex        =   12
         Top             =   240
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Label lblCantPasajeros 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblCantPasajeros"
         Height          =   360
         Left            =   2640
         TabIndex        =   11
         Top             =   780
         Width           =   615
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad de pasajeros:"
         Height          =   240
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   2100
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "al"
         Height          =   240
         Left            =   3960
         TabIndex        =   9
         Top             =   360
         Width           =   165
      End
      Begin VB.Label lblFechaHasta 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblFechaHasta"
         Height          =   360
         Left            =   4320
         TabIndex        =   8
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label lblFechaDesde 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblFechaDesde"
         Height          =   360
         Left            =   2640
         TabIndex        =   7
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Período de ocupación:"
         Height          =   240
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   2040
      End
   End
   Begin Hotel_Nana.gaHOTELbarra gaHOTELbarra1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   6495
      Width           =   9510
      _ExtentX        =   16775
      _ExtentY        =   582
   End
   Begin VB.Menu mnuformulario 
      Caption         =   "&Formulario"
      Begin VB.Menu mnuformularioAceptar 
         Caption         =   "Salir"
         Shortcut        =   {F12}
      End
      Begin VB.Menu div1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuImprimirConsulta 
         Caption         =   "Imprimir"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuImprimirTodas 
         Caption         =   "Imprimir todas las habitaciones"
         Shortcut        =   ^T
      End
   End
   Begin VB.Menu mnuOrden 
      Caption         =   "&Ordenado por ..."
      Begin VB.Menu mnuOrdenPasa 
         Caption         =   "Por nombre de pasajero"
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu mnuOrenPais 
         Caption         =   "Por pais "
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu mnuOrdenFecha 
         Caption         =   "Por fecha de nacimiento"
         Shortcut        =   ^{F3}
      End
   End
End
Attribute VB_Name = "frmPasajerosHabitacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private qdf_principal As QueryDef
Private rst_p As Recordset

Public hab_cuenta As Long

Private Sub Form_Load()
    'Apariencia formulario
    mSubConfiguroFuentesControlesSistema Me
    
    'obtengo habitacion
    hab_cuenta = Val(frmIngHabitacion.txtNroHab.Text)
    cabezal_formulario
    
    'Ordeno predeterminadamente por nombre
    mnuOrdenPasa_Click
    
    'inicializo propiedades de controles
    mSubBloqueoControlFormulario Me.lblCantPasajeros, True
    mSubBloqueoControlFormulario Me.lblFechaDesde, True
    mSubBloqueoControlFormulario Me.lblFechaHasta, True
End Sub

Private Sub cabezal_formulario()
    Me.gaHOTELtitular1.CaminoBaseDeDatos = vardir
    Me.gaHOTELtitular1.NumeroHabitacion = hab_cuenta
End Sub

Private Sub subInicializoConsulta(cri1 As String)
    'Genero un recordset con los datos de la consulta
    Dim consulta As String
    
    'No utilizo el procedimeinto SQLpasajeros_habitacion ya que el mismo solo
    'debulve el nombre del pasajero
    consulta = "Select clientes.nombre_completo_titular, " & _
                "       paises.descri_pais, " & _
                "       clientes.fecha_nac_titular, " & _
                "       checkin.nrocorrcli, " & _
                "       checkin.fCheckDes, " & _
                "       checkin.fCheckHas " & _
                "From   checkin, clientes, paises " & _
                "Where  checkin.nrocorrcli = clientes.nrocorr " & _
                "and checkin.nrohab = " & hab_cuenta & _
                " and clientes.pais_reside_titular = paises.cod_pais " & _
                " Order by " & cri1
    Set qdf_principal = bdHOTEL.CreateQueryDef("")
    qdf_principal.SQL = consulta
    Set rst_p = qdf_principal.OpenRecordset(dbOpenSnapshot)
    
    rst_p.MoveLast  'obtengo cantidad de pasajeros
    'muestro información habitación
    Me.lblFechaDesde.Caption = rst_p("fcheckdes")
    Me.lblFechaHasta.Caption = rst_p("fcheckhas")
    Me.lblCantPasajeros.Caption = rst_p.RecordCount
    'NOTA: en realidad estoy mostrando la información del registro correspondiente al
    'último pasajero de la habitación, pero esto no es trascendente ya que la información
    'del período de alojamiento, simempre es la misma para todos los pasajeros de las habitaciones
    'ocupadas.
    If Not mFunDeterminoOcupacionValida(hab_cuenta) Then
        Image1.Visible = True
        lblAtencion.Visible = True
        lblAtencion.Caption = "Atención:" & Chr(10) & "Período de ocupación fuera del establecido."
    End If
End Sub

Private Sub subCabezalGrilla()
    gPasajeros.FormatString = _
    "|Nombre del pasajero                                                                            |" & _
    "Pais                         |" & _
    "Fecha de nac. "
End Sub
    
Private Sub subReOrdenoGrilla(criterio As Byte)
    Select Case criterio
        Case 1  'por nombre de pasajero
            subInicializoConsulta "clientes.nombre_completo_titular"
            subRecorroRegistro
            
        Case 2  'por pais
            subInicializoConsulta "descri_pais"
            subRecorroRegistro
            
        Case 3  'por fecha de nacimiento
            subInicializoConsulta "clientes.fecha_nac_titular"
            subRecorroRegistro
            
    End Select
End Sub

Private Sub subRecorroRegistro()
    'Recorro el RecorSet y muestro los datos en la grilla
    subLimpioGrilla
    
    rst_p.MoveFirst
    Do While Not rst_p.EOF
        gPasajeros.AddItem funCreoLinea
        rst_p.MoveNext
    Loop
End Sub

Private Function funCreoLinea()
    'Cada registro del RecordSet es una linea de la grilla
    Dim linea As String
    
    linea = Chr(9) & _
            rst_p("nombre_completo_titular") & _
            Chr(9) & rst_p("descri_pais") & _
            Chr(9) & rst_p("fecha_nac_titular")
    funCreoLinea = linea
End Function

Private Sub subLimpioGrilla()
    limpio_grilla gPasajeros
    subCabezalGrilla
End Sub
   
Private Sub Form_Unload(Cancel As Integer)
    'Descargo formulario de memoria
    Set frmPasajerosHabitacion = Nothing
End Sub

'*************************************************
'*
'* Imprimir reporte
'*
'*************************************************

Private Sub botImprimir_Click()
    'Imprimo solo la habitación seleccionada
    If mfunAplicoConfImp(2, 16) = 1 Then
        subArmoReportePasajerosHab 1, hab_cuenta
    End If
End Sub

Private Sub botImprimirTodas_Click()
    'Imprimo todas las habitaciones ocupadas actualmente
    If mfunAplicoConfImp(2, 16) = 1 Then
        subArmoReportePasajerosHab 2
    End If
End Sub

Private Sub subArmoReportePasajerosHab(tipoReporte As Byte, Optional nrohab As Long)
    '----------------------------------------------------------------------------------
    'Configuro el control Crystal genérico,obtiene datos y emite el listado
    'de pasajeros por habitación
    '-------------------------------------------------------------------------------
    'Parámetros.
    '   Entrada:    [tipoReporte] determina que tipo de información que muestro.
    '               1 = muestro solo la habitación seleccionada
    '               2 = muestro todas las habitaciones ocupadas actualmente
    '
    '               [nroHab Optional] en caso de que se este listando pasajeros
    '               de una sola habitación, la misma se pasa como parámetro.
    '-------------------------------------------------------------------------------
    Dim descTipoFiltro As String
    
    'inicializo control data
     subInicializoControlData frmMAIN.Data1CrystalReport
    
    'NOTA: esta consulta debe de ser la misma que se utilizo para crear el reporte,
    'por ese motivo no debe de ser modificada en aspectos tales como campos a mostrar y
    'tablas utilizadas.
    
    'selecciono todos los pasajeros que se encuentran alojados en el hotel.
    'si tipoReporte=1, entonces solo selecciono los de una habitación determinada.
    
    frmMAIN.Data1CrystalReport.RecordSource = _
    "select * from checkin,habitaciones,tipo_habitaciones,clientes,paises where " & _
    "checkin.nrohab = habitaciones.nrohab and " & _
    "checkin.nroCorrCli = clientes.nroCorr and " & _
    "habitaciones.tipohab = tipo_habitaciones.tipohab and " & _
    "clientes.pais_reside_titular = paises.cod_pais "
    
    'verifico si tengo que filtrar para una habitación determinada
    If tipoReporte = 1 Then
        frmMAIN.Data1CrystalReport.RecordSource = frmMAIN.Data1CrystalReport.RecordSource & _
        "and checkin.nrohab = " & nrohab
        'no muestro la sección de resumen en el listado y mantengo el valor predeterminado
        'para los demas parámetros de la sección
        frmMAIN.CrystalReport1.SectionFormat(0) = "SUMMARY;F;X;X;X;X;X;X"
        'establesco descripción a mostrar en cabezal del reporte
        descTipoFiltro = "Hab.: " & nrohab
    Else
        'establesco descripción a mostrar en cabezal del reporte
        descTipoFiltro = "Completo"
    End If

    'ejecuto consulta control data
    frmMAIN.Data1CrystalReport.Refresh
    
    'como de aquí en más voy a trabajar directamente con el control data
    'me aseguro de que se allan encontrado reservas
    If frmMAIN.Data1CrystalReport.Recordset.RecordCount > 0 Then
        'determino el nombre del reporte a utilizar
        frmMAIN.CrystalReport1.ReportFileName = vardir2 + "rptPhab.rpt"
    
        'inicializo fórmulas
        With frmMAIN.CrystalReport1
            .Formulas(0) = "cabVersion = '" & mFunImpVersionAplicacion & "'"        'nombre y versión aplicación
            .Formulas(1) = "cabRegistro = '" & mFunImpRegistro & "'"                'hotel propietario de la aplicación
            .Formulas(2) = "cabHoraImp = ' " & CStr(Time()) & "'"                   'hora actual
            .Formulas(3) = "cabTipoFiltro = '" & descTipoFiltro & "'"
        End With
        'genero listado
        frmMAIN.CrystalReport1.Action = 1
        'aviso de confirmación de impresión de reporte
        mSubMensaje 4, 139  'se imprimieron los pasajeros por habitación
        'inicializo fórmulas
        mSubInicializoFormulas 3
        'establesco nuevamente valores predeterminados para la sección resumen
        frmMAIN.CrystalReport1.SectionFormat(0) = "SUMMARY;T;X;X;X;X;X;X"
    Else
        'aviso de que no hay gastos para imprimir
        mSubMensaje 3, 9
    End If
End Sub

'*************************************************
'*
'* Fin procedimientos de impresión
'*
'*************************************************

Private Sub mnuOrdenPasa_Click()
    'Ordeno por nonbre pasajero
    subReOrdenoGrilla 1
    mSubMuestroIcono gPasajeros, 1
    subDesmarcoTodas
    mnuOrdenPasa.Checked = True
End Sub

Private Sub mnuOrenPais_Click()
    'Ordeno por pais
    subReOrdenoGrilla 2
    mSubMuestroIcono gPasajeros, 2
    subDesmarcoTodas
    mnuOrenPais.Checked = True
End Sub

Private Sub mnuOrdenFecha_Click()
    'Oreno por fechas de nacimiento
    subReOrdenoGrilla 3
    mSubMuestroIcono gPasajeros, 3
    subDesmarcoTodas
    mnuOrdenFecha.Checked = True
End Sub

Private Sub subDesmarcoTodas()
    'Desmarco todas las opciones del menu, de esta manera solo
    'se puede ver marcada la opción por la cual esta ordenada la grilla en ese momento.
    mnuOrdenPasa.Checked = False
    mnuOrenPais.Checked = False
    mnuOrdenFecha.Checked = False
End Sub

Private Sub mnuFormularioAceptar_Click()
    'Equivale a presionar el boton salir o F12
    botSalir_Click
End Sub

Private Sub mnuImprimirTodas_Click()
    'Equivale a presionar el boton de imprimir todas las habitaciones
    botImprimirTodas_Click
End Sub

Private Sub mnuImprimirConsulta_Click()
    'Equivale a presionar el boton de imprimir o Ctrol+I
    botImprimir_Click
End Sub

Private Sub botSalir_Click()
    Unload Me
    frmIngHabitacion.Show 1
End Sub

'******************************************
'*
'*  Asistencia a usuarios
'*
'******************************************

Private Sub botImprimir_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 51
End Sub

Private Sub botSalir_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 1, 1
End Sub

Private Sub botImprimirTodas_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 211
End Sub

Private Sub botSalir_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub botImprimir_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub botImprimirTodas_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

