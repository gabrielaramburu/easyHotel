VERSION 5.00
Object = "{EB8C3860-8603-11D0-A35D-7062E4000000}#1.1#0"; "VCBND.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmCotizaciones 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cotizaciones"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Ingreso valor"
      Height          =   4575
      Left            =   120
      TabIndex        =   19
      Top             =   0
      Width           =   3615
      Begin VB.CommandButton botNuevoValor 
         Caption         =   "&Agregar"
         Height          =   375
         Left            =   2160
         TabIndex        =   6
         Top             =   4080
         Width           =   1215
      End
      Begin VB.TextBox txtCotizacion 
         Height          =   375
         Left            =   240
         MaxLength       =   10
         TabIndex        =   5
         Top             =   3480
         Width           =   1335
      End
      Begin VB.ComboBox cboMonedas 
         Height          =   360
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   600
         Width           =   3135
      End
      Begin VcBndCtl.VcCalCombo fechaCot 
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   2640
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _0              =   $"frmCotizaciones.frx":0000
         _1              =   $"frmCotizaciones.frx":0409
         _2              =   $"frmCotizaciones.frx":0812
         _3              =   "-A@@@)f@@@E@@@A@@@@@@@@@'@@@E@@@A@@@@@@@@@)h@@@E@@@A@@@@@@@@@,456D"
         _count          =   4
         _ver            =   2
      End
      Begin VB.Label lblFechaCotActual 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblFechaCotActual"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label lblCotizacionActual 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblCotizacionActual"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   21
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "&Valor de la nueva cotización"
         Height          =   240
         Left            =   240
         TabIndex        =   4
         Top             =   3240
         Width           =   2520
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "&Fecha de la nueva cotización"
         Height          =   240
         Left            =   240
         TabIndex        =   2
         Top             =   2400
         Width           =   2610
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&Monedas"
         Height          =   240
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Cotización actual"
         Height          =   240
         Left            =   240
         TabIndex        =   20
         Top             =   1080
         Width           =   1515
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Consultar "
      Height          =   4575
      Left            =   3840
      TabIndex        =   18
      Top             =   0
      Width           =   5535
      Begin VB.ComboBox cboTipoOrdenacion 
         Height          =   360
         ItemData        =   "frmCotizaciones.frx":0C1B
         Left            =   3960
         List            =   "frmCotizaciones.frx":0C25
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   2040
         Width           =   1455
      End
      Begin MSDBGrid.DBGrid gCotizaciones 
         Bindings        =   "frmCotizaciones.frx":0C42
         Height          =   3855
         Left            =   240
         OleObjectBlob   =   "frmCotizaciones.frx":0C52
         TabIndex        =   16
         Top             =   600
         Width           =   3615
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   495
         Left            =   3960
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   3000
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VcBndCtl.VcCalCombo fDesde 
         Height          =   375
         Left            =   3960
         TabIndex        =   8
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _0              =   $"frmCotizaciones.frx":160A
         _1              =   $"frmCotizaciones.frx":1A13
         _2              =   $"frmCotizaciones.frx":1E1C
         _3              =   "-A@@@)f@@@E@@@A@@@@@@@@@'@@@E@@@A@@@@@@@@@)h@@@E@@@A@@@@@@@@@,456D"
         _count          =   4
         _ver            =   2
      End
      Begin VB.CommandButton botConsultar 
         Caption         =   "&Consultar"
         Height          =   375
         Left            =   4200
         TabIndex        =   13
         Top             =   4080
         Width           =   1215
      End
      Begin VcBndCtl.VcCalCombo fHasta 
         Height          =   375
         Left            =   3960
         TabIndex        =   10
         Top             =   1320
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _0              =   $"frmCotizaciones.frx":2225
         _1              =   $"frmCotizaciones.frx":262E
         _2              =   $"frmCotizaciones.frx":2A37
         _3              =   "-@A@@@)f@@@E@@@A@@@@@@@@@'@@@E@@@A@@@@@@@@@)h@@@E@@@A@@@@@@@@@,456D"
         _count          =   4
         _ver            =   2
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "O&rdenadar"
         Height          =   240
         Left            =   3960
         TabIndex        =   11
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Fec&ha final"
         Height          =   240
         Left            =   3960
         TabIndex        =   9
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "F&echa inicial"
         Height          =   240
         Left            =   3960
         TabIndex        =   7
         Top             =   360
         Width           =   1125
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "C&otizaciones"
         Height          =   240
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   1155
      End
   End
   Begin VB.CommandButton botSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   8160
      TabIndex        =   14
      Top             =   4680
      Width           =   1215
   End
   Begin Hotel_Nana.gaHOTELbarra gaHOTELbarra1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   17
      Top             =   5115
      Width           =   9510
      _ExtentX        =   16775
      _ExtentY        =   582
   End
   Begin VB.Menu mnuFormulario 
      Caption         =   "&Formulario"
      Begin VB.Menu mnuFormularioAgregar 
         Caption         =   "Agregar "
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuFormularioSalir 
         Caption         =   "Salir"
         Shortcut        =   {F12}
      End
   End
End
Attribute VB_Name = "frmCotizaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub botConsultar_Click()
    'Realizo consulta de cotizaciones
    
    'valido fecha inicial
    If Not IsDate(Me.fDesde.Value) Then
        mSubMensaje 3, 1
        Me.fDesde.SetFocus
    Else
        'valido fecha final
        If Not IsDate(Me.fHasta.Value) Then
            mSubMensaje 3, 1
            Me.fHasta.SetFocus
        Else
            'valido que las fechas ingresadas sena corectas
            If Me.fDesde > Me.fHasta Then
                mSubMensaje 3, 3
            Else
                'ejecuto consulta
                subEjecutoConsulta 2, Me.cboMonedas.ItemData(Me.cboMonedas.ListIndex), _
                Me.fDesde.Value, Me.fHasta.Value
            End If
        End If
    End If
End Sub

Private Sub cboMonedas_Click()
    'Cuando cambio el elemento seleccionado del combo de monedas
    'actualizo el formulario.
    
    subInicializoControles
    'ejecuto consulta mostrando todas la cotizaciones de la moneda por defecto
    subEjecutoConsulta 1, Me.cboMonedas.ItemData(Me.cboMonedas.ListIndex)
End Sub

Private Sub Form_Load()
    'inicializo control data
    subInicializoControlData Me.Data1
    'cargo combo de monedas
    carga_tipo_moneda Me.cboMonedas
    'por defecto muestro moneda dólares
    posiciono_combo Me.cboMonedas, 1
End Sub

Private Sub botNuevoValor_Click()
    'Agrego o modifico la cotización de una moneda determinada, para una fecha determinada
    
    'valido los datos ingresados
    If funValidoDatos Then
        If mFunMensaje(4, 122) Then
            'verifico si existe cotización para la fecha
            If funExisteCotizacion(Me.cboMonedas.ItemData(cboMonedas.ListIndex), Me.fechaCot.Value) Then
                'si existe cotización: modifico el valor
                tbCOTIZACIONES.Edit
                    tbCOTIZACIONES("valorCot") = Me.txtCotizacion
                tbCOTIZACIONES.Update
            Else
                'sino existe cotización: agrego un nuevo registro
                tbCOTIZACIONES.AddNew
                    tbCOTIZACIONES("codmoneda") = Me.cboMonedas.ItemData(cboMonedas.ListIndex)
                    tbCOTIZACIONES("fechaCot") = Me.fechaCot.Value
                    tbCOTIZACIONES("valorCot") = Me.txtCotizacion
                tbCOTIZACIONES.Update
            End If
            'ejecuto consulta
            subEjecutoConsulta 1, Me.cboMonedas.ItemData(cboMonedas.ListIndex), Me.fechaCot
            'inicializo controles para nuevo ingreso
            subInicializoControles
        End If
    End If
End Sub

Private Sub subEjecutoConsulta(tipoConsulta As Byte, moneda As Byte, Optional fd As Date, Optional fh As Date)
    'Realizo consulta y muestro en grilla
    '---------------------------------------------------------------------------------------
    'Parámetros.
    '   Entrada:    [tipoConsulta] 1= ejecuto la consulta después de ingresar o
    '                              modificar una cotización (no trabaja con rango de fechas)
    '                              2= ejecuto la consulta con un rango de fechas completo
    '               [moneda]    moneda a la que voy a mostrar la cotización
    '               [fd]        fecha desde
    '               [fh]        fecha hasta
    '               Estos parámetros son opcionales porque no simpre utilizo fechas
    '               para listar las cotizaciones.
    '--------------------------------------------------------------------------------------
    Dim consulta As String
    'consulta general
    consulta = "select cotizaciones.fechacot," & _
                        "cotizaciones.valorcot, " & _
                        "monedas.descmoneda " & _
                        "from monedas,cotizaciones " & _
                        "where monedas.codmoneda = cotizaciones.codmoneda " & _
                        " and monedas.codmoneda = " & moneda

    Select Case tipoConsulta
        Case 1
            'ordeno en forma descendente
            consulta = consulta & " and cotizaciones.fechacot = " & fechaSQL(m_FechaSis)
                        
        Case 2
            'selecciono cotizaciones dentro de un rango de fechas ingresado
            consulta = consulta & " and cotizaciones.fechacot >= " & fechaSQL(fd) & _
                                    " and cotizaciones.fechacot <= " & fechaSQL(fh) & _
                                    " Order by cotizaciones.fechacot " & funObtengoOrden
    End Select
    Data1.RecordSource = consulta
    Data1.Refresh
    'configuro cabezal consulta
    gCotizaciones.Columns(0).Caption = "Fecha"
    gCotizaciones.Columns(0).Width = 900
    gCotizaciones.Columns(1).Caption = "Valor"
    gCotizaciones.Columns(1).Width = 1000
    gCotizaciones.Columns(2).Caption = "Moneda"
    gCotizaciones.Columns(2).Width = 1500
End Sub

Private Sub subInicializoControles()
    'Inicializo los controles para un nuevo ingreso
    Me.txtCotizacion = ""
    'por defecto trabajo con la fecha de la aplicación
    Me.fechaCot.Value = m_FechaSis
    'obtengo ultima cotización
    Me.lblCotizacionActual.Caption = mFunObtengoUltimaCotizacion(1, Me.cboMonedas.ItemData(cboMonedas.ListIndex), m_FechaSis)
    Me.lblFechaCotActual.Caption = mFunObtengoUltimaCotizacion(2, Me.cboMonedas.ItemData(cboMonedas.ListIndex), m_FechaSis)
    'inicializo los controles de ingreso de fecha de consulta
    Me.fDesde.Value = Null
    Me.fHasta.Value = Null
    'por defecto ordeno en forma ascendete
    Me.cboTipoOrdenacion.ListIndex = 0
End Sub

Private Function funObtengoOrden() As String
    'Devuelve un string con el valor del combo de ordenación
    'el cual se utiliza para determinar el orden de ordenación de la consulta de cotizaciones
    '-----------------------------------------------------------------------------------------
    'Parámetros.
    '   Salida: Desc, si el valor del combo es Descendente
    '           Asc, si el valor del combo es Ascendente
    '------------------------------------------------------------------------------------------
    If Me.cboTipoOrdenacion.Text = "Descendente" Then
        funObtengoOrden = "Desc"
    Else
        If Me.cboTipoOrdenacion.Text = "Ascendente" Then
            funObtengoOrden = "Asc"
        Else
            'no hay ningun elemento seleccionado en el combo,
            'ordeno en forma Ascendente
            funObtengoOrden = "Asc"
        End If
    End If
End Function

Private Function funExisteCotizacion(moneda As Byte, fecha As Date) As Boolean
    'Determina si la cotización de una moneda, para una fecha determinada, existe.
    '----------------------------------------------------------------------------
    'Parámetros.
    '   Entrada:    [moneda] Moneda seleccionada: 0 moneda nacional.
    '                                             1 dólares.
    '               [fecha] Fecha de la cotización
    '
    '   Salida: True, existe el registro con clave moneda,fecha
    '           False, no existe el registro con clave moneda,fecha
    '-----------------------------------------------------------------------------
    tbCOTIZACIONES.Index = "pkCotizaciones"
    tbCOTIZACIONES.Seek "=", moneda, fecha
    If Not tbCOTIZACIONES.NoMatch Then
        'existe cotización
        funExisteCotizacion = True
    Else
        'no existe cotización
        funExisteCotizacion = False
    End If
End Function

Private Function funValidoDatos() As Boolean
    'Determino si se ingresaron los datos suficientes como
    'para ingresar o modificar una cotización
    '-----------------------------------------------------------------------------------
    'Parámetros.
    '   Salida:
    '           True, se ingreso fecha y valor
    '           False, no se ingreso fecha
    '           False, no se ingresó valor
    '           False, si la fecha de la cotización es mayor a la fecha de la aplicación
    '           False, si la cotización es 0
    '------------------------------------------------------------------------------------
    'por defecto asumo que está todo bien
    funValidoDatos = True
    'verifico si se ingresó fecha
    If Not IsDate(Me.fechaCot.Value) Then
        'aviso: no se ingresó fecha
        mSubMensaje 4, 120
        funValidoDatos = False
        'le doy el focus al control correspondiente
        Me.fechaCot.SetFocus
    Else
        If Me.fechaCot.Value > m_FechaSis Then
            'aviso: la fecha de la cotización debe de ser menor igual a la fecha actual
            mSubMensaje 4, 123
            funValidoDatos = False
            'le doy el focus al control correspondiente
            Me.fechaCot.SetFocus
        Else
            'verifico si se ingresó cotización
            If Trim(Me.txtCotizacion) = Empty Then
                'aviso: no se ingresó cotización
                mSubMensaje 4, 121
                funValidoDatos = False
                'le doy el focus al control correspondiente
                Me.txtCotizacion.SetFocus
            Else
                If Me.txtCotizacion.Text <= 0 Then
                    'aviso: la cotización debe de ser mayor a 0
                    mSubMensaje 4, 128
                    funValidoDatos = False
                    'le doy el focus al control correspondiente
                    Me.txtCotizacion.SetFocus
                End If
            End If
        End If
    End If
End Function

Private Sub mnuFormularioAgregar_Click()
    'Agrego o modifico valor cotización
    botNuevoValor_Click
End Sub

Private Sub mnuFormularioSalir_Click()
    'Cierro formulario
    botSalir_Click
End Sub

Private Sub txtCotizacion_KeyPress(KeyAscii As Integer)
    'Valido solo ingreso de números
    ValidoNum KeyAscii, True, True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCotizaciones = Nothing
End Sub

Private Sub botSalir_Click()
    Unload Me
End Sub

'*****************************************************
'
'   Asistencia a usuarios
'
'*****************************************************
Private Sub cboMonedas_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 195
End Sub

Private Sub fechaCot_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 196
End Sub

Private Sub txtCotizacion_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 197
End Sub

Private Sub botNuevoValor_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 198
End Sub

Private Sub fDesde_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 199
End Sub

Private Sub fHasta_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 200
End Sub

Private Sub botConsultar_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 201
End Sub

Private Sub botSalir_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 1, 1
End Sub

Private Sub cboTipoOrdenacion_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 202
End Sub

Private Sub cboMonedas_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub fechaCot_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub txtCotizacion_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub botNuevoValor_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub fDesde_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub fHasta_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub botConsultar_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub botSalir_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub cboTipoOrdenacion_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

