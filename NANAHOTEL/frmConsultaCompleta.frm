VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmConsultaCompleta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta Completa"
   ClientHeight    =   7680
   ClientLeft      =   900
   ClientTop       =   675
   ClientWidth     =   10365
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   10365
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton botImprimir 
      Height          =   375
      Left            =   7680
      Picture         =   "frmConsultaCompleta.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Tag             =   "Imprimir"
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton botSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   9000
      TabIndex        =   5
      Top             =   6960
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6735
      Left            =   120
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   120
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   11880
      _Version        =   327680
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Detalle"
      TabPicture(0)   =   "frmConsultaCompleta.frx":0942
      Tab(0).ControlCount=   1
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      TabCaption(1)   =   "Resumen"
      TabPicture(1)   =   "frmConsultaCompleta.frx":095E
      Tab(1).ControlCount=   2
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame3"
      Tab(1).Control(1).Enabled=   0   'False
      Begin VB.Frame Frame2 
         Caption         =   "Resumen por estado"
         Height          =   3735
         Left            =   -74760
         TabIndex        =   10
         Top             =   480
         Width           =   9615
         Begin VB.Image Image3 
            Height          =   105
            Left            =   240
            Picture         =   "frmConsultaCompleta.frx":097A
            Stretch         =   -1  'True
            Top             =   2400
            Width           =   9090
         End
         Begin VB.Label lblPorHotelOcupacion 
            AutoSize        =   -1  'True
            Caption         =   "%"
            Height          =   240
            Left            =   8880
            TabIndex        =   24
            Top             =   3000
            Width           =   180
         End
         Begin VB.Label lblBarraBloq 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1680
            TabIndex        =   23
            Top             =   1680
            Visible         =   0   'False
            Width           =   6975
         End
         Begin VB.Label lblBarraOcu 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1680
            TabIndex        =   22
            Top             =   1080
            Visible         =   0   'False
            Width           =   6975
         End
         Begin VB.Label lblBarraRes 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1680
            TabIndex        =   21
            Top             =   480
            Visible         =   0   'False
            Width           =   6975
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Bloqueada"
            Height          =   240
            Left            =   240
            TabIndex        =   20
            Top             =   1680
            Width           =   1005
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Ocupadas"
            Height          =   240
            Left            =   240
            TabIndex        =   19
            Top             =   1080
            Width           =   945
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Reservadas"
            Height          =   240
            Left            =   240
            TabIndex        =   18
            Top             =   480
            Width           =   1125
         End
         Begin VB.Label lblPorOcu 
            AutoSize        =   -1  'True
            Caption         =   "%"
            Height          =   240
            Left            =   8880
            TabIndex        =   17
            Top             =   1080
            Width           =   180
         End
         Begin VB.Label lblPorBloq 
            AutoSize        =   -1  'True
            Caption         =   "%"
            Height          =   240
            Left            =   8880
            TabIndex        =   16
            Top             =   1680
            Width           =   180
         End
         Begin VB.Label lblPorRes 
            AutoSize        =   -1  'True
            Caption         =   "%"
            Height          =   240
            Left            =   8880
            TabIndex        =   15
            Top             =   480
            Width           =   180
         End
         Begin VB.Label lblBarraOcupacion 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1680
            TabIndex        =   14
            Top             =   3000
            Visible         =   0   'False
            Width           =   6975
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Ocupación"
            Height          =   240
            Left            =   240
            TabIndex        =   13
            Top             =   3000
            Width           =   975
         End
         Begin VB.Line Line3 
            X1              =   5160
            X2              =   5160
            Y1              =   2880
            Y2              =   3360
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Vacío"
            Height          =   240
            Left            =   1440
            TabIndex        =   12
            Top             =   2640
            Width           =   525
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Lleno"
            Height          =   240
            Left            =   8400
            TabIndex        =   11
            Top             =   2640
            Width           =   495
         End
         Begin VB.Label Label2 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1680
            TabIndex        =   25
            Top             =   3000
            Width           =   6975
         End
         Begin VB.Label Label3 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1680
            TabIndex        =   26
            Top             =   1680
            Width           =   6975
         End
         Begin VB.Label Label6 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1680
            TabIndex        =   27
            Top             =   1080
            Width           =   6975
         End
         Begin VB.Label Label7 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1680
            TabIndex        =   28
            Top             =   480
            Width           =   6975
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Resumen por situación "
         Height          =   1935
         Left            =   -74760
         TabIndex        =   9
         Top             =   4440
         Width           =   9615
         Begin VB.Label lblBarraSucias 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1680
            TabIndex        =   36
            Top             =   1200
            Visible         =   0   'False
            Width           =   6975
         End
         Begin VB.Label lblBarraLimpias 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1680
            TabIndex        =   35
            Top             =   600
            Visible         =   0   'False
            Width           =   6975
         End
         Begin VB.Image Image2 
            Height          =   480
            Left            =   120
            Picture         =   "frmConsultaCompleta.frx":0D0D
            Top             =   1080
            Width           =   480
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   120
            Picture         =   "frmConsultaCompleta.frx":1F7F
            Top             =   360
            Width           =   480
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Sucias"
            Height          =   240
            Left            =   720
            TabIndex        =   34
            Top             =   1200
            Width           =   615
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Limpias"
            Height          =   240
            Left            =   720
            TabIndex        =   33
            Top             =   600
            Width           =   705
         End
         Begin VB.Label lblPorSucias 
            AutoSize        =   -1  'True
            Caption         =   "%"
            Height          =   240
            Left            =   8880
            TabIndex        =   32
            Top             =   1200
            Width           =   180
         End
         Begin VB.Label lblPorLimpias 
            AutoSize        =   -1  'True
            Caption         =   "%"
            Height          =   240
            Left            =   8880
            TabIndex        =   31
            Top             =   600
            Width           =   180
         End
         Begin VB.Label Label9 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1680
            TabIndex        =   30
            Top             =   600
            Width           =   6975
         End
         Begin VB.Label Label8 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1680
            TabIndex        =   29
            Top             =   1200
            Width           =   6975
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   6255
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   9735
         Begin VB.CommandButton botProcesar 
            Caption         =   "&Procesar"
            Height          =   375
            Left            =   7920
            TabIndex        =   2
            Top             =   120
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
            Left            =   2160
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   120
            Width           =   1815
         End
         Begin MSFlexGridLib.MSFlexGrid gListado 
            Height          =   5535
            Left            =   120
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   600
            Width           =   9015
            _ExtentX        =   15901
            _ExtentY        =   9763
            _Version        =   393216
            FocusRect       =   2
            HighLight       =   2
            AllowUserResizing=   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "&Tipo de Habitación"
            Height          =   240
            Left            =   120
            TabIndex        =   0
            Top             =   180
            Width           =   1725
         End
      End
   End
   Begin Hotel_Nana.gaHOTELbarra gaHOTELbarra1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   7350
      Width           =   10365
      _ExtentX        =   18283
      _ExtentY        =   582
   End
   Begin VB.Menu mnuFormulario 
      Caption         =   "&Formulario"
      Begin VB.Menu mnuFormularioProcesar 
         Caption         =   "Procesar"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuFormularioAceptar 
         Caption         =   "Salir"
         Shortcut        =   {F12}
      End
      Begin VB.Menu div1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFormularioImprimir 
         Caption         =   "Imprimir"
         Shortcut        =   ^I
      End
   End
   Begin VB.Menu mnuVer 
      Caption         =   "&Ver información de ..."
      Begin VB.Menu mnuVerResumen 
         Caption         =   "Resumen"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuVerDetalle 
         Caption         =   "Detalle"
         Shortcut        =   {F6}
      End
   End
   Begin VB.Menu mnuOrden 
      Caption         =   "&Ordenado por ..."
      Begin VB.Menu mnuOrdenNro 
         Caption         =   "Por número de habitación"
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu mnuOrdenTipo 
         Caption         =   "Por tipo de habitación"
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu mnuOrdenEstado 
         Caption         =   "Por estado de habitación"
         Shortcut        =   ^{F3}
      End
      Begin VB.Menu mnuOrdenSitu 
         Caption         =   "Por situación"
         Shortcut        =   ^{F4}
      End
   End
End
Attribute VB_Name = "frmConsultaCompleta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Declaración de constantes
Private Const cLargoBarraGrafica As Integer = 6975   'determina el largo de las barras las gráficas

'Declaración de variables utilizadas para realzar la gráfica de estado
Private totReservadas As Integer    'contador de habitaciones reservadas
Private totOcupadas As Integer      'contador de habitaciones ocupadas
Private totBloqueadas As Integer    'contador de habitaciones bloqueadas

'Declaración de variables utilizadas para realizar gráfica de situaciones
Private totLimpias As Integer
Private totSucias As Integer

Private primeraVezResumen As Boolean    'utilizada para optimizar la realización de las gráficas
Private qdf_principal As QueryDef
Private rst_p As Recordset

Private Sub Form_Load()
    'Apariencia formulario
    mSubConfiguroFuentesControlesSistema Me
    'cargo combo de tipo de habitaciones
    cboTipo_habitacion.AddItem ("(Todas)")
    carga_tipo_hab Me.cboTipo_habitacion
    cboTipo_habitacion.ListIndex = 0
        
    'inicializo varible globales
    primeraVezResumen = True    'la primera vez que ingrese al tabs de resumen,realizo la gráfica
    totReservadas = 0
    totOcupadas = 0
    totBloqueadas = 0
    totLimpias = 0
    totSucias = 0
    
    'Inicialmente ordeno por número de habitación
    mnuOrdenNro_Click
End Sub

Private Sub inicializo_consulta(cri1 As String, Optional cri2 As String)
    'Genero un RecordSet con los datos de la consulta
    Dim consulta As String
    consulta = "Select habitaciones.nrohab, " & _
                "tipo_habitaciones.descripcion, " & _
                "habitaciones.tipohab, " & _
                "habitaciones.situacionhab " & _
                "From habitaciones,tipo_habitaciones " & _
                "Where habitaciones.tipohab = tipo_habitaciones.tipohab " & _
                "Order by " & cri1 & cri2
    
    Set qdf_principal = bdHOTEL.CreateQueryDef("")
    qdf_principal.SQL = consulta
    
    Set rst_p = qdf_principal.OpenRecordset(dbOpenSnapshot)
End Sub

Private Sub recorro_registro()
    'Recorro el RecorSet y muestro los datos en la grilla
    Dim habOcupada As Long
    subLimpioGrilla
    rst_p.MoveFirst
    Do While Not rst_p.EOF
        'verifico si la habitación es del tipo seleccionado
        If rst_p("tipohab") = cboTipo_habitacion.ItemData(cboTipo_habitacion.ListIndex) _
        Or cboTipo_habitacion.Text = "(Todas)" Then
            gListado.AddItem creo_linea
            
            'muestro icono de situación
            gListado.col = 4
            gListado.Row = gListado.Rows - 1
            gListado.CellPictureAlignment = 1
            If rst_p("situacionhab") = 1 Then   'limpia
                Set gListado.CellPicture = frmMAIN.ImageList1.ListImages(1).Picture
            End If
            If rst_p("situacionhab") = 2 Then  'sucia
                Set gListado.CellPicture = frmMAIN.ImageList1.ListImages(2).Picture
            End If
            'es importante determinar si la habitación esta ocupada fuera del período,
            'es decir, que no se fue del hotel en la fecha de egreso determinada.
            'Si es así, lo indico en la consulta.
            
            'solo lo verifico para las habitaciones ocupadas
            If gListado.TextMatrix(gListado.Row, 3) = "Ocupada" Then
                'verifico si la ocupación es válida
                habOcupada = CLng(gListado.TextMatrix(gListado.Row, 1))
                If Not mFunDeterminoOcupacionValida(habOcupada) Then
                    'la ocupación no es válida, es decir, no se le realizó el checkoout
                    'a la habitación.
                    'muestro ícono
                    gListado.col = 3
                    gListado.CellPictureAlignment = 7
                    Set gListado.CellPicture = frmMAIN.ImageList1.ListImages(7).Picture
                End If
            End If
        End If
        rst_p.MoveNext
    Loop
End Sub

Private Sub botProcesar_Click()
    'Es utilizado para ejecutra la consulta luego de cambiar
    'el combo de tipo de habitaciones
    recorro_registro
End Sub

Private Sub creo_cabezal_grilla()
    'Creo cabezal de la grilla
    gListado.FormatString = "      |Nro. de habitación|Tipo de habitación|Estado|Situación"
    'Estable los anchos de las columnas uniformemente
    mSubAparienciaGrilla gListado, 700
End Sub

Private Function creo_linea()
    'Cada registro del RecordSet es una linea de la grilla
    Dim linea As String
    Dim estadoHab As String
    Dim situacionHab As String
    'obtengo el estado de la habitación
    estadoHab = mFunObtengoEstadoHab(rst_p("nrohab"))
    situacionHab = mFunObtengoSituacionHab(rst_p("situacionhab"))
    linea = Chr(9) & _
            rst_p("nrohab") & _
            Chr(9) & rst_p("descripcion") & _
            Chr(9) & estadoHab & _
            Chr(9) & "         " & situacionHab
    creo_linea = linea
    'obtengo información para gráfica
    
    'verifico que la información procesada corresponda a todas las habitaciones del hotel
    'por defecto ejecuto esta conulta (todas las habitaciones) al iniciar la consulta
    If Me.cboTipo_habitacion.Text = "(Todas)" Then
        'cuento los totales de los estados de las habitaciones
        Select Case estadoHab
            Case "Reservada"
                totReservadas = totReservadas + 1
            Case "Ocupada"
                totOcupadas = totOcupadas + 1
            Case "Bloqueada"
                totBloqueadas = totBloqueadas + 1
        End Select
        'cuento los totales de las situaciones de las habitaciones
        Select Case situacionHab
            Case "Limpia"
                totLimpias = totLimpias + 1
            Case "Sucia"
                totSucias = totSucias + 1
        End Select
    End If
End Function

Private Sub subReOrdenoGrilla(criterio As Byte)
    Select Case criterio
        Case 1  'por número de habitación
            inicializo_consulta "habitaciones.nrohab"
            recorro_registro
            
        Case 2  'por tipo+número
            inicializo_consulta "habitaciones.tipohab", ",habitaciones.nrohab"
            recorro_registro
            
        Case 3  'por estado
            'ordeno utilizando métodos de grilla
            'ya que el campo estado no esta disponible en el recordeset.
            'No esta disponible porque este dato se obtiene mediante una función
            'que devuelve el estado actual de la habitación.
            
            ordeno_por_short
            
        Case 4  'por situacion
            inicializo_consulta "habitaciones.situacionhab", ",habitaciones.nrohab"
            recorro_registro
    End Select
End Sub

Private Sub subLimpioGrilla()
    limpio_grilla gListado
    creo_cabezal_grilla
End Sub

Private Sub ordeno_por_short()
    gListado.col = 3                'ordeno por estado
    gListado.Row = gListado.RowSel  'ordeno todas las filas
    gListado.Sort = 5
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Descargo formulario de memoria
    Set frmConsultaCompleta = Nothing
End Sub

Private Sub mnuOrdenNro_Click()
    'Ordeno por número de habitacion
    subReOrdenoGrilla 1
    mSubMuestroIcono gListado, 1
    subDesmarcoTodas
    mnuOrdenNro.Checked = True
End Sub

Private Sub mnuOrdenTipo_Click()
    'Ordeno por tipo
    subReOrdenoGrilla 2
    mSubMuestroIcono gListado, 2
    subDesmarcoTodas
    mnuOrdenTipo.Checked = True
End Sub

Private Sub mnuOrdenEstado_Click()
    'Ordeno por estado
    subReOrdenoGrilla 3
    mSubMuestroIcono gListado, 3
    subDesmarcoTodas
    mnuOrdenEstado.Checked = True
End Sub

Private Sub mnuOrdenSitu_Click()
    'Ordeno por situación
    subReOrdenoGrilla 4
    mSubMuestroIcono gListado, 4
    subDesmarcoTodas
    mnuOrdenSitu.Checked = True
End Sub

Private Sub subDesmarcoTodas()
    'Desmarco todas las opciones del menu, de esta manera solo
    'se puede ver marcada la opción por la cual esta ordenada la grilla en ese momento.
    mnuOrdenNro.Checked = False
    mnuOrdenTipo.Checked = False
    mnuOrdenEstado.Checked = False
    mnuOrdenSitu.Checked = False
End Sub

Private Sub botSalir_Click()
    Unload Me
End Sub

Private Sub mnuFormularioAceptar_Click()
    'Esta opción equivale a presionar el boton de aceptar o la tecla F12
    botSalir_Click
End Sub

Private Sub mnuVerDetalle_Click()
    'Equivale a presionar el boton de F5
    Me.SSTab1.Tab = 1   'tabs de detalle
End Sub

Private Sub mnuVerResumen_Click()
    'Equivale a presionar el boton de F5
    Me.SSTab1.Tab = 0   'tabs de resumen
End Sub

Private Sub mnuFormularioProcesar_Click()
    'Equivale a presionar el boton de procesar
    botProcesar_Click
End Sub

'****************************************************************
'
'   Graficas
'
'*****************************************************************

Private Sub SSTab1_Click(PreviousTab As Integer)
    
    Select Case SSTab1.Tab
        Case 1  'tabs de resumen
            'solo realizo las gráficas la primera vez que ingreso al tabs
            If primeraVezResumen Then
                primeraVezResumen = False
                subRealizoGraficaEstado
                subRealizoGraficaSituacion
            End If
    End Select
End Sub

Private Sub subRealizoGraficaSituacion()
    'Realizao la gráfca de situaciones,con la información obtenida al generar la consulta
    Dim porLimpias As Single
    Dim porSucias As Single
    Dim totHabHotel As Integer  'total de habitaciones del hotel
    'asigno el total de habitaciones del hotel a cada celda
    totHabHotel = mFunObtengoTotHabHotel
    
    'calculo porcentajes
    porLimpias = funCalculoPorcentaje(totHabHotel, totLimpias)
    porSucias = funCalculoPorcentaje(totHabHotel, totSucias)
    
    'muestro barra de limpias
    subDibujoBarra Me.lblBarraLimpias, porLimpias, totLimpias
    'muestro barra de súcias
    subDibujoBarra Me.lblBarraSucias, porSucias, totSucias
    'Pinto lineas de barras
    lblBarraLimpias.BackColor = const_color_limpias
    lblBarraSucias.BackColor = const_color_sucias
    'muestro porcentaje al final de la gráfica
    lblPorLimpias.Caption = Format(porLimpias, "#0.#") & " %"
    lblPorSucias.Caption = Format(porSucias, "#0.#") & " %"
    'modifico el color del texto de las barras y aque el balnco no se ve bien sobre un fondo claro
    lblBarraLimpias.ForeColor = &H80000012  'negro
    lblBarraSucias.ForeColor = &H80000012   'negro
    
End Sub

Private Sub subRealizoGraficaEstado()
    'Realizo la gráfica de estados, con la información obtenida al generar la consulta
    Dim porRes As Single
    Dim porOcu As Single
    Dim porBloq As Single
    Dim porNoDisponibles As Single
    
    Dim totHabHotel As Integer  'total de habitaciones del hotel
    Dim totHabHotelNoDisponibles As Integer 'total de habitaciones del hotel no disponibles
    
    'asigno el total de habitaciones del hotel a cada celda
    totHabHotel = mFunObtengoTotHabHotel
    
    'calculo procentajes para cada estado
    porRes = funCalculoPorcentaje(totHabHotel, totReservadas)
    porOcu = funCalculoPorcentaje(totHabHotel, totOcupadas)
    porBloq = funCalculoPorcentaje(totHabHotel, totBloqueadas)
    
    'muestro barra de reservadas
    subDibujoBarra Me.lblBarraRes, porRes, totReservadas
    'muestro barra de ocupadas
    subDibujoBarra Me.lblBarraOcu, porOcu, totOcupadas
    'muestro barra de bloqueadas
    subDibujoBarra Me.lblBarraBloq, porBloq, totBloqueadas
        
    'Pinto lineas de barras
    lblBarraRes.BackColor = const_color_reservada
    lblBarraOcu.BackColor = const_color_ocupada
    lblBarraBloq.BackColor = const_color_bloqueada
    lblBarraOcupacion.BackColor = const_color_ocupacion
    
    'muestro porcentaje al final de la gráfica
    lblPorRes.Caption = Format(porRes, "#0.#") & " %"
    lblPorOcu.Caption = Format(porOcu, "#0.#") & " %"
    lblPorBloq.Caption = Format(porBloq, "#0.#") & " %"
    
    'realizo barra de ocupación del hotel
    totHabHotelNoDisponibles = totReservadas + totOcupadas + totBloqueadas
    porNoDisponibles = funCalculoPorcentaje(totHabHotel, totHabHotelNoDisponibles)
    subDibujoBarra Me.lblBarraOcupacion, porNoDisponibles, totHabHotelNoDisponibles
    lblPorHotelOcupacion.Caption = Format(porNoDisponibles, "#0.#") & " %"
End Sub

Private Function funCalculoPorcentaje(totHabHotel As Integer, totHabEstado As Integer) As Single
    'Calculo el porcentaje para cada tipo de estado, con relación al total de habitaciones
    'del hotel
    '---------------------------------------------------------------------------------------------
    'Parámetros.
    '   Entrada:    [totHabHotel]   Total de habitaciones del hotel
    '               [totHabEstado]  Total de habitaciones para cada estado
    '   Salida      porcentaje de habitaciones del hotel que tiene asignado un estado determinado
    '----------------------------------------------------------------------------------------------
    'Realizo regla de tres
    Dim porcentaje As Single
    porcentaje = totHabEstado
    porcentaje = porcentaje * 100
    porcentaje = porcentaje / totHabHotel
    funCalculoPorcentaje = porcentaje
End Function

Private Sub subDibujoBarra(barra As Label, porcentaje As Single, cantHab As Integer)
    'Muestro la barra de un largo determinado
    barra.Visible = True
    barra.Width = (cLargoBarraGrafica * porcentaje) / 100
    barra.Caption = cantHab 'cantidad de habitaciones que conforman el porcentaje
    barra.ForeColor = mConstSisColor_Blanco
    barra.FontBold = True
End Sub

'***************************************************************
'*
'*      Asistencia al usuario
'*
'***************************************************************

Private Sub cboTipo_habitacion_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 203
End Sub

Private Sub botImprimir_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 20
End Sub

Private Sub botSalir_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 1, 1
End Sub

Private Sub botProcesar_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 204
End Sub

Private Sub botProcesar_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub botSalir_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub botImprimir_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub cboTipo_habitacion_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

