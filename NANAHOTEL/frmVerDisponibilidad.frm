VERSION 5.00
Object = "{EB8C3860-8603-11D0-A35D-7062E4000000}#1.1#0"; "VCBND.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmVerDisponibilidad 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cuadro de disponibilidad"
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
   Begin VB.Frame Frame2 
      Caption         =   "Gráfica de disponibilidad  "
      Height          =   2175
      Left            =   120
      TabIndex        =   14
      Top             =   5280
      Width           =   11655
      Begin MSFlexGridLib.MSFlexGrid gGrafica 
         Height          =   1095
         Left            =   4440
         TabIndex        =   15
         Top             =   480
         Visible         =   0   'False
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   1931
         _Version        =   393216
      End
      Begin VB.Label lblPorOcu 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   240
         Left            =   10440
         TabIndex        =   27
         Top             =   870
         Width           =   180
      End
      Begin VB.Label lblPorNoAsig 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   240
         Left            =   10440
         TabIndex        =   26
         Top             =   1590
         Width           =   180
      End
      Begin VB.Label lblPorBloq 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   240
         Left            =   10440
         TabIndex        =   25
         Top             =   1230
         Width           =   180
      End
      Begin VB.Label lblPorRes 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   240
         Left            =   10440
         TabIndex        =   24
         Top             =   510
         Width           =   180
      End
      Begin VB.Label lblBarraNoAsig 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1680
         TabIndex        =   23
         Top             =   1560
         Visible         =   0   'False
         Width           =   8535
      End
      Begin VB.Label lblBarraBloq 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1680
         TabIndex        =   22
         Top             =   1200
         Visible         =   0   'False
         Width           =   8535
      End
      Begin VB.Label lblBarraOcu 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1680
         TabIndex        =   21
         Top             =   840
         Visible         =   0   'False
         Width           =   8535
      End
      Begin VB.Label lblBarraRes 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1680
         TabIndex        =   20
         Top             =   480
         Visible         =   0   'False
         Width           =   8535
      End
      Begin VB.Line Line2 
         X1              =   1680
         X2              =   10200
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Line Line1 
         X1              =   1680
         X2              =   1680
         Y1              =   360
         Y2              =   1920
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "No asignadas"
         Height          =   240
         Left            =   120
         TabIndex        =   19
         Top             =   1680
         Width           =   1275
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Bloqueada"
         Height          =   240
         Left            =   120
         TabIndex        =   18
         Top             =   1275
         Width           =   1005
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Ocupadas"
         Height          =   240
         Left            =   120
         TabIndex        =   17
         Top             =   885
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Reservadas"
         Height          =   240
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   1125
      End
   End
   Begin Hotel_Nana.gaHOTELbarra gaHOTELbarra1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   7605
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   582
   End
   Begin VB.Frame Frame1 
      Height          =   5175
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   11655
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
         Left            =   4680
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton botSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   375
         Left            =   10320
         TabIndex        =   8
         Top             =   680
         Width           =   1215
      End
      Begin VB.CommandButton botImprimir 
         Height          =   375
         Left            =   9000
         Picture         =   "frmVerDisponibilidad.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         Tag             =   "Imprimir"
         Top             =   680
         Width           =   1215
      End
      Begin VB.CommandButton botProcesar 
         Caption         =   "&Procesar"
         Height          =   375
         Left            =   7680
         TabIndex        =   6
         Tag             =   "Procesar"
         Top             =   680
         Width           =   1215
      End
      Begin MSFlexGridLib.MSFlexGrid gDisponibilidad 
         Height          =   3735
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   6588
         _Version        =   393216
         TextStyle       =   3
         FocusRect       =   2
         HighLight       =   0
      End
      Begin VcBndCtl.VcCalCombo desde 
         Height          =   375
         Left            =   840
         TabIndex        =   1
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _0              =   $"frmVerDisponibilidad.frx":0942
         _1              =   $"frmVerDisponibilidad.frx":0D4B
         _2              =   $"frmVerDisponibilidad.frx":1154
         _3              =   "-E@@@A@@@@@@@@@'@@@E@@@A@@@@@@@@@)h@@@E@@@A@@@@@@@@@,4E7F"
         _count          =   4
         _ver            =   2
      End
      Begin VcBndCtl.VcCalCombo hasta 
         Height          =   375
         Left            =   840
         TabIndex        =   3
         Top             =   675
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _0              =   $"frmVerDisponibilidad.frx":155D
         _1              =   $"frmVerDisponibilidad.frx":1966
         _2              =   $"frmVerDisponibilidad.frx":1D6F
         _3              =   "-@@E@@@A@@@@@@@@@'@@@E@@@A@@@@@@@@@)h@@@E@@@A@@@@@@@@@,4E7F"
         _count          =   4
         _ver            =   2
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "T&otal de habitaciones disponibles"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   1100
         Width           =   2355
      End
      Begin ComctlLib.ImageList ImageList1 
         Left            =   6960
         Top             =   240
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   18
         ImageHeight     =   18
         MaskColor       =   12632256
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   1
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmVerDisponibilidad.frx":2178
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "&Desde"
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   300
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "&Hasta"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   735
         Width           =   420
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "&Tipo de Habitación"
         Height          =   195
         Left            =   2640
         TabIndex        =   4
         Top             =   300
         Width           =   1935
      End
      Begin VB.Label lmovimiento 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   10080
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Menu mnuFormulario 
      Caption         =   "&Formulario"
      Begin VB.Menu mnuFormularioAceptar 
         Caption         =   "Aceptar"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuDiv1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFormularioProcesar 
         Caption         =   "Procesar"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuFormularioImprimir 
         Caption         =   "Imprimir"
         Shortcut        =   ^I
      End
   End
End
Attribute VB_Name = "frmVerDisponibilidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cant_dias As Long
Private Const ColorTodasOcupadas = &HFF&    'rojo
Private Const constLargoBarra = 8535

Private totHabHotel As Integer  'utilizada para realizar la gráfica

'Variables de configuración
Private color1Semana As String
Private color2Semana As String
Private colorAño As String
Private colorMes As String
Private IluminacionMes As Boolean
Private IluminacionAño As Boolean
Private IluminacionSemanal As Boolean
Private muestroIconoOcupada As Boolean
Private LargoCelda As Integer
Private AnchoCelda As Integer
Private tamañoFuenteCaracteres As Byte
Private tamañoFuenteDigitos As Byte
Private alinIcono As Byte
Private alinFuente As Byte

'Variables del recordset
Private qdf_principal As QueryDef
Private rst_principal As Recordset

Private Sub botImprimir_Click()
    'Imprimo información
End Sub

Private Sub Form_Load()
    'Apariencia formulario
    mSubConfiguroFuentesControlesSistema Me
    
    'cargo tipo habitación
    cboTipo_habitacion.AddItem ("(Todas)")
    carga_tipo_hab frmVerDisponibilidad.cboTipo_habitacion
    
    cboTipo_habitacion.ListIndex = 0
    
    'cargo fecha de inicio por defecto
    desde.Value = m_FechaSis
    subConfiguroFormulario
    
    'Asigno el total de habitaciones del hotel a cada celda
    totHabHotel = mFunObtengoTotHabHotel
End Sub

Private Sub subConfiguroFormulario()
    'Configuro formulario de acuerdo a los valores preestablecidos
    
    'obtengo ancho de celda
    If mFunPosicionoParaGrabar(2, 1) Then
        AnchoCelda = tbSISTEMA_CONF_FORMULARIOS("1ValorNumerico")
    End If
    
    'obtengo medida del largo de grilla
    If mFunPosicionoParaGrabar(2, 2) Then
        LargoCelda = tbSISTEMA_CONF_FORMULARIOS("1ValorNumerico")
    End If
    
    'cantidad de días
    If mFunPosicionoParaGrabar(2, 3) Then
        hasta.Value = desde.Value + tbSISTEMA_CONF_FORMULARIOS("1ValorNumerico")
    End If
    
    'Configuro iluminación de mes
    If mFunPosicionoParaGrabar(2, 4) Then
        IluminacionMes = CBool(tbSISTEMA_CONF_FORMULARIOS("1Valorbol"))
        colorMes = tbSISTEMA_CONF_FORMULARIOS("1ValorNumerico")
    End If
    
    'Configuro iluminacion año
    If mFunPosicionoParaGrabar(2, 5) Then
        IluminacionAño = CBool(tbSISTEMA_CONF_FORMULARIOS("1Valorbol"))
        colorAño = tbSISTEMA_CONF_FORMULARIOS("1ValorNumerico")
    End If
    
    'configuro iluminacion semanal
    If mFunPosicionoParaGrabar(2, 6) Then
        IluminacionSemanal = CBool(tbSISTEMA_CONF_FORMULARIOS("1Valorbol"))
        color1Semana = tbSISTEMA_CONF_FORMULARIOS("1ValorNumerico")
        color2Semana = tbSISTEMA_CONF_FORMULARIOS("2ValorNumerico")
    End If
    
    'muestro icono ocupada
    If mFunPosicionoParaGrabar(2, 7) Then
        muestroIconoOcupada = CBool(tbSISTEMA_CONF_FORMULARIOS("1ValorBol"))
    End If
    
    'tamaño fuente dígitos
    If mFunPosicionoParaGrabar(2, 8) Then
        tamañoFuenteDigitos = tbSISTEMA_CONF_FORMULARIOS("1ValorNumerico")
    End If
    
    'tamaño fuente caracteres
    If mFunPosicionoParaGrabar(2, 9) Then
        tamañoFuenteCaracteres = tbSISTEMA_CONF_FORMULARIOS("1ValorNumerico")
    End If
    
    'alineación icono
    If mFunPosicionoParaGrabar(2, 10) Then
        alinIcono = funObtengoAlin(tbSISTEMA_CONF_FORMULARIOS("1ValorNumerico"))
    End If

    'alineación fuente
    If mFunPosicionoParaGrabar(2, 11) Then
        alinFuente = funObtengoAlin(tbSISTEMA_CONF_FORMULARIOS("1ValorNumerico"))
    End If
End Sub

Private Function funObtengoAlin(tipo As Byte)
    'Devuelbo la constante de centrado correspondiente al valor que se seleccionó
    'en la pantalla de configuración
    Select Case tipo
        Case 0  'izq
            funObtengoAlin = 1
        Case 1  'cen
            funObtengoAlin = 4
        Case 2  'der
            funObtengoAlin = 7
    End Select
End Function

Private Sub botProcesar_Click()
    'Muestro información en grilla
    If IsDate(desde) And IsDate(hasta) Then
        If desde.Value >= m_FechaSis Then
            If desde.Value <= hasta.Value Then
                subBarraProgreso 1
                subArmoGrilla
                subRealizoGrafica   'la grafica tiene que estar dibujada para poder almacenar los
                                    'datos correspondientes mientras recorro los archivos
                                    'es decir mientras ejecuto subObtengoInformacion
                subBarraProgreso 2
                subObtengoInformación
                subRealizoBarras    'una vez obtenidos los datos realizo la gráfica y la muestro
            Else
                'la fecha hasta no puede ser menor a la fecha desde
                mSubMensaje 3, 3
                desde.SetFocus
            End If
        Else
            'la fecha no puede ser menor a la del día de hoy
            mSubMensaje 3, 2
            desde.SetFocus
        End If
    End If
End Sub

Private Sub subBarraProgreso(valor As Long)

    'Muestro la barra de progreso a medida que voy ejecutando las operaciones
    '9 es el total de operaciones que realizo
    Me.gaHOTELbarra1.Progreso 0, 9, valor
    If valor = 9 Then
        Me.gaHOTELbarra1.ProgresoFin
    End If
End Sub
'*************************************************************************************
'*
'*  Armo grilla, dependiendo de variables de configuración
'*
'*
'*************************************************************************************

Private Sub subArmoGrilla()
    'inicializo grilla
    limpio_grilla gDisponibilidad
    subGeneroColumnas
    subGeneroFilas
    subMuestroColores
    subBarraProgreso 1
End Sub

Private Sub subGeneroColumnas() 'mismo procedimiento que en CuadroHab
    'Genera las columnas de la grilla dependiendo del rango de fechas
    'que se ingrese
    'Inserta en el cabezal de cada columna la fecha correspondiente
    'NOTA: debido a que el largo de las celdas puede ser menor al ocupado por la información
    'de la fecha, es necesario primero dar formato a la celda determinando su largo, para
    'después recien llenar las celdas con la información de la fecha, que de todas maneras
    'aparecerá cortada a los ojos del usuario, pero manteniendo internamente el valor
    'correcto, el cual permite obtener información como principio de mes, de año, etc.
    Dim fecha_aux As String
    Dim cabezal As String
    Dim celda As String
    Dim i As Integer
    Dim cantDias As Integer
    Dim fecha As Date
    
    i = 1
    cabezal = "                         |"  'dejo la primer columna libre
    fecha = desde.Value
    'calculo cantidad de dias (columnas de la grilla)
    cantDias = (hasta.Value - desde.Value) + 1
    Do While i <= cantDias
        'obtengo columnas del mismo ancho
        celda = "                                        "   '40 caracteres
        celda = Mid(celda, 1, LargoCelda)
        cabezal = cabezal & celda & "|"
        fecha_aux = fecha_aux & Format(fecha, "dddd dd mmm") & Chr(9)
        fecha = fecha + 1
        i = i + 1
    Loop
    'elimino el último caracter del string, el cual genera una columna vacía.
    gDisponibilidad.FormatString = Mid(cabezal, 1, Len(cabezal) - 1)
    'selecciono celdas e inicializo con fechas
    marco_celdas_grilla gDisponibilidad, 1, gDisponibilidad.Cols - 1, 0, 0
    gDisponibilidad.Clip = fecha_aux
End Sub

Private Sub subGeneroFilas()
    'Genera las filas de la grilla
    
    tbTIPO_HABITACIONES.MoveFirst
    tbTIPO_HABITACIONES.Index = "i_tipo_hab"
    Do While Not tbTIPO_HABITACIONES.EOF
        'verifico el tipo de habitación
        If tbTIPO_HABITACIONES("tipohab") = cboTipo_habitacion.ItemData(cboTipo_habitacion.ListIndex) _
        Or cboTipo_habitacion.Text = "(Todas)" Then
            gDisponibilidad.AddItem tbTIPO_HABITACIONES("descripcion")
            gDisponibilidad.RowHeight(gDisponibilidad.Rows - 1) = AnchoCelda
            'almaceno el número total de habitación de cada tipo en la fila correspondiente a
            'ese tipo de habitación.
            marco_celdas_grilla gDisponibilidad, 1, gDisponibilidad.Cols - 1, gDisponibilidad.Rows - 1, gDisponibilidad.Rows - 1
            gDisponibilidad.Text = tbTIPO_HABITACIONES("total_por_tipo")
            'configuro tamaño de dígitos
            gDisponibilidad.CellFontSize = tamañoFuenteDigitos
            'configuro alineación de los dígitos
            gDisponibilidad.CellAlignment = alinFuente
        End If
        tbTIPO_HABITACIONES.MoveNext
    Loop
    'al terminar de formar la grilla configuro tamaño de caracteres
    marco_celdas_grilla gDisponibilidad, 0, 0, 0, gDisponibilidad.Rows - 1
    gDisponibilidad.CellFontSize = tamañoFuenteCaracteres
End Sub

Private Sub subMuestroColores()
    'a) Cambia de colores el cabezal de la grilla cada una semana, con
    'el objetivo de facilitar la lectura de la grilla
    'b) Cambia de colores la columna donde comienza cada mes
    'c) Cambia de color la columna donde comienza cada año
    
    Dim cambiocolor As Boolean
    Dim colorSemana As String
    Dim i As Integer
    
    cambiocolor = False
    gDisponibilidad.Row = 0
    i = 1
       
    colorSemana = color1Semana  'inicializo el color de la primera semana
    'recorro la grilla
    Do While i < gDisponibilidad.Cols
        'cambio color de cabezal grilla, para poder identificar claramente
        'la nueva semana
        gDisponibilidad.col = i
        If IluminacionMes Then
            subIluminoMes colorMes
        End If
        If IluminacionAño Then
            subIluminoAño colorAño
        End If
        
        If IluminacionSemanal Then
            'Cambio de color el cabezal de la grilla,
            'intercalando dos colores distintos para cada semana
            gDisponibilidad.Row = 0 'los cambio se efectúan en el cabezal de la grilla
            If Mid(gDisponibilidad.Text, 1, 1) = "D" Then   'si es comienzo de semana
                If cambiocolor Then
                    colorSemana = color1Semana
                    cambiocolor = False
                Else
                    colorSemana = color2Semana
                    cambiocolor = True
                End If
            End If
            gDisponibilidad.CellBackColor = colorSemana
        End If
        i = i + 1
    Loop
End Sub

Private Sub subIluminoMes(colorMes As String)
    'Cambio de color la columna correspondiente al primer día del mes
    'obtengo parte número de la fecha
    Dim fecha As String
    fecha = corto_strMedio(gDisponibilidad.Text, " ")    'esta linea falla gabriel!!!!!
    If CByte(fecha) = 1 Then
        marco_celdas_grilla gDisponibilidad, gDisponibilidad.col, gDisponibilidad.col, 1, 1
        gDisponibilidad.TextMatrix(1, gDisponibilidad.col) = "nuevo mes"
        gDisponibilidad.CellBackColor = colorMes
    End If
End Sub

Private Sub subIluminoAño(colorAño As String)
    'Cambio de color la columna correspondiente al primer día del año
    If corto_strDer(gDisponibilidad.TextMatrix(0, gDisponibilidad.col), " ") = "01 Ene" Then
        marco_celdas_grilla gDisponibilidad, gDisponibilidad.col, gDisponibilidad.col, 1, 1
        gDisponibilidad.TextMatrix(1, gDisponibilidad.col) = "nuevo año"
        gDisponibilidad.CellBackColor = colorAño
    End If
End Sub

'*******************************************************************
'*
'*      Obtengo información y actualizo grilla
'*
'********************************************************************

Private Sub subObtengoInformación()
    'Obtengo reservas, ocupadas, bloqueadas, en el rango de fechas correspondiente
    subObtengoReservas
    subBarraProgreso 3
    subObtengoOcupadas
    subBarraProgreso 4
    subObtengoBloqueadas
    subBarraProgreso 5
End Sub

Private Sub subObtengoReservas()
    'Obtengo las reservas
    'El procedimiento funciona igual que el utilizado en cuadro de situación
    
    tbRESERVAS.Index = "i_res_fhas"
    tbRESERVAS.Seek ">", desde.Value    'como el último día de la reserva no se muestra en
                                        'la grilla no tiene sentido recorrer las reserva
                                        '>= a desde.value
    If Not tbRESERVAS.NoMatch Then
        'Recorro las reservas que esten habilitadas segun sea fechahasta >= desdeIngresada
        Do While Not tbRESERVAS.EOF
            'Discrimino las reservas que ingresan posterior a la fecha hastaIngresada
            If tbRESERVAS("fechaing") <= hasta.Value Then
                tbHAB_RESERVAS.Index = "ihab_reserva"
                tbHAB_RESERVAS.Seek ">=", tbRESERVAS("nroreserva"), 1
                Do While Not tbHAB_RESERVAS.EOF
                    'Recorro las habitaciones de cada reserva seleccionada
                    If Not tbHAB_RESERVAS.NoMatch Then
                        If tbHAB_RESERVAS("nroreserva") = tbRESERVAS("nroreserva") Then
                            If busco_habitaTF(tbHAB_RESERVAS("nrohabitacion")) Then
                                'Si pertenece al tipo
                                If tbHABITACIONES("tipohab") = cboTipo_habitacion.ItemData(cboTipo_habitacion.ListIndex) _
                                Or cboTipo_habitacion.Text = "(Todas)" Then
                                    'Tengo que discriminar de esta manera para
                                    'no mostrar las reservas no show.
                                    'la reserva ingresa hoy
                                    If tbRESERVAS("fechaing") = m_FechaSis Then
                                        If Not funEstaOcupada Then    'muestro
                                            'si la reserva ingresa hoy puede ser que ya lo
                                            'halla hecho por ese motivo no la muestro
                                            'ya que aparecerá como ocupada
                                            subRestoEnGrilla tbHAB_RESERVAS("nrohabitacion"), _
                                            tbRESERVAS("fechaing"), tbRESERVAS("fechaegr"), 1
                                        End If
                                    End If
                                    'es una reserva futura
                                    If tbRESERVAS("fechaing") > m_FechaSis Then
                                       subRestoEnGrilla tbHAB_RESERVAS("nrohabitacion"), _
                                        tbRESERVAS("fechaing"), tbRESERVAS("fechaegr"), 1
                                    End If
                                End If
                            Else
                                'reserva no asignada
                                subRestoEnGrilla 0, tbRESERVAS("fechaing"), _
                                                tbRESERVAS("fechaegr"), 4, tbHAB_RESERVAS("tipohabitacion")
                            End If
                        Else
                            Exit Do
                        End If
                    End If
                    tbHAB_RESERVAS.MoveNext
                Loop
            End If
            tbRESERVAS.MoveNext
        Loop
    End If
End Sub

Private Function funEstaOcupada()   'idem al utilizado en cuadro de habitaciones
    'Determina si la habitacion de la reserva correspondiente ya llegó al hotel
    'Existen diferencias entre utilizar este procedimiento y el procedimeinto
    'que busca si uan habitación esta libre o ocupada.
    'La diferencia radica en que una ocupación puede pasar el límite de la fecha de egreso
    'originando que cuando llegue el día de ocupación de otra reserva, la habitación este
    'ocupada pero no con los pasajeros de la reserva correcta.
    funEstaOcupada = False

    If busco_ReservaHabita_checkin(tbHAB_RESERVAS("nroreserva"), tbHAB_RESERVAS("nrohabitacion")) Then
        'la habitación de la reserva correspondiente ya ingreso al hotel
        funEstaOcupada = True
    End If
End Function

Private Sub subObtengoOcupadas() 'idem al utilizado en cuadro de habitacion
    'Muestro las habitaciones ocupadas
    'Como una habitación puede tener alojados más de un pasajero, tengo que realizar
    'un corte de control.
    Dim habAnt As Long
    Dim fechaDes As Date    'Es necesario declarar estas variables
    Dim fechaHas As Date    'ya que para la última habitación ocupada, el procedimiento
                            'llega a fin de archivo.
    
    tbCHECKIN.Index = "i_habitacion"
    tbCHECKIN.Seek ">=", 0
    'Recorro todas las habitaciones alojadas, si pertenecen al tipo correcto
    'las muestro en la grilla.
    Do While Not tbCHECKIN.EOF
        If busco_habitaTF(tbCHECKIN("nrohab")) Then
            'verifico si la habitación es del tipo a trabajar
            If tbHABITACIONES("tipohab") = cboTipo_habitacion.ItemData(cboTipo_habitacion.ListIndex) _
            Or cboTipo_habitacion.Text = "(Todas)" Then
                'datos de la ocupación
                habAnt = tbCHECKIN("nrohab")
                fechaDes = tbCHECKIN("fcheckdes")
                fechaHas = tbCHECKIN("fcheckhas")
                Do While Not tbCHECKIN.EOF
                    If tbCHECKIN("nrohab") = habAnt Then
                        tbCHECKIN.MoveNext
                    Else
                        Exit Do
                    End If
                Loop
                'termine recorrer todos los pasajeros de la habitación (o fin de archivo)
                'resto 1 en la grilla
                subRestoEnGrilla habAnt, fechaDes, fechaHas, 2
            Else
                'si el tipo no es el correcto paso al proximo registro
                tbCHECKIN.MoveNext
            End If
        Else
            'si la habitación no existe paso al proximo registro
            'esto nunca debería de ocurrir
            tbCHECKIN.MoveNext
        End If
    Loop
End Sub

Private Sub subObtengoBloqueadas()  'idem procedimiento de cuadro de habitaciones
    'Muestra las habitaciones bloqueadas con linea verde
    tbBLOQUEO_HAB.Index = "i_bloq_fh2"
    tbBLOQUEO_HAB.Seek ">=", desde.Value
    'Recorro tbBLOQUEO_HAB discriminado los bloqueos que no esten en fecha
    If Not tbBLOQUEO_HAB.NoMatch Then
        Do While Not tbBLOQUEO_HAB.EOF
            'confirmo que el bloqueo sea para una habitación del tipo seleccionado
            If busco_habitaTF(tbBLOQUEO_HAB("hab_bloq")) Then
                If tbHABITACIONES("tipohab") = cboTipo_habitacion.ItemData(cboTipo_habitacion.ListIndex) _
                Or cboTipo_habitacion.Text = "(Todas)" Then
                    If tbBLOQUEO_HAB("fdesdebloq") <= hasta.Value Then
                       subRestoEnGrilla tbBLOQUEO_HAB("hab_bloq"), _
                        tbBLOQUEO_HAB("fdesdebloq"), tbBLOQUEO_HAB("fhastabloq"), 3
                    End If
                End If
            End If
            tbBLOQUEO_HAB.MoveNext
        Loop
    End If
End Sub

Private Sub subRestoEnGrilla(hab As Long, des As Date, has As Date, tipo As Byte, _
                            Optional tipoHab As Integer)
    'A)Resta 1 a el total de habitaciones del tipo determinado, y en el rango de fechas
    'que determine la reserva, ocupación o bloqueo.
    'B)Como la grilla de la grafica es una copia de la de disponibilidad utilizo este
    'precedimiento para obtener las columnas inicial y final, con la cual trabajar
    'en la grilla de la grafica.
    
    'El parámetro [tipoHab] se utiliza cuando trabajo con reservas no asignadas, ya que
    'para este caso el número de la habitación siempre es 0.
    
    Dim col As Integer
    Dim col2 As Integer
    Dim dibujoLinea As Boolean
    Dim tipoHabAux As Integer
    
    dibujoLinea = True
    'obtengo columna inicial = col
    col = des - desde.Value
    If col < 0 Then
        'en caso de que la información comienze antes de la primer columna,
        'tomo como primer columna la columna 0
        col = 0
    End If
    col = col + 1   'la primera columna esta reservada para mostrar las habitaciones
                    
    'obtengo columna final = col2
    col2 = has - desde.Value
    If col2 <= 0 Then
        'si el valor de col2 es 0,la linea no se dibuja.
        dibujoLinea = False
    End If
    If col2 > gDisponibilidad.Cols - 1 Then
        'en caso de que no tenga suficiente columnas para mostrar toda la información
        'asumo como última columna a la última columna de la grilla
        col2 = gDisponibilidad.Cols - 1
    End If
    'no le sumo 1 a col2 ya que siempre se dibuja una celda menos en relación con
    'la cantidad de días mostrar

    If dibujoLinea Then
        If tipo = 4 Then   'no asignadas
            'ya tengo el tipo de la habitación
            'me paro en la fila correspondiente a la habitación a procesar
            gDisponibilidad.Row = funObtengoFila(tipoHab)
            subMuestroEnCelda col, col2, gDisponibilidad.Row
        Else
            'obtengo el tipo de la habitación
            tipoHabAux = mFunObtengoTipoHab(hab)
            'me paro en la fila correspondiente a la habitación a procesar
            gDisponibilidad.Row = funObtengoFila(tipoHabAux)
            subMuestroEnCelda col, col2, gDisponibilidad.Row
        End If
        
        'la misma información que muestro en la celda de disponobilidad
        'la utilizo para cargar la grilla de la grafica
        subCargoGrillaGrafica col, col2, tipo
    End If
End Sub

Private Sub subMuestroEnCelda(coli As Integer, colf As Integer, fila As Integer)
    'Recorro las celdas correspondientes y disminuyo en 1 el total de habitaciones
    'libres
    'No puedo utilizar el procedimiento de seleccion de visual ya que
    'no funciona, debido a que es necesario tratar todas las celdas del rango
    'individualmente.
    Dim i As Double
    i = coli
    gDisponibilidad.Row = fila
    Do While i <= colf
        gDisponibilidad.col = i
        'al total de habitaciones de ese tipo, en ese día le resto 1
        gDisponibilidad.Text = Val(gDisponibilidad.Text) - 1
        If Val(gDisponibilidad.Text) = 0 Then
            gDisponibilidad.CellForeColor = ColorTodasOcupadas
            'muestro icono
            If muestroIconoOcupada Then
                'no muestro cantidad
                gDisponibilidad.Text = ""
                'alineo icono
                gDisponibilidad.CellPictureAlignment = alinIcono
                Set gDisponibilidad.CellPicture = ImageList1.ListImages(1).Picture
            End If
        End If
        i = i + 1
    Loop
End Sub

Private Function funObtengoFila(tipoHab As Integer) As Integer
    'Recorre la primer columna de la grilla y se posiciona en el tipo correspondiente
    'a la habitación
    '----------------------------------------------------------------------------------
    'Parámetros.
    '   Entrada [tipoHab]   Tipo de la habitación con la cual estoy trabjando.
    
    '   Para poder calcular el total de habitaciones disponibles, es necesario posicionarse
    '   en la grilla, sobre la fila que representa al tipo de habitación.
    '   Cuando estoy trabajando con las habitaciones ocupadas, reservadas y bloquedas, tengo
    '   acceso (así lo condiciona el diseño de la BD) al número de habitación, pero cuando
    '   estoy trabajando con las reservas no asignadas, este valor siempre es 0, por lo que
    '   tengo que recurrir al tipo de la habitación, dato que no tengo disponible para las
    '   ocupadas, reservadas y bloqueadas.
    '   Esto se soluciona obteniendo el tipo de habitación antes de llamar a esta función.
    '
    '   Salida  Número de fila correspondiente al tipo de habitación con la que se está
    '           trabajando.
    '------------------------------------------------------------------------------------
    
    Dim tipoHabDesc As String
    Dim indice As Byte
    'busco la descripción del tipo de habitación a mostrar
    tipoHabDesc = mFun_BuscoDescriTipoHab(CLng(tipoHab))
    indice = 1
    gDisponibilidad.col = 0
    Do While indice < gDisponibilidad.Rows
        gDisponibilidad.Row = indice
        If gDisponibilidad.Text = tipoHabDesc Then
            funObtengoFila = gDisponibilidad.Row
            Exit Do
        End If
        indice = indice + 1
    Loop
End Function

'*************************************************************
'*
'*  Realizo gráfica
'*
'*************************************************************

Private Sub subCargoGrillaGrafica(coli As Integer, colf As Integer, tipo As Byte)
    'Recorro las celdas correspondientes y aumento en 1 el total de habitaciones
    'no disponibles
    Dim i As Long
    i = coli
    gGrafica.Row = tipo + 1 '+1 se debe a que la fila 0 es el cabezal y la 1 se deja vacía
                            'por lo que se cominza a trabajar en la fila2
    Do While i <= colf
        gGrafica.col = i
        gGrafica.Text = Val(gGrafica.Text) + 1
        i = i + 1
    Loop
End Sub

Private Sub subRealizoGrafica()
    'Realizo gráfica con los datos obtenidos
    
    'inicializo grilla
    limpio_grilla gGrafica
    subGeneroColumnasGrafica
    subGeneroFilasGrafica
End Sub

Private Sub subGeneroColumnasGrafica()
    'genero las mismas cantidad de columnas que la gráfica original
    'Genera las columnas de la grilla dependiendo del rango de fechas
    'que se ingrese
    
    Dim cabezal As String
    Dim celda As String
    Dim i As Long
    Dim cantDias As Integer
    
    i = 1
    cabezal = "   |"  'dejo la primer columna libre
    'calculo cantidad de dias (columnas de la grilla)
    cantDias = (hasta.Value - desde.Value) + 1
    Do While i <= cantDias
        cabezal = cabezal & "  |"
        i = i + 1
    Loop
    gGrafica.FormatString = Mid(cabezal, 1, Len(cabezal) - 1)
End Sub

Private Sub subGeneroFilasGrafica()
    'Genero una fila para cada barra de la gráfica
    
    gGrafica.AddItem "1"    'reservadas
    gGrafica.AddItem "2"    'ocupadas
    gGrafica.AddItem "3"    'bloqueadas
    gGrafica.AddItem "4"    'noasignadas
End Sub

Private Sub subRealizoBarras()
    'Con la información obtenida en la grilla genero las líneas de la gráfica
        
    'Pinto lineas de barras
    lblBarraRes.BackColor = const_color_reservada
    lblBarraOcu.BackColor = const_color_ocupada
    lblBarraBloq.BackColor = const_color_bloqueada
    lblBarraNoAsig.BackColor = const_color_noAsignada
    subMuestroBarra 1   'res
    subBarraProgreso 6
    subMuestroBarra 2   'ocu
    subBarraProgreso 7
    subMuestroBarra 3   'bloq
    subBarraProgreso 8
    subMuestroBarra 4   'noasig
    subBarraProgreso 9
End Sub

Private Function subMuestroBarra(tipo As Byte)
    'Calculo el porcentaje de reservas, ocupaciones, bloqueos y no asignadas
    Dim i As Long
    Dim totHabNoDisponibles As Integer
    Dim porcentaje As Single
    
    'recorro grilla para obtener valores
    gGrafica.Row = tipo + 1 'me posiciono en la fila correspondiente al tipo dato
                        ' que deseo obtener 1=res, 2= ocu , 3=bloq, 4 = noasig
    
    i = 1   'empiezo siempre desde la priemer columna
    totHabNoDisponibles = 0
    Do While i < gGrafica.Cols
        gGrafica.col = i
        'en la grilla de grafica almaceno el total de habitaciones no disponibles
        totHabNoDisponibles = totHabNoDisponibles + Val(gGrafica.Text)
        i = i + 1
    Loop
    'totHabHotel            = 100%
    'totHabNoDisponibles    = x
    i = i - 1   'le resto 1 a i porque cuando sale del bucle, tiene
                ' un 1 de más
    
    'Esto es algo interesante:
    'si realizo totHabNoDisponibles * 100 el resultado genera desbordamiento para
    'totHabNodisponibles=367, al parecer 36700, que es el resultado de la operación
    'primero se almacena en esta varibable para después recién pasar a porcentaje,
    'originando un error ya que el tipo integer no soporta este número
    porcentaje = totHabNoDisponibles
    porcentaje = porcentaje * 100
    porcentaje = porcentaje / (totHabHotel * i)
    porcentaje = Format(porcentaje, "##.0")
    
    Select Case tipo
        Case 1  'reservadas
            subDibujoBarra lblBarraRes, porcentaje, totHabNoDisponibles
            lblPorRes.Caption = porcentaje & "%"
        Case 2  'ocupadas
            subDibujoBarra lblBarraOcu, porcentaje, totHabNoDisponibles
            lblPorOcu.Caption = porcentaje & "%"
        Case 3  'bloqueadas
            subDibujoBarra lblBarraBloq, porcentaje, totHabNoDisponibles
            lblPorBloq.Caption = porcentaje & "%"
        Case 4  'noasignadas
            subDibujoBarra lblBarraNoAsig, porcentaje, totHabNoDisponibles
            lblPorNoAsig.Caption = porcentaje & "%"
    End Select
End Function

Private Sub subDibujoBarra(barra As Label, porcentaje As Single, cantHab As Integer)
    'Muestro la barra de un largo determinado
    barra.Visible = True
    barra.Width = (constLargoBarra * porcentaje) / 100
    'barra.Caption = cantHab 'cantidad de habitaciones que conforman el porcentaje
    'a pedido del nano no muestro este dato
    barra.ForeColor = mConstSisColor_Blanco
    barra.FontBold = True
End Sub

Private Sub botSalir_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Descargo formulario de memoria
    Set frmVerDisponibilidad = Nothing
End Sub

Private Sub mnuFormularioAceptar_Click()
    'Equivale a presionar el boton aceptar o la tecla F12
    botProcesar_Click
End Sub

Private Sub mnuFormularioImprimir_Click()
    'Equivale a procesar el boton de imprimir o la tecla Ctrol+I
    If Me.botImprimir.Enabled = True Then
        botImprimir_Click
    End If
End Sub

Private Sub mnuFormularioProcesar_Click()
    'Equivale a presionar el boton de procesar o la tecla F9
    If Me.botProcesar.Enabled = True Then
        botProcesar_Click
    End If
End Sub

'**************************************************
'*
'*  Asistencia a usuarios
'*
'**************************************************

Private Sub desde_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 134
End Sub

Private Sub hasta_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 135
End Sub

Private Sub cboTipo_habitacion_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 136
End Sub

Private Sub botProcesar_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 137
End Sub

Private Sub botImprimir_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 138
End Sub

Private Sub botSalir_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 1, 1
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

Private Sub desde_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub hasta_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub cboTipo_habitacion_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub


