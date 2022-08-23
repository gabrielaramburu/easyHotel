VERSION 5.00
Object = "{EB8C3860-8603-11D0-A35D-7062E4000000}#1.1#0"; "VCBND.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmCuadroHab 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cuadro de situación"
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
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   7605
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   582
   End
   Begin VB.Frame Frame1 
      Height          =   7455
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   11655
      Begin VB.CommandButton botImprimir 
         Height          =   375
         Left            =   9000
         Picture         =   "frmCuadroHab.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         Tag             =   "Imprimir"
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton botSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   375
         Left            =   10320
         TabIndex        =   8
         Top             =   960
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
         Left            =   4680
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton botProcesar 
         Caption         =   "&Procesar"
         Height          =   375
         Left            =   7680
         TabIndex        =   6
         Tag             =   "Procesar"
         Top             =   960
         Width           =   1215
      End
      Begin VcBndCtl.VcCalCombo desde 
         Height          =   375
         Left            =   1200
         TabIndex        =   1
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _0              =   $"frmCuadroHab.frx":0942
         _1              =   $"frmCuadroHab.frx":0D4B
         _2              =   $"frmCuadroHab.frx":1154
         _3              =   "-@@@E@@@A@@@@@@@@@'@@@E@@@A@@@@@@@@@)h@@@E@@@A@@@@@@@@@,4E7F"
         _count          =   4
         _ver            =   2
      End
      Begin MSFlexGridLib.MSFlexGrid gHabitacion 
         Height          =   5895
         Left            =   120
         TabIndex        =   10
         Top             =   1440
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   10398
         _Version        =   393216
         Cols            =   3
         Redraw          =   -1  'True
         AllowBigSelection=   -1  'True
         FocusRect       =   2
         HighLight       =   0
         GridLines       =   0
         MousePointer    =   2
      End
      Begin VcBndCtl.VcCalCombo hasta 
         Height          =   375
         Left            =   1200
         TabIndex        =   3
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _0              =   $"frmCuadroHab.frx":155E
         _1              =   $"frmCuadroHab.frx":1967
         _2              =   $"frmCuadroHab.frx":1D70
         _3              =   "-@@@E@@@A@@@@@@@@@'@@@E@@@A@@@@@@@@@)h@@@E@@@A@@@@@@@@@,4E7F"
         _count          =   4
         _ver            =   2
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "&Grilla de situación"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   1245
      End
      Begin ComctlLib.ImageList ImageList1 
         Left            =   3360
         Top             =   600
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   12
         ImageHeight     =   10
         MaskColor       =   12632256
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   8
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmCuadroHab.frx":217A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmCuadroHab.frx":2334
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmCuadroHab.frx":24EE
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmCuadroHab.frx":26A8
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmCuadroHab.frx":2862
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmCuadroHab.frx":2A1C
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmCuadroHab.frx":2BD6
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmCuadroHab.frx":2D90
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "&Tipo de Habitación"
         Height          =   240
         Left            =   2880
         TabIndex        =   4
         Top             =   300
         Width           =   1725
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "&Hasta"
         Height          =   240
         Left            =   120
         TabIndex        =   2
         Top             =   780
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&Desde"
         Height          =   240
         Left            =   120
         TabIndex        =   0
         Top             =   300
         Width           =   615
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
      Begin VB.Menu mnuProcesarConsulta 
         Caption         =   "Procesar "
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuFormularioImprimir 
         Caption         =   "Imprimir"
         Shortcut        =   ^I
      End
   End
   Begin VB.Menu mnuVer 
      Caption         =   "&Ir a..."
      Begin VB.Menu mnuVerDisponibilidad 
         Caption         =   "Ver disponibilidad..."
         Shortcut        =   ^D
      End
   End
End
Attribute VB_Name = "frmCuadroHab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cant_dias As Long

'Variables de configuración
Private PosicionNoAsignadas As Byte
Private color1Semana As String
Private color2Semana As String
Private colorAño As String
Private colorMes As String
Private IluminacionMes As Boolean
Private IluminacionAño As Boolean
Private IluminacionSemanal As Boolean
Private LargoCelda As Integer
Private AnchoCelda As Integer
Private PpioFinLinea As Boolean

Private Sub botImprimir_Click()
    'Imprimo información
End Sub

Private Sub Form_Load()
    'Apariencia formulario
    mSubConfiguroFuentesControlesSistema Me
    
    'cargo tipo habitación
    cboTipo_habitacion.AddItem ("(Todas)")
    carga_tipo_hab frmCuadroHab.cboTipo_habitacion
    
    cboTipo_habitacion.ListIndex = 0
    
    'cargo fecha de inicio por defecto
    desde.Value = m_FechaSis
    subConfiguroFormulario
End Sub

Private Sub subBarraProgreso(valor As Long)
    'Muestro la barra de progreso a medida que voy ejecutando las operaciones
    '5 es el total de operaciones que realizo
    Me.gaHOTELbarra1.Progreso 0, 4, valor
    If valor = 4 Then
        Me.gaHOTELbarra1.ProgresoFin
    End If
End Sub

Private Sub subConfiguroFormulario()
    'Configuro formulario de acuerdo a los valores preestablecidos
    
    'obtengo ancho de celda
    If mFunPosicionoParaGrabar(1, 1) Then
        AnchoCelda = tbSISTEMA_CONF_FORMULARIOS("1ValorNumerico")
    End If
    
    'obtengo medida del largo de grilla
    If mFunPosicionoParaGrabar(1, 2) Then
        LargoCelda = tbSISTEMA_CONF_FORMULARIOS("1ValorNumerico")
    End If
    
    'obtengo posición no asignadas
    If mFunPosicionoParaGrabar(1, 3) Then
        PosicionNoAsignadas = tbSISTEMA_CONF_FORMULARIOS("1ValorNumerico")
    End If
    
    'Configuro líneas divisorias
    If mFunPosicionoParaGrabar(1, 4) Then   'mostrar líneas divisorias
        gHabitacion.GridLines = tbSISTEMA_CONF_FORMULARIOS("1ValorBol")
    End If
    
    'Configuro fecha hasta
    If mFunPosicionoParaGrabar(1, 5) Then
        hasta.Value = desde.Value + tbSISTEMA_CONF_FORMULARIOS("1ValorNumerico")
    End If
    
    'Configuro iluminación de mes
    If mFunPosicionoParaGrabar(1, 6) Then
        IluminacionMes = CBool(tbSISTEMA_CONF_FORMULARIOS("1Valorbol"))
        colorMes = tbSISTEMA_CONF_FORMULARIOS("1ValorNumerico")
    End If
    
    'Configuro iluminacion año
    If mFunPosicionoParaGrabar(1, 7) Then
        IluminacionAño = CBool(tbSISTEMA_CONF_FORMULARIOS("1Valorbol"))
        colorAño = tbSISTEMA_CONF_FORMULARIOS("1ValorNumerico")
    End If
        
    'configuro iluminacion semanal
    If mFunPosicionoParaGrabar(1, 8) Then
        IluminacionSemanal = CBool(tbSISTEMA_CONF_FORMULARIOS("1Valorbol"))
        color1Semana = tbSISTEMA_CONF_FORMULARIOS("1ValorNumerico")
        color2Semana = tbSISTEMA_CONF_FORMULARIOS("2ValorNumerico")
    End If
    
    'configuro imagen en ppio y fin de linea
    If mFunPosicionoParaGrabar(1, 9) Then
        PpioFinLinea = CBool(tbSISTEMA_CONF_FORMULARIOS("1Valorbol"))
    End If
    
End Sub

Private Sub botProcesar_Click()
    'Muestro información en grilla
    If IsDate(desde) And IsDate(hasta) Then
        If desde.Value >= m_FechaSis Then
            If desde.Value <= hasta.Value Then
                subBarraProgreso 1
                subArmoGrilla
                subMuestroDiasReservados
                subBarraProgreso 2
                subMuestroDiasOcupados
                subBarraProgreso 3
                subMuestroDiasBloqueados
                subBarraProgreso 4
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

'*************************************************************************************
'*
'*  Armo grilla, dependiendo de variables de configuración
'*
'*
'*************************************************************************************

Private Sub subArmoGrilla()
    'inicializo grilla
    limpio_grilla gHabitacion
    subGeneroColumnas
    subGeneroFilas
    subMuestroColores
End Sub

Private Sub subGeneroColumnas()
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
    cabezal = "         |"  'dejo la primer columna libre
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
    gHabitacion.FormatString = Mid(cabezal, 1, Len(cabezal) - 1)
    'selecciono celdas y inicializo con fechas
    marco_celdas_grilla gHabitacion, 1, gHabitacion.Cols - 1, 0, 0
    gHabitacion.Clip = fecha_aux
End Sub

Private Sub subGeneroFilas()
    'Genera las filas de la grilla
    'Realiza un corte de control del archivo tbHABITACIONES por tipo,
    'cuando se cambia de habitación se genera una linea divisoria para facilitar
    'la lectura de la grilla
    Dim muestro_linea As Boolean
    Dim tipo_ant As Integer
    
    gHabitacion.col = 0
    tbHABITACIONES.MoveFirst
    tbHABITACIONES.Index = "i_tipohab"
    Do While Not tbHABITACIONES.EOF
        muestro_linea = False
        tipo_ant = tbHABITACIONES("tipohab")
        Do While Not tbHABITACIONES.EOF
            If tipo_ant = tbHABITACIONES("tipohab") Then
                If tbHABITACIONES("tipohab") = cboTipo_habitacion.ItemData(cboTipo_habitacion.ListIndex) _
                Or cboTipo_habitacion.Text = "(Todas)" Then
                    gHabitacion.AddItem (tbHABITACIONES("nrohab"))
                    gHabitacion.RowHeight(gHabitacion.Rows - 1) = AnchoCelda
                    gHabitacion.Row = gHabitacion.Rows - 1
                    gHabitacion.CellFontSize = 12
                    muestro_linea = True
                End If
            Else
                Exit Do
            End If
            tbHABITACIONES.MoveNext
        Loop
        'muestro línea divisoria
        If muestro_linea Then
            gHabitacion.AddItem ("")
            gHabitacion.RowHeight(gHabitacion.Rows - 1) = 50
        End If
    Loop
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
    gHabitacion.Row = 0
    i = 1
       
    colorSemana = color1Semana  'inicializo el color de la primera semana
    'recorro la grilla
    Do While i < gHabitacion.Cols
        'cambio color de cabezal grilla, para poder identificar claramente
        'la nueva semana
        gHabitacion.col = i
        If IluminacionMes Then
            subIluminoMes colorMes
        End If
        If IluminacionAño Then
            subIluminoAño colorAño
        End If
        
        If IluminacionSemanal Then
            'Cambio de color el cabezal de la grilla,
            'intercalando dos colores distintos para cada semana
            gHabitacion.Row = 0 'los cambio se efectúan en el cabezal de la grilla
            If Mid(gHabitacion.Text, 1, 1) = "D" Then   'si es comienzo de semana
                If cambiocolor Then
                    colorSemana = color1Semana
                    cambiocolor = False
                Else
                    colorSemana = color2Semana
                    cambiocolor = True
                End If
            End If
            gHabitacion.CellBackColor = colorSemana
        End If
        i = i + 1
    Loop
End Sub

Private Sub subIluminoMes(colorMes As String)
    'Cambio de color la columna correspondiente al primer día del mes
    'obtengo parte número de la fecha
    Dim fecha As String
    fecha = corto_strMedio(gHabitacion.Text, " ")    'esta linea falla gabriel!!!!!
    If CByte(fecha) = 1 Then
        marco_celdas_grilla gHabitacion, gHabitacion.col, gHabitacion.col, 1, 1
        gHabitacion.TextMatrix(1, gHabitacion.col) = "nuevo mes"
        gHabitacion.CellBackColor = colorMes
    End If
End Sub

Private Sub subIluminoAño(colorAño As String)
    'Cambio de color la columna correspondiente al primer día del año
    If corto_strDer(gHabitacion.TextMatrix(0, gHabitacion.col), " ") = "01 Ene" Then
        marco_celdas_grilla gHabitacion, gHabitacion.col, gHabitacion.col, 1, 1
        gHabitacion.TextMatrix(1, gHabitacion.col) = "nuevo año"
        gHabitacion.CellBackColor = colorAño
    End If
End Sub

'***************************************************************************************
'*
'*      Proceso información y muestro
'*
'***************************************************************************************
Private Sub subMuestroDiasOcupados()
    'Muestro las habitaciones ocupadas con linea azul
    'Como una habitación puede tener alojados más de un pasajero, tengo que realizar
    'un corte de control.
    Dim habAnt As Integer
    Dim fechaDesAnt As Date    'Es necesario declarar estas variables
    Dim fechaHasAnt As Date    'ya que para la última habitación ocupada, el procedimiento
    Dim resAnt As Long      'llega a fin de archivo.
    
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
                fechaDesAnt = tbCHECKIN("fcheckdes")
                fechaHasAnt = tbCHECKIN("fcheckhas")
                resAnt = tbCHECKIN("nroreserva")
                Do While Not tbCHECKIN.EOF
                    If tbCHECKIN("nrohab") = habAnt Then
                        tbCHECKIN.MoveNext
                    Else
                        Exit Do
                    End If
                Loop
                'termine recorrer todos los pasajeros de la habitación (o fin de archivo)
                subDibujoLineaEnGrilla habAnt, fechaDesAnt, fechaHasAnt, _
                resAnt, const_color_ocupada
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

Private Sub subMuestroDiasReservados()
    'Muestro las habitaciones reservadas con linea roja
    Dim color As Long
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
                            'determino color linea
                            color = const_color_reservada
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
                                            'halla echo por ese motivo no la muestro
                                            'ya que aparecerá como linea ocupada
                                            subDibujoLineaEnGrilla tbHAB_RESERVAS("nrohabitacion"), _
                                            tbRESERVAS("fechaing"), tbRESERVAS("fechaegr"), _
                                            tbRESERVAS("nroreserva"), color
                                        End If
                                    End If
                                    'es una reserva futura
                                    If tbRESERVAS("fechaing") > m_FechaSis Then
                                       subDibujoLineaEnGrilla tbHAB_RESERVAS("nrohabitacion"), _
                                        tbRESERVAS("fechaing"), tbRESERVAS("fechaegr"), _
                                        tbRESERVAS("nroreserva"), color
                                    End If
                                End If
                            Else    'reserva no asignada
                                color = const_color_noAsignada
                                'valido tipo de habitación
                                If tbHAB_RESERVAS("tipohabitacion") = cboTipo_habitacion.ItemData(cboTipo_habitacion.ListIndex) _
                                Or cboTipo_habitacion.Text = "(Todas)" Then
                                    'no es necesario discriminar si la reserva ingresa hoy o
                                    'no, ya que realizo lo mismo en ambos casos.
                                    subDibujoLineaEnGrilla 0, _
                                    tbRESERVAS("fechaing"), tbRESERVAS("fechaegr"), _
                                    tbRESERVAS("nroreserva"), color, True, tbHAB_RESERVAS("nrocorr")
                                End If
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

Private Function funEstaOcupada()
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

Private Sub subMuestroDiasBloqueados()
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
                       subDibujoLineaEnGrilla tbBLOQUEO_HAB("hab_bloq"), _
                        tbBLOQUEO_HAB("fdesdebloq"), tbBLOQUEO_HAB("fhastabloq"), tbBLOQUEO_HAB("nrocorr_bloq"), const_color_bloqueada
                    End If
                End If
            End If
            tbBLOQUEO_HAB.MoveNext
        Loop
    End If
End Sub

Private Sub subDibujoLineaEnGrilla(hab As Integer, des As Date, has As Date, _
                                   iden As Long, color As Long, Optional noAsignada As Boolean, _
                                   Optional corrHab As Long)
    'Genera linea en la grilla, toma en cuenta la fecha desde y hasta y el tipo.
    'Roja= Reserva,    Azul=Ocupada,    Verde=Bloqueada
    
    Dim col As Integer
    Dim col2 As Integer
    Dim dibujoLinea As Boolean
    Dim muestroIconoFinLinea As Boolean
    
    dibujoLinea = True
    muestroIconoFinLinea = True
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
    If col2 > gHabitacion.Cols - 1 Then
        'en caso de que no tenga suficiente columnas para mostrar toda la información
        'asumo como última columna a la última columna de la grilla
        col2 = gHabitacion.Cols - 1
        muestroIconoFinLinea = False
    End If
    'no le sumo 1 a col2 ya que siempre se dibuja una celda menos en relación con
    'a la cantidad de días mostrar

    If dibujoLinea Then
        If noAsignada = True Then
             'creo una nueva linea en la grilla
            gHabitacion.Row = funObtengoFilaNoAsignada(corrHab)
        Else
            'me paro en la fila correspondiente a la habitación a procesar
            gHabitacion.Row = funObtengoFila(hab)
        End If
        'selecciono celdas en la fila determinada y en el rango de columna calculado
        marco_celdas_grilla gHabitacion, col, col2, gHabitacion.Row, gHabitacion.Row
        'pinto celdas
        gHabitacion.CellBackColor = color
        gHabitacion.CellForeColor = color
        gHabitacion.Text = funAsignoIdentificacionCeldas(iden, color)
        'el primer día muestro información
        subMuestroPrimerDia col, color
        'el último día también muestro información
        subMuestroUltimoDia col2, color, muestroIconoFinLinea
    End If
End Sub

Private Sub subMuestroPrimerDia(col As Integer, color As Long)
    'Muestro información en el primer día de la línea dibujada
    Dim indiceImagen As Byte
    
    gHabitacion.col = col   'me posiciono en grilla
    
    'muestro número de reserva
    If color = const_color_reservada Or color = const_color_noAsignada Then
        gHabitacion.CellAlignment = 4   'alineo el texto al medio de la celda
        gHabitacion.CellFontBold = True 'muestro en negrita
        'solo en reserva muestro información dentro de la primer celda
        gHabitacion.CellForeColor = &H80000005  'color del texto = blanco
    End If
    
    'oculto número de bloqueo
    If color = const_color_bloqueada Then
        gHabitacion.CellForeColor = color
    End If
    
    If PpioFinLinea Then
        'muestro icono de inicio
        Select Case color
            Case const_color_bloqueada
                indiceImagen = 6
            Case const_color_ocupada
                indiceImagen = 4
            Case const_color_reservada
                indiceImagen = 2
            Case const_color_noAsignada
                indiceImagen = 8
        End Select
        gHabitacion.CellPictureAlignment = 1    'alineo imagen a la izquierda
        Set gHabitacion.CellPicture = ImageList1.ListImages(indiceImagen).Picture
    End If
End Sub

Private Sub subMuestroUltimoDia(col As Integer, color As Long, muestroIconoFinLinea As Boolean)
    'Muestro información en el último día de la grilla
    Dim indiceImagen As Byte
    
    If muestroIconoFinLinea Then
        'esto determina si la línea continúa más alla de los límites de la grilla
        'en ese caso no muestro indicativo de fín de linea
        gHabitacion.col = col   'me posiciono en grilla
        
        'muestro número de reserva
        If color = const_color_reservada Or color = const_color_noAsignada Then
            gHabitacion.CellAlignment = 4   'alineo el texto al medio de la celda
            gHabitacion.CellFontBold = True 'muestro en negrita
            'solo en reserva muestro información dentro de la primer celda
            gHabitacion.CellForeColor = &H80000005  'color del texto = blanco
        End If
        
        'oculto número de bloqueo
        If color = const_color_bloqueada Then
            gHabitacion.CellForeColor = color
        End If
        
        If PpioFinLinea Then
            'muestro icono de inicio
            Select Case color
                Case const_color_bloqueada
                    indiceImagen = 5
                Case const_color_ocupada
                    indiceImagen = 3
                Case const_color_reservada
                    indiceImagen = 1
                Case const_color_noAsignada
                    indiceImagen = 7
            End Select
            gHabitacion.CellPictureAlignment = 7    'alineo imagen a la derecha
            Set gHabitacion.CellPicture = ImageList1.ListImages(indiceImagen).Picture
        End If
    End If
End Sub

Private Function funAsignoIdentificacionCeldas(iden As Long, tipo As Long)
    'Da formato al número de la reserva, dividiéndole en dos partes.
    'El color de la celda me sirve para identificar si se trata de una reserva
    'sino no realizo nada ya que no se muestra en la grilla ningún número para los demás casos.
    Dim aux As String
    If tipo = const_color_reservada Or const_color_noAsignada Then    'asigno número de reserva
        aux = Mid(Str(iden), 1, 5)
        aux = aux & "-" & Mid(Str(iden), 6, 10)
    End If
    If tipo = const_color_bloqueada Then
        aux = iden
    End If
    funAsignoIdentificacionCeldas = aux
End Function

Private Function funObtengoFila(hab As Integer)
    'Se posiciona en la fila de la girra que corresponda
    'a la habitación que se pasa como parámetro
    Dim indice As Byte
    funObtengoFila = 0
    indice = 1
    gHabitacion.col = 0
    Do While indice < gHabitacion.Rows
        gHabitacion.Row = indice
        If Trim(gHabitacion.Text) = Trim(Str(hab)) Then
            funObtengoFila = gHabitacion.Row
            Exit Do
        End If
        indice = indice + 1
    Loop
End Function

Private Function funObtengoFilaNoAsignada(corrHab As Long)
    'Cuando muestro una reserva no asignada es necesario crear una nueva fila,
    'esta puede ser creada al principio o al final de la grilla.
    If PosicionNoAsignadas = 0 Then  'ppio de grilla
        gHabitacion.AddItem corrHab, 2
        gHabitacion.RowHeight(2) = AnchoCelda   'la nueva fila creada tiene que ser del
        'ancho configurado por el usuario
        funObtengoFilaNoAsignada = 2    'seguda fila de la grilla ya que la 1 se deja sin usar
                                        'para mayor prolijidad
    Else
        gHabitacion.AddItem corrHab, gHabitacion.Rows - 1
        gHabitacion.RowHeight(gHabitacion.Rows - 2) = AnchoCelda   'la nueva fila creada tiene que ser del
        'ancho configurado por el usuario
        funObtengoFilaNoAsignada = gHabitacion.Rows - 2
    End If
End Function

Private Sub Form_Unload(Cancel As Integer)
    'Descargo formulario de memoria
    Set frmCuadroHab = Nothing
End Sub

Private Sub gHabitacion_Click()
    'Permite consultar información extra al hacer doble clik sobre una linea
    'determinada.
    
    'determino sobre que grilla hice click por el color del fondo de la
    'misma
    Select Case gHabitacion.CellBackColor
        Case const_color_reservada
            tipoAccionCuadroHabInf = 1
            frmCuadroHabInf.Show 1
        Case const_color_ocupada
            tipoAccionCuadroHabInf = 2
            frmCuadroHabInf.Show 1
        Case const_color_bloqueada
            tipoAccionCuadroHabInf = 3
            frmCuadroHabInf.Show 1
        Case const_color_noAsignada
            tipoAccionCuadroHabInf = 4
            frmCuadroHabInf.Show 1
    End Select
End Sub

Private Sub gHabitacion_KeyPress(KeyAscii As Integer)
    'Realizo lo mismo que si hago click
    gHabitacion_Click
End Sub

Private Sub botSalir_Click()
    Unload Me
End Sub

Private Sub mnuFormularioAceptar_Click()
    'Esta opción del menú es lo mismo que presionar la tecla F12 o el boton de aceptar
    botSalir_Click
End Sub

Private Sub mnuFormularioImprimir_Click()
    'Es lo mismo que presionar Ctrol+I o el boton de imprimir
    botImprimir_Click
End Sub

Private Sub mnuProcesarConsulta_Click()
    'Esta opción del menú es lo mismo que presionar el boton procesar o la tecla F9
    botProcesar_Click
End Sub

Private Sub mnuVerDisponibilidad_Click()
    'Abre el formulario de Ver Disponibilidad
    OprEjecutada = 22
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        frmVerDisponibilidad.Show 1
    End If
End Sub

'*************************************************************************
'*
'*  Asistencia de usuarios
'*
'*************************************************************************

Private Sub desde_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 25
End Sub

Private Sub hasta_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 26
End Sub

Private Sub cboTipo_habitacion_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 27
End Sub

Private Sub botProcesar_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 28
End Sub

Private Sub botImprimir_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 29
End Sub

Private Sub botSalir_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 1, 1
End Sub

Private Sub gHabitacion_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 30
End Sub

Private Sub botSalir_LostFocus()
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

Private Sub hasta_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub desde_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub gHabitacion_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

