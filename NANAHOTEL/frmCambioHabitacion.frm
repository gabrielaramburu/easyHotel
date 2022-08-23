VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCambioHabitacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cambio habitacion"
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
   Begin Hotel_Nana.gaHOTELbarra gaHOTELbarra1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   6495
      Width           =   9510
      _ExtentX        =   16775
      _ExtentY        =   582
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   6240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   ""
      Top             =   1560
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Frame Frame2 
      Caption         =   "Pasajeros"
      Height          =   4455
      Left            =   120
      TabIndex        =   9
      Top             =   1920
      Width           =   9255
      Begin VB.CommandButton botConfirmar 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   7920
         TabIndex        =   7
         Tag             =   "Aceptar"
         Top             =   3960
         Width           =   1215
      End
      Begin VB.Data Data2 
         Caption         =   "Data2"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   6120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   2  'Snapshot
         RecordSource    =   ""
         Top             =   0
         Visible         =   0   'False
         Width           =   3135
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   3495
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   6165
         _Version        =   327680
         Style           =   1
         TabHeight       =   520
         TabCaption(0)   =   "Habitación desde"
         TabPicture(0)   =   "frmCambioHabitacion.frx":0000
         Tab(0).ControlCount=   4
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lblHabitacionDesde"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "MSFlexGrid1"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "botTodos"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "botCambiar"
         Tab(0).Control(3).Enabled=   0   'False
         TabCaption(1)   =   "Habitación hacia"
         TabPicture(1)   =   "frmCambioHabitacion.frx":001C
         Tab(1).ControlCount=   2
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "gHacia"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "lblHabitacionHacia"
         Tab(1).Control(1).Enabled=   0   'False
         TabCaption(2)   =   "Resultado operación"
         TabPicture(2)   =   "frmCambioHabitacion.frx":0038
         Tab(2).ControlCount=   2
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "lblResultado"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "gNuevaHubicacion"
         Tab(2).Control(1).Enabled=   0   'False
         Begin VB.CommandButton botCambiar 
            Caption         =   "&Cambiar"
            Height          =   375
            Left            =   7680
            TabIndex        =   23
            Top             =   720
            Width           =   1215
         End
         Begin MSFlexGridLib.MSFlexGrid gHacia 
            Bindings        =   "frmCambioHabitacion.frx":0054
            Height          =   2295
            Left            =   -74880
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   720
            Width           =   7455
            _ExtentX        =   13150
            _ExtentY        =   4048
            _Version        =   393216
            FixedCols       =   0
            GridLines       =   0
         End
         Begin VB.CommandButton botTodos 
            Caption         =   "Sel. &todos"
            Height          =   375
            Left            =   7680
            TabIndex        =   18
            Top             =   2640
            Width           =   1215
         End
         Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
            Bindings        =   "frmCambioHabitacion.frx":0064
            Height          =   2295
            Left            =   120
            TabIndex        =   21
            Top             =   720
            Width           =   7455
            _ExtentX        =   13150
            _ExtentY        =   4048
            _Version        =   393216
            FocusRect       =   2
            HighLight       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin MSFlexGridLib.MSFlexGrid gNuevaHubicacion 
            Height          =   2535
            Left            =   -74880
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   720
            Width           =   8295
            _ExtentX        =   14631
            _ExtentY        =   4471
            _Version        =   393216
            Rows            =   0
            FixedRows       =   0
            FixedCols       =   0
            GridLines       =   0
         End
         Begin VB.Label lblResultado 
            AutoSize        =   -1  'True
            Caption         =   "&Resultado del cambio"
            Height          =   195
            Left            =   -74880
            TabIndex        =   5
            Top             =   480
            Width           =   1530
         End
         Begin VB.Label lblHabitacionHacia 
            AutoSize        =   -1  'True
            Caption         =   "lblHabitacionHacia"
            Height          =   195
            Left            =   -74880
            TabIndex        =   20
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label lblHabitacionDesde 
            AutoSize        =   -1  'True
            Caption         =   "lblHabitacionDesde"
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   480
            Width           =   1380
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Habitaciones que intervienen en el cambio"
      Height          =   1815
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Width           =   9255
      Begin VB.CommandButton botAyudaHabTodas 
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7080
         TabIndex        =   25
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton botAyudaHabOcupadas 
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   24
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox txtHabDesde 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1200
         MaxLength       =   4
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtHabHacia 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   5880
         MaxLength       =   4
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton botProcesar 
         Caption         =   "&Procesar"
         Height          =   375
         Left            =   7920
         TabIndex        =   4
         Tag             =   "Procesar"
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label LTit2Hacia 
         Caption         =   "LTit2Hacia"
         Height          =   255
         Left            =   4920
         TabIndex        =   15
         Top             =   1440
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VB.Label LTit1Hacia 
         Caption         =   "LTit1Hacia"
         Height          =   255
         Left            =   4920
         TabIndex        =   14
         Top             =   1200
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VB.Label LTit2Desde 
         Caption         =   "LTit2Desde"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   1440
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VB.Label LhabHacia 
         Caption         =   "LhabHacia"
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
         Left            =   4920
         TabIndex        =   12
         Top             =   840
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VB.Label LTit1Desde 
         Caption         =   "LTit1Desde"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1200
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VB.Label LhabDesde 
         Caption         =   "LhabDesde"
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
         TabIndex        =   10
         Top             =   840
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "&Hacia"
         Height          =   195
         Left            =   4920
         TabIndex        =   2
         Top             =   450
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&Desde"
         Height          =   195
         Left            =   240
         TabIndex        =   0
         Top             =   450
         Width           =   465
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
         Caption         =   "Procesar          F9"
      End
      Begin VB.Menu mnuFormularioCambiar 
         Caption         =   "Cambiar           F9"
      End
   End
   Begin VB.Menu mnuVer 
      Caption         =   "&Ver información de..."
      Begin VB.Menu mnuVerDesde 
         Caption         =   "Habitación desde"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuVerHasta 
         Caption         =   "Habitación hasta"
         Shortcut        =   {F6}
      End
   End
End
Attribute VB_Name = "frmCambioHabitacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const cEspaciosBlanco As String = "        "    'utilizada para armar la grilla
                                                        'de resultados

Private esta_ocupada_hacia As Boolean
Private desde_quedo_libre As Boolean
Private fd_nueva_hab As Date
Private fh_nueva_hab As Date

'Este procedimiento se encarga de cambiar pasajeros de una habitacion (DESDE) a otra (HACIA)
'La habitación DESDE siempre debe de estar ocupada
'Podemos encontar diferentes variantes:
'1) Si la habitación HACIA está libre, se creará un nuevo checkin con la fechas desde y hasta
' correspondiete a la habitacion DESDE.
'Los titulares y la tarifa de la habitación HACIA serán los mismos que los de la DESDE.
'2) Si la habitación HACIA está ocupada, simplemnete se alojarán en la habitación respetando
'el período de fechas de ésta habitación (HACIA)
'En todos los casos si la habitación DESDE queda vacía se hará un checkout, traspasándose los
'gastos correspondientes a la habitación HACIA

Private Sub Form_Load()
    'Apariencia formulario
    mSubConfiguroFuentesControlesSistema Me
    'apariencia del formulario
    subEstablescoApariencia
    
    'Doy formato a la grilla antes de cargar el resultado de la consulta
    MSFlexGrid1.FormatString = "      | Nombre Pasajeros  " & _
    "                                                                   " & _
    "                                                            |         "
    MSFlexGrid1.ColWidth(2) = 0
    'inicializo controles data
    subInicializoControlData Me.Data1       'utilizado para trabajar con la grilla de pasajeros
                                            'de la habitación desde.
    subInicializoControlData Me.Data2       'utilizado para trabajar con la grilla de pasajeros
                                            'de la habitación hacia.
                                    
    'antes de realizar el cambio no permito trabajar con la ficha de reultados
    Me.SSTab1.TabEnabled(2) = False
    'por defecto establesco mo ficha predeterminada la 0
    Me.SSTab1.Tab = 0
    botones True
    subMuestroControles False
End Sub

Private Sub botProcesar_Click()
    If valido_habitaciones Then
        If valido_habitacion_hacia Then
            botones False
            subMuestroControles True
            genero_consulta
            'Cargo información de la habitación hacia
            'en la segunda ficha
            subCargoInfHabitacionHacia
        End If
    End If
End Sub

Private Sub genero_consulta()
    'Selecciono todos los pasajeros de la habitación DESDE
    
    SQLpasajeros_habitacion Val(txtHabDesde.Text), Data1
    
    MSFlexGrid1.FormatString = "      | Nombre Pasajeros  " & _
    "                                                                   " & _
    "                                      |         "
    'oculto la columna que contiene el número de cliente
    MSFlexGrid1.ColWidth(2) = 0
    'muestro el tipo y número de la habitación en la etiqueta correspondiente
    lblHabitacionDesde.Caption = "&Pasajeros de la habitación " & Me.txtHabDesde.Text & " " & Me.LhabDesde.Caption
End Sub

Private Sub subCargoInfHabitacionHacia()
    'Cargo el nombre de los pasajeros (y el total de los mismos)
    'alojados en la habitación hacia.

    'ejecuto consulta
    SQLpasajeros_habitacion CLng(txtHabHacia.Text), Data2
    'estableszo propiedades de la grilla
    gHacia.FormatString = " Nombre Pasajeros  " & _
    "                                                                   " & _
    "                                      |         "
    'oculto la columna que contiene el número de cliente
    gHacia.ColWidth(1) = 0
    'muestro el tipo y número de la habitación en la etiqueta correspondiente
    lblHabitacionHacia.Caption = "&Pasajeros de la habitación " & Me.txtHabHacia.Text & " " & Me.LhabHacia.Caption
End Sub

Private Sub subMuestroResultado()
    'Muestro información que sirve para verificar los pasos efectuados por el procedimiento.
    
    'oculto fichas que no se usan más y muestro la ficha de resultados
    Me.SSTab1.TabEnabled(0) = False
    Me.SSTab1.TabEnabled(1) = False
    Me.SSTab1.TabEnabled(2) = True
    'me posiciono en la ficha de resultado
    Me.SSTab1.Tab = 2
    'no permito las opciones del menú
    Me.mnuVer.Enabled = False
    'no permito ejecutar cambiar nuevamente ni seleccionar todos
    Me.mnuFormularioCambiar.Enabled = False
    Me.botCambiar.Enabled = False
    Me.botTodos.Enabled = False
    
    'cambio ancho de la única columna
    Me.gNuevaHubicacion.ColWidth(0) = Me.gNuevaHubicacion.Width + 500
    
    'obengo nuevamente los pasajeros de la habitación desde
    SQLpasajeros_habitacion CLng(txtHabDesde.Text), Data1
    'obtengo los pasajeros de la habitación hasta
    SQLpasajeros_habitacion CLng(txtHabHacia.Text), Data2
    
    'creo fila indicando cominezo de la habitación
    subMuestroInicioHab 0
    'recorro los pasajeros de la habitación desde y muestro en grilla
    If Not funRecorroPasajerosYMuestro(Data1) Then
        'la habitación desde quedo vacía
        Me.gNuevaHubicacion.AddItem cEspaciosBlanco & "La habitación quedo vacía."
        Me.gNuevaHubicacion.AddItem cEspaciosBlanco & "Los gastos de esta habitación " & _
                                    " se pasan a la habitación " & Me.txtHabHacia.Text & " " & Me.LhabHacia.Caption
    End If
    'creo fila indicando cominezo de la habitación
    subMuestroInicioHab 1
    'recorro los pasajeros de la habitación hacia y muestro en grilla
    If funRecorroPasajerosYMuestro(Data2) Then
       'siempre voy a encontrar pasajero en la habitación hacia
       'aquí no realizo nada
       If Not esta_ocupada_hacia Then
            'la habitación hasta estaba libre
            Me.gNuevaHubicacion.AddItem cEspaciosBlanco & "Se realizo checkin a la habitación " & _
                                                        Me.txtHabHacia.Text & " " & Me.LhabHacia.Caption
            Me.gNuevaHubicacion.AddItem cEspaciosBlanco & "Los titulares y la tarifa de esta habitación, pasan a ser los misos " & _
                                                        "que los de la habitación " & Me.txtHabDesde.Text & " " & Me.LhabDesde.Caption
       End If
    End If
    'aviso de cambio de habitación correcto
    mSubMensaje 4, 30
End Sub

Private Sub subMuestroInicioHab(tipoHab As Byte)
    'Creo línea en la grilla con información de la habitación.
    Dim descHab As String
    If tipoHab = 0 Then    'muestro habitación desde
        descHab = "Habitación " & txtHabDesde.Text & " " & LhabDesde.Caption
    Else                   'muestro habitación hacia
        descHab = "Habitación " & txtHabHacia.Text & " " & LhabHacia.Caption
    End If
    Me.gNuevaHubicacion.AddItem descHab
    Me.gNuevaHubicacion.Row = Me.gNuevaHubicacion.Rows - 1  'me ubico en la última fila para
                                                            'poder cambiar propiedades de la celda recien creada
    Me.gNuevaHubicacion.CellFontBold = True
End Sub

Private Function funRecorroPasajerosYMuestro(control As Data) As Boolean
    'Recorro el control data y por cada registro, creo una nueva línea en la grilla
    
    'por defecto devuelvo false
    funRecorroPasajerosYMuestro = False
    'veridico si tengo registros en el recordeset
    If control.Recordset.RecordCount > 1 Then
        control.Recordset.MoveFirst
        Do While Not control.Recordset.EOF
            'creo línea en la grilla
            Me.gNuevaHubicacion.AddItem cEspaciosBlanco & _
                                        control.Recordset.Fields(0).Value
            Me.gNuevaHubicacion.Row = _
                                    Me.gNuevaHubicacion.Rows - 1  'me ubico en la última fila para
                                                                  'poder cambiar propiedades de la celda recien creada
            'muestro los pasajero en negrita
            Me.gNuevaHubicacion.CellFontBold = True
            
            'si creo una línea es porque existe por lo menos 1 pasajero
            funRecorroPasajerosYMuestro = True
            control.Recordset.MoveNext
        Loop
    End If
End Function

Private Sub botCambiar_Click()
    Dim res As Byte
    
    If hay_pasajeros_marcados Then
        'aviso deconfirmación de cambio de habitación
        res = mFunMensaje(4, 23)
        If res = True Then
            If esta_ocupada_hacia Then
                'si la hab. HACIA esta OCUPADA
                muevo_pasajeros_desde_to_hacia 1
                
                'No cambio titulares habitación hacia
            Else
                'si la hab. HACIA esta LIBRE
                muevo_pasajeros_desde_to_hacia 0
        
                'Los titulares de la habitacion HACIA son los mismos
                'que la habitación DESDE y la tarifa también
                cambio_titulares_tarifa
                
            End If
            
            If desde_quedo_libre Then
                'si la habitación desde quedó libre
                inicializo_habitacion txtHabDesde.Text
                cambio_situacion txtHabDesde.Text, 2    'sucia
                
            End If
            'muestro resultado de la operación
            subMuestroResultado
            'grabo bitácora
            GraboBitacora "Hab.des " & txtHabDesde.Text & " Hab.has " & txtHabHacia.Text
        End If
    Else
        'debe de marcar pasajero para continuar
        mSubMensaje 4, 24
    End If
End Sub

Private Function hay_pasajeros_marcados()
    Dim i As Integer
    hay_pasajeros_marcados = False
    i = 1
    Do While i < MSFlexGrid1.Rows
        MSFlexGrid1.Row = i
        'si el pasajero está marcado
        If MSFlexGrid1.CellBackColor = mSisColor_15FilaSeleccionada Then
            hay_pasajeros_marcados = True
            Exit Do
        End If
        i = i + 1
    Loop
End Function

Private Sub cambio_titulares_tarifa()
    'Cuando la habitación HACIA esta vacia, no tiene asignada
    'ni titular ni tarifa, por lo que es necesario asigarles
    'los valores correspondientes a la habitación DESDE.
    
    Dim tcu As Integer
    Dim tca As Integer
    Dim tce As Integer
    
    Dim tu As Long
    Dim ta As Long
    Dim te As Long
    Dim tarifa As Double
    
    If busco_habitaTF(txtHabDesde.Text) Then
        tcu = tbHABITACIONES("tipocuenta_unica")
        tca = tbHABITACIONES("tipocuenta_aloja")
        tce = tbHABITACIONES("tipocuenta_extra")
        tu = tbHABITACIONES("titular_unica")
        ta = tbHABITACIONES("titular_aloja")
        te = tbHABITACIONES("titular_extra")
        tarifa = tbHABITACIONES("tarifa")
    End If
    
    If busco_habitaTF(txtHabHacia.Text) Then
        tbHABITACIONES.Edit
            tbHABITACIONES("tipocuenta_unica") = tcu
            tbHABITACIONES("tipocuenta_aloja") = tca
            tbHABITACIONES("tipocuenta_extra") = tce
            tbHABITACIONES("titular_unica") = tu
            tbHABITACIONES("titular_aloja") = ta
            tbHABITACIONES("titular_extra") = te
            tbHABITACIONES("tarifa") = tarifa
        tbHABITACIONES.Update
    End If
End Sub

Private Sub muevo_pasajeros_desde_to_hacia(estado_hab As Byte)
    'Recorro la grilla de pasajeros, con los que estan marcados
    'y cambio los mismos de habitación

    Dim i As Integer
    Dim nrocli As Long
    Dim fdes_aux As Date
    Dim fhas_aux As Date
    
    desde_quedo_libre = True
    i = 1
    MSFlexGrid1.Row = 1
    Do While i < MSFlexGrid1.Rows
        MSFlexGrid1.Row = i
        MSFlexGrid1.col = 1
        'si el pasajero está marcado
        If MSFlexGrid1.CellBackColor = mSisColor_15FilaSeleccionada Then
            MSFlexGrid1.col = 2
            nrocli = MSFlexGrid1.Text   'obtengo cliente
                
            If estado_hab = 1 Then  'ocupada
                'antes de poscisionarme en el registro correspondiente
                'a la hab. DESDE debo obtener las fechas fdes y fhas de la
                'Hab.HACIA
                'NOTA: ----> espero poder poscisionarme en forma directa
                'por una clave secundaria con duplicados <-----
                
                If busco_habita_checkin(Val(txtHabHacia.Text)) Then
                    fdes_aux = tbCHECKIN("fcheckdes")
                    fhas_aux = tbCHECKIN("fcheckhas")
                End If
            End If
            
            'me posiciono en el resgistro corresponiente al pasajero
            'de la hab. DESDE
            'NOTA: ----> espero poder modificar una clave primaria sin armar
            'relajo <----
            tbCHECKIN.Index = "i_checkin"
            tbCHECKIN.Seek "=", txtHabDesde.Text, nrocli
            If Not tbCHECKIN.NoMatch Then
                tbCHECKIN.Edit
                tbCHECKIN("nrohab") = txtHabHacia.Text
                If estado_hab = 1 Then  'HACIA ocupada
                    tbCHECKIN("fcheckdes") = fdes_aux
                    tbCHECKIN("fcheckhas") = fhas_aux
                End If
                tbCHECKIN.Update
            End If
        Else
            'si encuentro un pasajero que no esté marcado ya se que la habitación
            'no quedo libre.
            desde_quedo_libre = False
        End If
        i = i + 1
    Loop
End Sub

Private Sub mens_error(tipo As Byte)
    Select Case tipo
        Case 1
            'la habitación HACIA se encuentra reservada"
            mSubMensaje 4, 25
        Case 3
            'la habitación HACIA se encuentra bloqueada"
            mSubMensaje 4, 26
    End Select
End Sub

Private Function valido_habitaciones()
    'Controlo que existan las habitaciones; si existen muestro
    'los datos correspondientes y continúo con la ejecución del programa
    valido_habitaciones = True
            
    If Trim(txtHabDesde.Text) = Empty Or Trim(txtHabHacia.Text) = Empty Then
        'debe de ingresar dos habitaciones"
        mSubMensaje 4, 27
        valido_habitaciones = False
        txtHabDesde.SetFocus
    Else
        If Val(txtHabDesde.Text) = Val(txtHabHacia.Text) Then
            'las habitaciones son iguales
            mSubMensaje 4, 28
            valido_habitaciones = False
            txtHabDesde.SetFocus
            Exit Function
        End If
        
        'Habitación DESDE
        If busco_habitaTF(Val(txtHabDesde.Text)) Then
            'para la habitación DESDE además valido que este ocupada
            If busco_habita_checkin(Val(txtHabDesde.Text)) Then
                'muestro datos habitación
                muestro_titular LTit1Desde, LTit2Desde
                muestro_tipo LhabDesde
            Else
                'no hay pasajeros hospedados en esa habitación
                mSubMensaje 4, 29
                txtHabDesde.SetFocus
                valido_habitaciones = False
                Exit Function
            End If
        Else
            'no existe habitación desde
            mSubMensaje 4, 17
            txtHabDesde.SetFocus
            valido_habitaciones = False
            Exit Function
        End If
        
        'Habitación HACIA
        If busco_habitaTF(Val(txtHabHacia.Text)) Then
            esta_ocupada_hacia = busco_habita_checkin(txtHabHacia.Text)
            'muestro datos habitación
            'Si la habitación está ocupada muestro los titulares, si está libre muestro
            'la situación
            If esta_ocupada_hacia Then
                muestro_titular LTit1Hacia, LTit2Hacia
            Else
                muestro_estado_situacion
            End If
            muestro_tipo LhabHacia
        Else
            'no existe habitación
            mSubMensaje 4, 17
            txtHabHacia.SetFocus
            valido_habitaciones = False
        End If
    End If
End Function

Private Function valido_habitacion_hacia()
    'Para la habitación HACIA es necesario validar
    'que la misma este disponible: sin reserva y no bloqueada, para el período
    'que se desea ocupar (el período de la hab. DESDE)
    
    valido_habitacion_hacia = False
    'Si esta ocupada no hago nada
    If Not esta_ocupada_hacia Then
        obtengo_fechas_nueva_hab 'obtengo período de habitación DESDE
        'verifico que HASTA no esté reservada
        If Not habitacion_reservada(Val(txtHabHacia.Text), fd_nueva_hab, fh_nueva_hab) Then
            'verifico que HASTA no esté bloqueada
            If Not habitacion_bloqueada(txtHabHacia.Text, fd_nueva_hab, fh_nueva_hab) Then
                valido_habitacion_hacia = True
            Else
                mens_error 3
            End If
        Else
            mens_error 1
        End If
    
    Else
        valido_habitacion_hacia = True
    End If
End Function

Private Sub obtengo_fechas_nueva_hab()
    If busco_habita_checkin(txtHabDesde.Text) Then
        'la fecha desde es independiente de la fecha de la habitación desde ya
        'que el nuevo período de alojamiento de la nueva habitación,
        'siempre será a partir de la fecha actual
        
        fd_nueva_hab = m_FechaSis
        fh_nueva_hab = tbCHECKIN("fcheckhas")
    End If
End Sub

Private Sub muestro_estado_situacion()
    LTit1Hacia.Caption = "Estado: LIBRE"
    If busco_estado_habTF(2, tbHABITACIONES("situacionhab")) Then
        LTit2Hacia.Caption = "Situación: " & tbTIPO_ESTADO_HAB("descri")
    End If
End Sub

Private Sub muestro_titular(tit1 As Label, tit2 As Label)
    If tbHABITACIONES("titular_unica") <> 0 Then 'unico titular
        tit1.Caption = "Titular único :" & busco_titular_hab(tbHABITACIONES("nrohab"), "unica")
        tit1.Visible = True
    Else
        tit1.Caption = "Titualr aloja.:" & busco_titular_hab(tbHABITACIONES("nrohab"), "aloja")
        tit2.Caption = "Titular extras:" & busco_titular_hab(tbHABITACIONES("nrohab"), "extra")
        tit1.Visible = True
        tit2.Visible = True
    End If
End Sub

Private Sub muestro_tipo(tipo As Label)
    Dim tipo_hab As String
    'cargo tipo y número de habitación
    'obtengo tipo habitación
    If busco_tipo_habTF(tbHABITACIONES("tipohab")) Then
        tipo_hab = tbTIPO_HABITACIONES("descripcion")
    End If
    tipo.Caption = "Suite " & tipo_hab
    tipo.Visible = True
End Sub

Private Sub botTodos_Click()
    'Selecciono todas las filas (pasajeros) de la grilla.
    
    'verifico que existan pasajeros
    If MSFlexGrid1.Rows > 1 Then
         marco_celdas_grilla Me.MSFlexGrid1, 1, 1, 1, Me.MSFlexGrid1.Rows - 1
        'marco del color determinado
        MSFlexGrid1.CellBackColor = mSisColor_15FilaSeleccionada
        MSFlexGrid1.CellForeColor = mSisColor_19FilaSeleccionadaTexto
    End If
End Sub

Private Sub botones(x As Boolean)
    'True=muestro antes de procesar
    'False=muestro despues de procear
    If x Then
        Me.botTodos.Enabled = False
        Me.botCambiar.Enabled = False
        Me.mnuFormularioCambiar.Enabled = False
        Me.SSTab1.TabEnabled(0) = False
        Me.SSTab1.TabEnabled(1) = False
        Me.mnuVer.Enabled = False
    Else
        Me.botTodos.Enabled = True
        Me.botCambiar.Enabled = True
        Me.botAyudaHabOcupadas.Enabled = False
        Me.botAyudaHabTodas.Enabled = False
        Me.txtHabDesde.BackColor = mSisColor_18ControlesNoHabilitados
        Me.txtHabHacia.BackColor = mSisColor_18ControlesNoHabilitados
        Me.txtHabDesde.Locked = True
        Me.txtHabHacia.Locked = True
        Me.txtHabDesde.TabStop = False
        Me.txtHabHacia.TabStop = False
        Me.mnuFormularioProcesar.Enabled = False
        Me.mnuFormularioCambiar.Enabled = True
        Me.SSTab1.TabEnabled(0) = True
        Me.SSTab1.TabEnabled(1) = True
        Me.mnuVer.Enabled = True
        Me.botProcesar.Enabled = False
    End If
End Sub

Private Sub subMuestroControles(x As Boolean)
    'Muestro o oculto las grillas y las etiquetas que se encuentran en las fichas
    'antes de presionar el boton de procesar.
    Me.lblHabitacionDesde.Visible = x
    Me.lblHabitacionHacia.Visible = x
    Me.gHacia.Visible = x
    Me.MSFlexGrid1.Visible = x
    Me.botCambiar.Visible = x
    Me.botTodos.Visible = x
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Descargo formulario de memoria
    Set frmCambioHabitacion = Nothing
End Sub

Private Sub MSFlexGrid1_DblClick()
    'Selecciono grilla con el mouse
    marco_grilla MSFlexGrid1, 1, 1
End Sub

Private Sub MSFlexGrid1_KeyPress(KeyAscii As Integer)
    'Permito la selección de pasajeros con la tecla Enter
    marco_grilla MSFlexGrid1, 1, 1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'Intercepto la tecla F9
    If KeyCode = vbKeyF9 Then
        Form_KeyPress (KeyCode)
    Else
        If KeyCode = vbKeyF1 Then
            'verifico que los botones esten activos
            If Me.botAyudaHabOcupadas.Enabled = True And Me.botAyudaHabTodas.Enabled = True Then
                'determino que ayuda estoy necesitando, ya que existen dos botone:
                'uno para las habitaciones ocupadas y otro para todas las habitaciones.
                If Me.ActiveControl.Name = "txtHabDesde" Then
                    'estoy posicionado sobre el control de ingreso de la habitación
                    'desde
                    'llamo ayuda de habitacione ocupadas
                    botAyudaHabOcupadas_Click
                Else
                    If Me.ActiveControl.Name = "txtHabHacia" Then
                        'estoy posicionado sobre el control de ingreso de la habitación
                        'hacia
                        'llamo ayuda de todas las habitacione
                        botAyudaHabTodas_Click
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'La tecla F9 tiene dos funciones dependiendo del estado del formulario:
    'puede equivaler a presionar el boton procesar
    'o puede equivaler a preionar el boton cambiar
    'Para ambos casos su uso es intuitivo para el usuario así que no hay problemas.
        
    Select Case KeyAscii
        Case vbKeyF9
            If Me.botProcesar.Enabled Then
                botProcesar_Click
            Else
                If botCambiar.Enabled Then
                    botCambiar_Click
                End If
            End If
        Case vbKeyEscape
            Unload Me
    End Select
End Sub

Private Sub botAyudaHabOcupadas_Click()
    'Muestro todas las habitaciones ocupadas del hotel.
    'Se utiliza para ingresar la habitación desde que siempre tiene que estar ocupada.
    Me.txtHabDesde.Text = mFunBusqueda(9)
End Sub

Private Sub botAyudaHabTodas_Click()
    'Muestro todas las habitaciones del hotel
    'Se utiliza para ingresar la habitación hacias, la cual puede estar
    'libre o ocupada.
    Me.txtHabHacia.Text = mFunBusqueda(8)
End Sub

Private Sub botConfirmar_Click()
    Unload Me
End Sub

Private Sub mnuFormularioAceptar_Click()
    'Equivale a presionar el boton de aceptar
    botConfirmar_Click
End Sub

Private Sub mnuFormularioCambiar_Click()
    'Equivale a preionar el boton de cambiar o la tecla F9
    botCambiar_Click
End Sub

Private Sub mnuFormularioProcesar_Click()
    'Equivale a presionar el boton procesar o la tecla F9
    botProcesar_Click
End Sub

Private Sub mnuVerDesde_Click()
    'Selecciono la primer ficha
    Me.SSTab1.Tab = 0
End Sub

Private Sub mnuVerHasta_Click()
    'Selecciono la segunda ficha
    Me.SSTab1.Tab = 1
End Sub

Private Sub txtHabDesde_KeyPress(KeyAscii As Integer)
    'Permito solo el ingreso de números
    ValidoNum KeyAscii, False, False
End Sub

Private Sub txtHabHacia_KeyPress(KeyAscii As Integer)
    'Permito solo el ingreso de números
    ValidoNum KeyAscii, False, False
End Sub

Private Sub subEstablescoApariencia()
    'Determino la apariencia de ciertos controles configurables
End Sub

'******************************************************************************
'*
'*  Asistencia al usuario
'*
'******************************************************************************

Private Sub txtHabDesde_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 9
End Sub

Private Sub txtHabDesde_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub txtHabHacia_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 10
End Sub
    
Private Sub txtHabHacia_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub
    
Private Sub botProcesar_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 1, 4
End Sub

Private Sub botProcesar_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub MSFlexGrid1_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 12
End Sub

Private Sub MSFlexGrid1_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub botCambiar_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 11
End Sub

Private Sub botCambiar_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub botTodos_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 13
End Sub

Private Sub botTodos_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub botConfirmar_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 1, 1
End Sub

Private Sub botConfirmar_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

