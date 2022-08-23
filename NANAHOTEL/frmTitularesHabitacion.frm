VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmTitularesHabitacion 
   Caption         =   "Form1"
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11250
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6510
   ScaleWidth      =   11250
   StartUpPosition =   2  'CenterScreen
   Begin Hotel_Nana.gaHOTELtitular gaHOTELtitular1 
      Height          =   1335
      Left            =   120
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   0
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   2355
      BackColor       =   12632256
   End
   Begin VB.Frame Frame8 
      Height          =   1335
      Left            =   4320
      TabIndex        =   38
      Top             =   0
      Width           =   6615
      Begin Hotel_Nana.gaHOTELtipo gaHOTELtipo1 
         Height          =   300
         Left            =   120
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   240
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   529
         BackColor       =   12632256
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   8760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   "select cliente.nombre_completo_titular,checkin.nrocorrcli from checkin, clientes where checkin.nrocorrcli = clientes.nrocorr"
      Top             =   1440
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Frame Frame6 
      Caption         =   "Tipo de cuenta habitación "
      Height          =   4695
      Left            =   120
      TabIndex        =   19
      Top             =   1440
      Width           =   11055
      Begin VB.Frame Frame7 
         Caption         =   "Tarifa de la habitación "
         Height          =   735
         Left            =   120
         TabIndex        =   37
         Top             =   3600
         Width           =   4935
         Begin VB.TextBox txttarifa 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   3720
            MaxLength       =   6
            TabIndex        =   31
            Top             =   205
            Width           =   1095
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Ta&rifa U$S"
            Height          =   240
            Left            =   2520
            TabIndex        =   30
            Top             =   285
            Width           =   960
         End
      End
      Begin VB.CommandButton botAceptar 
         Height          =   375
         Left            =   8400
         Picture         =   "frmTitularesHabitacion.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   33
         Tag             =   "Aceptar"
         Top             =   4200
         Width           =   1215
      End
      Begin VB.CommandButton botCancelar 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   9720
         Picture         =   "frmTitularesHabitacion.frx":08B6
         Style           =   1  'Graphical
         TabIndex        =   32
         Tag             =   "Cancelar"
         Top             =   4200
         Width           =   1215
      End
      Begin VB.CommandButton botConfirmarCuenta 
         Caption         =   "C&onfirmar cuenta"
         Height          =   375
         Left            =   3240
         TabIndex        =   2
         Top             =   600
         Width           =   1815
      End
      Begin VB.OptionButton chkCuentasSeparadas 
         Caption         =   "Cuentas &separadas"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   2415
      End
      Begin VB.OptionButton chkCuentaUnica 
         Caption         =   "C&uenta única"
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   2415
      End
      Begin MSFlexGridLib.MSFlexGrid gTitulares 
         Height          =   1575
         Left            =   120
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1920
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   2778
         _Version        =   393216
         Rows            =   0
         FixedRows       =   0
         FixedCols       =   0
         GridLines       =   0
      End
      Begin TabDlg.SSTab sstTipoTitular 
         Height          =   3255
         Left            =   5280
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   240
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   5741
         _Version        =   327680
         Style           =   1
         Tabs            =   5
         TabsPerRow      =   5
         TabHeight       =   520
         TabCaption(0)   =   "Seleccione tipo"
         TabPicture(0)   =   "frmTitularesHabitacion.frx":1178
         Tab(0).ControlCount=   1
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame1"
         Tab(0).Control(0).Enabled=   0   'False
         TabCaption(1)   =   "Propio Pax"
         TabPicture(1)   =   "frmTitularesHabitacion.frx":1194
         Tab(1).ControlCount=   1
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame2"
         Tab(1).Control(0).Enabled=   0   'False
         TabCaption(2)   =   "Otro Pax"
         TabPicture(2)   =   "frmTitularesHabitacion.frx":11B0
         Tab(2).ControlCount=   1
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame3"
         Tab(2).Control(0).Enabled=   0   'False
         TabCaption(3)   =   "Emp./Age."
         TabPicture(3)   =   "frmTitularesHabitacion.frx":11CC
         Tab(3).ControlCount=   1
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Frame4"
         Tab(3).Control(0).Enabled=   0   'False
         TabCaption(4)   =   "Otros"
         TabPicture(4)   =   "frmTitularesHabitacion.frx":11E8
         Tab(4).ControlCount=   1
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "Frame5"
         Tab(4).Control(0).Enabled=   0   'False
         Begin VB.Frame Frame1 
            Height          =   2775
            Left            =   120
            TabIndex        =   29
            Top             =   360
            Width           =   5415
            Begin VB.OptionButton chkTitularUnico 
               Caption         =   "Titular ú&nico"
               Height          =   255
               Left            =   240
               TabIndex        =   3
               Top             =   720
               Width           =   2055
            End
            Begin VB.OptionButton chkTitularExtras 
               Caption         =   "Titular gastos e&xtras"
               Height          =   255
               Left            =   240
               TabIndex        =   5
               Top             =   1680
               Width           =   2175
            End
            Begin VB.OptionButton chkTitularAlojamiento 
               Caption         =   "Titular gastos &alojamiento"
               Height          =   255
               Left            =   240
               TabIndex        =   4
               Top             =   1200
               Width           =   2895
            End
            Begin VB.Label lblTipoCuentaSelTabs 
               AutoSize        =   -1  'True
               Caption         =   "Seleccione tipo de titular"
               Height          =   195
               Left            =   240
               TabIndex        =   36
               Top             =   240
               Width           =   1740
            End
         End
         Begin VB.Frame Frame2 
            Height          =   2775
            Left            =   -74880
            TabIndex        =   28
            Top             =   360
            Width           =   5415
            Begin VB.CommandButton botConfirmar 
               Caption         =   "&Confirmar"
               Height          =   375
               Index           =   0
               Left            =   4080
               TabIndex        =   8
               Top             =   2280
               Width           =   1215
            End
            Begin MSFlexGridLib.MSFlexGrid gPropioPax 
               Bindings        =   "frmTitularesHabitacion.frx":1204
               Height          =   1695
               Left            =   120
               TabIndex        =   7
               Top             =   480
               Width           =   5175
               _ExtentX        =   9128
               _ExtentY        =   2990
               _Version        =   393216
               Rows            =   1
               Cols            =   4
               AllowBigSelection=   0   'False
               FocusRect       =   0
               ScrollBars      =   0
               SelectionMode   =   1
               AllowUserResizing=   1
               Appearance      =   0
               MousePointer    =   2
               FormatString    =   $"frmTitularesHabitacion.frx":1214
            End
            Begin VB.Label lblPasajeros 
               AutoSize        =   -1  'True
               Caption         =   "lblPasajeros"
               Height          =   195
               Left            =   120
               TabIndex        =   6
               Top             =   240
               Width           =   840
            End
         End
         Begin VB.Frame Frame3 
            Height          =   2775
            Left            =   -74880
            TabIndex        =   26
            Top             =   360
            Width           =   5415
            Begin VB.TextBox txtOtroPax 
               Enabled         =   0   'False
               Height          =   360
               Left            =   120
               TabIndex        =   27
               TabStop         =   0   'False
               Top             =   540
               Width           =   4455
            End
            Begin VB.CommandButton botAyuda 
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
               Index           =   0
               Left            =   4680
               Style           =   1  'Graphical
               TabIndex        =   10
               Top             =   540
               Width           =   555
            End
            Begin VB.CommandButton botConfirmar 
               Caption         =   "&Confirmar"
               Height          =   375
               Index           =   1
               Left            =   4080
               TabIndex        =   11
               Top             =   2280
               Width           =   1215
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "&Pasajero del hotel"
               Height          =   195
               Left            =   120
               TabIndex        =   9
               Top             =   240
               Width           =   1260
            End
         End
         Begin VB.Frame Frame4 
            Height          =   2775
            Left            =   -74880
            TabIndex        =   24
            Top             =   360
            Width           =   5415
            Begin VB.CommandButton botAyuda 
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
               Index           =   1
               Left            =   4680
               Style           =   1  'Graphical
               TabIndex        =   13
               Top             =   533
               Width           =   555
            End
            Begin VB.TextBox txtEmpAge 
               Enabled         =   0   'False
               Height          =   360
               Left            =   120
               TabIndex        =   25
               TabStop         =   0   'False
               Top             =   540
               Width           =   4455
            End
            Begin VB.CommandButton botConfirmar 
               Caption         =   "&Confirmar"
               Height          =   375
               Index           =   2
               Left            =   4080
               TabIndex        =   14
               Top             =   2280
               Width           =   1215
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "&Empresa o agencia"
               Height          =   195
               Left            =   120
               TabIndex        =   12
               Top             =   240
               Width           =   1365
            End
         End
         Begin VB.Frame Frame5 
            Height          =   2775
            Left            =   -74880
            TabIndex        =   22
            Top             =   360
            Width           =   5415
            Begin VB.CommandButton botAyuda 
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
               Index           =   2
               Left            =   4680
               Style           =   1  'Graphical
               TabIndex        =   16
               Top             =   533
               Width           =   555
            End
            Begin VB.TextBox txtOtros 
               Enabled         =   0   'False
               Height          =   360
               Left            =   120
               TabIndex        =   23
               TabStop         =   0   'False
               Top             =   540
               Width           =   4455
            End
            Begin VB.CommandButton botConfirmar 
               Caption         =   "&Confirmar"
               Height          =   375
               Index           =   3
               Left            =   4080
               TabIndex        =   17
               Top             =   2280
               Width           =   1215
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "C&lientes no alojados"
               Height          =   195
               Left            =   120
               TabIndex        =   15
               Top             =   240
               Width           =   1410
            End
         End
      End
      Begin VB.Label lblTipoCuentaSel 
         AutoSize        =   -1  'True
         Caption         =   "lblTipoCuentaSel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   35
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "&Titulares"
         Height          =   195
         Left            =   120
         TabIndex        =   34
         Top             =   1680
         Width           =   600
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   5040
         Y1              =   1200
         Y2              =   1200
      End
   End
   Begin Hotel_Nana.gaHOTELbarra gaHOTELbarra1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   6180
      Width           =   11250
      _ExtentX        =   19844
      _ExtentY        =   582
   End
   Begin VB.Menu mnuFormulario 
      Caption         =   "&Formulario"
      Begin VB.Menu mnuFormularioAceptar 
         Caption         =   "Aceptar"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuFormularioCancelar 
         Caption         =   "Cancelar"
      End
      Begin VB.Menu div1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFormularioConfirmarTit 
         Caption         =   "Confirmar titular"
         Shortcut        =   {F9}
      End
   End
   Begin VB.Menu mnuVer 
      Caption         =   "&Ver información de..."
      Begin VB.Menu mnuVerTipo 
         Caption         =   "Tipo de titular"
      End
      Begin VB.Menu mnuVerPropioPax 
         Caption         =   "Propio pax"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuVerOtro 
         Caption         =   "Otro pax"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuVerEmpAgen 
         Caption         =   "Emp./Agen."
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuVerOtros 
         Caption         =   "Otros"
         Shortcut        =   {F8}
      End
   End
   Begin VB.Menu mnuBuscar 
      Caption         =   "&Buscar"
      Begin VB.Menu mnuBuscarTit 
         Caption         =   "ayuda de..."
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmTitularesHabitacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Declaro propiedades del formulario
Public propTipoAccionFormularioTitular As Byte  'determina el tipo de acción para lo cual es llamado este
                                                'formulario
                                                '   1 walkin
                                                '   2 checkin sin asignación
                                                '   3 checkin con asignación
                                                '   4 cambio de titular

Public propHabCuenta As Long                    'habitación con la cual estoy trabajando

                                                'Utilizadas en el procedimiento grabo reserva.
Public propHabNroReserva As Long                'Estas dos propiedades sirven para formar la clave del
Public propHabNroCorr As Long                   'archivo de HAB_RESERVAS,al cual se accede para grabar el
                                                'número de habitación.

'Declaro variables generale spara este formulario
Private gTipoCuenta As Byte         'determina el tipo de cuenta de la habitación
                                    '0 = cuenta única
                                    '1 = cuenta separadas
                            
'Declaración de constantes
Private Const cEspacios As String = "           "       'sirve para mostrar los titulares
                                                        'en la grilla

Private Sub Form_Activate()
    'NOTA: Por algún motivo que no puede determinar
    'esta línea de código da error (si la pongo en el evento load
    
    'inicializo control data
    subInicializoControlData Me.Data1
End Sub

Private Sub Form_Load()
    'Apariencia formulario
    mSubConfiguroFuentesControlesSistema Me
    
    'inicializo control de cabezal formulario
    subCabezalFormulario
    'inicializo apariencia incial del formulario
    subInicializoAparienciaInicial
    'configuro apariencia dependiendo de la tarea que realiza el mismo
    subInicialioAparienciaSegunTarea
End Sub

'******************************************************
'*
'*  Procedimientos que manejan el formulario
'*
'******************************************************

Private Sub subCabezalFormulario()
    'Inicializo propiedades del control que se muestra al principio del formulario.
    
    'verifico para que llamo al formulario
    If propTipoAccionFormularioTitular = 4 Then     'lo llamo para cambiar el titular
        Me.gaHOTELtitular1.CaminoBaseDeDatos = vardir
        Me.gaHOTELtitular1.NumeroHabitacion = propHabCuenta
        'cambio tamaño del cabezal del formuario
        Me.gaHOTELtitular1.Width = 11055
        'oculto frame que no utlizo para este caso
        Frame8.Visible = False
    Else                                            'lo llamo para asignar titular
        Me.gaHOTELtipo1.CaminoBaseDeDatos = vardir
        Me.gaHOTELtipo1.NumeroHabitacion = propHabCuenta
        'cambio tamaño y posicióndel cabezal del formulario
        Frame8.Left = 120
        Frame8.Top = 0
        Frame8.Width = 11055
        'oculto control que no utilizo
        Me.gaHOTELtitular1.Visible = False
    End If
End Sub

Private Sub BotAyuda_Click(Index As Integer)
    'Muestro ayuda
    Select Case Index
        Case 0  'otro pax
            'muestra todos los pasajeros hospedados en el hotel.
            Me.txtOtroPax.Tag = Val(mFunBusqueda(2))
            If busco_clienteTF(Me.txtOtroPax.Tag) Then
                Me.txtOtroPax.Text = tbCLIENTES("nombre_completo_titular")
            End If
            
        Case 1  'emp agencia
            'muestra todas las empresas del hotel
            Me.txtEmpAge.Tag = Val(mFunBusqueda(3))
            If busco_empTF(Me.txtEmpAge.Tag) Then
               Me.txtEmpAge.Text = tbEMPRESAS("nomEmp")
            End If
            
        Case 2  'otro
            'muestra todos los pasajeros que no están hospedados
            Me.txtOtros.Tag = Val(mFunBusqueda(7))
            If busco_clienteTF(Me.txtOtros.Tag) Then
                Me.txtOtros.Text = tbCLIENTES("nombre_completo_titular")
            End If
        
    End Select
    'le paso el focus al boton de confirmar
    Me.botConfirmar(Index).SetFocus
End Sub

Private Sub botConfirmar_Click(Index As Integer)
    'Se confirma un titular
    If funValidoDatosTitular Then
        'Cargo información en grilla de titulares
        subCargoTitular
        'configuro ancho de columnas
        gTitulares.ColWidth(0) = gTitulares.Width - 255
        gTitulares.ColWidth(1) = 0
        'habilito el tabs 0 para mejorar interface
        Me.sstTipoTitular.Tab = 0
        'le doy el focus a los controles de selección
        If Me.chkTitularAlojamiento.Visible = True Then
            Me.chkTitularAlojamiento.SetFocus
        Else
            Me.chkTitularUnico.SetFocus
        End If
    End If
End Sub

Private Function funValidoDatosTitular()
    'Valido que al momento de confirmar el titular se halla seleccionado (F1) el mismo.
    Dim codErr As Byte
    Dim codDescMsg As Integer
    Dim focus As Object
    
    'por defecto asumo que todo esta bien
    funValidoDatosTitular = True
    codErr = 0
    
    Select Case Me.sstTipoTitular.Tab
        Case 1  'propio pax
            'en este caso no hay que validar nada ya que la grilla de seleccioón
            'de propioPax siempre tiene una fila seleccionada
        Case 2  'otro pax
            If Me.txtOtroPax.Tag = Empty Then
                codErr = 1
            End If
            
        Case 3  'emp ag
            If Me.txtEmpAge.Tag = Empty Then
                codErr = 2
            End If
            
        Case 4  'otro
            If Me.txtOtros.Tag = Empty Then
                codErr = 3
            End If
    End Select
    'verifico si encontre errores
    If codErr > 0 Then
        Select Case codErr
            Case 1
                'Debe de seleccionar un pasajero
                codDescMsg = 70
                Set focus = Me.botAyuda(0)
            Case 2
                'Debe de seleccionar una empresa o agencia
                codDescMsg = 71
                Set focus = Me.botAyuda(1)
            Case 3
                codDescMsg = 72
                'Debe de seleccionar un cliente del hotel
                Set focus = Me.botAyuda(2)
        End Select
        funValidoDatosTitular = False
        mSubMensaje 4, codDescMsg
        'le doy el focus al boton de ayuda
        focus.SetFocus
        Set focus = Nothing
    End If
End Function

Private Sub subCargoTitular()
    'Crea tres líneas en la grilla de titulares o modifica las existentes
    'Cada línea contiene información del titular recién seleccionado.
    Dim fila As Byte
    
    'creo las líneas en la grilla
    subCreoLineasGrilla
        
    'obtengo fila donde debo de cargar los datos
    fila = funObtengoFila
    'en la primer fila cargo tipo de titular(relacionado con la cuenta)
    gTitulares.TextMatrix(fila, 0) = funObtengoDescTipoTitular  'columna 0
    'en la segunda fila cargo tipo de titular(relacionado con el origen)
    gTitulares.TextMatrix(fila + 1, 0) = cEspacios & funObtengoDescTipoTitularOrigen    'columna 0
    gTitulares.TextMatrix(fila + 1, 1) = funObtengoTipoTitularOrigen                    'columna 1
    'en la tercer fila cargo nombre del titular
    gTitulares.TextMatrix(fila + 2, 0) = cEspacios & funObtengoNombreTitular
    'en la tercer fila columna 2 cargo el número del titular
    gTitulares.TextMatrix(fila + 2, 1) = funObtengoNumeroTitular
    'cambio atributos del texto para mejorar presentación
    subMejoroPresentacionGrilla fila
End Sub

Private Function funObtengoNombreTitular() As String
    'Dependiendo del origen del titular, devuelvo el nombre establecido para el mismo
    Select Case Me.sstTipoTitular.Tab
        Case 1  'propio pax
            funObtengoNombreTitular = Me.gPropioPax.TextMatrix(gPropioPax.Row, 1)
        Case 2  'otro pax
            funObtengoNombreTitular = Me.txtOtroPax.Text
        Case 3  'emp agencia
            funObtengoNombreTitular = Me.txtEmpAge.Text
        Case 4  'otros
            funObtengoNombreTitular = Me.txtOtros.Text
    End Select
End Function

Private Function funObtengoNumeroTitular() As String
    'Dependiendo del origen del titular, devuelvo el número establecido para el mismo
    Select Case Me.sstTipoTitular.Tab
        Case 1  'propio pax
            funObtengoNumeroTitular = gPropioPax.TextMatrix(gPropioPax.Row, 2)
        Case 2  'otro pax
            funObtengoNumeroTitular = txtOtroPax.Tag
        Case 3  'emp agencia
            funObtengoNumeroTitular = txtEmpAge.Tag
        Case 4  'otros
            funObtengoNumeroTitular = txtOtros.Tag
    End Select
End Function

Private Sub subCreoLineasGrilla()
    'Dependiendo de la cantidad de filas que tenga la grilla voy a saber si estoy
    'modificando un titular o creando uno nuevo.
    'Dependiendo del tipo de cuenta seleccionado, depende la información cargada en la grilla.
    'Si el tipo de cuenta es único,
            'las filas 0,1 y 2 son para el titular único.
            'la grilla contendrá 3 filas
    'Si el tipo de cuenta es separadas,
            'las filas 0,1 y 2 son para el titular del alojamiento
            'las filas 3,4, y 5 son para el titular de gastos extras
            'la grilla contendrá 6 filas

    If Me.chkTitularUnico.Value = True Or _
    Me.chkTitularAlojamiento.Value = True Then  'titular única o titular de alojamiento
        If Me.gTitulares.Rows > 0 Then
            'no hago nada porque estoy modificando
        Else
            'creo las tres primeras líneas
            Me.gTitulares.AddItem ""
            Me.gTitulares.AddItem ""
            Me.gTitulares.AddItem ""
        End If
    End If
    
    If Me.chkTitularExtras.Value = True Then                      'titular extras
        If Me.gTitulares.Rows > 3 Then
            'no hago nada porque estoy modificando
        Else
            If Me.gTitulares.Rows = 3 Then
                'si es igual a tres es porque ya creo el titular de alojamiento
                'y ahora estoy creando el titular de extras
                Me.gTitulares.AddItem ""
                Me.gTitulares.AddItem ""
                Me.gTitulares.AddItem ""
            Else
                If Me.gTitulares.Rows = 0 Then
                'es porque primero estoy creando el titular de extras antes que el de
                'alojamiento
                    Me.gTitulares.AddItem ""
                    Me.gTitulares.AddItem ""
                    Me.gTitulares.AddItem ""
                    Me.gTitulares.AddItem ""
                    Me.gTitulares.AddItem ""
                    Me.gTitulares.AddItem ""
                End If
            End If
        End If
    End If
End Sub

Private Function funObtengoFila() As Byte
    'Dependiendo del tipo de titular a cargar, depende la fila donde esta información
    'comienza a cargarse en la grilla.
    If Me.chkTitularUnico.Value = True Or _
        Me.chkTitularAlojamiento.Value = True Then
            funObtengoFila = 0
    Else
        If Me.chkTitularExtras.Value = True Then
            funObtengoFila = 3
        End If
    End If
End Function

Private Function funObtengoDescTipoTitular() As String
    'Devuelve un string con la descripción del tipo de titular recién creado.
    If Me.chkTitularUnico.Value Then
        funObtengoDescTipoTitular = "Titular único"
    Else
        If Me.chkTitularAlojamiento.Value Then
            funObtengoDescTipoTitular = "Titular del alojamiento"
        Else
            If Me.chkTitularExtras.Value Then
                funObtengoDescTipoTitular = "Titular de los gastos extras"
            End If
        End If
    End If
End Function

Private Function funObtengoDescTipoTitularOrigen() As String
    'Devuelve un string con la descripción del titular recien creado,
    'correspondiente al origen del mismo.
    Select Case Me.sstTipoTitular.Tab
        Case 1
            funObtengoDescTipoTitularOrigen = "Propio Pax"
        Case 2
            funObtengoDescTipoTitularOrigen = "OtroPax"
        Case 3
            funObtengoDescTipoTitularOrigen = "Empresa o Agencia"
        Case 4
            funObtengoDescTipoTitularOrigen = "Otros"
    End Select
End Function

Private Function funObtengoTipoTitularOrigen() As Byte
    'Devuelve 1 si el titular es el propio pax
    '2 si es otro pax
    '3 si es empresa o agencia
    '4 si es otro
    
    funObtengoTipoTitularOrigen = Me.sstTipoTitular.Tab
End Function

Private Sub botConfirmarCuenta_Click()
    'Cuando el usuario confirma un tipo de cuenta
    'inicializo la variable que determna que tipo de cuenta se seleccionó.
    'No puedo tomar como referencia el valor de los controles chk ya que los mismos
    'pueden ser modificados por el usuario luego de seleccionar una opción.
    
    If Me.chkCuentaUnica.Value = True Then
        gTipoCuenta = 0
    End If
    If Me.chkCuentasSeparadas.Value = True Then
        gTipoCuenta = 1
    End If
    'modifico apariecia controles formulario
    subInicializoAparienciaDeTrabajo
End Sub

Private Sub subInicializoAparienciaInicial()
    'Establese las propiedades de los controles del formulario al iniciar
    'su ejecución.
    'Es independiente de la tarea que realiza el formulario.
    
    Me.lblTipoCuentaSelTabs.Visible = False
    Me.lblTipoCuentaSel.Caption = "No hay tipo de cuenta seleccionada."
    Me.chkTitularAlojamiento.Visible = False
    Me.chkTitularExtras.Visible = False
    Me.chkTitularUnico.Visible = False
    Me.sstTipoTitular.TabEnabled(1) = False
    Me.sstTipoTitular.TabEnabled(2) = False
    Me.sstTipoTitular.TabEnabled(3) = False
    Me.sstTipoTitular.TabEnabled(4) = False
    Me.Frame2.Enabled = False
    Me.Frame3.Enabled = False
    Me.Frame4.Enabled = False
    Me.Frame5.Enabled = False
    Me.mnuBuscar.Enabled = False
End Sub

Private Sub subInicializoAparienciaDeTrabajo()
    'Establece las propiedades de los controles luego de haber seleccionado
    'un tipo de cuenta
    Dim descTipoCuenta As String
    Select Case gTipoCuenta
        Case 0  'unica
            'oculto titulares de tipo de cuenta separadas
            Me.chkTitularAlojamiento.Visible = False
            Me.chkTitularExtras.Visible = False

            Me.chkTitularUnico.Visible = True
            'luego de seleccionar tipo de cuenta le doy el focus al control de
            'selección de tipo de titular
            Me.chkTitularUnico.SetFocus
            descTipoCuenta = "Cuenta única"
        Case 1  'separadas
            'oculto titular de cuenta única
            Me.chkTitularUnico.Visible = False
            
            Me.chkTitularAlojamiento.Visible = True
            Me.chkTitularExtras.Visible = True
            'luego de seleccionar tipo de cuenta le doy el focus al control de
            'selección de tipo de titular
            Me.chkTitularAlojamiento.SetFocus
            descTipoCuenta = "Cuenta separadas"
    End Select
    
    Me.sstTipoTitular.TabEnabled(1) = True
    Me.sstTipoTitular.TabEnabled(2) = True
    Me.sstTipoTitular.TabEnabled(3) = True
    Me.sstTipoTitular.TabEnabled(4) = True
    Me.Frame2.Enabled = True
    Me.Frame3.Enabled = True
    Me.Frame4.Enabled = True
    Me.Frame5.Enabled = True

    Me.lblTipoCuentaSel.Caption = descTipoCuenta
    Me.lblTipoCuentaSelTabs.Visible = True
    'limpio grilla de titulares
    Me.gTitulares.Clear
    'habilito el primer tabs
    Me.sstTipoTitular.Tab = 0
End Sub

Private Sub subInicialioAparienciaSegunTarea()
    'Inicializo apariencia del formulario dependiendo de la tarea que realiza el mismo
    Dim x As Boolean
    Dim descTitulo As String
    
    Select Case propTipoAccionFormularioTitular
        Case 1
            x = True
            descTitulo = "Walkin: selección de titular"
            'obtengo la tarifa del tipo de habitación
            subObtengoTarifaTipoHab (propHabCuenta)
            
        Case 2 To 3
            x = True
            descTitulo = "Checkin: selección de titular"
            'obtengo la tarifa del tipo de habitación
            subObtengoTarifaTipoHab (propHabCuenta)
        Case 4
            x = False
            descTitulo = "Cambio de titular"
    End Select
    'muetro o oculto ingreso de tarifas
    Frame7.Visible = x
    'cambio titulo del formualario
    Me.Caption = descTitulo
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Descargo formulario de memoria
    Set frmTitularesHabitacion = Nothing
End Sub

Private Sub sstTipoTitular_Click(PreviousTab As Integer)
    'Este evento se ejecuta cuando se cambia la ficha de titulares
    
    'cambio título de la opción de ayuda en el menu de opciones
    subCambioMenu sstTipoTitular.Tab
    Select Case sstTipoTitular.Tab
        Case 1  'propio pax
            'inicializo etiqueta de grilla
            Me.lblPasajeros.Caption = "&Pasajero/s de la habitación " & CStr(propHabCuenta)
            'verifico si ya cargue la grilla con anterioridad
            If Me.gPropioPax.Rows = 1 Then
                'si no cargue la cargo por única vez
                'cargo con los pasajeros de la habitación
                SQLpasajeros_habitacion propHabCuenta, Data1
                'creo nuevamente el cabezal de la grilla
                subCreoCabezalGrilla
            End If
            'le doy el focus a la grilla para facilitar la selección con tecla enter
            gPropioPax.SetFocus
            
        Case 2  'otro pax
                'incializo control
                Me.txtOtroPax.Text = Empty
                Me.txtOtroPax.Tag = Empty
                'le doy el focus al botón de ayuda
                botAyuda(0).SetFocus
                'bloqueo control de ingreso
                mSubBloqueoControlFormulario Me.txtOtroPax, True
                
        Case 3  'emp. o agencia
                'inicializo control
                Me.txtEmpAge.Text = Empty
                Me.txtEmpAge.Tag = Empty
                'le doy el focus al botón de ayuda
                botAyuda(1).SetFocus
                'bloqueo control de ingreso
                mSubBloqueoControlFormulario Me.txtEmpAge, True
                
        Case 4  'otro
                'inicializo control
                Me.txtOtros.Text = Empty
                Me.txtOtros.Tag = Empty
                'le doy el focus al botón de ayuda
                botAyuda(2).SetFocus
                'bloqueo control de ingreso
                mSubBloqueoControlFormulario Me.txtOtros, True
    End Select
End Sub

Private Sub subCambioMenu(op As Byte)
    'Cambio el título de la opción de ayuda en el menu de opciones
    Select Case op
        Case 0  'ficha de selección
            Me.mnuBuscar.Enabled = False
        Case 1  'ficha de pasajeros
            Me.mnuBuscar.Enabled = False
        Case 2  'ficha de otro pax
            Me.mnuBuscar.Enabled = True
            Me.mnuBuscarTit.Caption = "Otro pax..."
        Case 3  'ficha de emp agencia
            Me.mnuBuscar.Enabled = True
            Me.mnuBuscarTit.Caption = "Agencia o empresa..."
        Case 4  'ficha de otros clientes
            Me.mnuBuscar.Enabled = True
            Me.mnuBuscarTit.Caption = "Otros clientes..."
    End Select
End Sub

Private Sub subCreoCabezalGrilla()
    'Creo el cabezal de la grilla después de ejecutar la consulta
    Me.gPropioPax.FormatString = "   |Nombre pasajero                                                    |   "
    'oculto la segunda columna ya que en la misma se almacena el número de cliente
    Me.gPropioPax.ColWidth(2) = 0
End Sub

Private Sub subMejoroPresentacionGrilla(fila As Byte)
    'Cambio los atributos del texto de la grilla para mejorar presentación
    
    'cambio la fuente del nombre de titular a negrita
    Me.gTitulares.Row = fila + 2
    Me.gTitulares.col = 0
    Me.gTitulares.CellFontBold = True
End Sub

Private Sub txttarifa_KeyPress(KeyAscii As Integer)
    'Valido solo se ingresen números
    ValidoNum KeyAscii, True, True
End Sub

Private Sub mnuFormularioAceptar_Click()
    'Equivale a presionar la tecla F12 o el boton de aceptar
    botAceptar_Click
End Sub

Private Sub mnuFormularioCancelar_Click()
    'Equivale a presionar boton cancelar o la tecla Esc
    botCancelar_Click
End Sub

Private Sub mnuFormularioConfirmarTit_Click()
    'Equivale a presionar la tecla F9
    'Dependiendo de la ficha que este visible, dependerá la ayuda que se ejecute
    If Me.sstTipoTitular.Tab > 0 Then   'no ejecuto para la ficha de selección de tipo
        'verifico si se permite trabajar con la ficha
        If Me.sstTipoTitular.TabEnabled(Me.sstTipoTitular.Tab) Then
            'como el índice del boton no coincide con el índice del tabs
            'tengo que restarle 1 al índice del tab
            botConfirmar_Click (Me.sstTipoTitular.Tab - 1)
        End If
    End If
End Sub

Private Sub mnuBuscarTit_Click()
    'Dependiendo de la ficha que este visible, dependerá la ayuda que se ejecute
    If Me.sstTipoTitular.Tab > 1 Then   'no ejecuto para ficha de selección de tipo
                                        'ni de propio pax
        'verifico si se permite trabajar con la ficha
        If Me.sstTipoTitular.TabEnabled(Me.sstTipoTitular.Tab) Then
            'como el índice del boton no coincide con el índice del tabs
            'tengo que restarle 2 al índice del tab
            BotAyuda_Click (Me.sstTipoTitular.Tab - 2)
        End If
    End If
End Sub

Private Sub mnuVerEmpAgen_Click()
    'Muestro ficha de agencia empresa. Equivale a tecla F7
    If Me.sstTipoTitular.TabEnabled(3) Then
        Me.sstTipoTitular.Tab = 3
    End If
End Sub

Private Sub mnuVerOtro_Click()
    'Muestro ficha de otro pax.Equivale a tecla F6
    If Me.sstTipoTitular.TabEnabled(2) Then
        Me.sstTipoTitular.Tab = 2
    End If
End Sub

Private Sub mnuVerOtros_Click()
    'Muestro ficha de otros clientes.Equivale a tecla F8
    If Me.sstTipoTitular.TabEnabled(4) Then
        Me.sstTipoTitular.Tab = 4
    End If
End Sub

Private Sub mnuVerPropioPax_Click()
    'Muestro ficha de propio pax.Equivale a tecla F5
    If Me.sstTipoTitular.TabEnabled(1) Then
        Me.sstTipoTitular.Tab = 1
    End If
End Sub

Private Sub mnuVerTipo_Click()
    'Muestro ficha de seleccionar tipo
    Me.sstTipoTitular.Tab = 0
End Sub

Private Sub botCancelar_Click()
    'Si cancelo esta operación y el formulario es llamdo desde checkin,
    'muestro nuevamente dicho formulario, ya que en realidad estoy dando un paso atras.
    If Me.propTipoAccionFormularioTitular = 1 Or _
        Me.propTipoAccionFormularioTitular = 2 Or _
        Me.propTipoAccionFormularioTitular = 3 Then
        Unload Me
        frmCheck_in.Show 1
    Else
        'aviso de confirmación de salir
        If mFunMensaje(4, 78) Then
            'Si estoy cambiando un titular, regreso al formulario de
            'ingreso de habitaciones
            Unload Me
            frmIngHabitacion.Show 1
        End If
    End If
End Sub

'******************************************************
'*
'*  Procedimientos que realizan el cambio de titular
'*
'******************************************************

Private Sub botAceptar_Click()
    'El usuario confirma el cambio de titular
    If funValidoTitulares Then
        If busco_habitaTF(propHabCuenta) Then
            'pido confirmación del usuario
            If mFunMensaje(4, 79) Then
                'ejecuto los procedimientos correspondientes
                'dependiendo de la accion para la cual es llamdo este formulario,
                Select Case propTipoAccionFormularioTitular
                    Case 1 To 3 'nuevo titular (walkin,checkn sin asignación, checkin con asignacion)
                        'grabo tarifa para la habitación
                        subGraboTarifaHabitacion
                        'modifico reserva de la habitación
                        subGraboReserva
                    Case 4      'cambio titulares
                        'cambio los gastos
                        subCambioGastos
                        subInicializoTitularesAnteriores
                End Select
                    
                'asigno nuevos titulares a la habitación
                If funAsignoNuevosTitularesHabitacion Then
                    'la asignación se realizó correctamente
                    'grabo bitácora
                    
                    'aviso de asignación de titular
                    mSubMensaje 4, 73
                    Unload Me
                    
                    'descargo el resto de los formularios
                    Unload frmCheck_in
                    Unload frmCargaReserva  'este formulario no es cargado cuando se efectúa un walkin
                                            'pero para simplificar el código, se descarga de todas maneras.
                    If propTipoAccionFormularioTitular = 4 Then
                        'Si estoy cambiando un titular, regreso al formulario de
                        'ingreso de habitaciones
                        frmIngHabitacion.Show 1
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Function funAsignoNuevosTitularesHabitacion() As Boolean
    'Con la información que se encuentra en la grilla de titulares, grabo el archivo de
    'habitaciones.
    Dim titAloja As Long
    Dim titExtras As Long
    
    funAsignoNuevosTitularesHabitacion = True
    'determino el tipo de cuenta seleccionada.
    If gTipoCuenta = 0 Then         'cuenta única
        subGraboTitularHabitacion gTitulares.TextMatrix(2, 1), _
                                    gTitulares.TextMatrix(1, 1), 1
    Else
        If gTipoCuenta = 1 Then     'cuenta separada
            titAloja = CLng(gTitulares.TextMatrix(2, 1))
            titExtras = CLng(gTitulares.TextMatrix(5, 1))
            'verifico si los titulares de alojamiento y de extras son iguales
            If titAloja = titExtras Then
                'El titular para la cuanta de alojamiento y de gastos extras, es la misma
                'persona. Cambie uno de los dos titulares o seleccione tipo de cuentas única
                mSubMensaje 4, 74
                    funAsignoNuevosTitularesHabitacion = False
            Else
                'titular alojamiento
                subGraboTitularHabitacion titAloja, _
                                        gTitulares.TextMatrix(1, 1), 2
                'titular extras
                subGraboTitularHabitacion titExtras, _
                                         gTitulares.TextMatrix(4, 1), 3
            End If
        End If
    End If
End Function

Private Sub subGraboTitularHabitacion(tit As Long, cue As Byte, tipo As Byte)
    'tit: contiene el número que identifica a el titular.
    'cue: contiene el tipo del titular para cada tipo de cuenta, ejemplo:
        '1 propio pax
        '2 otro pax
        '3 emp agencia
        '4 otro
        '0 nada (ver explicación de porque se puede dar este valor en subInicializoTitularesAnteriores)
    'tipo: contiene el tipo de titular que estoy modificando, ejemplo:
        '1 titular unico
        '2 titular alojamiento
        '3 titular extras
        
    tbHABITACIONES.Edit
    Select Case tipo
        Case 1
            tbHABITACIONES("titular_unica") = tit
            tbHABITACIONES("tipocuenta_unica") = cue
        Case 2
            tbHABITACIONES("tipocuenta_aloja") = cue
            tbHABITACIONES("titular_aloja") = tit
        Case 3
            tbHABITACIONES("tipocuenta_extra") = cue
            tbHABITACIONES("titular_extra") = tit
    End Select
    tbHABITACIONES.Update
End Sub

Private Function funValidoTitulares()
    'Valido que se hallan ingresado los titulares antes de seguir adelante.
    'La validación se realiza mediante dos formas:
    'primero se valida que las filas de la grilla para cada caso, esten creadas.
    'segundo, si es así, se valida que contengan datos.
    
    Dim codErr As Byte
    Dim codDescErr As Integer
    
    'por defecto asumo que estan bien
    funValidoTitulares = True
    codErr = 0
    If gTipoCuenta = 0 Then             'cuenta única
        'solo valido si se ingreso titular únicio
        If Me.gTitulares.Rows > 0 Then
            If Me.gTitulares.TextMatrix(2, 1) = Empty Then
                'no hay titular ingresado
                codErr = 1
            End If
        Else
            'no hay titular ingresado
            codErr = 1
        End If
    Else
        If gTipoCuenta = 1 Then         'cuentas separadas
            If gTitulares.Rows > 0 Then
                'valido si se ingreso titular de alojamiento
                If Me.gTitulares.TextMatrix(2, 1) = Empty Then
                    'no hay titular de alojamiento ingresado
                    codErr = 2
                Else
                    If gTitulares.Row > 2 Then
                        If Me.gTitulares.TextMatrix(5, 1) = Empty Then
                            'no hay titular de gastos extras ingresado
                            codErr = 3
                        End If
                    Else
                        'no hay titular de gastos extras ingresado
                        codErr = 3
                    End If
                End If
            Else
                'no hay titular de alojamiento ingresado
                codErr = 2
            End If
        End If
    End If
    'valido si hay error
    If codErr > 0 Then
        Select Case codErr
            Case 1
                'debe de ingresar titular único
                codDescErr = 75
            Case 2
                'debe de ingresar titular de alojamiento
                codDescErr = 76
            Case 3
                'debe de ingresasr titular de gastos extras
                codDescErr = 77
        End Select
        funValidoTitulares = False
        mSubMensaje 4, codDescErr
        'establesco la ficha de selección como visible
        Me.sstTipoTitular.Tab = 0
    End If
End Function

'**********************************************************************
'*
'*  Procedimientos efectuados cuando asigno nuevo titular
'*  proptipoAccionFormularioTitulares = 1, 2 o 3
'*
'***********************************************************************

Private Sub subGraboReserva()
    'Existen los casos en los que la reserva se realiza sin asignar habitación, dejando
    'para realizar la misma, en el momento de realizar el checkin.
    'Cuando confirmo el Checkin tengo que:
    '1)grabar el número de habitación en el archivo HAB_RESERVA,
    '2)y cambiar el estado a Asignada;
    'Esto es importante para que en el caso de consultar la reserva,
    'me indique que nro. de habitación fue asignado ya que originalmente se muestra la
    'descripción de no asignada.
    'También (y esto es lo más importante) al ingresar por checkin con el número de reserva
    'ya no aparecerá como una reserva no asignada en checkin, por lo que no se permitirá
    'asignarle nuevamente otro número de habitación.
    'No reaizo esta operación si llamo a este formulario después de:
    '   realizar un walkin                  1
    '   checkin con habitación asignada     3
    '   cambio de titular                   4
    Dim nro_corr As Long
    
    'verifico si la operación es correcta
    If propTipoAccionFormularioTitular = 2 Then 'checkin sin asignación
        tbHAB_RESERVAS.Index = "ihab_reserva"
        tbHAB_RESERVAS.Seek "=", propHabNroReserva, propHabNroCorr
        If Not tbHAB_RESERVAS.NoMatch Then
            tbHAB_RESERVAS.Edit
                tbHAB_RESERVAS("nrohabitacion") = propHabCuenta
                tbHAB_RESERVAS("descri_estado") = "Asignada"
            tbHAB_RESERVAS.Update
        End If
    End If
End Sub

Private Sub subGraboTarifaHabitacion()
    'Cuando realizo el checkin de un habitación, el último paso es asignarle un tipo de cuenta,
    'es decir, un titular, en este último paso también existe la posibilidad de cambiar el precio de la tarifa, a otro
    'diferente al preestablecido para el tipo de habitación con la cual se esta trabajando.
    tbHABITACIONES.Edit
        tbHABITACIONES("tarifa") = Val(txttarifa.Text)
    tbHABITACIONES.Update
End Sub

Private Sub subObtengoTarifaTipoHab(hab As Long)
    'Obtiene la tarifa correspondiente al tipo de la habitación a la cual se
    'le realiza el ingreso.
    
    'busco habitación
    If busco_habitaTF(propHabCuenta) Then
        'busco tarifa del tipo
        txttarifa.Text = mFunBuscoTarifaHab(tbHABITACIONES("tipohab"))
    End If
End Sub

'**********************************************************************
'*
'*  Procedimientos efectuados cuando asigno cambio de titular
'*  proptipoAccionFormularioTitulares = 4
'*
'***********************************************************************

Private Sub subInicializoTitularesAnteriores()
    'Si estoy cambiando de titular es necesario inicializar los registros que
    'contienen información de los titulares anteriores.
    'Cuando se inicializan los tres titulares como en este caso se esta indicando que
    'la habitación no tiene titular asignado.
    'Además de este caso (en el cual la habitación no tiene asignado titular) se puede dar:
    'a)   tbHABITACIONES("tipocuenta_unica") <> 1,2,3 o 4
    '     tbHABITACIONES("tipocuenta_extra") = 0
    '     tbHABITACIONES("tipocuenta_aloja") = 0            La habtación tiene tipo de cuenta unica.
    
    'b)   tbHABITACIONES("tipocuenta_unica") = 0
    '     tbHABITACIONES("tipocuenta_extra") <> 1,2,3 o 4
    '     tbHABITACIONES("tipocuenta_aloja") <> 1,2,3 o 4    La habtación tiene tipo de cuenta separadas.
    
    
    'inicializo titular único
    subGraboTitularHabitacion 0, 0, 1
    'inicializo titular aloja
    subGraboTitularHabitacion 0, 0, 2
    'iniciaizo titular extras
    subGraboTitularHabitacion 0, 0, 3
End Sub

Private Sub subCambioGastos()
    'Los gastos de la habitación que antes pertenecían a otro titular
    'ahora se cambian para el nuevo.
    Dim consulta1 As String
    Dim consulta2 As String
    If gTipoCuenta = 0 Then                          'cuenta única
        'cambio extras
         consulta1 = "UPDATE cuentas_extra SET titular_cuenta =" & funObtengoTitUnica & _
                    " WHERE habitacion_cuenta = " & propHabCuenta
        'cambio aloja
         consulta2 = "UPDATE cuentas_aloja SET titular_aloja =" & funObtengoTitUnica & _
                    " WHERE habitacion_cuenta_aloja = " & propHabCuenta
    Else
        If gTipoCuenta = 1 Then                        'cuenta separadas
            'cambio extras
             consulta1 = "UPDATE cuentas_extra SET titular_cuenta =" & funObtengoTitExtra & _
                        " WHERE habitacion_cuenta = " & propHabCuenta
            'cambio aloja
             consulta2 = "UPDATE cuentas_aloja SET titular_aloja =" & funObtengoTitAloja & _
                        " WHERE habitacion_cuenta_aloja = " & propHabCuenta
        End If
    End If
    bdHOTEL.Execute consulta1
    bdHOTEL.Execute consulta2
End Sub

Private Function funObtengoTitUnica() As String
    'Devuelvo el número de titular unica ingresado en la grilla
    funObtengoTitUnica = Me.gTitulares.TextMatrix(2, 1)
End Function

Private Function funObtengoTitExtra() As String
    'Devuelvo el número de titular extra ingresado en la grilla
    funObtengoTitExtra = Me.gTitulares.TextMatrix(5, 1)
End Function

Private Function funObtengoTitAloja() As String
    'Devuelvo el número de titular aloja ingresado en la grilla
    funObtengoTitAloja = Me.gTitulares.TextMatrix(2, 1)
End Function

'**********************************************************
'*
'*  Asistencia a usuarios
'*
'**********************************************************
    
Private Sub chkCuentaUnica_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 139
End Sub

Private Sub botAceptar_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 1, 2
End Sub

Private Sub chkCuentasSeparadas_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 139
End Sub

Private Sub botConfirmarCuenta_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 140
End Sub

Private Sub botCancelar_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 1, 3
End Sub

Private Sub txttarifa_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 144
End Sub

Private Sub chkTitularUnico_GotFocus()
    'Al darle el focus a este control,visualizo el tabs correspondiente.
    Me.sstTipoTitular.Tab = 0
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 141
End Sub

Private Sub chkTitularAlojamiento_GotFocus()
    'Al darle el focus a este control,visualizo el tabs correspondiente.
    Me.sstTipoTitular.Tab = 0
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 141
End Sub

Private Sub chkTitularExtras_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 141
End Sub

Private Sub gPropioPax_GotFocus()
    'Al darle el focus a este control,visualizo el tabs correspondiente.
    Me.sstTipoTitular.Tab = 1
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 142
End Sub

Private Sub botConfirmar_GotFocus(Index As Integer)
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 143
End Sub

Private Sub botAyuda_GotFocus(Index As Integer)
    'Cada vez que le doy el focus a estos controles, visualizo el tab correspondiente.
    Select Case Index
        Case 0  'otro pax
            Me.sstTipoTitular.Tab = 2
            mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 145
        Case 1  'empresa o agencia
            Me.sstTipoTitular.Tab = 3
            mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 146
        Case 2  'otro cliente
            Me.sstTipoTitular.Tab = 4
            mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 147
    End Select
End Sub

Private Sub botAyuda_LostFocus(Index As Integer)
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub botConfirmar_LostFocus(Index As Integer)
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub botCancelar_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub botAceptar_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub chkCuentaUnica_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub chkCuentasSeparadas_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub botConfirmarCuenta_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub txttarifa_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub gPropioPax_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub chkTitularAlojamiento_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub chkTitularExtras_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub chkTitularUnico_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

