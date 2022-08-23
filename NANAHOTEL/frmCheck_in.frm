VERSION 5.00
Object = "{EB8C3860-8603-11D0-A35D-7062E4000000}#1.1#0"; "VCBND.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCheck_in 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Check-In"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   11910
   ClipControls    =   0   'False
   ControlBox      =   0   'False
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
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   7605
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   582
   End
   Begin VB.Frame Frame4 
      Caption         =   "Datos Check-In"
      Height          =   2775
      Left            =   120
      TabIndex        =   54
      Top             =   0
      Width           =   4335
      Begin TabDlg.SSTab SSTab1 
         Height          =   2415
         Left            =   120
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   240
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   4260
         _Version        =   327680
         Style           =   1
         Tabs            =   2
         TabHeight       =   520
         ForeColor       =   16711680
         TabCaption(0)   =   "Check-In"
         TabPicture(0)   =   "frmCheck_in.frx":0000
         Tab(0).ControlCount=   1
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame1"
         Tab(0).Control(0).Enabled=   0   'False
         TabCaption(1)   =   "Walk-In"
         TabPicture(1)   =   "frmCheck_in.frx":001C
         Tab(1).ControlCount=   1
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame5"
         Tab(1).Control(0).Enabled=   0   'False
         Begin VB.Frame Frame5 
            BorderStyle     =   0  'None
            Height          =   1935
            Left            =   -74880
            TabIndex        =   65
            Top             =   360
            Width           =   3855
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
               Left            =   1320
               Style           =   2  'Dropdown List
               TabIndex        =   6
               Top             =   840
               Width           =   1815
            End
            Begin VB.TextBox txtHabWalk 
               BackColor       =   &H80000016&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   375
               Left            =   1320
               Locked          =   -1  'True
               TabIndex        =   66
               TabStop         =   0   'False
               Top             =   1440
               Width           =   735
            End
            Begin VB.CommandButton botHabitacionWalk 
               Caption         =   "&Selección"
               Height          =   375
               Left            =   2160
               TabIndex        =   7
               Top             =   1440
               Width           =   1215
            End
            Begin VB.CommandButton botSeguirWalk 
               Appearance      =   0  'Flat
               Caption         =   ">"
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
               Left            =   3480
               TabIndex        =   8
               Top             =   1440
               Width           =   375
            End
            Begin VcBndCtl.VcCalCombo fEgreso 
               Height          =   375
               Left            =   1320
               TabIndex        =   4
               Top             =   240
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   661
               _0              =   $"frmCheck_in.frx":0038
               _1              =   $"frmCheck_in.frx":0441
               _2              =   $"frmCheck_in.frx":084A
               _3              =   "-@B@@@@@%@@@C@@@@@@@D@@@A@@@)f@@@E@@@A@@@@@@@@@'@@@E@@@A@@@@@@@@@)h@@@E@@@A@@@@@@@@@,1E1B"
               _count          =   4
               _ver            =   2
            End
            Begin VB.Label Label19 
               Caption         =   "&Tipo de habitación"
               Height          =   480
               Left            =   120
               TabIndex        =   5
               Top             =   720
               Width           =   1200
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "&Egreso"
               Height          =   240
               Left            =   120
               TabIndex        =   3
               Top             =   300
               Width           =   660
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Habitación"
               Height          =   240
               Left            =   120
               TabIndex        =   49
               Top             =   1500
               Width           =   975
            End
         End
         Begin VB.Frame Frame1 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   1935
            Left            =   120
            TabIndex        =   58
            Top             =   450
            Width           =   3855
            Begin MSFlexGridLib.MSFlexGrid gHabitaciones 
               Height          =   1050
               Left            =   0
               TabIndex        =   0
               Top             =   360
               Width           =   3735
               _ExtentX        =   6588
               _ExtentY        =   1852
               _Version        =   393216
               Rows            =   0
               Cols            =   4
               FixedRows       =   0
               FixedCols       =   0
               BackColorBkg    =   16777215
               GridColor       =   8421504
               FocusRect       =   0
               GridLines       =   0
               GridLinesFixed  =   0
               ScrollBars      =   2
               SelectionMode   =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin VB.TextBox fechaingreso 
               BackColor       =   &H80000000&
               ForeColor       =   &H00C00000&
               Height          =   285
               Left            =   840
               Locked          =   -1  'True
               TabIndex        =   61
               TabStop         =   0   'False
               Top             =   0
               Width           =   975
            End
            Begin VB.TextBox fechaegreso 
               BackColor       =   &H80000000&
               ForeColor       =   &H00C00000&
               Height          =   285
               Left            =   2760
               Locked          =   -1  'True
               TabIndex        =   60
               TabStop         =   0   'False
               Top             =   0
               Width           =   975
            End
            Begin VB.TextBox txtHabCheck 
               BackColor       =   &H80000016&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   375
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   59
               TabStop         =   0   'False
               Top             =   1455
               Width           =   855
            End
            Begin VB.CommandButton botSelHabCheck 
               Caption         =   "&Selección"
               Height          =   375
               Left            =   2040
               TabIndex        =   1
               Top             =   1440
               Width           =   1215
            End
            Begin VB.CommandButton botSeguir 
               Caption         =   ">"
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
               Left            =   3360
               TabIndex        =   2
               Top             =   1440
               Width           =   375
            End
            Begin VB.Label Label13 
               Caption         =   "Habitación"
               Height          =   255
               Left            =   0
               TabIndex        =   64
               Top             =   1530
               Width           =   1095
            End
            Begin VB.Label Label1 
               Caption         =   "Ingreso"
               Height          =   255
               Left            =   0
               TabIndex        =   63
               Top             =   0
               Width           =   735
            End
            Begin VB.Label Label2 
               Caption         =   "Egreso"
               Height          =   255
               Left            =   2040
               TabIndex        =   62
               Top             =   15
               Width           =   735
            End
         End
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Pasajeros alo&jados por habitación"
      Height          =   2775
      Left            =   4560
      TabIndex        =   52
      Top             =   0
      Width           =   7215
      Begin MSFlexGridLib.MSFlexGrid dbgrid1 
         Bindings        =   "frmCheck_in.frx":0C53
         Height          =   2055
         Left            =   240
         TabIndex        =   53
         Top             =   360
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   3625
         _Version        =   393216
         Cols            =   4
         AllowBigSelection=   0   'False
         FocusRect       =   2
         HighLight       =   0
         ScrollBars      =   2
         SelectionMode   =   1
         Appearance      =   0
         MousePointer    =   2
         FormatString    =   $"frmCheck_in.frx":0C63
      End
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   10200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   "select cliente.nombre_completo_titular,checkin.nrocorrcli from checkin, clientes where checkin.nrocorrcli = clientes.nrocorr"
      Top             =   4080
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos extras"
      Height          =   4815
      Left            =   120
      TabIndex        =   50
      Top             =   2880
      Width           =   11655
      Begin VB.ComboBox cboTipoDocu 
         Height          =   360
         Left            =   3960
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   3120
         Width           =   975
      End
      Begin VB.TextBox fechanac 
         Height          =   285
         Left            =   5280
         MaxLength       =   15
         TabIndex        =   28
         Top             =   2640
         Width           =   1095
      End
      Begin VB.ComboBox cboNacionalidad 
         Height          =   360
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   3120
         Width           =   1815
      End
      Begin VB.TextBox txt1er_Nom_titular 
         Height          =   315
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   10
         Top             =   240
         Width           =   3255
      End
      Begin VB.TextBox txt2do_Nom_titular 
         Height          =   315
         Left            =   6480
         MaxLength       =   50
         TabIndex        =   12
         Top             =   240
         Width           =   3375
      End
      Begin VB.TextBox txt1er_Ape_titular 
         Height          =   315
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   14
         Top             =   720
         Width           =   3255
      End
      Begin VB.TextBox txt2do_Ape_titular 
         Height          =   315
         Left            =   6480
         MaxLength       =   50
         TabIndex        =   16
         Top             =   720
         Width           =   3375
      End
      Begin VB.CommandButton botayuda 
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
         Height          =   825
         Left            =   10200
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton botConfirmaCheckin 
         Caption         =   "Continuar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   10200
         TabIndex        =   47
         Tag             =   "Continuar"
         Top             =   3840
         Width           =   1215
      End
      Begin VB.CommandButton botCancelar 
         Cancel          =   -1  'True
         Caption         =   "Salir"
         Height          =   375
         Left            =   10200
         TabIndex        =   48
         Top             =   4320
         Width           =   1215
      End
      Begin VB.TextBox txtRuc 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   7320
         MaxLength       =   12
         TabIndex        =   43
         Top             =   3120
         Width           =   2535
      End
      Begin VB.CommandButton botConfirmaIngreso 
         Caption         =   "C. pasajero"
         Height          =   375
         Left            =   10200
         TabIndex        =   46
         Top             =   3360
         Width           =   1215
      End
      Begin VB.ComboBox cboEstadoCivil 
         Height          =   360
         Left            =   8160
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   2640
         Width           =   1695
      End
      Begin VB.ComboBox cboSexo 
         Height          =   360
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   2640
         Width           =   1815
      End
      Begin VB.TextBox txtEmail 
         Height          =   285
         Left            =   7320
         MaxLength       =   26
         TabIndex        =   39
         Top             =   2160
         Width           =   2535
      End
      Begin VB.TextBox faxcheckin 
         Height          =   285
         Left            =   4200
         MaxLength       =   15
         TabIndex        =   24
         Top             =   2160
         Width           =   2175
      End
      Begin VB.TextBox txtTele 
         Height          =   285
         Left            =   1440
         MaxLength       =   15
         TabIndex        =   22
         Top             =   2160
         Width           =   1815
      End
      Begin VB.TextBox txtCodPostal 
         Height          =   285
         Left            =   8280
         MaxLength       =   15
         TabIndex        =   35
         Top             =   1200
         Width           =   1575
      End
      Begin VB.ComboBox cboPais 
         Height          =   315
         Left            =   7320
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   1680
         Width           =   2535
      End
      Begin VB.TextBox txtLocalidad 
         Height          =   285
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   20
         Top             =   1680
         Width           =   4935
      End
      Begin VB.TextBox txtDirecion 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1440
         MaxLength       =   55
         TabIndex        =   18
         Top             =   1200
         Width           =   4935
      End
      Begin VB.TextBox txtDocu 
         Height          =   360
         Left            =   4920
         MaxLength       =   16
         TabIndex        =   33
         Top             =   3120
         Width           =   1455
      End
      Begin VB.TextBox txtObservaciones 
         Height          =   855
         Left            =   240
         MaxLength       =   255
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         ScrollBars      =   2  'Vertical
         TabIndex        =   45
         Top             =   3840
         Width           =   9615
      End
      Begin VB.Label nacionalidadcheck 
         Caption         =   "&Nacionalidad"
         Height          =   240
         Left            =   240
         TabIndex        =   29
         Top             =   3150
         Width           =   1095
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "&1er Nombre"
         Height          =   240
         Left            =   240
         TabIndex        =   9
         Top             =   300
         Width           =   1065
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "1er &Apellido"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   780
         Width           =   825
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "&2do Nombre"
         Height          =   195
         Left            =   5040
         TabIndex        =   11
         Top             =   270
         Width           =   870
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "2&do Apellido"
         Height          =   195
         Left            =   5040
         TabIndex        =   15
         Top             =   750
         Width           =   870
      End
      Begin VB.Label Label9 
         Caption         =   "&R.U.C"
         Height          =   255
         Index           =   0
         Left            =   6600
         TabIndex        =   42
         Top             =   3135
         Width           =   615
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Estado ci&vil"
         Height          =   240
         Left            =   6600
         TabIndex        =   40
         Top             =   2677
         Width           =   1035
      End
      Begin VB.Label fechanaccheck 
         AutoSize        =   -1  'True
         Caption         =   "Fec&ha nacimiento"
         Height          =   240
         Left            =   3480
         TabIndex        =   27
         Top             =   2670
         Width           =   1590
      End
      Begin VB.Label Label11 
         Caption         =   "Label9"
         Height          =   15
         Left            =   480
         TabIndex        =   51
         Top             =   5040
         Width           =   495
      End
      Begin VB.Label sexocheck 
         Caption         =   "Se&xo"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   2670
         Width           =   615
      End
      Begin VB.Label otrocheck 
         Caption         =   "E-&mail"
         Height          =   255
         Left            =   6600
         TabIndex        =   38
         Top             =   2175
         Width           =   615
      End
      Begin VB.Label faxcheck 
         AutoSize        =   -1  'True
         Caption         =   "Fa&x"
         Height          =   240
         Left            =   3480
         TabIndex        =   23
         Top             =   2160
         Width           =   330
      End
      Begin VB.Label Telcheck 
         AutoSize        =   -1  'True
         Caption         =   "Teléf&ono"
         Height          =   240
         Left            =   240
         TabIndex        =   21
         Top             =   2175
         Width           =   810
      End
      Begin VB.Label Codpostal 
         AutoSize        =   -1  'True
         Caption         =   "Cod.Po&stal"
         Height          =   240
         Left            =   7200
         TabIndex        =   34
         Top             =   1200
         Width           =   990
      End
      Begin VB.Label paischeck_in 
         Caption         =   "&País"
         Height          =   255
         Left            =   6600
         TabIndex        =   36
         Top             =   1710
         Width           =   495
      End
      Begin VB.Label poblacioncheck_in 
         AutoSize        =   -1  'True
         Caption         =   "&Localidad"
         Height          =   240
         Left            =   240
         TabIndex        =   19
         Top             =   1695
         Width           =   900
      End
      Begin VB.Label direcheck_in 
         AutoSize        =   -1  'True
         Caption         =   "D&irección"
         Height          =   240
         Left            =   240
         TabIndex        =   17
         Top             =   1215
         Width           =   855
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Do&c."
         Height          =   240
         Left            =   3480
         TabIndex        =   31
         Top             =   3135
         Width           =   420
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "O&bservaciones"
         Height          =   240
         Index           =   1
         Left            =   240
         TabIndex        =   44
         Top             =   3600
         Width           =   1380
      End
   End
   Begin VB.Menu mnuFormulario 
      Caption         =   "&Formulario"
      Begin VB.Menu mnuFormularioContinuar 
         Caption         =   "Continuar"
         Enabled         =   0   'False
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuFormularioSalir 
         Caption         =   "Salir"
      End
      Begin VB.Menu mnuDiv1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFormularioconfirmar 
         Caption         =   "Confirmar habitación                        F9"
      End
      Begin VB.Menu mnuFormularioConfirmarPasajero 
         Caption         =   "Confirmar pasajero                           F9"
      End
   End
   Begin VB.Menu mnuIr 
      Caption         =   "&Ir a..."
      Begin VB.Menu mnuIrVerDispo 
         Caption         =   "Ver disponibilidad"
         Shortcut        =   ^D
      End
   End
   Begin VB.Menu mnuBuscar 
      Caption         =   "Buscar..."
      Enabled         =   0   'False
      Begin VB.Menu mnuBuscarClientes 
         Caption         =   "Clientes..."
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmCheck_in"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private nrocliente As Long          'Utilizada para trabajar con los clientes que ingresan
                                    'a las habitaciones
Private ayuda As Boolean

Private Sub Form_Activate()
    'Cambio la propiedad enabled de los controles de la ficha que no se usa
    'para facilitar la interfaz del usuario. Si no se hace esto se le está
    'dando el focus a controles que nunca se acceden.
        
    If tipo_accion_checkin = 1 Then       'check-in
        SSTab1.Tab = 0
        SSTab1.TabEnabled(1) = False
        'cambio propiedad enabled de los controles del walkin
        Me.Frame5.Enabled = False
    Else                                    'walk-inOcupada o walk-inLibre
        SSTab1.Tab = 1
        SSTab1.TabEnabled(0) = False
        'cambio propiedad enabled de los controles del checkin
        Me.Frame1.Enabled = False
    End If
End Sub

Private Sub Form_Load()
    Dim res_aux As String
    
    'Apariencia formulario
    mSubConfiguroFuentesControlesSistema Me

    configuracion_apariencia

    ayuda = False
    'dbgrid1.Font.Bold = True
   
    'inicializo control data
    subInicializoControlData Me.Data2
    Data2.RecordSource = "select cliente.nombre_completo_titular,checkin.nrocorrcli from checkin, clientes where checkin.nrocorrcli = clientes.nrocorr"
    
    carga_tipo_sexo cboSexo
    carga_tipo_estadocivil cboEstadoCivil
    carga_tipo_nacionalidad cboNacionalidad
    carga_tipo_pais cboPais
    mSubCargoComboConstantes 2, Me.cboTipoDocu
    'cargo tipo habitación
    carga_tipo_hab frmCheck_in.cboTipo_habitacion

    desabilito_datos_cliente True
    
    If tipo_accion_checkin = 1 Then 'checkin
        res_aux = Mid(Str(nro_reserva), 1, 5) + "-" + Mid(Str(nro_reserva), 6, 10)
        frmCheck_in.Caption = "Check-In: se está utilizando la reserva Nº " & res_aux
    
        fechaingreso.Text = tbRESERVAS("fechaing")
        fechaegreso.Text = tbRESERVAS("fechaegr")
    
        cargo_ventana_habitaciones
        gHabitaciones_Click
    End If
    
    If tipo_accion_checkin = 2 Then 'walkinL
        frmCheck_in.Caption = "Walk-In"
    End If
    
    If tipo_accion_checkin = 3 Then 'walkinO
        frmCheck_in.Caption = "Walk-In a una habitación ocupada"
        'en un walkinO no es necesario ingresar fecha de ingreso
        'ni ver cuadro de disponibilidad
        mSubBloqueoControlFormulario Me.fEgreso, True
        Me.mnuIr.Enabled = False
    End If
    cargo_ventana_cli   'no mover de aqui, sino se produce error de too few parameters!!!!
End Sub

Private Sub botSeguir_Click()
    Dim valido As Boolean
    Dim nro_hab As Long
    
    valido = True
    If obtengo_nro_hab_ventanaHab = 0 Then
        If Val(txtHabCheck.Text) = 0 Then
            'debe de seleccionar habitacion
            mSubMensaje 4, 62
            valido = False
        End If
    End If
    
    If valido Then
        'obtengo número de habitación
        nro_hab = obtengo_nro_hab_ventanaHab
        If nro_hab = 0 Then
            nro_hab = txtHabCheck.Text
        End If
        'si la habitación a la cual voy a ingresar a los pasajeros está ocupada
        'es necesario determinar si el período de ocupación es el correcto.
        'Puede que no lo sea, en el caso de que la habitación a la cual se desea ingresar
        'no halla quedado libre (realizado checkout) en la fecha establecida.
        If Not mFunDeterminoOcupacionValida(nro_hab) Then
            'el período NO esta dentro de lo previsto
            mSubMensaje 4, 130, _
            "Esto se debe a que no se realizó el Check-Out en la fecha establecida."
        Else
            'el período esta dentro del establecido o la habitación esta libre.
            desabilito_datos_cliente False
            desabilito_habitacion 0
            limpio_cliente
            'le doy el focus al primer control de ingreso de clientes
            'para mejorar la interface.
            Me.txt1er_Nom_titular.SetFocus
        End If
    End If
End Sub

Private Sub BotAyuda_Click()
    Dim nrocli_aux As String
    nrocli_aux = mFunBusqueda(1)    'todos los pasajeros
    nrocliente = Val(nrocli_aux)
    If Val(nrocliente) <> 0 Then
        ayuda = True
        If busco_clienteTF(nrocliente) Then
            desabilito_datos_cliente (False)
            cargo_datos_formulario
        End If
    End If
End Sub

Private Sub botCancelar_Click()
    subBorroPasajerosChekin
    If MsgBox("¿Desea salir del formulario de ingreso?", vbOKCancel + vbQuestion, "Confirmación de salir") = vbOK Then
        Unload Me
    End If
End Sub

Private Sub subBorroPasajerosChekin()
    'Si la habitación que estoy trabajando no tiene titular
    'es porque todavía no complete todo el proceso (no asigne titular)
    'por ese motivo tengo que borrar los pasajeros ingresados.
    
    Dim hab_aux As Long
    Dim consulta As String
    
    'es necesario confirmar si tiene titular ya que puedo estar trabajando con
    'un ingreso de pasajeros y en ese caso no tengo que borrar a nadie
    If tipo_accion_checkin = 1 Then 'checkin
        hab_aux = obtengo_nro_hab_formulario
        If tiene_titular(hab_aux) = False Then
            consulta = "DELETE FROM checkin WHERE nrohab = " & Str(hab_aux)
            bdHOTEL.Execute consulta
        End If
    End If
    
    'no es necesario confirmar si tiene titular ya que en un walkinL,
    'siempre será la primera vez que ingreso pasajreos para la habitación
    If tipo_accion_checkin = 2 Then 'walkinL
        hab_aux = Val(txtHabWalk.Text)
        consulta = "DELETE FROM checkin where nroreserva = " + Str(nro_reserva) & " and"
        consulta = consulta & " nrohab = " & Str(hab_aux)
        bdHOTEL.Execute consulta
    End If
End Sub

Private Sub botConfirmaCheckin_Click()
    Dim nroHabSeleccionada As Long
    'verifico si hay pasajeros seleccionados
    If dbgrid1.Rows > 1 Then
        If obtengo_nro_hab_formulario = 0 Then
            'debe de ingresar número de habitación
            mSubMensaje 4, 62
        Else
            If tipo_accion_checkin = 1 Then 'checkin
                Me.Hide
                'obtengo número de habitación desde lista.
                nroHabSeleccionada = CLng(Me.gHabitaciones.TextMatrix(Me.gHabitaciones.row, 0))
                'verifico si se asigno habitación
                If nroHabSeleccionada = 0 Then
                    'no se asigno habitación
                    frmTitularesHabitacion.propTipoAccionFormularioTitular = 2  'checkin sin habitación asignada
                    frmTitularesHabitacion.propHabCuenta = CLng(Me.txtHabCheck.Text)
                    frmTitularesHabitacion.propHabNroReserva = nro_reserva
                    frmTitularesHabitacion.propHabNroCorr = CLng(Me.gHabitaciones.TextMatrix(Me.gHabitaciones.row, 3))
                    
                Else
                    'si se asignó habitación
                    frmTitularesHabitacion.propTipoAccionFormularioTitular = 3  'checkin con habitación asignada
                    frmTitularesHabitacion.propHabCuenta = nroHabSeleccionada
                End If
                frmTitularesHabitacion.Show 1
            Else
                If tipo_accion_checkin = 2 Then 'walkin libre
                    Me.Hide
                    'obtengo número de habitación desde lista.
                    nroHabSeleccionada = CLng(Me.txtHabWalk.Text)
                    'si se asignó habitación
                    frmTitularesHabitacion.propTipoAccionFormularioTitular = 1  'walkin
                    frmTitularesHabitacion.propHabCuenta = nroHabSeleccionada
                    frmTitularesHabitacion.Show 1
                End If
            End If
        End If
    Else
        'debe ingresar al menos un pasajero a la habitación
        mSubMensaje 4, 63
    End If
End Sub

Private Function tiene_titular(hab As Long)
    tiene_titular = False
    If busco_habitaTF(hab) Then
        'si no tiene titular
        If tbHABITACIONES("titular_unica") = 0 And tbHABITACIONES("titular_aloja") = 0 Then
            tiene_titular = False
        Else
            tiene_titular = True
        End If
    End If
End Function

Private Sub botConfirmaIngreso_Click()
    Dim linea As String
    Dim nro_hab As Long
    Dim fdes_auxWO As Date
    Dim fhas_auxWO As Date
    If valido_datos Then
        If Not funClienteAlojado Then
            'aviso de confirmación de pasajero al hotel
            If mFunMensaje(4, 68) Then
                'obtengo número de habitación
                nro_hab = obtengo_nro_hab_formulario
                
                If tipo_accion_checkin = 3 Then 'walkinO
                    'Cuando ingreso un pasajero a una habitacion ocupada (walkinO)
                    'debo primero buscar el período de ocupación de esa habitacion (fdes y fhas)
                    'para poder asignarselo al nuevo pasajero
                    If busco_habita_checkin(nro_hab) Then
                        fdes_auxWO = tbCHECKIN("fcheckdes")
                        fhas_auxWO = tbCHECKIN("fcheckhas")
                    End If
                End If

                'actualizo datos cliente. Tengo que realizar este proceso antes de
                'crear un registro en el archivo checkin, para obtener el número de
                'cliente
                grabo_datos_clientes
                
                tbCHECKIN.AddNew
                'grabo registro en checkin
                tbCHECKIN("nrohab") = nro_hab
                tbCHECKIN("nroreserva") = nro_reserva
                tbCHECKIN("nrocorrcli") = nrocliente
                tbCHECKIN("horainghab") = Time
                tbCHECKIN("finghab") = m_FechaSis
                Select Case tipo_accion_checkin
                    Case 1  'checkin
                        tbCHECKIN("fcheckdes") = fechaingreso.Text
                        tbCHECKIN("fcheckhas") = fechaegreso.Text
                        
                    Case 2  'walkinL
                        tbCHECKIN("fcheckdes") = m_FechaSis
                        tbCHECKIN("fcheckhas") = fEgreso.Text
                        
                    Case 3  'walkinO
                        tbCHECKIN("fcheckdes") = fdes_auxWO
                        tbCHECKIN("fcheckhas") = fhas_auxWO
                        
                End Select
                tbCHECKIN.Update
                        
                cambiar_walkin (False)
                'habilito o no el boton de continuar
                boton_titular nro_hab
            End If
        End If
        ayuda = False
        'actualizo la ventana de pasajeros
        cargo_ventana_cli
        limpio_cliente
    End If
End Sub

Private Sub cargo_ventana_cli()
    Dim consulta
    Dim hab_aux As Long
    hab_aux = obtengo_nro_hab_formulario
    
    SQLpasajeros_habitacion hab_aux, Data2
    
    'propiedades de la grilla
    dbgrid1.FormatString = " |Nombres de los pasajeros que se hospedan en la habitación " & Str(hab_aux) & "                      |                    "
    dbgrid1.ColWidth(2) = 0             'oculto columna donde tengo el numero de pasajero
End Sub

Private Function funClienteAlojado() As Boolean
    Dim nrohab As Long
    'Si el cliente ya se ingresó en el Checkin que estoy trabajando,
    'permito modificar sus datos desde aquí. (checkin)
    
    funClienteAlojado = False
    nrohab = obtengo_nro_hab_formulario
    If busco_titular_checkinTF(nrohab, nrocliente) Then
        funClienteAlojado = True
        'permito modificar los datos de un cliente
        grabo_datos_clientes
        'aviso de que se modificaron los datos del cliente
        mSubMensaje 4, 69
        Exit Function
    End If
    
    'valido que el cliente no este alojado en el hotel
    tbCHECKIN.Index = "i_checkin_cli"
    tbCHECKIN.Seek ">=", nrocliente
    If Not tbCHECKIN.NoMatch Then
        If tbCHECKIN("nrocorrcli") = nrocliente Then
            funClienteAlojado = True
            'el cliente ya está alojado en el hotel
            mSubMensaje 4, 64
        End If
    End If
End Function

Private Sub botSelHabCheck_Click()
    If IsDate(fechaingreso.Text) And IsDate(fechaegreso.Text) Then
          tipo_accion_SeleccionHab = 2
          frmReservaSeleHab.Show 1
    End If
End Sub

Private Sub dbgrid1_DblClick()
    Dim cli As Long
    dbgrid1.col = 2
    cli = Val(dbgrid1.Text)
    If busco_clienteTF(cli) Then
        nrocliente = cli
        ayuda = True
        cargo_datos_formulario
    End If
End Sub

Private Sub dbgrid1_KeyPress(KeyAscii As Integer)
    'Permito la selección de pasajeros con la tecla enter
    If KeyAscii = vbKeyReturn Then
        'simulo que hize un click sobre la grilla
        dbgrid1_DblClick
    End If
End Sub

Private Sub cargo_ventana_habitaciones()
    'Cargo en la ventana de habitaciones, todas las habitaciones
    'que pertenecen a la reserva con la cual se está trabajando.
    'En la ventana se carga:
    '1) número de habitación
    '2) descripción del tipo
    '3) tipo de habitación
    '4) nrr. correlativo de la habitación en HAB_RESERVA
    
    tbHAB_RESERVAS.MoveFirst
    tbHAB_RESERVAS.Index = "ihab_reserva"
    tbHAB_RESERVAS.Seek ">=", nro_reserva, 1
    If Not tbHAB_RESERVAS.NoMatch Then  'si existe alguna habitacion
        Do While Not tbHAB_RESERVAS.EOF
            If tbHAB_RESERVAS("nroreserva") = nro_reserva Then
                'agrego linea a la grilla
                Me.gHabitaciones.AddItem ""
                'cargo nro. de habitación
                Me.gHabitaciones.TextMatrix(Me.gHabitaciones.Rows - 1, 0) = _
                    tbHAB_RESERVAS("nrohabitacion")
                    
                'cargo descripción del tipo de habitación
                Me.gHabitaciones.TextMatrix(Me.gHabitaciones.Rows - 1, 1) = _
                    CStr(tbHAB_RESERVAS("nrohabitacion")) & "  " & "Suite " & _
                    tbHAB_RESERVAS("nomtipohabitacion")
                    
                'cargo tipo de habitación
                Me.gHabitaciones.TextMatrix(Me.gHabitaciones.Rows - 1, 2) = _
                    tbHAB_RESERVAS("tipohabitacion")
                    
                'cargo correlativo de la habitación
                Me.gHabitaciones.TextMatrix(Me.gHabitaciones.Rows - 1, 3) = _
                    tbHAB_RESERVAS("nrocorr")
            Else
                Exit Do
            End If
            tbHAB_RESERVAS.MoveNext
        Loop
    End If
    'establesco propiedades de las columnas
    Me.gHabitaciones.ColWidth(0) = 0
    Me.gHabitaciones.ColWidth(1) = Me.gHabitaciones.Width
    Me.gHabitaciones.ColWidth(2) = 0
    Me.gHabitaciones.ColWidth(3) = 0
    'centro información hacia la izquierda
    marco_celdas_grilla Me.gHabitaciones, 1, 1, 0, Me.gHabitaciones.Rows - 1
    Me.gHabitaciones.CellAlignment = 1
    'por defecto selecciono la primer fila
    marco_celdas_grilla Me.gHabitaciones, 1, 1, 0, 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Descargo formulario de memoria
    Set frmCheck_in = Nothing
End Sub

Private Sub gHabitaciones_Click()
    Dim nro_hab As Long
    Dim x As Boolean
    
    cargo_ventana_cli
    limpio_cliente
    'si no tiene nro. habitación asignada habilito control hab.
    nro_hab = obtengo_nro_hab_ventanaHab
    If nro_hab <> 0 Then    'si tiene nro. habitacion oculto
        x = False
    Else                    'si no tiene nro. habitacion muestro
        x = True
    End If
    Label13.Visible = x
    txtHabCheck.Visible = x
    botSelHabCheck.Visible = x
End Sub

Private Sub boton_titular(hab As Long)
    'busco si la habitación ya tiene titular
    If tiene_titular(hab) = True Then
        'tiene titular, quiere decir que estoy haciendo un Walkin Ocupada
        'o un Checkin ya alojado.
        botConfirmaCheckin.Enabled = False
        Me.mnuFormularioContinuar.Enabled = False
    Else
        'la habitación no tiene titular
        If hab <> 0 Then
            'cuento cantidad de filas(pasajeros) de la grilla de pasajeros
            If dbgrid1.Rows >= 1 Then
                'existe por lo menos 1 pasajero
                botConfirmaCheckin.Enabled = True
                Me.mnuFormularioContinuar.Enabled = True
            Else
                'no se han seleccionado pasajeros
                botConfirmaCheckin.Enabled = False
                Me.mnuFormularioContinuar.Enabled = False
            End If
        Else
            botConfirmaCheckin.Enabled = False
            Me.mnuFormularioContinuar.Enabled = False
        End If
    End If
End Sub

Private Sub desabilito_habitacion(tipo As Byte)
    Dim color As String
    color = &H80000016      'gris
    If tipo = 0 Then 'checkin
        gHabitaciones.Enabled = False
        botSelHabCheck.Enabled = False
        botSeguir.Enabled = False
        Me.mnuFormularioconfirmar.Enabled = False
        
        gHabitaciones.BackColor = color
        botSelHabCheck.BackColor = color
        
    Else            'walkin
        fEgreso.Enabled = False
        cboTipo_habitacion.Enabled = False
        botHabitacionWalk.Enabled = False
        botSeguirWalk.Enabled = False
        Me.mnuFormularioconfirmar.Enabled = False
        
        fEgreso.BackColor = color
        cboTipo_habitacion.BackColor = color
        botHabitacionWalk.BackColor = color
    End If
End Sub

Private Function obtengo_nro_hab_ventanaHab()
    obtengo_nro_hab_ventanaHab = Me.gHabitaciones.TextMatrix(Me.gHabitaciones.row, 0)
End Function

Private Function obtengo_nro_hab_formulario()
    Dim hab As Long
    If tipo_accion_checkin = 1 Then 'checkin
        hab = CLng(obtengo_nro_hab_ventanaHab)
        If hab = 0 Then
            hab = Val(txtHabCheck.Text)
        End If
    Else
        hab = Val(txtHabWalk.Text)
    End If
    obtengo_nro_hab_formulario = hab
End Function


'********************************************
'* Procedimientos exclusivos de Walkin      *
'********************************************

Private Sub botHabitacionWalk_Click()
    If cboTipo_habitacion.ListIndex <> -1 Then
        If tipo_accion_checkin = 2 Then 'walkinL
            If funValidoFechaEgreso Then
                  tipo_accion_SeleccionHab = 3
                  frmReservaSeleHab.Show 1
            End If
        Else
            'walkinO
            tipo_accion_SeleccionHab = 4
            frmReservaSeleHab.Show 1
        End If
    End If
End Sub

Private Function funValidoFechaEgreso()
    'Valido que la fecha de egreso sea la correcta
    funValidoFechaEgreso = True
    If IsDate(fEgreso.Text) Then
        If fEgreso.Value <= m_FechaSis Then
            'la fecha de egreso no puede ser menor o igual a la fecha de hoy
            mSubMensaje 3, 65
            funValidoFechaEgreso = False
            fEgreso.SetFocus
        End If
    Else
        funValidoFechaEgreso = False
    End If
End Function

Private Sub botSeguirWalk_Click()
    If valido_ingreso_walk Then
        If tipo_accion_checkin = 2 Then 'walkinL
            
        Else                            'walkinO
            cargo_ventana_cli
        End If
        
        desabilito_datos_cliente False
        desabilito_habitacion 1
        limpio_cliente
        'le doy el focus al primer control de ingreso de clientes
        'para mejorar la interface.
        Me.txt1er_Nom_titular.SetFocus
    End If
End Sub

Private Function valido_ingreso_walk()
    valido_ingreso_walk = True
    If tipo_accion_checkin = 2 Then     'walkinL
        If IsDate(fEgreso.Text) Then
            If fEgreso.Value <= m_FechaSis Then
                valido_ingreso_walk = False
                'la fecha de egreso no pude ser menor o igual a la de hoy
                mSubMensaje 4, 65
                fEgreso.SetFocus
                Exit Function
            End If
        Else
            valido_ingreso_walk = False
            'formato de fecha de egreso incorrecto
            mSubMensaje 3, 1
            fEgreso.SetFocus
            Exit Function
        End If
    End If
    If Val(txtHabWalk.Text) = 0 Then
        'debe de seleccionar habitación
        mSubMensaje 4, 62
        botSeguirWalk.SetFocus
        valido_ingreso_walk = False
    End If
End Function

Private Sub cambiar_walkin(x As Boolean)
    txtHabWalk.Enabled = x
    fEgreso.Enabled = x
    cboTipo_habitacion.Enabled = x
End Sub

Private Sub fEgreso_Change()
    cboTipo_habitacion.ListIndex = -1
    txtHabWalk.Text = ""
End Sub

'********************************************
'* Procedimientos para manejo de clientes   *
'********************************************

Private Sub desabilito_datos_cliente(x As Boolean)
    Dim color As String
    Dim xInvertida As Boolean
    
    txt1er_Nom_titular.Locked = x
    txt2do_Nom_titular.Locked = x
    txt1er_Ape_titular.Locked = x
    txt2do_Ape_titular.Locked = x
    txtDirecion.Locked = x
    txtLocalidad.Locked = x
    cboPais.Locked = x
    txtCodPostal.Locked = x
    txtTele.Locked = x
    faxcheckin.Locked = x
    cboSexo.Locked = x
    cboNacionalidad.Locked = x
    cboEstadoCivil.Locked = x
    txtDocu.Locked = x
    cboTipoDocu.Locked = x
    txtRuc.Locked = x
    txtObservaciones.Locked = x
    txtEmail.Locked = x
    
    If x Then
        color = mSisColor_18ControlesNoHabilitados
        xInvertida = False
    Else
        xInvertida = True
        color = &H80000005  'blanco
    End If
    
    fechanac.Enabled = xInvertida
    botConfirmaIngreso.Enabled = xInvertida
    botayuda.Enabled = xInvertida
    'opción del menu
    Me.mnuBuscar.Enabled = xInvertida
    Me.mnuFormularioConfirmarPasajero.Enabled = xInvertida
          
    txt1er_Nom_titular.BackColor = color
    txt2do_Nom_titular.BackColor = color
    txt1er_Ape_titular.BackColor = color
    txt2do_Ape_titular.BackColor = color
    txtDirecion.BackColor = color
    txtLocalidad.BackColor = color
    cboPais.BackColor = color
    txtCodPostal.BackColor = color
    txtTele.BackColor = color
    faxcheckin.BackColor = color
    cboSexo.BackColor = color
    cboNacionalidad.BackColor = color
    cboEstadoCivil.BackColor = color
    txtDocu.BackColor = color
    cboTipoDocu.BackColor = color
    txtRuc.BackColor = color
    txtObservaciones.BackColor = color
    txtEmail.BackColor = color
    fechanac.BackColor = color
    
    txt1er_Nom_titular.TabStop = xInvertida
    txt2do_Nom_titular.TabStop = xInvertida
    txt1er_Ape_titular.TabStop = xInvertida
    txt2do_Ape_titular.TabStop = xInvertida
    txtDirecion.TabStop = xInvertida
    txtLocalidad.TabStop = xInvertida
    cboPais.TabStop = xInvertida
    txtCodPostal.TabStop = xInvertida
    txtTele.TabStop = xInvertida
    faxcheckin.TabStop = xInvertida
    cboSexo.TabStop = xInvertida
    cboNacionalidad.TabStop = xInvertida
    cboEstadoCivil.TabStop = xInvertida
    txtDocu.TabStop = xInvertida
    cboTipoDocu.TabStop = xInvertida
    txtRuc.TabStop = xInvertida
    txtObservaciones.TabStop = xInvertida
    txtEmail.TabStop = xInvertida
    fechanac.TabStop = xInvertida
End Sub

Private Function valido_datos()
    Dim fecha_aux As Date

    valido_datos = True
    If txt1er_Ape_titular.Text = "" Then
        'debe ingresar al menos 1er apellido para continuar
        mSubMensaje 4, 66
        txt1er_Ape_titular.SetFocus
        valido_datos = False
        Exit Function
    End If
    
    If cboSexo.ListIndex = -1 Then
        'debe ingresar sexo para continuar
        mSubMensaje 4, 67
        cboSexo.SetFocus
        valido_datos = False
        Exit Function
    End If
    
    'verifico si se ingresó fecha
    If Trim(fechanac.Text) <> Empty Then
        'verifico que el formato sea el correcto.
        If IsDate(fechanac.Text) Then
            'verifico que este dentro del rango permitido
            fecha_aux = fechanac.Text
            If fecha_aux > m_FechaSis Then
                'la fecha de nacimiento no puede ser mayor a la fecha de hoy
                mSubMensaje 3, 5
                fechanac.SetFocus
                valido_datos = False
                Exit Function
            End If
        Else
            'el formato de la fecha de nacimiento no es válido
            mSubMensaje 3, 1
            fechanac.SetFocus
            valido_datos = False
            Exit Function
        End If
    End If
End Function

Private Sub gHabitaciones_SelChange()
    'Cuando cambio las filas con las teclas de selección,
    'asumo que stoy haciendo click con el mouse, sobre la nueva fila seleccionada.
    gHabitaciones_Click
End Sub

Private Sub txt1er_ape_titular_LostFocus()
    Dim aux As String
    aux = StrConv(txt1er_Ape_titular, 1)
    txt1er_Ape_titular = aux
    'asistencia a usuario
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub txt1er_Nom_titular_LostFocus()
    Dim aux As String
    aux = StrConv(txt1er_Nom_titular, 2)
    aux = StrConv(aux, 3)
    txt1er_Nom_titular = aux
    'asistencia a usuario
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub txt2do_ape_titular_LostFocus()
    Dim aux As String
    aux = StrConv(txt2do_Ape_titular, 1)
    txt2do_Ape_titular = aux
    'asistencia a usuario
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub txt2do_Nom_titular_LostFocus()
    Dim aux As String
    aux = StrConv(txt2do_Nom_titular, 2)
    aux = StrConv(aux, 3)
    txt2do_Nom_titular = aux
    'asistencia a usuario
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub limpio_cliente()
    Me.txt1er_Ape_titular.Text = ""
    Me.txt1er_Nom_titular.Text = ""
    Me.txt2do_Ape_titular.Text = ""
    Me.txt2do_Nom_titular.Text = ""
    txtDirecion.Text = ""
    txtLocalidad.Text = ""
    txtCodPostal.Text = ""
    txtTele.Text = ""
    faxcheckin.Text = ""
    If cboPais.ListCount > 0 Then cboPais.ListIndex = 0
    If cboSexo.ListCount > 0 Then cboSexo.ListIndex = 0
    If cboNacionalidad.ListCount > 0 Then cboNacionalidad.ListIndex = 0
    If cboEstadoCivil.ListCount > 0 Then cboEstadoCivil.ListIndex = 0
    txtDocu.Text = ""
    txtRuc.Text = ""
    txtObservaciones.Text = ""
    txtEmail.Text = ""
    fechanac.Text = ""
    nrocliente = 0
End Sub

Private Sub cargo_datos_formulario()
    posiciono_combo frmCheck_in.cboPais, tbCLIENTES("pais_reside_titular")
    posiciono_combo frmCheck_in.cboSexo, tbCLIENTES("sexo_titular")
    posiciono_combo frmCheck_in.cboNacionalidad, tbCLIENTES("nacionalidad_titular")
    posiciono_combo frmCheck_in.cboEstadoCivil, tbCLIENTES("estado_civil_titular")
    posiciono_combo frmCheck_in.cboTipoDocu, tbCLIENTES("tipoDocu_titular")
    
    txt1er_Nom_titular = tbCLIENTES("primer_nom_titular")
    txt2do_Nom_titular = tbCLIENTES("segundo_nom_titular")
    txt1er_Ape_titular = tbCLIENTES("primer_ape_titular")
    txt2do_Ape_titular = tbCLIENTES("segundo_ape_titular")
    txtDirecion = tbCLIENTES("direccion_titular")
    txtLocalidad = tbCLIENTES("localidad_titular")
    txtCodPostal = tbCLIENTES("cod_postal_titular")
    txtTele = tbCLIENTES("tel_titular")
    faxcheckin = tbCLIENTES("fax_titular")
    txtDocu = tbCLIENTES("documento_titular")
    txtObservaciones = tbCLIENTES("observaciones_titular")
    txtRuc = tbCLIENTES("ruc_titular")
    txtEmail = tbCLIENTES("email_titular")
    If Not IsNull(tbCLIENTES("fecha_nac_titular")) Then fechanac.Text = tbCLIENTES("fecha_nac_titular")
End Sub

Private Sub grabo_datos_clientes()
    Dim nom_aux As String
    If ayuda = False Then
        'no se utilizó la ayuda para ingresar los datos del cliente,
        'y no se seleccionó un cliente de la grilla
        'es decir se esta ingresando un nuevo cliente
        tbCLIENTES.AddNew
        nrocliente = funObtengoProxCliente
        tbCLIENTES("nrocorr") = nrocliente
    Else
        tbCLIENTES.Edit
    End If
    
    tbCLIENTES("primer_nom_titular") = txt1er_Nom_titular.Text
    tbCLIENTES("segundo_nom_titular") = txt2do_Nom_titular.Text
    tbCLIENTES("primer_ape_titular") = txt1er_Ape_titular.Text
    tbCLIENTES("segundo_ape_titular") = txt2do_Ape_titular.Text
    tbCLIENTES("direccion_titular") = txtDirecion.Text
    tbCLIENTES("localidad_titular") = txtLocalidad.Text
    
    tbCLIENTES("pais_reside_titular") = cboPais.ItemData(cboPais.ListIndex)
    tbCLIENTES("sexo_titular") = cboSexo.ItemData(cboSexo.ListIndex)
    tbCLIENTES("nacionalidad_titular") = cboNacionalidad.ItemData(cboNacionalidad.ListIndex)
    tbCLIENTES("estado_civil_titular") = cboEstadoCivil.ItemData(cboEstadoCivil.ListIndex)
    tbCLIENTES("tipoDocu_titular") = cboTipoDocu.ItemData(cboTipoDocu.ListIndex)
                    
    tbCLIENTES("cod_postal_titular") = txtCodPostal.Text
    tbCLIENTES("tel_titular") = txtTele.Text
    tbCLIENTES("fax_titular") = faxcheckin.Text
    tbCLIENTES("documento_titular") = txtDocu.Text
    tbCLIENTES("ruc_titular") = txtRuc.Text
    tbCLIENTES("observaciones_titular") = txtObservaciones.Text
    nom_aux = txt1er_Nom_titular.Text & " " & txt2do_Nom_titular.Text & " " & txt1er_Ape_titular.Text & " " & txt2do_Ape_titular.Text
    tbCLIENTES("nombre_completo_titular") = nom_aux
    tbCLIENTES("email_titular") = txtEmail.Text
    If IsDate(fechanac.Text) Then tbCLIENTES("fecha_nac_titular") = fechanac.Text
    tbCLIENTES.Update
End Sub

Private Function funObtengoProxCliente() As Long
    'Devuelve el valor del campo tbPARAMETROS("nrocliente"), el cual determina
    'el número de cliente a asignar al pasajero que se ingresa desde CHECKIN.
    'Calcula el próximo número de cliente.
    '--------------------------------------------------------------------------------
    'Parámetros.
    '   Salida: Valor del campo tbPARAMETROS("nrocliente").
    '           Si este valor es 0, quiere decir que todavía no existe ningún cliente
    '           en la base de datos por lo que se devuelve 1.
    '--------------------------------------------------------------------------------
    Dim nrocli As Long
    nrocli = tbPARAMETROS("nrocliente")
    If nrocli = 0 Then
        'no existen clientes
        nrocli = 1
    Else
        'existen clientes
        nrocli = tbPARAMETROS("nrocliente")
    End If
    funObtengoProxCliente = nrocli
    
    'Calculo el próximo número de cliente
    tbPARAMETROS.Edit
    tbPARAMETROS("nrocliente") = nrocli + 1
    tbPARAMETROS.Update
End Function

Private Sub configuracion_apariencia()
    'Determina la apariencia del los elemento configurables del formulario
    Me.gHabitaciones.ForeColor = mSisColor_10CheckinSeleccionHab
End Sub

Private Sub mnuFormularioContinuar_Click()
    'Equivale a presionar la tecla F12 o el boton de continuar
    If Me.botConfirmaCheckin.Enabled = True Then
        botConfirmaCheckin_Click
    End If
End Sub

Private Sub mnuFormularioSalir_Click()
    'Equivale a presionar la tecla Esc o el boton de salir
    botCancelar_Click
End Sub

Private Sub mnuBuscarClientes_Click()
    'Equivale a presionar el boton de ayuda o la tecla F1
    BotAyuda_Click
End Sub

Private Sub mnuFormularioConfirmar_Click()
    'Equivale a presionar el boton de seleccionar habitación del tabs visible (checkin o walkin)
    If Me.SSTab1.Tab = 0 Then      'chekin
        botSeguir_Click
    Else
        If Me.SSTab1.Tab = 1 Then   'walkin
            botSeguirWalk_Click
        End If
    End If
End Sub

Private Sub mnuFormularioConfirmarPasajero_Click()
    'Equivale a presionar el boton de confirmar pasajero
    botConfirmaIngreso_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'Como hay dos opciones del menú con la misma tecla de función (F9) debo de
    'implementar esta tecla manualmente.
    If KeyAscii = vbKeyF9 Then
        If Me.mnuFormularioConfirmarPasajero.Enabled = True Then
            mnuFormularioConfirmarPasajero_Click
        Else
            If Me.mnuFormularioconfirmar.Enabled = True Then
                mnuFormularioConfirmar_Click
            End If
        End If
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'Intercepto la tecla F9
    If KeyCode = vbKeyF9 Then
        Form_KeyPress (KeyCode)
    End If
End Sub

'****************************************************
'*
'*    Asistencia a usuarios.
'*
'****************************************************

Private Sub gHabitaciones_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 87
End Sub

Private Sub botSelHabCheck_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 86
End Sub

Private Sub botSeguir_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 88
End Sub

Private Sub fEgreso_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 89
End Sub

Private Sub cboTipo_habitacion_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 90
End Sub

Private Sub botHabitacionWalk_GotFocus()
    'Este boton realiza dos tareas diferentes, dependiendo del uso del formulario.
    If tipo_accion_checkin = 2 Then 'walkinL
        'Muestro asistencia para cunado estoy realizando un walkinL
        mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 113
    End If
    If tipo_accion_checkin = 3 Then 'walkinO
        'Muestro asistencia para cunado estoy realizando un walkinO
        mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 85
    End If
End Sub

Private Sub botSeguirWalk_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 88
End Sub

Private Sub txt1er_Nom_titular_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 91
End Sub

Private Sub txt2do_Nom_titular_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 92
End Sub

Private Sub txt1er_Ape_titular_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 93
End Sub

Private Sub txt2do_Ape_titular_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 94
End Sub

Private Sub txtCodPostal_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 103
End Sub

Private Sub txtDirecion_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 95
End Sub

Private Sub txtDocu_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 102
End Sub

Private Sub txtEmail_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 104
End Sub

Private Sub txtLocalidad_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 96
End Sub

Private Sub txtObservaciones_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 108
End Sub

Private Sub txtRuc_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 106
End Sub

Private Sub txtTele_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 97
End Sub

Private Sub cboSexo_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 99
End Sub

Private Sub cboNacionalidad_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 100
End Sub

Private Sub faxcheckin_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 98
End Sub

Private Sub fechanac_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 101
End Sub

Private Sub cboPais_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 107
End Sub

Private Sub botConfirmaIngreso_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 111
End Sub

Private Sub botConfirmaCheckin_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 112
End Sub

Private Sub botCancelar_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 1, 1
End Sub

Private Sub cboEstadoCivil_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 105
End Sub

Private Sub dbgrid1_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 109
End Sub

Private Sub cboTipoDocu_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 212
End Sub

Private Sub dbgrid1_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub botCancelar_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub botConfirmaIngreso_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub gHabitaciones_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub botConfirmaCheckin_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub cboEstadoCivil_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub faxcheckin_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub fechanac_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub cboPais_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub txtTele_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub cboSexo_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub cboNacionalidad_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub txtDirecion_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub txtDocu_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub txtEmail_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub txtLocalidad_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub txtCodPostal_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub txtRuc_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub botSeguirWalk_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub cboTipo_habitacion_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub botHabitacionWalk_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub botSeguir_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub botSelHabCheck_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub fEgreso_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub txtObservaciones_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub cboTipoDocu_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

