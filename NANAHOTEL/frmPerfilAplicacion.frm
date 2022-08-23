VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmPerfilAplicacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Definición del perfil de la aplicación"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11340
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   11340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton botAceptar 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   9960
      TabIndex        =   32
      Top             =   6840
      Width           =   1215
   End
   Begin TabDlg.SSTab ssTab1 
      Height          =   6615
      Left            =   120
      TabIndex        =   33
      Top             =   120
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   11668
      _Version        =   327680
      Style           =   1
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Aviso"
      TabPicture(0)   =   "frmPerfilAplicacion.frx":0000
      Tab(0).ControlCount=   1
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      TabCaption(1)   =   "Tipo de habitaciones"
      TabPicture(1)   =   "frmPerfilAplicacion.frx":001C
      Tab(1).ControlCount=   1
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame7"
      Tab(1).Control(0).Enabled=   0   'False
      TabCaption(2)   =   "Habitaciones"
      TabPicture(2)   =   "frmPerfilAplicacion.frx":0038
      Tab(2).ControlCount=   1
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame3"
      Tab(2).Control(0).Enabled=   0   'False
      TabCaption(3)   =   "Habitaciones General"
      TabPicture(3)   =   "frmPerfilAplicacion.frx":0054
      Tab(3).ControlCount=   1
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame2"
      Tab(3).Control(0).Enabled=   0   'False
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   5895
         Left            =   -74880
         TabIndex        =   55
         Top             =   480
         Width           =   10695
         Begin VB.Frame Frame11 
            Caption         =   "Agregar"
            Height          =   1335
            Left            =   5160
            TabIndex        =   61
            Top             =   1080
            Width           =   5535
            Begin VB.TextBox txtDescReg 
               Height          =   375
               Left            =   240
               MaxLength       =   50
               TabIndex        =   29
               Top             =   600
               Width           =   3855
            End
            Begin VB.CommandButton botAgregarReg 
               Caption         =   "&Agregar"
               Height          =   375
               Left            =   4200
               TabIndex        =   30
               Top             =   600
               Width           =   1215
            End
            Begin VB.Label lblDescReg 
               AutoSize        =   -1  'True
               Caption         =   "lblDescReg"
               Height          =   240
               Left            =   240
               TabIndex        =   28
               Top             =   360
               Width           =   1080
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Eliminar"
            Height          =   1935
            Left            =   5160
            TabIndex        =   56
            Top             =   2520
            Width           =   5535
            Begin VB.TextBox txtCodRegEli 
               Height          =   375
               Left            =   240
               MaxLength       =   10
               TabIndex        =   58
               Top             =   600
               Width           =   1815
            End
            Begin VB.TextBox txtDescRegEli 
               Height          =   375
               Left            =   240
               MaxLength       =   50
               TabIndex        =   57
               Top             =   1320
               Width           =   3855
            End
            Begin VB.CommandButton botEliminarReg 
               Caption         =   "&Eliminar"
               Height          =   375
               Left            =   4200
               TabIndex        =   31
               Top             =   1320
               Width           =   1215
            End
            Begin VB.Label lblCodRegEli 
               AutoSize        =   -1  'True
               Caption         =   "lblCodRegEli"
               Height          =   240
               Left            =   240
               TabIndex        =   60
               Top             =   360
               Width           =   1200
            End
            Begin VB.Label lblDescRegEli 
               AutoSize        =   -1  'True
               Caption         =   "lblDescRegEli"
               Height          =   240
               Left            =   240
               TabIndex        =   59
               Top             =   1080
               Width           =   1305
            End
         End
         Begin VB.ComboBox cboTipoTrabajo 
            Height          =   360
            ItemData        =   "frmPerfilAplicacion.frx":0070
            Left            =   1800
            List            =   "frmPerfilAplicacion.frx":007A
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   120
            Width           =   3615
         End
         Begin MSFlexGridLib.MSFlexGrid gRegistros 
            Height          =   2895
            Left            =   240
            TabIndex        =   27
            Top             =   1200
            Width           =   4695
            _ExtentX        =   8281
            _ExtentY        =   5106
            _Version        =   393216
            Rows            =   1
            FixedCols       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin VB.Label lblRegistrosDefinidos 
            AutoSize        =   -1  'True
            Caption         =   "lblRegistrosDefinidos"
            Height          =   240
            Left            =   240
            TabIndex        =   26
            Top             =   960
            Width           =   1935
         End
         Begin VB.Label lblRegistrosNoDefinidos 
            AutoSize        =   -1  'True
            Caption         =   "lblRegistrosNoDefinidos"
            Height          =   240
            Left            =   240
            TabIndex        =   62
            Top             =   4200
            Visible         =   0   'False
            Width           =   2205
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "&Trabajar con:"
            Height          =   240
            Left            =   120
            TabIndex        =   24
            Top             =   180
            Width           =   1200
         End
         Begin VB.Image Image2 
            Height          =   105
            Left            =   0
            Picture         =   "frmPerfilAplicacion.frx":009F
            Top             =   600
            Width           =   10650
         End
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   6015
         Left            =   240
         TabIndex        =   44
         Top             =   480
         Width           =   10575
         Begin VB.Frame Frame10 
            Caption         =   "Agregar habitación"
            Height          =   1935
            Left            =   3840
            TabIndex        =   53
            Top             =   0
            Width           =   6615
            Begin VB.CommandButton botAgregarHab 
               Caption         =   "&Agregar"
               Height          =   375
               Left            =   5160
               TabIndex        =   15
               Top             =   1320
               Width           =   1215
            End
            Begin VB.TextBox txtNumeroHab 
               Height          =   375
               Left            =   360
               MaxLength       =   10
               TabIndex        =   14
               Top             =   1320
               Width           =   1575
            End
            Begin VB.ComboBox cboTipoHabitacion 
               Height          =   360
               Left            =   360
               Style           =   2  'Dropdown List
               TabIndex        =   12
               Top             =   600
               Width           =   3135
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "&Número"
               Height          =   240
               Left            =   360
               TabIndex        =   13
               Top             =   1080
               Width           =   720
            End
            Begin VB.Label Label7 
               Caption         =   "&Tipo de habitación"
               Height          =   255
               Left            =   360
               TabIndex        =   11
               Top             =   360
               Width           =   2535
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Agregar grupo de habitaciones "
            Height          =   1935
            Left            =   3840
            TabIndex        =   52
            Top             =   2040
            Width           =   6615
            Begin VB.CommandButton botAgregarGrupo 
               Caption         =   "&Crear grupo"
               Height          =   375
               Left            =   5160
               TabIndex        =   22
               Top             =   1320
               Width           =   1215
            End
            Begin VB.TextBox txtCantHab 
               Height          =   375
               Left            =   3840
               MaxLength       =   2
               TabIndex        =   21
               Top             =   600
               Width           =   1575
            End
            Begin VB.TextBox txtNumeroHabInicial 
               Height          =   375
               Left            =   240
               MaxLength       =   10
               TabIndex        =   19
               Top             =   1320
               Width           =   1575
            End
            Begin VB.ComboBox cboTipoHabitacionGrupo 
               Height          =   360
               Left            =   240
               Style           =   2  'Dropdown List
               TabIndex        =   17
               Top             =   600
               Width           =   3135
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "Canti&dad "
               Height          =   240
               Left            =   3840
               TabIndex        =   20
               Top             =   360
               Width           =   855
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "Nú&mero inicial"
               Height          =   240
               Left            =   240
               TabIndex        =   18
               Top             =   1080
               Width           =   1275
            End
            Begin VB.Label Label9 
               Caption         =   "T&ipo de habitaciónes"
               Height          =   255
               Left            =   240
               TabIndex        =   16
               Top             =   360
               Width           =   2535
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Eliminar habitación"
            Height          =   1815
            Left            =   3840
            TabIndex        =   45
            Top             =   4080
            Width           =   6615
            Begin VB.CommandButton botEliminarHabitacion 
               Caption         =   "Elimina&r"
               Height          =   375
               Left            =   5160
               TabIndex        =   23
               Top             =   1200
               Width           =   1215
            End
            Begin VB.ComboBox cboTipoHabEliminar 
               Height          =   360
               Left            =   240
               Style           =   2  'Dropdown List
               TabIndex        =   48
               Top             =   1200
               Width           =   3135
            End
            Begin VB.TextBox txtNumeroHabEliminar 
               Height          =   375
               Left            =   240
               MaxLength       =   10
               TabIndex        =   47
               Top             =   480
               Width           =   1575
            End
            Begin ComctlLib.ProgressBar ProgressBar1 
               Height          =   255
               Left            =   2040
               TabIndex        =   46
               Top             =   540
               Visible         =   0   'False
               Width           =   4335
               _ExtentX        =   7646
               _ExtentY        =   450
               _Version        =   327682
               Appearance      =   1
               Max             =   10
            End
            Begin VB.Label Label13 
               Caption         =   "Tipo de habitación"
               Height          =   255
               Left            =   240
               TabIndex        =   51
               Top             =   960
               Width           =   2535
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               Caption         =   "Número"
               Height          =   240
               Left            =   240
               TabIndex        =   50
               Top             =   240
               Width           =   720
            End
            Begin VB.Label lblVerificacion 
               Caption         =   "lblVerificacion"
               Height          =   240
               Left            =   2040
               TabIndex        =   49
               Top             =   240
               Visible         =   0   'False
               Width           =   4245
            End
         End
         Begin MSFlexGridLib.MSFlexGrid gHabitaciones 
            Height          =   5415
            Left            =   0
            TabIndex        =   10
            Top             =   240
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   9551
            _Version        =   393216
            Rows            =   1
            FixedCols       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "&Habitaciones definidas"
            Height          =   240
            Left            =   0
            TabIndex        =   9
            Top             =   0
            Width           =   2070
         End
         Begin VB.Label lblAvisoDeNoHabitaciones 
            AutoSize        =   -1  'True
            Caption         =   "No existen habitaciones definidas"
            Height          =   240
            Left            =   0
            TabIndex        =   54
            Top             =   5760
            Width           =   3015
         End
      End
      Begin VB.Frame Frame7 
         BorderStyle     =   0  'None
         Height          =   5895
         Left            =   -74880
         TabIndex        =   36
         Top             =   480
         Width           =   10815
         Begin VB.Frame Frame8 
            Caption         =   "Agregar tipo "
            Height          =   2415
            Left            =   5040
            TabIndex        =   37
            Top             =   240
            Width           =   5655
            Begin VB.CommandButton botAgregarTipo 
               Caption         =   "A&gregar"
               Height          =   375
               Left            =   4320
               TabIndex        =   7
               Top             =   1920
               Width           =   1215
            End
            Begin VB.TextBox txtTarifaInicialTipo 
               Height          =   375
               Left            =   240
               MaxLength       =   10
               TabIndex        =   6
               Top             =   1320
               Width           =   1695
            End
            Begin VB.TextBox txtDescTipo 
               Height          =   375
               Left            =   240
               MaxLength       =   50
               TabIndex        =   4
               Top             =   600
               Width           =   3855
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "T&arifa inicial"
               Height          =   240
               Left            =   240
               TabIndex        =   5
               Top             =   1080
               Width           =   1080
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "&Descripción del tipo"
               Height          =   240
               Left            =   240
               TabIndex        =   3
               Top             =   360
               Width           =   1785
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "Eliminar tipo "
            Height          =   2415
            Left            =   5040
            TabIndex        =   38
            Top             =   2760
            Width           =   5655
            Begin VB.CommandButton botEliminartipo 
               Caption         =   "&Eliminar"
               Height          =   375
               Left            =   4320
               TabIndex        =   8
               Top             =   1920
               Width           =   1215
            End
            Begin VB.TextBox txtDescTipoAEliminar 
               Height          =   375
               Left            =   240
               MaxLength       =   50
               TabIndex        =   40
               Top             =   1320
               Width           =   3855
            End
            Begin VB.TextBox txtCodigoTipo 
               Height          =   375
               Left            =   240
               MaxLength       =   10
               TabIndex        =   39
               Top             =   600
               Width           =   1815
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Descripción del tipo"
               Height          =   240
               Left            =   240
               TabIndex        =   42
               Top             =   1080
               Width           =   1785
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Código"
               Height          =   240
               Left            =   240
               TabIndex        =   41
               Top             =   360
               Width           =   660
            End
         End
         Begin MSFlexGridLib.MSFlexGrid gTiposHab 
            Height          =   4455
            Left            =   120
            TabIndex        =   2
            Top             =   360
            Width           =   4695
            _ExtentX        =   8281
            _ExtentY        =   7858
            _Version        =   393216
            Rows            =   1
            Cols            =   3
            FixedCols       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin VB.Label lblAvisoDeExistenciaDeTipos 
            AutoSize        =   -1  'True
            Caption         =   "No existen tipos definidos de habitaciones"
            Height          =   240
            Left            =   120
            TabIndex        =   43
            Top             =   4920
            Visible         =   0   'False
            Width           =   4620
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "&Tipos definidos"
            Height          =   240
            Left            =   120
            TabIndex        =   1
            Top             =   120
            Width           =   1395
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Advertencia "
         Height          =   5895
         Left            =   -74760
         TabIndex        =   34
         Top             =   480
         Width           =   10575
         Begin VB.CommandButton botConfirmar 
            Caption         =   "&Continuar"
            Height          =   375
            Left            =   9240
            TabIndex        =   0
            Top             =   5400
            Width           =   1215
         End
         Begin VB.Image Image1 
            Height          =   2535
            Left            =   240
            Picture         =   "frmPerfilAplicacion.frx":0432
            Stretch         =   -1  'True
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label lblAviso 
            Caption         =   "lblAviso"
            Height          =   2535
            Left            =   2160
            TabIndex        =   35
            Top             =   600
            Width           =   8055
         End
      End
   End
End
Attribute VB_Name = "frmPerfilAplicacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub botAceptar_Click()
    Unload Me
End Sub

Private Sub botConfirmar_Click()
    'autorizo tabs de configuración
    Me.ssTab1.TabEnabled(1) = True
    Me.ssTab1.TabEnabled(2) = True
    Me.ssTab1.TabEnabled(3) = True
    'me posiciono sobre el tabs de tipos de habitaciones
    Me.ssTab1.Tab = 1
End Sub

Private Sub Form_Activate()
    'por defecto muestro el tabs 0 (aviso al usuario)
    Me.ssTab1.Tab = 0
End Sub

Private Sub Form_Load()
    'Cargo etiqueta de aviso
    Me.lblAviso.Caption = "Este formulario contiene información que " & _
                        "determina el funcionamiento de la aplicación. " & Chr(10) & Chr(10) & _
                        "Cambiar dicha información en forma indevida, " & _
                        "puede originar que la alicación no funcione como ested desea. " & Chr(10) & Chr(10) & _
                        "Si no esta seguro de como funciona este formulario, o como puede " & _
                        "afectar a su aplicación los posibles cambios, solicite ayuda técnica antes " & _
                        "de continuar."
                        
    'inicialmente no permito trabajr con tabs de configuracion
    Me.ssTab1.TabEnabled(1) = False
    Me.ssTab1.TabEnabled(2) = False
    Me.ssTab1.TabEnabled(3) = False
    'configuro cotroles
    subInicializoControles
    'habilito controles del tabs
    subHabilitoTab 0
End Sub

Private Sub subInicializoControles()
    'Inicializo las propiedades de algunos controles del formulario
    
    'bloqueo controles en tab de tipo de habitación
    mSubBloqueoControlFormulario Me.txtCodigoTipo, True
    mSubBloqueoControlFormulario Me.txtDescTipoAEliminar, True
    'bloqueo controles en tabs de habitaciones
    mSubBloqueoControlFormulario Me.cboTipoHabEliminar, True
    mSubBloqueoControlFormulario Me.txtNumeroHabEliminar, True
    'bloqueo controles en tabs de habitaciones general
    mSubBloqueoControlFormulario Me.txtDescRegEli, True
    mSubBloqueoControlFormulario Me.txtCodRegEli, True
    
    'propiedades de grilla de tipos de habitaciones
    Me.gTiposHab.BackColorSel = mSisColor_15FilaSeleccionada
    Me.gTiposHab.ForeColorSel = mSisColor_19FilaSeleccionadaTexto
    'propiedades de grilla de habitaciones
    Me.gHabitaciones.BackColorSel = mSisColor_15FilaSeleccionada
    Me.gHabitaciones.ForeColorSel = mSisColor_19FilaSeleccionadaTexto
    'propiedades de grilla de regitros (motivos de bloqueo y situaciones)
    Me.gRegistros.BackColorSel = mSisColor_15FilaSeleccionada
    Me.gRegistros.ForeColorSel = mSisColor_19FilaSeleccionadaTexto
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmPerfilAplicacion = Nothing
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    'Determino que tabs esta visible
    Select Case ssTab1.Tab
        Case 0  'aviso
            'habilito controles del tabs
            subHabilitoTab 0
        Case 1  'tipo de habitaciones
            'habilito controles del tabs
            subHabilitoTab 1
            subInicializoIngresoTipos
            subMuestroTipoHabitaciones
            
        Case 2  'habitaciones
            'habilito controles del tabs
            subHabilitoTab 2
            subInicializoIngresoHab
            subCargoCombosTipoHab
            subMuestroHabitaciones
        Case 3  'habitaciones motivos de bloqueo y situaciones
            'habilito controles del tabs
            subHabilitoTab 3
            'esto desencadena el evento click del combo
            Me.cboTipoTrabajo.ListIndex = 0
    End Select
    Me.Refresh
End Sub

Private Sub subHabilitoTab(nroTab As Byte)
    'Para evitar que se le pueda dar el focus por medio del teclado a
    'controles de tabs que no estan activos, bloqueo todos los controles de los
    'tabs menos del que esta activo actualmente.
    Frame1.Enabled = False  'tabs de aviso
    Frame7.Enabled = False  'tabs de tipos de habitaciones
    Frame3.Enabled = False  'tabs de habitaciones
    Frame2.Enabled = False  'tabs de registros
    
    Select Case nroTab
        Case 0  'tabs de aviso
            Frame1.Enabled = True
        Case 1  'tabs de tipos de habitaciones
            Frame7.Enabled = True
        Case 2  'tabs de habitaciones
            Frame3.Enabled = True
        Case 3  'tabs de registros
            Frame2.Enabled = True
    End Select
End Sub

Private Sub subInicializoGrilla(grilla As MSFlexGrid)
    'Inicializo grilla
    grilla.Clear
    grilla.Rows = 2
    grilla.FixedRows = 1
End Sub

'***********************************************************************
'*
'*  Tipos de habitaciones
'*
'***********************************************************************

Private Sub subInicializoIngresoTipos()
    'Limpia los controles de ingreso de tipos de habitaciones
    Me.txtCodigoTipo.Text = Empty
    Me.txtDescTipo.Text = Empty
    Me.txtDescTipoAEliminar.Text = Empty
    Me.txtTarifaInicialTipo.Text = Empty
End Sub

Private Sub subMuestroTipoHabitaciones()
    'Muestra los tipos ya definidos
    Dim consulta As String
    Dim qdfTipoHab As QueryDef
    Dim rstTipoHab As Recordset
    
    consulta = "Select tipohab,descripcion,tarifa " & _
                "from TIPO_HABITACIONES " & _
                "Order by descripcion"
                
    'ejecuto consuta
    Set qdfTipoHab = bdHOTEL.CreateQueryDef("")
    qdfTipoHab.SQL = consulta
    Set rstTipoHab = qdfTipoHab.OpenRecordset(dbOpenSnapshot)
    'cargo recordset a grilla
    subCargoGrillaTipoHab rstTipoHab
    
    Set qdfTipoHab = Nothing
    Set rstTipoHab = Nothing
End Sub

Private Sub subCargoGrillaTipoHab(tiposHab As Recordset)
    'Recorre el recordset con la información de los diferentes tipos
    'de habitaciones del hotel y lo muestra en la grilla de tipos.
    
    'inicializo grilla
    subInicializoGrilla Me.gTiposHab
    'verifico si existen registros
    If tiposHab.RecordCount > 0 Then
        'recorro recordset
        tiposHab.MoveFirst
        Do While Not tiposHab.EOF
            'agrego línea a grilla de tipos
            gTiposHab.AddItem _
                tiposHab("tipohab") & Chr(9) & _
                tiposHab("descripcion") & Chr(9) & _
                tiposHab("tarifa")
                tiposHab.MoveNext
        Loop
        'si existen registros, no muestro etiqueta de aviso
        Me.lblAvisoDeExistenciaDeTipos.Visible = False
        'permito trabajar con los tabs de habitaciones y habitaciones general
        Me.ssTab1.TabEnabled(2) = True 'habitaciones
        Me.ssTab1.TabEnabled(3) = True 'habitaciones general
    Else
        'si no existen registros, muestro etiqueta de aviso
        Me.lblAvisoDeExistenciaDeTipos.Visible = True
        'no permito trabajar con los tabs de habitaciones y habitaciones general
        Me.ssTab1.TabEnabled(2) = False 'habitaciones
        Me.ssTab1.TabEnabled(3) = False 'habitaciones general
    End If
    'establesco cabezal de la grilla
    gTiposHab.FormatString = " Código | Descripción     | Tarifa"
    'ajusto tamaño columna descripción
    gTiposHab.ColWidth(1) = gTiposHab.Width - _
                        gTiposHab.ColWidth(0) - _
                        gTiposHab.ColWidth(2) - 500
                        '500 es aprox. el ancho de la scrollBars
End Sub

Private Sub botAgregarTipo_Click()
    'Agrego un nuevo tipo de habitación
    Dim nroCorrTipoHab As Integer
    'valido que se ingrese descripción del tipo de habitación
    If Trim(Me.txtDescTipo.Text) <> Empty Then
        'obtengo próximo número correlativo libre
        nroCorrTipoHab = mFunObtengoNroCorrTipoHabitaciones
        'aviso de confirmación de operación
        If mFunMensaje(4, 90) Then
            'creo un nuevo registro
            tbTIPO_HABITACIONES.AddNew
                tbTIPO_HABITACIONES("tipoHab") = nroCorrTipoHab
                tbTIPO_HABITACIONES("descripcion") = Me.txtDescTipo.Text
                tbTIPO_HABITACIONES("tarifa") = Val(Me.txtTarifaInicialTipo.Text)
            tbTIPO_HABITACIONES.Update
            'inicializo controles de ingreso y actualizo grilla
            subInicializoIngresoTipos
            subMuestroTipoHabitaciones
            'grabo bitacora
            GraboBitacora "Nuevo tipo habitación: " & nroCorrTipoHab
        End If
    Else
        'aviso de ingreso de tipo
        mSubMensaje 4, 106
    End If
End Sub

Private Sub txtCodigoTipo_KeyPress(KeyAscii As Integer)
    'Solo permito el ingreso de números
    ValidoNum KeyAscii, False, False
End Sub

Private Sub txtTarifaInicialTipo_KeyPress(KeyAscii As Integer)
    'Solo permito el ingreso de números
    ValidoNum KeyAscii, True, True
End Sub

Private Sub gTiposHab_DblClick()
    Dim tipoSel As Long
    tipoSel = Val(Me.gTiposHab.TextMatrix(Me.gTiposHab.Row, 0))
    'Busco tipo seleccionado
    If busco_tipo_habTF(tipoSel) Then
        'si existe muestro código
        Me.txtCodigoTipo.Text = tbTIPO_HABITACIONES("tipohab")
        'muestro descripción
        Me.txtDescTipoAEliminar.Text = tbTIPO_HABITACIONES("descripcion")
        'ilumino fila en grilla
        marco_celdas_grilla Me.gTiposHab, 0, 2, Me.gTiposHab.Row, Me.gTiposHab.Row
    End If
End Sub

Private Sub botEliminartipo_Click()
    'Elimino un tipo de habitación
    If funValidoEliminacionTipoHab Then
        'aviso de confirmación de operación
        If mFunMensaje(4, 91) Then
            'busco tipo seleccionado
            If busco_tipo_habTF(CLng(Me.txtCodigoTipo.Text)) Then
                'elimino tipo de habitación
                tbTIPO_HABITACIONES.Delete
                'grabo bitácora
                GraboBitacora "Eliminación tipo habitación " & Me.txtCodigoTipo.Text

                'inicializo controles de ingreso y actualizo grilla
                subInicializoIngresoTipos
                subMuestroTipoHabitaciones
            End If
        End If
    End If
End Sub

Private Function funValidoEliminacionTipoHab() As Boolean
    'Determina si un tipo de habitación se puede eliminar.
    'Para que un tipo de habitación se pueda borrar, no debe de tener asociado
    'ninguna habitación.
    '---------------------------------------------------------------------------
    'Parámetros.
    '   Salida:     True, si el tipo de habitación se puede borrar
    '               False, si el tipo de habitación tiene habitaciones asignadas
    '               False, si no se seleccionó un tipo de habitación
    '----------------------------------------------------------------------------
    'por defecto asumo que puedo eliminar el tipo
    funValidoEliminacionTipoHab = True
    
    'verifico si se seleccionó tipo
    If Me.txtCodigoTipo.Text = Empty Then
        'debe de seleccionar código de tipo de habitación
        mSubMensaje 4, 92
        funValidoEliminacionTipoHab = False
    Else
        'verifico que el tipo no este referenciado en la tabla habitaciones
        tbHABITACIONES.Index = "i_tipohab"
        tbHABITACIONES.Seek ">=", CLng(Me.txtCodigoTipo.Text)
        If Not tbHABITACIONES.NoMatch Then
            If tbHABITACIONES("tipohab") = CLng(Me.txtCodigoTipo.Text) Then
                'existe 1 o más habitaciones con el tipo seleccionado
                mSubMensaje 4, 93
                funValidoEliminacionTipoHab = False
            End If
        End If
    End If
End Function

'***********************************************************************
'*
'*  Definición Habitaciones
'*
'***********************************************************************

Private Sub subInicializoIngresoHab()
    'Inicializa los controles que permiten el ingreso de habitaciones
    Me.txtCantHab.Text = Empty
    Me.txtNumeroHab.Text = Empty
    Me.txtNumeroHabEliminar.Text = Empty
    Me.txtNumeroHabInicial.Text = Empty
    
End Sub

Private Sub subCargoCombosTipoHab()
    'Carga los combos de tipos de habitaciones
    
    'limpio combos
    Me.cboTipoHabEliminar.Clear
    Me.cboTipoHabitacion.Clear
    Me.cboTipoHabitacionGrupo.Clear
                                                'Utilizados para:
    carga_tipo_hab Me.cboTipoHabEliminar        'habitación a eliminar
    carga_tipo_hab Me.cboTipoHabitacionGrupo    'grupo de habitaciones
    carga_tipo_hab Me.cboTipoHabitacion         'nueva habitación
    
    'por defecto, muestro primer elemento de los combos
    Me.cboTipoHabEliminar.ListIndex = 0
    Me.cboTipoHabitacion.ListIndex = 0
    Me.cboTipoHabitacionGrupo.ListIndex = 0
End Sub

Private Sub txtNumeroHab_KeyPress(KeyAscii As Integer)
    'valido el ingreso de números
    ValidoNum KeyAscii, False, False
End Sub

Private Sub txtCantHab_KeyPress(KeyAscii As Integer)
    'valido el ingreso de números
    ValidoNum KeyAscii, False, False
End Sub

Private Sub txtNumeroHabInicial_KeyPress(KeyAscii As Integer)
    'valido el ingreso de números
    ValidoNum KeyAscii, False, False
End Sub

Private Sub subMuestroHabitaciones()
    'Muestra las habitaciones ya definidos
    Dim consulta As String
    Dim qdfHab As QueryDef
    Dim rstHab As Recordset
    
    consulta = "Select HABITACIONES.nrohab,TIPO_HABITACIONES.descripcion " & _
                "From HABITACIONES,TIPO_HABITACIONES " & _
                "Where HABITACIONES.tipohab = TIPO_HABITACIONES.tipohab " & _
                "Order by HABITACIONES.nrohab"
                
    'ejecuto consuta
    Set qdfHab = bdHOTEL.CreateQueryDef("")
    qdfHab.SQL = consulta
    Set rstHab = qdfHab.OpenRecordset(dbOpenSnapshot)
    'cargo recordset a grilla
    subCargoGrillaHabitaciones rstHab
    
    Set qdfHab = Nothing
    Set rstHab = Nothing
End Sub

Private Sub subCargoGrillaHabitaciones(rstHab As Recordset)
    'Recorre el recordset con la información de las diferentes
    'habitaciones del hotel y lo muestra en la grilla de habitaciones
    
    'inicializo grilla
    subInicializoGrilla Me.gHabitaciones
    'verifico si existen registros
    If rstHab.RecordCount > 0 Then
        'recorro recordset
        rstHab.MoveFirst
        Do While Not rstHab.EOF
            'agrego línea a grilla de tipos
            gHabitaciones.AddItem _
                rstHab("nroHab") & Chr(9) & _
                rstHab("descripcion")
                rstHab.MoveNext
        Loop
        'si existen registros, no muestro etiqueta de aviso
        Me.lblAvisoDeNoHabitaciones.Visible = False
    Else
        'si no existen registros, muestro etiqueta de aviso
        Me.lblAvisoDeNoHabitaciones.Visible = True
    End If
    'establesco cabezal de la grilla
    gHabitaciones.FormatString = " Numero | Tipo       "
    'ajusto tamaño columna descripción
    gHabitaciones.ColWidth(1) = gHabitaciones.Width - _
                        gHabitaciones.ColWidth(0) - 500
                        '500 es aprox. el ancho de la scrollBars

End Sub

Private Sub botAgregarHab_Click()
    'Creo una nueva habitación
    
    'valido que la habitación sea correcta
    If funValidoNuevaHabitacion Then
        'aviso de confirmación de la operación
        If mFunMensaje(4, 96) Then
            'creo nueva habitación
            tbHABITACIONES.AddNew
                tbHABITACIONES("nrohab") = CLng(Me.txtNumeroHab.Text)
                tbHABITACIONES("tipohab") = Me.cboTipoHabitacion.ItemData(Me.cboTipoHabitacion.ListIndex)
                tbHABITACIONES("ubicacionhab") = ""
                tbHABITACIONES("tipocuenta_unica") = 0
                tbHABITACIONES("tipocuenta_aloja") = 0
                tbHABITACIONES("tipocuenta_extra") = 0
                tbHABITACIONES("titular_unica") = 0
                tbHABITACIONES("titular_aloja") = 0
                tbHABITACIONES("titular_extra") = 0
                tbHABITACIONES("tarifa") = 0
                tbHABITACIONES("situacionhab") = 1      'establesco la situación de la habitación a limpia por defecto.
                tbHABITACIONES("fechasituacionhab") = Date
            tbHABITACIONES.Update
            'calculo total de tipos de habitaciones
            subCalculoTotalHabitacionesPorTipo Me.cboTipoHabitacion.ItemData(Me.cboTipoHabitacion.ListIndex)
            'grabo bitácora
            GraboBitacora "Nueva habitación " & Me.txtNumeroHab.Text

            'inicializo controles para nuevo ingreso
            subInicializoIngresoHab
            subMuestroHabitaciones
        End If
    End If
End Sub

Private Sub subCalculoTotalHabitacionesPorTipo(tipoHab As Byte)
    'En el archivo de tipos de habitaciones existe un campo que almacena la cantidad de
    'habitaciones de cada tipo, este valor se modifica en este procedimiento.
    'Este campo es utilizado en el formulario frmVerDisponibilidad, con el fin de inicializar
    'las celdas de la grilla con el total de habitaciones por tipo.
    '------------------------------------------------------------------------------------------
    'Parámetros.
    '   Entrada [tipoHab] tipo de habitación a la que se debe acceder
    '
    '   NOTA: se espera mejorar el rendimiento de la consulta de disponibilidad al tener este valor
    '           ya calculado.
    '---------------------------------------------------------------------------------------------
    Dim consulta As String
    Dim qdfTotHab As QueryDef
    Dim rstTotHab As Recordset
    Dim totHab As Byte
    
    consulta = "Select nrohab " & _
                "From HABITACIONES " & _
                "Where HABITACIONES.tipohab = " & tipoHab
                
                
    'ejecuto consuta
    Set qdfTotHab = bdHOTEL.CreateQueryDef("")
    qdfTotHab.SQL = consulta
    Set rstTotHab = qdfTotHab.OpenRecordset(dbOpenSnapshot)
    'cuento el total de habitacines del tipo determinado
    If rstTotHab.RecordCount > 0 Then
        rstTotHab.MoveLast
    End If
    totHab = rstTotHab.RecordCount
    
    If busco_tipo_habTF(CLng(tipoHab)) Then
        'modifico el valor del campo correspondiente
        tbTIPO_HABITACIONES.Edit
            tbTIPO_HABITACIONES("total_por_tipo") = totHab
        tbTIPO_HABITACIONES.Update
    End If
    Set qdfTotHab = Nothing
    Set rstTotHab = Nothing
End Sub

Private Function funValidoNuevaHabitacion() As Boolean
    'Determino si la habitación a ingresas es correcta
    '-------------------------------------------------------------------------
    'Parámetros.
    '   Salida: True, si se ingresó un valor y si la habitación no existe
    '           False, la habitación ya existe
    '           False,  no se ingresó número de habitación
    '-------------------------------------------------------------------------
    'por defecto asumo que puedo ingresar la habitación
    funValidoNuevaHabitacion = True
    
    'valido que se ingrese un valor
    If Me.txtNumeroHab.Text = Empty Then
        'aviso de datos incompletos
        mSubMensaje 4, 94
        funValidoNuevaHabitacion = False
    Else
        'valido que la habitación no exista
        If busco_habitaTF(CLng(Me.txtNumeroHab.Text)) Then
            'aviso de habitación existente
            mSubMensaje 4, 95
            funValidoNuevaHabitacion = False
        End If
    End If
End Function

Private Sub gHabitaciones_Click()
    'Muestro datos de la habitación a borrar
    Dim nrohab As Long
    
    'obtengo habitación de la fila seleccionada
    nrohab = Val(Me.gHabitaciones.TextMatrix(Me.gHabitaciones.Row, 0))
    'busco habitación
    If busco_habitaTF(nrohab) Then
        'muestro número de habitación
        Me.txtNumeroHabEliminar.Text = tbHABITACIONES("nroHab")
        'muestro tipo de habitación
        posiciono_combo Me.cboTipoHabEliminar, tbHABITACIONES("tipohab")
        'ilumino fila en grilla
        marco_celdas_grilla Me.gHabitaciones, 0, 1, Me.gHabitaciones.Row, Me.gHabitaciones.Row
    End If
End Sub

Private Sub botEliminarHabitacion_Click()
    'Elimino habitación
    Dim habEliminar As Long
    'aviso de inicio de proceso de validación
    'ya que el mismo puede demorar
    If mFunMensaje(4, 97) Then
        'valido que se ingrese habitación
        If Me.txtNumeroHabEliminar.Text <> Empty Then
            habEliminar = Val(Me.txtNumeroHabEliminar.Text)
            'valido que la habitación se pueda eliminar
            If funValidoEliminacionHabitacion(habEliminar) Then
                'me posiciono en el registro de la habitación a borrar
                If busco_habitaTF(habEliminar) Then
                    'elimino habitación
                    tbHABITACIONES.Delete
                    'calculo total de tipos de habitaciones
                    subCalculoTotalHabitacionesPorTipo _
                    Me.cboTipoHabEliminar.ItemData(Me.cboTipoHabEliminar.ListIndex)
                    
                    'aviso de eliminación ok
                    mSubMensaje 4, 100
                    'grabo bitácora
                    GraboBitacora "Eliminación habitación : " & habEliminar
                End If
            End If
        Else
            'aviso de ingreso de habitación
            mSubMensaje 4, 99
        End If
        'oculto barra de progreso
        Me.ProgressBar1.Visible = False
        Me.lblVerificacion.Visible = False
        'inicializo controles para nuevo ingreso
        subInicializoIngresoHab
        subMuestroHabitaciones
    End If
End Sub

Private Function funValidoEliminacionHabitacion(habEli As Long) As Boolean
    'Determina si la habitación se puede eliminar
    '--------------------------------------------------------------------------
    'Parámetros.
    '   Entrada:
    '           [habEli]    habitación que deseo eliminar
    '   Salida:
    '           True, si se cumplen todas las condiciones siguientes:
    '           no existe habitación en BLOQUEO_HAB
    '           no existe habitación en CHECKIN
    '           no existe habitación en CHECKOUT
    '           no existe habitación en CIERREDIARIO
    '           no existe habitación en CUENTAS_ALOJA
    '           no existe habitación en CUENTAS_EXTRAS
    '           no existe habitación en HAB_ANULADAS
    '           no existe habitación en HAB_RESERVADAS
    '           no existe habitación en HAB_RESERVA_AUX
    '           no existe habitación en SITUACION_HIS
    '
    '           False, si no se cumple 1 o más de las condiciones anteriores
    '----------------------------------------------------------------------------
    
    'por efecto asumo que no se puede eliminar la habitación
    funValidoEliminacionHabitacion = False
    
    'muestro barra de progreso
    Me.ProgressBar1.Value = 0
    Me.ProgressBar1.Visible = True
    Me.lblVerificacion.Visible = True
    Me.lblVerificacion = "Verificando: " & "archivo de bloqueos."
    'no permito trabajar con ningún control
    Me.Enabled = False
    If funVerificoExistenciaArchivo(tbBLOQUEO_HAB, tbBLOQUEO_HAB("hab_bloq"), habEli) Then
        'paso 1 de 10 completo
        Me.ProgressBar1.Value = 1
        Me.lblVerificacion = "Verificando: " & "archivo de ocupación."
        If funVerificoExistenciaArchivo(tbCHECKIN, tbCHECKIN("nroHab"), habEli) Then
            'paso 2 de 10 completo
            Me.ProgressBar1.Value = 2
            Me.lblVerificacion = "Verificando: " & "archivo de ocupaciones pasadas."
            If funVerificoExistenciaArchivo(tbCHECKOUT, tbCHECKOUT("nroHab"), habEli) Then
                'paso 3 de 10 completo
                Me.ProgressBar1.Value = 3
                Me.lblVerificacion = "Verificando: " & "archivo de cierres diarios."
                If funVerificoExistenciaArchivo(tbCIERRE_DIARIO, tbCIERRE_DIARIO("hab_cierre"), habEli) Then
                    'paso 4 de 10 completo
                    Me.ProgressBar1.Value = 4
                    Me.lblVerificacion = "Verificando: " & "archivo de gastos alojamiento."
                    If funVerificoExistenciaArchivo(tbCUENTAS_ALOJA, tbCUENTAS_ALOJA("habitacion_cuenta_aloja"), habEli) Then
                        'paso 5 de 10 completo
                        Me.ProgressBar1.Value = 5
                        Me.lblVerificacion = "Verificando: " & "archivo de gastos extras."
                        If funVerificoExistenciaArchivo(tbCUENTAS, tbCUENTAS("habitacion_cuenta"), habEli) Then
                            'paso 6 de 10 completo
                            Me.ProgressBar1.Value = 6
                            Me.lblVerificacion = "Verificando: " & "archivo de reservas anuladas."
                            If funVerificoExistenciaArchivo(tbHAB_ANULADAS, tbHAB_ANULADAS("nroHabitacion"), habEli) Then
                                'paso 7 de 10 completo
                                Me.ProgressBar1.Value = 7
                                Me.lblVerificacion = "Verificando: " & "archivo de reservas."
                                If funVerificoExistenciaArchivo(tbHAB_RESERVAS, tbHAB_RESERVAS("nroHabitacion"), habEli) Then
                                    'paso 8 de 10 completo
                                    Me.ProgressBar1.Value = 8
                                    Me.lblVerificacion = "Verificando: " & "archivo de reservas auxiliar."
                                    If funVerificoExistenciaArchivo(tbHAB_RESERVAS_AUX, tbHAB_RESERVAS_AUX("nroHabitacion"), habEli) Then
                                        'paso 9 de 10 completo
                                        Me.ProgressBar1.Value = 9
                                        Me.lblVerificacion = "Verificando: " & "archivo histórico de situaciones."
                                        If funVerificoExistenciaArchivo(tbSITUACION_HIS, tbSITUACION_HIS("nroHab_situ"), habEli) Then
                                            'paso 10 de 10 completo
                                            Me.ProgressBar1.Value = 10
                                            'se puede eliminar la habitación
                                            funValidoEliminacionHabitacion = True
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    If funValidoEliminacionHabitacion = False Then
        'aviso de no eliminación
        mSubMensaje 4, 98
    End If
    'al terminar el proceso permito trabajar con el formulario
    Me.Enabled = True
End Function

Private Function funVerificoExistenciaArchivo(tabla As Recordset, _
                                                campo As Field, _
                                                valor As Long) As Boolean

    'Determino si existe un valor en un archivo.
    '-----------------------------------------------------------------------
    'Parámetros.
    '   Entreda:
    '               [tabla] tabla en la que tengo que verificar
    '               [campo] campo de la tabla que me interesa verificar
    '               [valor] valor que me estoy buscando
    '
    '   Salida:     True, si encuentro valor
    '               False, si no encuentro valor
    '               False, si se produce un error de ejecución
    '-----------------------------------------------------------------------
    On Error GoTo error
    'por defecto asumo que se puede eliminar
    funVerificoExistenciaArchivo = True
    
    'posiciono en el principio de la tabla
    If Not tabla.BOF Then
        tabla.MoveFirst
    End If
    
    'recorro tabla
    Do While Not tabla.EOF
        'comparo valor con un campo del registro actual
        If campo.Value = valor Then
            'existe el valor
            funVerificoExistenciaArchivo = False
            Exit Function
        End If
        tabla.MoveNext
    Loop
    'NOTA: esta línea se puede borrar
    mSubEspera 1
Exit Function
error:
    'no permito eliminar habitación
    funVerificoExistenciaArchivo = False
End Function

Private Sub botAgregarGrupo_Click()
    'Creo un grupo de habitaciones partiendo de una habitación inicial.
    'Crear grupos de habitaciones facilita la tarea de inicialización de hoteles,
    'con un número importante de habitaciones.
    
    'verifico si es posible crear el grupo
    Dim habInicial As Long
    Dim cantHab As Byte
    
    habInicial = Val(Me.txtNumeroHabInicial.Text)
    cantHab = Val(Me.txtCantHab.Text)
    
    'valido que el grupo se pueda crear
    If funValidoNuevasHabitaciones(habInicial, cantHab) Then
        'aviso de confirmación de nuevo grupo
        If mFunMensaje(4, 104) Then
            'creo nuevas habitaciones
            subCreoNuevasHabitaciones habInicial, cantHab
            
            'calculo total de tipos de habitaciones
            subCalculoTotalHabitacionesPorTipo Me.cboTipoHabitacionGrupo.ItemData(Me.cboTipoHabitacionGrupo.ListIndex)

            'inicializo controles para nuevo ingreso
            subInicializoIngresoHab
            subMuestroHabitaciones
            'aviso de creación de habitaciones ok
            mSubMensaje 4, 105
            'grabo bitácora
            GraboBitacora "Nuevo grupo habitaciones: " & "Ini. " & habInicial & " Tot. " & cantHab
        End If
    End If
End Sub

Private Sub subCreoNuevasHabitaciones(habInicial As Long, cantHab As Byte)
    'Crea nuevas habitaciones del hotel.
    '--------------------------------------------------------------------------------
    'Parámetros.
    '   Entrada:
    '       [habInicial] el número de habitación inicial, a partir del cual se crea
    '                    la primera habitación del grupo de habitaciones.
    '
    '       [cantHab]   cantidad de habitaciones que se crean a partir del número
    '                   inicial.
    '--------------------------------------------------------------------------------
    Dim contHab As Byte
    Dim nuevaHab As Long
    
    contHab = 0
    nuevaHab = habInicial
    Do While contHab < cantHab
        nuevaHab = habInicial + contHab
        'busco nueva habitación creada
        If Not busco_habitaTF(nuevaHab) Then
            'si no existe la habitación la creo
            tbHABITACIONES.AddNew
                tbHABITACIONES("nrohab") = nuevaHab
                tbHABITACIONES("tipohab") = _
                    Me.cboTipoHabitacionGrupo.ItemData(Me.cboTipoHabitacionGrupo.ListIndex)
                tbHABITACIONES("ubicacionhab") = ""
                tbHABITACIONES("tipocuenta_unica") = 0
                tbHABITACIONES("tipocuenta_aloja") = 0
                tbHABITACIONES("tipocuenta_extra") = 0
                tbHABITACIONES("titular_unica") = 0
                tbHABITACIONES("titular_aloja") = 0
                tbHABITACIONES("titular_extra") = 0
                tbHABITACIONES("tarifa") = 0
                tbHABITACIONES("situacionhab") = 1      'establesco la situación de la habitación a limpia por defecto.
                tbHABITACIONES("fechasituacionhab") = Date
            tbHABITACIONES.Update
        End If
        contHab = contHab + 1
        
    Loop
End Sub

Private Function funValidoNuevasHabitaciones(habInicial As Long, cantHab As Byte) As Boolean
    'Determina si es posible crear el grupo de nuevas habitaciónes
    '----------------------------------------------------------------------------------------
    'Parámetros.
    '   Entrada:
    '               [habInicial] el número de habitación inicial, a partir del cual se crea
    '                            la primera habitación del grupo de habitaciones
    '
    '               [cantHab]   cantidad de habitaciones que se crean a partir del número
    '                           inicial
    '   Salida:
    '               True, no existe ninguna de las habitaciones a crear
    '               False, existe al menos 1 de las habitaciones a crear
    '               False, si no se ingresó cantidad de habitaciones
    '               False, si no se ingresó habitación inicial.
    '----------------------------------------------------------------------------------------
    Dim contHab As Byte
    Dim nuevaHab As Long
    
    'por defecto asumo que se puede crear el grupo
    funValidoNuevasHabitaciones = True
    
    'valido que se haya ingresado cantidad de habitaciones
    If cantHab = 0 Then
        'aviso de ingreso de cantidad de habitaciones
        mSubMensaje 4, 102
        funValidoNuevasHabitaciones = False
    Else
        'valido que se haya ingresado habitación inicial
        If habInicial = 0 Then
            'aviso de ingreso de habitación inicial
            mSubMensaje 4, 103
            funValidoNuevasHabitaciones = False
        Else
            'valido que las habitaciones del grupo no existan
            contHab = 0
            Do While contHab < cantHab
                'creo nueva habitación
                nuevaHab = habInicial + contHab
                'busco nueva habitación creada
                If busco_habitaTF(nuevaHab) Then
                    'existe la habitación
                    funValidoNuevasHabitaciones = False
                    Exit Do
                End If
                contHab = contHab + 1
            Loop
                
            'verifico resultado
            If funValidoNuevasHabitaciones = False Then
                'muestro mensaje de grupo incorrecto.
                mSubMensaje 4, 101, CStr(nuevaHab)
            End If
        End If
    End If
End Function

'***********************************************************************
'*
'*  Definición de motivos de bloqueo y situaciones
'*
'***********************************************************************

Private Sub subInicializoRegistros()
    'Inicializo los controles del tabs de registros
    Me.txtDescReg.Text = Empty
    Me.txtDescRegEli.Text = Empty
    Me.txtCodRegEli.Text = Empty
End Sub

Private Sub subMuestroRegistros(tipoReg As Byte)
    'Cargo la grilla de registros con los registros del archvio
    'tbTIPO_ESTADO_HAB pertenecientes a un tipo determinado.
    '---------------------------------------------------------------------------------
    'Parámetros.
    '   Entrada:    [tipoReg]   tipo de registro a mostrar
    '                           1= motivos de bloqueo
    '                           2= situaciones
    '---------------------------------------------------------------------------------
            
    Dim consulta As String
    Dim qdfReg As QueryDef
    Dim rstReg As Recordset
    
    consulta = "Select tipo_cod,cod,descri " & _
                "From TIPO_ESTADO_HAB " & _
                "Where tipo_cod = " & tipoReg & _
                " Order by descri"
                
    'ejecuto consuta
    Set qdfReg = bdHOTEL.CreateQueryDef("")
    qdfReg.SQL = consulta
    Set rstReg = qdfReg.OpenRecordset(dbOpenSnapshot)
    'cargo recordset a grilla
    subCargoGrillaRegistros rstReg
    
    Set qdfReg = Nothing
    Set rstReg = Nothing
End Sub

Private Sub subCargoGrillaRegistros(rstReg As Recordset)
    'Recorre el recordset con la información de los diferentes
    'registros y cargo enla grilla de registros
    
    'inicializo grilla
    subInicializoGrilla Me.gRegistros
    'verifico si existen registros
    If rstReg.RecordCount > 0 Then
        'recorro recordset
        rstReg.MoveFirst
        Do While Not rstReg.EOF
            'agrego línea a grilla de tipos
            gRegistros.AddItem _
                rstReg("cod") & Chr(9) & _
                rstReg("descri")
                rstReg.MoveNext
        Loop
        'si existen registros, no muestro etiqueta de aviso
        Me.lblRegistrosNoDefinidos.Visible = False
    Else
        'si no existen registros, muestro etiqueta de aviso
        Me.lblRegistrosNoDefinidos.Visible = True
    End If
    'establesco cabezal de la grilla
    gRegistros.FormatString = " Código | Descripción       "
    'ajusto tamaño columna descripción
    gRegistros.ColWidth(1) = gRegistros.Width - _
                        gRegistros.ColWidth(0) - 500
                        '500 es aprox. el ancho de la scrollBars
End Sub

Private Sub cboTipoTrabajo_Click()
    'Cada vez que se cambie el valor del combo, tengo que:
    '   mostrar los nuevos tipos de registros
    '   inicializar los controles del tabs
    '   inicializar etiquetas
    
    'muestro tipos de registros
    subMuestroRegistros Me.cboTipoTrabajo.ItemData(Me.cboTipoTrabajo.ListIndex)
    'inicializo los controles del tabs
    subInicializoRegistros
    'inicializo etiquetas
    subInicializoEtiquetas Me.cboTipoTrabajo.ItemData(Me.cboTipoTrabajo.ListIndex)
End Sub

Private Sub subInicializoEtiquetas(tipoTrabajo As Byte)
    'Dependiendo del valor del combo de tipo de trabajo,
    'cambio el título de las etiquetas de los controles.
    '-----------------------------------------------------------------
    'Parámetros.
    '   Entrada:    [tipoTrabajo] 1= etiquetas de motivos de bloqueo
    '                             2= etiquetas de situaciones
    '-----------------------------------------------------------------
    Select Case tipoTrabajo
        Case 1
            Me.lblDescReg.Caption = "&Descripción del motivo de bloqueo"
            Me.lblCodRegEli.Caption = "Código del motivo de bloqueo"
            Me.lblDescRegEli.Caption = "Descripción del motivo de bloqueo"
            Me.lblRegistrosNoDefinidos.Caption = "No existen motivos de bloqueo definidos."
            Me.lblRegistrosDefinidos.Caption = "&Motivos de bloqueos definidos"
        Case 2
            Me.lblDescReg.Caption = "&Descripción de la situación"
            Me.lblCodRegEli.Caption = "Código de la situación"
            Me.lblDescRegEli.Caption = "Descripción de la situación"
            Me.lblRegistrosNoDefinidos.Caption = "No existen situaciones definidas."
            Me.lblRegistrosDefinidos.Caption = "S&ituaciones definidas"
    End Select
End Sub

Private Sub botAgregarReg_Click()
    'Agrego un nuevo registro del tipo determinado por el combo
    Dim nroCorrReg As Integer
    Dim tipoAviso As Integer   '108= confirma ingreso de motivo de bloqueo
                            '109= confirma ingreso de situación
    
    'determino tipo de aviso
    If Me.cboTipoTrabajo.ItemData(Me.cboTipoTrabajo.ListIndex) = 1 Then
        tipoAviso = 108
    Else
        tipoAviso = 109
    End If
    
    'valido que se ingrese descripción
    If Trim(Me.txtDescReg.Text) <> Empty Then
        'aviso de confirmación de la operación
        If mFunMensaje(4, tipoAviso) Then
            'obtengo próximo número libre
            nroCorrReg = mFunObtengoProxTipoEstado(Me.cboTipoTrabajo.ItemData(Me.cboTipoTrabajo.ListIndex))
            'creo nuevo registro
            tbTIPO_ESTADO_HAB.AddNew
                tbTIPO_ESTADO_HAB("tipo_cod") = Me.cboTipoTrabajo.ItemData(Me.cboTipoTrabajo.ListIndex)
                tbTIPO_ESTADO_HAB("cod") = nroCorrReg
                tbTIPO_ESTADO_HAB("descri") = Me.txtDescReg.Text
            tbTIPO_ESTADO_HAB.Update
            'grabo bitacora
            GraboBitacora "Nuevo " & Me.cboTipoTrabajo.Text & " " & nroCorrReg
            'esto desencadena el evento click del combo
            cboTipoTrabajo_Click
        End If
    Else
        mSubMensaje 4, 107
    End If
End Sub

Private Sub botEliminarReg_Click()
    'Elimino registro
    Dim tipoAviso As Integer   '114= confirma eliminación de motivo de bloqueo
                               '115= confirma eliminación de situación
                               
                               '116= eliminación motivo bloqueo correcta
                               '117= eliminación situación correcta
    Dim tipoReg As Byte
    Dim codReg As Long
    
    'obtengo tipo de registro
    tipoReg = Me.cboTipoTrabajo.ItemData(Me.cboTipoTrabajo.ListIndex)
    'obtengo código de registro
    codReg = Val(Me.txtCodRegEli.Text)
    
    'determino tipo de aviso
    If tipoReg = 1 Then
        tipoAviso = 114
    Else
        tipoAviso = 115
    End If
    
    'valido que se pueda eliminar el registro
    If funValidoEliminacionRegistro(tipoReg, codReg) Then
        'aviso de confirmación de operación
        If mFunMensaje(4, tipoAviso) Then
            'busco registro
            If busco_estado_habTF(tipoReg, codReg) Then
                'elimino registro
                tbTIPO_ESTADO_HAB.Delete
                'aviso de eliminación de registro
                mSubMensaje 4, tipoAviso + 2
                'grabo bitacora
                GraboBitacora "Eliminación " & Me.cboTipoTrabajo.Text & " " & codReg
                'esto desencadena el evento click del combo
                cboTipoTrabajo_Click
            End If
        End If
    End If
End Sub

Private Function funValidoEliminacionRegistro(tipoReg As Byte, codReg As Long) As Boolean
    'Determina si es posible eliminar el registro
    '--------------------------------------------------------------------------------------
    'Parámetros.
    '   Entrada:    [tipoReg] tipo del registro a eliminar
    '                           1=motivos de bloqueo
    '                           2=situaciones
    '
    '               [codReg]    código del registro a eliminar
    '
    '   Salida: True, si se seleccionó registro independientemente del tipo que sea
    '           Si es del tipo 1 (motivos de bloqueo)
    '                   True, si no existe en archivo tbBLOQUEO_HAB
    '                   False, si existe en archivo tbBLOQUEO_HAB
    '           - Si existe en el archivo de bloqueos, quiere decir que el motivo de bloqueo
    '           se utilizó o es utilizado por un bloqueo vigente.
    '
    '           Si es del tipo 2 (situaciones)
    '              True, si no existe en archvo tbHABITACIONES
    '              True, si no existe en archivo tbSITUACION_HIS
    '              False, si existe en archivo tbHABITACIONES
    '              False, si existe en archivo tbSITUACION_HIS
    '           - Si existe en el archivo de habitaciones, quiere decir que una habitación
    '           tiene asignada actualmente la situación a borrar
    '           - Si existe en el archivo de situaciones, quiere decir que una habitación
    '           tuvo asignada esa situación en algún momento.
    '--------------------------------------------------------------------------------------
    
    'por defecto asumo que se puede eliminar
    funValidoEliminacionRegistro = True
    
    'verifico que se halla seleccionado una fila de la grilla
    If codReg = 0 Then
        'aviso de selección de fila
        mSubMensaje 4, 110
        funValidoEliminacionRegistro = False
    Else
        Select Case tipoReg
            Case 1
                'verifico que no exista en tbBLOQUEO_HAB
                If Not funVerificoExistenciaArchivo(tbBLOQUEO_HAB, tbBLOQUEO_HAB("motivoBloq"), codReg) Then
                    'existe registro
                    'aviso de existencia de registro
                    mSubMensaje 4, 111
                    funValidoEliminacionRegistro = False
                End If
            Case 2
                'verifico que no exista en tbHABITACIONES
                If Not funVerificoExistenciaArchivo(tbHABITACIONES, tbHABITACIONES("situacionHab"), codReg) Then
                    'exite registro
                    'aviso de existencia de registro
                    mSubMensaje 4, 112
                    funValidoEliminacionRegistro = False
                Else
                    'verifico que no exista en tbSITUACION_HIS
                    If Not funVerificoExistenciaArchivo(tbSITUACION_HIS, tbSITUACION_HIS("situacion_situ"), codReg) Then
                        'exite registro
                        'aviso de existencia de registro
                        mSubMensaje 4, 113
                        funValidoEliminacionRegistro = False
                    End If
                End If
        End Select
    End If
End Function

Private Sub gRegistros_DblClick()
    'Selecciono registro a eliminar
    
    Dim codReg As Long
    'obtengo código de la fila seleccionada
    codReg = Val(Me.gRegistros.TextMatrix(Me.gRegistros.Row, 0))
    'verifico si puedo trabajar con estos códigos
    If funCodigosPermitidos(Me.cboTipoTrabajo.ItemData(Me.cboTipoTrabajo.ListIndex), codReg) Then
        'busco registro
        If busco_estado_habTF(Me.cboTipoTrabajo.ItemData(Me.cboTipoTrabajo.ListIndex), codReg) Then
            'muestro datos
            Me.txtCodRegEli.Text = tbTIPO_ESTADO_HAB("cod")
            Me.txtDescRegEli.Text = tbTIPO_ESTADO_HAB("descri")
            'ilumino fila en grilla
            marco_celdas_grilla Me.gRegistros, 0, 1, Me.gRegistros.Row, Me.gRegistros.Row
        End If
    End If
End Sub

Private Function funCodigosPermitidos(tipoCod As Byte, codReg As Long) As Boolean
    'Determino si puedo trabajar con los códigos seleccionados.
    'Cuando trabajo con SITUACIONES, tengo que validar que no se puedan eliminar
    'los registros 2,1 Limpia y 2,2 Sucia.
    'Estos registros se inicializan en la etapa de diseño, y el código del programa
    'asume que existen y que tienen dicho valor.(cuando se realiza el cambio de
    'situación de una habitación, de limpia a sucia y viceversa)
    '---------------------------------------------------------------------------------------
    'Parámetros:
    '   Entrada [tipoCod]   Tipo de registro del archivo tbTIPO_ESTADO_HAB. Me interesa
    '                       solamente si el tipo de registro es 2 (situaciones)
    '
    '           [codReg]    Para los registros de tipo 2, representa el tipo de situación.
    '                       De forma predefinida, la tabla contiene: 1 Limpia, 2 Sucia.
    '
    '   Salida  True, si estoy trabajando con tipoCod = 1 (motivos de bloqueo)
    '           True, si estoy trabajando con tipoCod = 2 (situaciones) y codReg es
    '               distinto de 1 y 2.
    '--------------------------------------------------------------------------------------
    'premito seleccionar registro de la grilla por defecto
    funCodigosPermitidos = True
    If tipoCod = 2 And (codReg = 1 Or codReg = 2) Then
        'no se permite trabajar con estas situaciones
        mSubMensaje 4, 118
        funCodigosPermitidos = False
    End If
End Function
