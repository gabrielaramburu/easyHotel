VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seguimiento de operaciones realizadas"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   11910
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   11910
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   1455
      Left            =   5760
      TabIndex        =   9
      Top             =   4800
      Visible         =   0   'False
      Width           =   2175
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   2  'Snapshot
         RecordSource    =   "select * from sistema_bitacora,sistema_operaciones"
         Top             =   960
         Visible         =   0   'False
         Width           =   1815
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Bindings        =   "frmMain.frx":0000
         Left            =   240
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
      End
      Begin ComctlLib.ImageList ImageList1 
         Left            =   720
         Top             =   240
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   10
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":0010
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":0562
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":0AB4
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":1006
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":1558
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":1AAA
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":1FFC
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":254E
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":2AA0
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":2FF2
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   135
      Left            =   120
      TabIndex        =   8
      Top             =   7725
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   238
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Frame Frame1 
      Height          =   7600
      Left            =   8200
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   3615
      Begin ComctlLib.Toolbar Toolbar1 
         Height          =   405
         Left            =   120
         TabIndex        =   4
         Top             =   180
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   714
         ButtonWidth     =   609
         ButtonHeight    =   609
         ImageList       =   "ImageList1"
         _Version        =   327682
         BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
            NumButtons      =   6
            BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.ToolTipText     =   "Ejecutar listado"
               Object.Tag             =   ""
               ImageIndex      =   3
            EndProperty
            BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.ToolTipText     =   "Imprimir listado"
               Object.Tag             =   ""
               ImageIndex      =   4
            EndProperty
            BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Style           =   3
               MixedState      =   -1  'True
            EndProperty
            BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.ToolTipText     =   "Nuevo listado"
               Object.Tag             =   ""
               ImageIndex      =   5
            EndProperty
            BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.ToolTipText     =   "Eliminar listado"
               Object.Tag             =   ""
               ImageIndex      =   6
            EndProperty
            BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.ToolTipText     =   "Establecer como predeterminado"
               Object.Tag             =   ""
               ImageIndex      =   8
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txtInfLst 
         BackColor       =   &H80000000&
         Height          =   2655
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   4800
         Width           =   3375
      End
      Begin VB.CommandButton botCerrar 
         Height          =   255
         Left            =   3270
         Picture         =   "frmMain.frx":3544
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   180
         Width           =   255
      End
      Begin ComctlLib.TreeView twListados 
         Height          =   3825
         Left            =   120
         TabIndex        =   0
         Top             =   600
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   6747
         _Version        =   327682
         Indentation     =   527
         LineStyle       =   1
         Style           =   7
         ImageList       =   "ImageList1"
         Appearance      =   1
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Información del listado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   4560
         Width           =   1935
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   2
      Top             =   7635
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   529
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   8819
            MinWidth        =   8819
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   9437
            MinWidth        =   9437
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Nombre del último listado ejecutado"
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Alignment       =   2
            TextSave        =   "25/09/02"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Fecha del día"
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ListView lwOpr 
      Height          =   7635
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11900
      _ExtentX        =   20981
      _ExtentY        =   13467
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Menu mnuListados 
      Caption         =   "&Listados"
      Begin VB.Menu mnuListadosEjecutar 
         Caption         =   "&Ejecutar"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuListadosImprimir 
         Caption         =   "&Imprimir"
         Shortcut        =   ^I
      End
      Begin VB.Menu div1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuListadosNuevo 
         Caption         =   "&Nuevo"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuListadosEliminar 
         Caption         =   "E&liminar"
      End
      Begin VB.Menu mnuListadosPredeterminado 
         Caption         =   "E&stablecer como predeterminado"
      End
   End
   Begin VB.Menu mnuVer 
      Caption         =   "&Ver"
      Begin VB.Menu mnuVerListados 
         Caption         =   "Exlorador de listados"
         Shortcut        =   ^L
      End
   End
   Begin VB.Menu mnuAyuda 
      Caption         =   "&Ayuda"
      Begin VB.Menu mnuAyudaAcercaDe 
         Caption         =   "Acerca de..."
      End
   End
   Begin VB.Menu mnuSalir 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents PantallaConfig As ConfiguroInicial
Attribute PantallaConfig.VB_VarHelpID = -1
Private WithEvents PidoClave As UsuarioMuestro
Attribute PidoClave.VB_VarHelpID = -1

Private AccesoPermitido As Boolean
Private continuar As Boolean
Private Const OprEjecutada = 60      'número de operación correspondiente a bitácora
                                
Private Sub botCerrar_Click()
    'Cierro el cuadro de listados y agrando lista
    mSubOcultoCuadroListados
    mnuVerListados.Checked = False
End Sub

Private Sub Form_Load()
    'verifico si esxiste otra instancia del programa en ejecución
    If Not funExisteOtraInstancia Then
        'Leo archivo de configuración
        subLeoArchivoConfiguracion
        
        If continuar Then
            'Abro base de datos
            mSubAbroBaseDeDatos
            
            'determino base de datos del control data
            Data1.DatabaseName = CaminoBaseDeDatos
            
            'verifico si es una aplicación válida
            If mFunAplicacionValida Then
                'pido autorización para entrar al programa
                subAutorizacion
        
                If AccesoPermitido Then
                    If mfunUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
                        'Determino si al usuario se le permite el ingreso a la
                        'aplicación
                        
                        'Muestra el frame con los listados
                        If tbSISTEMA_BITACORAparametros("cuadrolistados") = 1 Then
                            mSubMuestroListados
                            Me.mnuVerListados.Checked = True
                        End If
                        
                        'verifico si existe listado predeterminado
                        If tbSISTEMA_BITACORAparametros("ListadoPredeterminado") <> Null Then
                            'existe listado predeterminado
                            mSubEjecutoListado tbSISTEMA_BITACORAparametros("ListadoPredeterminado")
                        End If
                    Else
                        Unload Me
                        'no se puede ejecutar el programa ya que el usuario no esta
                        'autorizado a ingresar a la aplicación
                    End If
                Else
                    Unload Me
                    'no se puede ejecutar el programa ya que se canceló
                    'el formulario de contraseñas
                End If
            Else
                'no puedo ejecutar el programa ya que no tengo archivo
                'de configuración
                Unload Me
            End If
        Else
            'no es una aplicación válida
            Unload Me
        End If
    Else
        'no ejecuto aplicación
        Unload Me
    End If
End Sub

Private Sub subLeoArchivoConfiguracion()
    'Obtiene el tamaño de la pantalla del formulario principal
    'y estable el camino a la base de datos
    On Error GoTo errores
    Dim NumArch As Integer
    Dim linea As String
    
    continuar = True
    NumArch = FreeFile
    Open App.Path & "\BITACORA.TXT" For Input As NumArch
    
    'leo archivo
    Line Input #NumArch, linea
    CaminoBaseDeDatos = linea
    Line Input #NumArch, linea
    Resolucion = linea
    
    Close NumArch
errores:
    If Err.Number <> 0 Then
        Select Case Err.Number
            Case 53
                'Si no econtre el archivo quiere decir que la aplicación
                'se está ejecutando por primera vez, por lo tanto
                'ejecuto dll para crear nuevo archivo de configuración
                subPantallaConfiguracion
                Err.Number = 0
        End Select
    End If
End Sub

Private Sub subPantallaConfiguracion()
    'Llamo a dll de configuración
    
    Set PantallaConfig = New ConfiguroInicial
        PantallaConfig.MostrarPantallaConfigurar _
                "Bitacora.txt", _
                App.Path, _
                "Bitacora"
    Set PantallaConfig = Nothing
End Sub

Private Sub subAutorizacion()
    'Determino si el usuario puede ingresar a la aplicación
    
    AccesoPermitido = False
    'Valido acceso a la aplicación
    If tbSISTEMA_PARAMETROS("SisAdminTF") = 0 Then
        'Nunca definí perfiles de usuario, por ese motivo
        'no pido contraseña ninguna.
        AccesoPermitido = True
    Else
        'Tengo definido perfiles de usuarios por lo que
        'tengo que ingresar contraseña
        
        Set PidoClave = New UsuarioMuestro
            'Ejecuto dll para pedir contraseña
            PidoClave.MuestroUsuario tbSISTEMA_USUARIOS
        Set PidoClave = Nothing
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Grabo la configuración actual del sistema
    subGraboConfActual
    'si estoy ejecutando una versión demo
    If gEsUnaVersionDemo Then
        'muestro mensaje de versión demo al salir de la aplicación
        subMuestroAvisoVersionDemo
    End If

    'Descargo formulario de memoria
    Set frmMain = Nothing
End Sub

Private Sub mnuAyudaAcercaDe_Click()
    'muestro formulario de AcercaDe
    frmAcercaDeBitacora.Show 1
End Sub

Private Sub mnuListadosEjecutar_Click()
    'verifico que existan listados definidos
    If mFunExistenListados Then
        tipo_accion_selec = 1   'ejecutar
        frmSeleccionarLst.Show 1
    End If
End Sub

Private Sub mnuListadosEliminar_Click()
    'verifico que existan listados definidos
    If mFunExistenListados Then
        tipo_accion_selec = 3   'eliminar
        frmSeleccionarLst.Show 1
    End If
End Sub

Private Sub mnuListadosImprimir_Click()
    'verifico que existan listados definidos
    If mFunExistenListados Then
        tipo_accion_selec = 2   'imprimir
        frmSeleccionarLst.Show 1
    End If
End Sub

Private Sub mnuListadosNuevo_Click()
    frmListados.Show 1
End Sub

Private Sub mnuListadosPredeterminado_Click()
    'verifico que existan listados definidos
    If mFunExistenListados Then
        tipo_accion_selec = 4   'predeterminado
        frmSeleccionarLst.Show 1
    End If
End Sub

Private Sub mnuVerListados_Click()
    'verifico que existan listados definidos
    If mFunExistenListados Then
        If mnuVerListados.Checked = True Then
            'oculto listado
            botCerrar_Click
        Else
            'Muestro el cuadro de listados
            mSubMuestroListados
            mnuVerListados.Checked = True
        End If
    End If
End Sub

Private Sub PantallaConfig_NotificoClientes(boton As Byte)
    'Cuanso no existe el archivo de configuración ejecuto este
    'procedimiento.
    'Si confirmo la pantalla de configuración vuelvo a leer el
    'archivo sino termino la ejecución del programa.
    
    Dim NumArch As Integer
    NumArch = FreeFile
    Dim linea As String
    If boton = 1 Then 'aceptar
        Open App.Path & "\BITACORA.TXT" For Input As NumArch

        'leo archivo
        Line Input #NumArch, linea
        CaminoBaseDeDatos = linea
        Line Input #NumArch, linea
        Resolucion = linea
        
        Close NumArch
    Else
        continuar = False  'finalizo ejecución
    End If
End Sub

Private Sub PidoClave_NotificoCliente(usuario As String, boton As Byte)
    'Este evento se ejecuta cuando hago click
    'en algun boton del cuadro de díalogo de contraseña de clientes
    If boton = 1 Then   'aceptar
        AccesoPermitido = True
        m_UsuarioSisNom = usuario
    Else
        AccesoPermitido = False
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    On Error GoTo error
    Select Case Button.Index
        Case 1  'boton de ejecutar
            mSubEjecutoListado Me.twListados.SelectedItem.Text
            
        Case 2  'botn imprimir
            mSubImprimoListado Me.twListados.SelectedItem.Text
            
        Case 4  'nuevo listado
            mnuListadosNuevo_Click
            'Actualizo el cuadro de listados
            mSubMuestroListados
            
        Case 5  'boton de eliminar
            mSubEliminoListado Me.twListados.SelectedItem.Text
            'Actualizo el cuadro de listados
            mSubMuestroListados
            
        Case 6  'boton de predeterminar
            mSubPredeterminarListado Me.twListados.SelectedItem.Text
            'Actualizo el cuadro de listados
            mSubMuestroListados
    End Select
error:
    If Err.Number = 91 Then
        'No encontré la manera de determinar si realmente hay algun nodo
        'seleccionado, por este motivo tuve que recurir a esto.
        'Es decir, cuando aprieto un boton de la barra de tareas y
        'no hay ningún listado seleccionado se procude el error 91
        'que es interceptado
        MsgBox "Debe de seleccionar algún listado.", vbInformation
        Err.Clear
    End If
End Sub

Private Sub twListados_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then  'secundario
        PopupMenu Me.mnuListados
    End If
End Sub

Private Sub twListados_NodeClick(ByVal Node As ComctlLib.Node)
    'Cada vez que me posiciono sobre un nodo correspondiente
    'a un listado, muestro información sobre el mismo.
    mSubMuestroInfListado Node.Text
End Sub

Private Sub mnuSalir_Click()
    Unload Me
End Sub

