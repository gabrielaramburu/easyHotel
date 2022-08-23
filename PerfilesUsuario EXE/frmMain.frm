VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administrador de accesos de usuarios"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   11910
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   11910
   Begin Crystal.CrystalReport CrystalReport1 
      Bindings        =   "frmMain.frx":058A
      Left            =   3000
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      Connect         =   ";PWD=manyacapo;"
   End
   Begin VB.TextBox txtDescOpr 
      Height          =   1860
      Left            =   4605
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   5760
      Width           =   5175
   End
   Begin VB.Frame Frame2 
      Height          =   7605
      Left            =   2400
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   7605
      Left            =   4560
      MousePointer    =   9  'Size W E
      TabIndex        =   4
      Top             =   0
      Width           =   40
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   3
      Top             =   7605
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   582
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   9596
            MinWidth        =   9596
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   8731
            MinWidth        =   8731
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   2
            TextSave        =   "11:06 p.m."
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame3 
      Caption         =   "Frame1"
      Height          =   1575
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Visible         =   0   'False
      Width           =   4095
      Begin VB.Data Data1CrystalReport 
         Caption         =   "Data1CrystalReport"
         Connect         =   ";PWD=manyacapo;"
         DatabaseName    =   "C:\NANAHOTEL\hotel.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   405
         Left            =   1200
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   2  'Snapshot
         RecordSource    =   $"frmMain.frx":05A7
         Top             =   960
         Visible         =   0   'False
         Width           =   2655
      End
      Begin ComctlLib.ImageList ImageList1 
         Left            =   240
         Top             =   600
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
               Picture         =   "frmMain.frx":0705
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":0C57
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":1199
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":16EB
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":1C7D
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":21CF
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":2721
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":2C73
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":31C5
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":3717
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin ComctlLib.ListView lwDerecha 
      Height          =   5715
      Left            =   4605
      TabIndex        =   2
      Top             =   0
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   10081
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin ComctlLib.TreeView twUsuarios 
      Height          =   7605
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   13414
      _Version        =   327682
      Indentation     =   529
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin VB.Menu mnuUsuarios 
      Caption         =   "&Usuarios"
      Begin VB.Menu mnuUsuariosNuevo 
         Caption         =   "Nuevo"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuUsuariosCambiarContraseña 
         Caption         =   "Cambiar contraseña"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuUsuariosPermisos 
         Caption         =   "Permisos"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuUsuariosSeparador 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUsuariosEliminar 
         Caption         =   "Eliminar"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mnuOpe 
      Caption         =   "&Operaciones"
      Begin VB.Menu mnuOpeListar 
         Caption         =   "Listar"
         Shortcut        =   ^L
      End
   End
   Begin VB.Menu mnuVer 
      Caption         =   "&Ver"
      Begin VB.Menu mnuVerLista 
         Caption         =   "Lista"
      End
      Begin VB.Menu mnuVerDetalle 
         Caption         =   "Detalle"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuVerActualizar 
         Caption         =   "Actualizar"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuAyuda 
      Caption         =   "&Ayuda"
      Begin VB.Menu mnuAyudaAcercaDe 
         Caption         =   "A&cerca de..."
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
Private WithEvents PidoClave As NoMuestroUsuario
Attribute PidoClave.VB_VarHelpID = -1
Private WithEvents NuevoUsuarioS As Contraseñas
Attribute NuevoUsuarioS.VB_VarHelpID = -1
Private WithEvents NuevoAdmin As Contraseñas
Attribute NuevoAdmin.VB_VarHelpID = -1
Private WithEvents ModificoContra As Contraseñas
Attribute ModificoContra.VB_VarHelpID = -1
Private WithEvents EliminoUsuario As Contraseñas
Attribute EliminoUsuario.VB_VarHelpID = -1
Private WithEvents PantallaConfig As ConfiguroInicial
Attribute PantallaConfig.VB_VarHelpID = -1

Private AccesoPermitido As Boolean
Private continuar As Boolean
Private xIni As Single  'trabaja para mover la barra divisoria
Private posX As Single  'utilizada para controlar los márgenes

Private Usr As String   'trabaja con el menu flotante

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        mnuSalir_Click
    End If
End Sub

Private Sub Form_Load()
    'verifico que no exista otra instancia en ejecución de la aplicación
    If Not funExisteOtraInstancia Then
        'por defecto muestro el listview en forma de reporte
        m_TipoLista = 3
        
        'Leo archivo de configuración
        subLeoArchivoConfiguracion
        
        If continuar Then
            'Cambio el tamaño de la pantalla al especificado en el archivo de conf.
            subCambioTamanioPantalla
            subInicializoFormulario
            
            'Abro base de datos
            mSubAbroBaseDeDatos
                    
            If mFunAplicacionValida Then
                'pido autorización para entrar al programa
                subAutorizacion
            Else
                'no es una aplicación válida
                Unload Me
            End If
        Else
            Unload Me
        End If
    Else
        'no ejecuto aplicación
        Unload Me
    End If
End Sub

Private Sub subCambioTamanioPantalla()
    'Adapta el tamaño de la pantalla al establecido en el archivo de
    'configuración
    If Resolucion = "640x480" Then
        frmMain.Height = 7200
        frmMain.Width = 9600
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
    Open App.Path & "\PERFILESUSUARIOS.TXT" For Input As NumArch
    
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
                'ejecuto dll para crear nuevor archivo de configuración
                subPantallaConfiguracion
                Err.Number = 0
        End Select
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'si estoy ejecutando una versión demo
    If gEsUnaVersionDemo Then
        'muestro mensaje de versión demo al salir de la aplicación
        subMuestroAvisoVersionDemo
    End If

    'Descargo formulario de memoria
    Set frmMain = Nothing
End Sub

Private Sub mnuAyudaAcercaDe_Click()
    'Muestro formulario de AcercaDe
    frmAcercaDePerfilesUsuarios.Show 1
End Sub

Private Sub mnuOpeListar_Click()
    On Error Resume Next
    'verifico si hay usuarios definidos
    If mFunExistenUsuariosDef Then
        'existen usuarios
        Load frmImpresion    'cargo el formulario de impresión
        frmImpresion.Show 1
    End If
End Sub

Private Sub mnuUsuariosPermisos_Click()
    On Error Resume Next
    'verifico si hay usuarios definidos
    If mFunExistenUsuariosDef Then
        'existen usuarios
        Load frmPermisos    'cargo el formulario de permisos
        If Len(Usr) > 0 Then
            frmPermisos.cboUsr.Text = Usr
        End If
        frmPermisos.Show 1
    End If
End Sub

Private Sub mnuVerActualizar_Click()
    'Actualiza el arbol de usuarios y op.
    
    'limpio arbol
    twUsuarios.Nodes.Clear
    subCargoArbol
    twUsuarios.Nodes(1).Expanded = True
End Sub

Private Sub mnuVerDetalle_Click()
    mnuVerLista.Checked = False
    mnuVerDetalle.Checked = True
    m_TipoLista = 3
    lwDerecha.View = m_TipoLista  'reporte
End Sub

Private Sub mnuVerLista_Click()
    mnuVerLista.Checked = True
    mnuVerDetalle.Checked = False
    m_TipoLista = 2
    lwDerecha.View = m_TipoLista 'lista
End Sub

Private Sub PantallaConfig_NotificoClientes(boton As Byte)
    'Cuanso no existe el archivo de configuración ejecuto este
    'procedimiento.
    'Si confitmo la pantall de configuración vuelvo a leer el
    'archivo sino termino la ejecución deñl programa.
    
    Dim NumArch As Integer
    NumArch = FreeFile
    Dim linea As String
    If boton = 1 Then 'aceptar
        Open App.Path & "\PERFILESUSUARIOS.TXT" For Input As NumArch

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

Private Sub subPantallaConfiguracion()
    'Llamo a dll de configuración
    
    Set PantallaConfig = New ConfiguroInicial
    PantallaConfig.MostrarPantallaConfigurar _
            "PerfilesUsuarios.txt", _
            App.Path, _
            "Perfiles de Usuario"
    Set PantallaConfig = Nothing
End Sub

Private Sub lwDerecha_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Muestro menu flotante si estoy trabajando con usuarios
    
    If Button = 2 Then  'boton secundario
        'verifico que tenga columnas
        If lwDerecha.ColumnHeaders.Count > 0 Then
            'verifico que este posicionado sobre cliente
            If lwDerecha.ColumnHeaders.Item(1).Text = "Nombre " Then
                Usr = FunObtengoUsrLista
                PopupMenu mnuUsuarios
            End If
        End If
    End If
End Sub

Private Function FunObtengoUsrLista()
    'Debulve el nombre del usuario que se encuentre seleccionado
    'en la lista de usuarios.
    
    FunObtengoUsrLista = ""
    'si estoy posicionado en el nodo de "usuarios"
    If twUsuarios.SelectedItem.Text = "Usuarios" Then
        'si tengo algun usuario seleccionado en la lista de la derecha
        FunObtengoUsrLista = lwDerecha.SelectedItem
    End If
End Function

Private Sub mnuUsuariosCambiarContraseña_Click()
    'Cambia la contraseña de un usuario
    
    'verifico si hay usuarios definidos
    If mFunExistenUsuariosDef Then
        'existen usuarios
        Set ModificoContra = New Contraseñas
            ModificoContra.ModificoUsuario tbSISTEMA_USUARIOS, Usr
        Set ModificoContra = Nothing
        Usr = ""
    End If
End Sub

Private Sub mnuUsuariosEliminar_Click()
    'Elimino un usuario del sistema
        
    'verifico si hay usuarios definidos
    If mFunExistenUsuariosDef Then
        'existen usuarios
        
        Set EliminoUsuario = New Contraseñas
            EliminoUsuario.EliminoUsuario tbSISTEMA_USUARIOS, tbSISTEMA_PERFILES, Usr
        Set EliminoUsuario = Nothing
        Usr = ""
        'Cuando elimino un usuario tengo que verificar si era el último usuario.
        'Esto se debe a que si elimino todos los usuarios es necesario desactivar
        'el control de usuarios en la aplicación principal, inicializando el campo
        'tbSISTEMA_PARAMETROS("SisAdminTF") a 0, ya que no tiene sentido realizar un
        'control de usuarios si no tengo usuarios definidos.
        
        'verifico si tengo usuarios
        If mFunExistenUsuariosDef(1) = False Then
            'no existen usuarios ya que acabo de eliminar el último
            tbSISTEMA_PARAMETROS.Edit
            tbSISTEMA_PARAMETROS("SisAdminTF") = 0
            tbSISTEMA_PARAMETROS.Update
        End If
    End If
End Sub

Private Sub mnuUsuariosNuevo_Click()
    'Crea un nuevo usuario en el sistema
    Set NuevoUsuarioS = New Contraseñas
        NuevoUsuarioS.NuevoUsuario tbSISTEMA_USUARIOS
    Set NuevoUsuarioS = Nothing
    
    'Después de crear un nuevo usuario, es necesario inicializar el campo
    'tbSISTEMA_PARAMETROS("SisAdminTF") a 1, con el objetivo de activar el control de
    'ingreso en la aplicación principal.
    'El motivo por lo que esta inicialización se realiza desde este lugar, es
    'que no es necesario tener activada la seguridad si no existen usuarios.
    
    'verifico si existen usuarios
    If tbSISTEMA_USUARIOS.RecordCount > 0 Then
        'existen usuarios
        
        'NOTA: para facilitar el código se inicializa el campo
        'cada vez que se crea un usuario, siendo necesario inicializar solo cuando
        'se crea el primero.
        
        'inicializo campo para activar control de usuarios
        tbSISTEMA_PARAMETROS.Edit
            tbSISTEMA_PARAMETROS("SisAdminTF") = 1
        tbSISTEMA_PARAMETROS.Update
    End If
End Sub

Private Sub NuevoAdmin_NotificoCliente(boton As Byte)
    'Este código se ejecuta al apretar un boton en el
    'formulario de nuevo administrador
    AccesoPermitido = True
    If boton = 2 Then 'cancelar
        AccesoPermitido = False
    End If
    Set NuevoAdmin = Nothing
End Sub

Private Sub PidoClave_NotificoCliente(usuario As String, boton As Byte)
    'Este código se ejecuta al apretar un boton
    'en el formulario de ingreso de usuario admin
    
    AccesoPermitido = True
    If boton = 2 Then 'cancelar
        AccesoPermitido = False
    End If
End Sub

Private Sub subAutorizacion()
    'Este código solo se ejecuta al iniciar la aplicación
    Dim f As Date
    
    'Valido acceso a la aplicación
    If IsNull(tbSISTEMA_PARAMETROS("SisAdmin")) Then
        'NOTA: Se puede dar el caso de que el campo sisAdminTF = 0
        'y el campo sisAdmin tenga un valor. Esto se puede originar por dos motivos:
        '   a) se ejecutó la aplicación pero no se creo ningún usuario todavía
        '   b) se eliminarón todos los usuarios
        'En cualquiera de estos caso no es necesario pedir nuevamente la definición del
        'usuario administrador.
        
        'Nunca ingresó a la aplicación
        
        'Ejecuto dll para pedir contraseña para usuario Admin
        Set NuevoAdmin = New Contraseñas
            NuevoAdmin.NuevoUsuarioAdmin tbSISTEMA_PARAMETROS
        Set NuevoAdmin = Nothing
    Else
        'Ya ingresé a la aplicación y definí contraseña para usuario admin
        Set PidoClave = New NoMuestroUsuario
            'Ejecuto dll para pedir contraseña
            PidoClave.MuestroAdmin tbSISTEMA_PARAMETROS("SisAdmin")
        Set PidoClave = Nothing
    End If
    If AccesoPermitido Then 'clave administrador correcta
        subCargoArbol
    Else
        Unload Me
    End If
End Sub

Private Sub subCargoArbol()
    'Arma árbol de usuarios y opciones
    'nodo principal
    twUsuarios.Nodes.Add , , "ndoPpal", "Administrador de usuarios", 6
    subCargoUsuarios
    subCargoOperaciones
End Sub

Private Sub subCargoUsuarios()
    'Creo los nodos correspondientes a los usuarios del sistema
        
    'nodo principal de usuario
    twUsuarios.Nodes.Add "ndoPpal", 4, "Usuarios", "Usuarios", 1
    'recorro archivo de usuarios
    tbSISTEMA_USUARIOS.Index = "iclaves"
    tbSISTEMA_USUARIOS.Seek ">=", ""
    Do While Not tbSISTEMA_USUARIOS.EOF
        If Not tbSISTEMA_USUARIOS.NoMatch Then
            'creo un nodo para cada usuario
            twUsuarios.Nodes.Add "Usuarios", 4, _
            tbSISTEMA_USUARIOS("NomUsr"), _
            tbSISTEMA_USUARIOS("NomUsr"), 5
        End If
        tbSISTEMA_USUARIOS.MoveNext
    Loop
End Sub

Private Sub subCargoOperaciones()
    'Creo los nodos correspondientes a las operaciones del sistema
    'Las operaciones que aparecen aquí son generalemente opciones del sistema
    
    'NOTA: este archivo siempre tendrá información (registros) ya que las operaciones
    'del sistema se cargan en la etapa de diseño del programa, y no se modifican luego.
    
    Dim imagen As Byte
    
    'nodo principal de operaciones
    twUsuarios.Nodes.Add "ndoPpal", 4, "Operaciones", "Operaciones", 4
    
    'recorro archivo de operaciones
    tbSISTEMA_OPERACIONES.MoveFirst
    tbSISTEMA_OPERACIONES.Index = "pk_operaciones"
    Do While Not tbSISTEMA_OPERACIONES.EOF
        If Not tbSISTEMA_OPERACIONES.NoMatch Then
            'creo un nodo para cada operación
            
            If tbSISTEMA_OPERACIONES("tipoOpr") = 1 Then
                imagen = 2
            Else
                imagen = 3
            End If
            twUsuarios.Nodes.Add "Operaciones", 4, _
            "Opr" & tbSISTEMA_OPERACIONES("CodOpr"), _
            tbSISTEMA_OPERACIONES("DescOpr"), imagen
        End If
        tbSISTEMA_OPERACIONES.MoveNext
    Loop
End Sub

Private Sub subInicializoFormulario()
    'Cambia el tamaño de los controles del formulario para
    'adaptarse al tamaño de la pantalla
        
    'Muestro listview
    'ancho formulario - abcho arbol - diferencia
    'generada por barra DIVISORIA
    
    
    lwDerecha.Width = Me.Width - _
                    twUsuarios.Width - 100
                    
    txtDescOpr.Width = Me.Width - twUsuarios.Width - 100
    
    subNoMuestroDescOpr
End Sub

'***********************************************************************
'*
'*      Procedimientos para controlar barra divisoria
'*
'***********************************************************************

Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    xIni = Frame1.Left
    Frame2.Visible = True
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then  'boton del mouse apretado
        Frame2.Top = Frame1.Top
        If (xIni + X) > 600 And (xIni + X) < 9000 Then 'controlo márgenes
            posX = X
            Frame2.Left = xIni + posX
        End If
    End If
End Sub

Private Sub Frame1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Frame2.Visible = False
    'controlo márgenes nuevamente para corregir pequeña diferencia.
    If (xIni + posX) > 600 And (xIni + posX) < 9000 Then
        Frame1.Left = xIni + posX
        'Cambio ancho de ventana de árbol
        twUsuarios.Width = Frame1.Left + 15    'los 15 son para mejorar vista
        'Muevo ventana de grilla
        lwDerecha.Left = Frame1.Left + 40
        lwDerecha.Width = frmMain.Width - lwDerecha.Left - 100 'y cambio tamaño
        'Muevo ventana de descripción de operaciones
        txtDescOpr.Left = Frame1.Left + 40
        txtDescOpr.Width = frmMain.Width - txtDescOpr.Left - 100 ' y cambio tamaño
    End If
End Sub
'*************************************************************************

Private Sub twUsuarios_NodeClick(ByVal Node As ComctlLib.Node)
    'Oculto venta de descripción de operaciones
    subNoMuestroDescOpr
    Select Case Node.Key
        Case "Operaciones"
            subBorroListV
            mSubMuestroOperaciones
            'Muestro cantidad de operaciones en barra de estado
            subMuestroLeyendaEnBarra 1, Node.Children
            
        Case "Usuarios"
            subBorroListV
            mSubMuestroUsuarios
            'Muestro cantidad de usuarios en barra de estado
            subMuestroLeyendaEnBarra 2, Node.Children
            
        Case "ndoPpal"
            subMuestroLeyendaEnBarra 6, 0
        Case Else
            'si estoy sobre una operación
            If Node.Parent.Key = "Operaciones" Then
                subBorroListV
                mSubMuestroUsuariosPermitidos
                'Muestro cantidad de usuarios permitidos para
                'la operación seleccionada
                subMuestroLeyendaEnBarra 3, 0
                'Muestro ventana de descripción
                subMuestroDescOpr
            End If
            'si estoy sobre un usuario
            If Node.Parent.Key = "Usuarios" Then
                subBorroListV
                mSubMuestroOperacionesPermitidas
                'Muestro cantidad de operaciones permitidad para
                'el usuario seleccionado
                subMuestroLeyendaEnBarra 4, 0
            End If
    End Select
End Sub

Private Sub subMuestroDescOpr()
    'Este procedimiento lo llamo unicamente cuandio estoy
    'sobre una operaciones, dentro del nodo de operaciones.
    
    Dim Opr As Integer
    lwDerecha.Height = 5715 'achico lista de operaciones
    txtDescOpr.Visible = True
    Opr = Val(Mid(frmMain.twUsuarios.SelectedItem.Key, 4))
    If funBuscoOperacionTF(Opr) Then
        'muestro descripción
        txtDescOpr.Text = tbSISTEMA_OPERACIONES("InfOpr")
    End If
End Sub

Private Sub subNoMuestroDescOpr()
    lwDerecha.Height = 7605 'agrando lista de operaciones
    txtDescOpr.Visible = False
End Sub

Private Sub subMuestroLeyendaEnBarra(elemento As Byte, Cant As Variant)
    'Muestra leyenda en barra de estado.
    'El mensaje depende del elemento seleccionado
    
    'Por defecto no muestro ningín icono
    Me.StatusBar1.Panels(2).Picture = ImageList1.ListImages(10).Picture
    Select Case elemento
        Case 1  'pricipal de operaciones
            Me.StatusBar1.Panels(1).Text = _
            "Total de operaciones del sistema " & Cant
            Me.StatusBar1.Panels(2).Text = ""
        Case 2  'principal de usuarios
            Me.StatusBar1.Panels(1).Text = _
            "Total de usuarios " & Cant
            Me.StatusBar1.Panels(2).Text = ""
        Case 3  'nodos de operaciones
            Me.StatusBar1.Panels(1).Text = _
            lwDerecha.ListItems.Count & " usuarios tienen acceso a esta operación"
            
            Me.StatusBar1.Panels(2).Text = frmMain.twUsuarios.SelectedItem.Text
            'Muestro ícono
            Me.StatusBar1.Panels(2).Picture = ImageList1.ListImages(8).Picture
        Case 4  'nodos de usuarios
            Me.StatusBar1.Panels(1).Text = _
            lwDerecha.ListItems.Count & " operaciones permitidas para este usuario"
            
            Me.StatusBar1.Panels(2).Text = frmMain.twUsuarios.SelectedItem.Text
            'Muestro icono
            Me.StatusBar1.Panels(2).Picture = ImageList1.ListImages(9).Picture
            
        Case 5 'click sobre lista derecha
            Me.StatusBar1.Panels(1).Text = ""

        Case 6  'nodo principal
            Me.StatusBar1.Panels(1).Text = ""
            Me.StatusBar1.Panels(2).Text = ""
    End Select
End Sub

Private Sub subBorroListV()
    lwDerecha.ColumnHeaders.Clear
    lwDerecha.ListItems.Clear
End Sub

Private Sub mnuSalir_Click()
    Unload Me
End Sub


