VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmImpresion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Impresión de reportes"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7920
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   7920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton botSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6600
      TabIndex        =   17
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton botImprimir 
      Height          =   375
      Left            =   5280
      Picture         =   "frmImpresion.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4800
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4575
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   8070
      _Version        =   327680
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "&Usuarios"
      TabPicture(0)   =   "frmImpresion.frx":0942
      Tab(0).ControlCount=   2
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame3(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      TabCaption(1)   =   "O&peraciones"
      TabPicture(1)   =   "frmImpresion.frx":095E
      Tab(1).ControlCount=   2
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame3(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame1"
      Tab(1).Control(1).Enabled=   0   'False
      Begin VB.Frame Frame1 
         Caption         =   "Determinar filtros de información "
         Height          =   1935
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   7455
         Begin VB.ComboBox cboUsrOpr 
            Height          =   360
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   1320
            Width           =   2775
         End
         Begin VB.ComboBox cboOpr1Nivel 
            Height          =   360
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   600
            Width           =   3855
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "U&suario a listar"
            Height          =   240
            Left            =   240
            TabIndex        =   10
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "&Operaciones de 1er. nivel"
            Height          =   240
            Left            =   240
            TabIndex        =   8
            Top             =   360
            Width           =   2295
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Configuración del listado de usuarios "
         Height          =   2055
         Index           =   1
         Left            =   120
         TabIndex        =   21
         Top             =   2400
         Width           =   7455
         Begin VB.CheckBox clickMostrarVistaPrevia 
            Caption         =   "Mostrar vista previa"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   14
            Top             =   1200
            Width           =   2295
         End
         Begin VB.CheckBox clickMostrarConfirmacion 
            Caption         =   "Mostrar mensaje confirmación"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   15
            Top             =   1560
            Width           =   3015
         End
         Begin VB.ComboBox cboImpresorasSis 
            Height          =   360
            Index           =   1
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   720
            Width           =   3855
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "I&mpresora predeterminada del listado"
            Height          =   240
            Index           =   1
            Left            =   240
            TabIndex        =   12
            Top             =   480
            Width           =   3375
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Configuración del listado de usuarios "
         Height          =   2055
         Index           =   0
         Left            =   -74880
         TabIndex        =   20
         Top             =   2400
         Width           =   7455
         Begin VB.ComboBox cboImpresorasSis 
            Height          =   360
            Index           =   0
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   720
            Width           =   3855
         End
         Begin VB.CheckBox clickMostrarConfirmacion 
            Caption         =   "Mostrar mensaje confirmación"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   7
            Top             =   1560
            Width           =   3015
         End
         Begin VB.CheckBox clickMostrarVistaPrevia 
            Caption         =   "Mostrar vista previa"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   6
            Top             =   1200
            Width           =   2295
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "I&mpresora predeterminada del listado"
            Height          =   240
            Index           =   0
            Left            =   240
            TabIndex        =   4
            Top             =   480
            Width           =   3375
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Determinar filtros de información "
         Height          =   1935
         Left            =   -74880
         TabIndex        =   19
         Top             =   360
         Width           =   7455
         Begin VB.ComboBox cboTipoAcceso 
            Height          =   360
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   1320
            Width           =   3855
         End
         Begin VB.ComboBox cboUsr 
            Height          =   360
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   600
            Width           =   2775
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "&Operaciones"
            Height          =   240
            Left            =   240
            TabIndex        =   2
            Top             =   1080
            Width           =   1170
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "U&suario a listar"
            Height          =   240
            Left            =   240
            TabIndex        =   0
            Top             =   360
            Width           =   1335
         End
      End
   End
   Begin VB.Menu mnuFormulario 
      Caption         =   "&Formulario"
      Begin VB.Menu mnuFormularioImprimir 
         Caption         =   "Imprimir"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuFormularioSalir 
         Caption         =   "&Salir"
      End
   End
End
Attribute VB_Name = "frmImpresion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim primeraVezUsuarios As Boolean
Dim primeraVezOperaciones As Boolean

Private Sub Form_Load()
    'controlan la inicialización de los controles de cada ficha
    primeraVezUsuarios = True
    primeraVezOperaciones = True
    SSTab1_Click (0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmImpresion = Nothing
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    'bloqueo frame de tabs.
    subBloqueoTabs
    'determino con que listado estoy trabajando
    Select Case Me.SSTab1.TabCaption(SSTab1.Tab)
        Case "&Usuarios"
            'permito trabajar con frame de usuarios
            Me.Frame2.Enabled = True   'frame usr. filtro
            Me.Frame3(0).Enabled = True    'frame usr. conf.
            If primeraVezUsuarios Then
                'inicializo combo de usuarios
                mSubCargoComboUsr Me.cboUsr, False
                'inicializo combo de operaciones
                mSubCargoComboTipoAccesos Me.cboTipoAcceso
                'cargo impresoras en combo
                mSubCargoImpresorasInstaladas Me.cboImpresorasSis(0)
                'inicializo configuración actual
                subInicializoConfiguracion 3, 1, 0
                primeraVezUsuarios = False
            End If
            
        Case "O&peraciones"
            'permito trabajar con frame de operaciones
            Me.Frame1.Enabled = True   'frame opr. fitro
            Me.Frame3(1).Enabled = True    'frame opr. conf.
            If primeraVezOperaciones Then
                primeraVezOperaciones = False
                'inicializo combo de usuarios
                mSubCargoComboUsr Me.cboUsrOpr, True
                'inicializo operaciones primer nivel
                mSubCargoComboOpr Me.cboOpr1Nivel, True, 1
                'cargo impresoras en combo
                mSubCargoImpresorasInstaladas Me.cboImpresorasSis(1)
    
                'inicio configuración actual del listado
                subInicializoConfiguracion 3, 1, 1
            End If
    End Select
    'esta línea es necesaria para provocar que se muestren todos los
    'controles (epsecialmente combos) cuando cambio de tabs la primera vez.
    Me.Refresh
End Sub

Private Sub subBloqueoTabs()
    '-----------------------------------------------------
    'Bloqueo los frames de todos los tabs.
    'El objetivo es mejorar la interface con usuario.
    '-----------------------------------------------------
    Me.Frame1.Enabled = False   'frame opr. fitro
    Me.Frame2.Enabled = False   'frame usr. filtro
    Me.Frame3(0).Enabled = False    'frame usr. conf.
    Me.Frame3(1).Enabled = False    'frame opr. conf.
End Sub

Private Sub mSubCargoComboTipoAccesos(comboBox As comboBox)
    '---------------------------------------------------------
    'Inicializo el combo con los diferentes tipos de accesos
    '---------------------------------------------------------
    comboBox.AddItem "Todas las operaciones del sistema"
    comboBox.ItemData(comboBox.NewIndex) = 1
    comboBox.AddItem "Operaciones a las que SI tiene acceso"
    comboBox.ItemData(comboBox.NewIndex) = 2
    comboBox.AddItem "Operaciones a las que NO tiene accceso"
    comboBox.ItemData(comboBox.NewIndex) = 3
    'por defecto posiciono en el primer elemento del combo
    comboBox.ListIndex = 0
End Sub

Private Sub subInicializoConfiguracion(tipoLis As Byte, codLis As Integer, tipoCtrol As Byte)
    '--------------------------------------------------------------------------------------------
    'Muestro los datos de configuración del listado de usuarios o de operaciones.
    'Estos valores se graban en la aplicación principal, dentro del formulario de configuración
    'del sistema. Desde esta aplicación es posible cambiar sus propiedades al momento de imprimir,
    'pero no actualizarlas.
    '---------------------------------------------------------------------------------------------
    'Parámetros.
    '   [tipoLis] tipo de listado   1 = facturas
    '                               2 = varios
    '                               3 = perfiles (solo se usa este en este procedimeinto)
    '                               4 = nocrystal
    '   [codLis] código de listado  1 = perfiles por usuarios
    '                               2 = perfiles por operaciones
    '   [tipoCtrol] controles con los cuales estoy trabajando; se debe a que existe en el
    '               formulario array de controles para determinar la configuración de
    '               formularios. 0 = controles lis. usuarios
    '                            1 = controles lis opr.
    '----------------------------------------------------------------------------------------------
    
    Dim impSis As String
    'declaración de variable para utilizar biblioteca impresion.dll
    Dim biblioImpresion As ImpresionGeneral
    Set biblioImpresion = New ImpresionGeneral
    
    'obtengo datos del listado desde archivo SISTEMA_LISTADOS
    Me.clickMostrarConfirmacion(tipoCtrol).Value = mFunObtengoDatosListados(3, 1, 4)
    Me.clickMostrarVistaPrevia(tipoCtrol).Value = mFunObtengoDatosListados(3, 1, 2)
    impSis = mFunObtengoDatosListados(3, 1, 1)
    
    'verifico si el listado tiene asignada una impresora
    If IsNull(impSis) Then
        impSis = ""
    End If
    
    'verifico si la impresora es una impresora del sistema
    If biblioImpresion.mFunExisteImpresoraInstalada(impSis) Then
        'la impresora del listado esta instalada en el sistema
        'muestro impresora en combo
        Me.cboImpresorasSis(tipoCtrol).Text = impSis
    Else
        'la impresora no esta instalda
        'muestro entonces la impresora del sistema
        Me.cboImpresorasSis(tipoCtrol).Text = Printer.DeviceName
    End If
    Set biblioImpresion = Nothing
End Sub

Private Sub botImprimir_Click()
    Dim oprTodas As Boolean
    'determino que reporte estoy imprimiendo
    Select Case Me.SSTab1.TabCaption(SSTab1.Tab)
        Case "&Usuarios"
            'aplico configuración del reporte
            If mfunAplicoConfImp(Me.cboImpresorasSis(0).Text, _
            Me.clickMostrarVistaPrevia(0).Value, _
            Me.clickMostrarConfirmacion(0).Value) = 1 Then
                subArmoReporteUsuarios Me.cboUsr.Text, Me.cboTipoAcceso.ItemData(Me.cboTipoAcceso.ListIndex)
            End If
        Case "O&peraciones"
            'aplico configuración del reporte
            If mfunAplicoConfImp(Me.cboImpresorasSis(1).Text, _
            Me.clickMostrarVistaPrevia(1).Value, _
            Me.clickMostrarConfirmacion(1).Value) = 1 Then
                'determino operaciones a listar
                If Me.cboOpr1Nivel.Text = "(Todas)" Then
                    oprTodas = True
                Else
                    oprTodas = False
                End If
                subArmoReporteOperaciones oprTodas, Me.cboOpr1Nivel.ItemData(Me.cboOpr1Nivel.ListIndex), _
                Me.cboUsrOpr.Text
            End If
    End Select
End Sub

Private Sub botSalir_Click()
    Unload Me
End Sub

Private Sub mnuFormularioImprimir_Click()
    botImprimir_Click
End Sub

Private Sub mnuFormularioSalir_Click()
    botSalir_Click
End Sub


'*******************************************************
'*
'*  Impresión de listados por usuarios.
'*
'*******************************************************

Private Sub subArmoReporteOperaciones(todasOpr As Boolean, opr1Nivel As Integer, nomUsr As String)
    '----------------------------------------------------------------------------------
    'Configuro el control Crystal genérico,obtiene datos y emite el listado
    'de operaciones
    '-------------------------------------------------------------------------------
    'Parámetros.
    '   Entrada
    '           [todasOpr]      false= imprimo una operación determinada
    '                           true = imprimo todas las operaciones
    '           [opr1Nivel]     operación de primer nivel que deseo filtrar
    '           [nomUsr]        nombre del usuario que voy a imprimir
    '-------------------------------------------------------------------------------
    Dim consulta As String
    'inicializo control data
    subInicializoControlData frmMain.Data1CrystalReport
    'select sistema_operaciones.codOpr,sistema_operaciones.descOpr,sistema_usuarios.NomUsr,sistema_operaciones.PerteneceA,'' as 'usr'  from sistema_operaciones,sistema_usuarios UNION select sistema_operaciones.codOpr,sistema_operaciones.descOpr,sistema_perfiles.NomUsr,sistema_operaciones.PerteneceA,1  from sistema_perfiles,sistema_operaciones  where sistema_perfiles.codOpr = sistema_operaciones.CodOpr
    consulta = "select sistema_operaciones.codOpr,sistema_operaciones.descOpr,sistema_usuarios.NomUsr,sistema_operaciones.PerteneceA,'' as 'usr' " & _
    " from sistema_operaciones,sistema_usuarios " & _
    funFiltroOpr(todasOpr, opr1Nivel, nomUsr) & _
    " UNION " & _
    "select sistema_operaciones.codOpr,sistema_operaciones.descOpr,sistema_perfiles.NomUsr,sistema_operaciones.PerteneceA,1 " & _
    " from sistema_perfiles,sistema_operaciones  " & _
    " where sistema_perfiles.codOpr = sistema_operaciones.CodOpr " & _
    funFiltroPerfiles(todasOpr, opr1Nivel, nomUsr) & _
    " order by sistema_operaciones.codopr "

        
    'ejecuto consulta control data
    frmMain.Data1CrystalReport.RecordSource = consulta
    frmMain.Data1CrystalReport.Refresh
    
    'como de aquí en más voy a trabajar directamente con el control data
    'me aseguro de que se allan encontrado registros
    If frmMain.Data1CrystalReport.Recordset.RecordCount > 0 Then
        'determino el nombre del reporte a utilizar
        frmMain.CrystalReport1.ReportFileName = m_vardirRpt & "rptOpr1.rpt"
        
        'inicializo fórmulas
        With frmMain.CrystalReport1
            .Formulas(0) = "cabVersion = '" & mFunImpVersionAplicacion & "'"        'nombre y versión aplicación
            .Formulas(1) = "cabRegistro = '" & mFunImpRegistro & "'"                'hotel propietario de la aplicación
            .Formulas(2) = "cabHoraImp = ' " & CStr(Time()) & "'"                   'hora actual
            .Formulas(3) = "parte1Opr = '" & Me.cboOpr1Nivel.Text & "'"
            .Formulas(4) = "parte1NomUsr = '" & nomUsr & "'"
        End With
        'genero listado
        frmMain.CrystalReport1.Action = 1
        'aviso de confirmación de impresión de reporte
        mSubMensaje 4, 142  'se imprimió listado de operaciones
        'inicializo fórmulas
        mSubInicializoFormulas 4
    Else
        'aviso de que no hay datos para imprimir
        mSubMensaje 3, 9
    End If
End Sub

Private Function funFiltroOpr(todasOpr As Boolean, opr1Nivel As Integer, nomUsr As String)
    '----------------------------------------------------------------------------------
    ' Aplico filtro en primera parte consulta.
    '-------------------------------------------------------------------------------
    'Parámetros.
    '   Entrada
    '           [todasOpr]      false= imprimo una operación determinada
    '                           true = imprimo todas las operaciones
    '           [opr1Nivel]     operación de primer nivel que deseo filtrar
    '-------------------------------------------------------------------------------
    Dim parteOprFiltro As String
    Dim parteUsrFiltro As String
    
    If todasOpr = False Then
        parteOprFiltro = " sistema_operaciones.perteneceA = " & opr1Nivel
    End If
    If nomUsr <> "(Todos)" Then
        parteUsrFiltro = " sistema_usuarios.NomUsr='" & nomUsr & "'"
    End If
    
    If todasOpr = True Then
        If nomUsr = "(Todos)" Then
            'todas las opr y usr.
            funFiltroOpr = " "
        Else
            'todas opr y usr determinado
            funFiltroOpr = " where " & parteUsrFiltro & " "
        End If
    Else
        If nomUsr = "(Todos)" Then
            'opr determinado y todos los usr
            funFiltroOpr = " where " & parteOprFiltro & " "
        Else
            'opr determinado y usr determinado
            funFiltroOpr = " where " & parteOprFiltro & " and " & parteUsrFiltro & " "
        End If
    End If
End Function

Private Function funFiltroPerfiles(todasOpr As Boolean, opr1Nivel As Integer, nomUsr As String)
    '----------------------------------------------------------------------------------
    ' Aplico filtro en segunda parte consulta
    '-------------------------------------------------------------------------------
    'Parámetros.
    '   Entrada
    '           [todasOpr]      false= imprimo una operación determinada
    '                           true = imprimo todas las operaciones
    '           [opr1Nivel]     operación de primer nivel que deseo filtrar
    '-------------------------------------------------------------------------------
    Dim parteOprFiltro As String
    Dim parteUsrFiltro As String
    
    If todasOpr = False Then
        parteOprFiltro = " sistema_operaciones.perteneceA = " & opr1Nivel
    End If
    If nomUsr <> "(Todos)" Then
        parteUsrFiltro = " sistema_perfiles.NomUsr='" & nomUsr & "'"
    End If
    
    If todasOpr = True Then
        If nomUsr = "(Todos)" Then
            'todas las opr y usr.
            funFiltroPerfiles = " "
        Else
            'todas opr y usr determinado
            funFiltroPerfiles = " and " & parteUsrFiltro & " "
        End If
    Else
        If nomUsr = "(Todos)" Then
            'opr determinado y todos los usr
            funFiltroPerfiles = " and " & parteOprFiltro & " "
        Else
            'opr determinado y usr determinado
            funFiltroPerfiles = " and " & parteOprFiltro & " and " & parteUsrFiltro & " "
        End If
    End If
End Function

Private Sub subArmoReporteUsuarios(nomUsr As String, tipoFiltroOpr As Byte)
    '----------------------------------------------------------------------------------
    'Configuro el control Crystal genérico,obtiene datos y emite el listado
    'de usuarios.
    '-------------------------------------------------------------------------------
    'Parámetros.
    '   Entrada
    '           [nomUsr]                nombre del usuario que voy a imprimir
    '           [tipoFiltroOpr]         1 = todas las opr.
    '                                   2 = opr. autorizadas
    '                                   3 = opr. no autorizadas
    '-------------------------------------------------------------------------------
    Dim nomLis As String
    Dim consulta As String
    
    'inicializo control data
    subInicializoControlData frmMain.Data1CrystalReport
    
    'determino que tipo de listado voy a realizar
    Select Case tipoFiltroOpr
        Case 1
            consulta = funListadoUsrTodas(nomUsr)
            nomLis = "rptUsr1.rpt"
        Case 2
            consulta = funListadoUsrAutorizadas(nomUsr)
            nomLis = "rptUsr2.rpt"
        Case 3
            consulta = funListadoUsrNoAutorizadas(nomUsr)
            nomLis = "rptUsr3.rpt"
    End Select
    
    frmMain.Data1CrystalReport.RecordSource = consulta
    'ejecuto consulta control data
    frmMain.Data1CrystalReport.Refresh
    
    'como de aquí en más voy a trabajar directamente con el control data
    'me aseguro de que se allan encontrado regristos
    If frmMain.Data1CrystalReport.Recordset.RecordCount > 0 Then
        'determino el nombre del reporte a utilizar
        frmMain.CrystalReport1.ReportFileName = m_vardirRpt & nomLis
        
        'inicializo fórmulas
        With frmMain.CrystalReport1
            .Formulas(0) = "cabVersion = '" & mFunImpVersionAplicacion & "'"        'nombre y versión aplicación
            .Formulas(1) = "cabRegistro = '" & mFunImpRegistro & "'"                'hotel propietario de la aplicación
            .Formulas(2) = "cabHoraImp = ' " & CStr(Time()) & "'"                   'hora actual
            .Formulas(3) = "parte1NomUsr = '" & nomUsr & "'"
        End With
        'genero listado
        frmMain.CrystalReport1.Action = 1
        'aviso de confirmación de impresión de reporte
        mSubMensaje 4, 141  'se imprimió listado de usuarios
        'inicializo fórmulas
        mSubInicializoFormulas 4
    Else
        'aviso de que no hay datos para imprimir
        mSubMensaje 3, 9
    End If
End Sub

Private Function funListadoUsrAutorizadas(nomUsr As String) As String
    '------------------------------------------------------------------------
    'Creo listado de operaciones autorizadas, para un usuario determinado.
    '------------------------------------------------------------------------
    'Prámetros.
    '   [nomUsr] nombre del usuario a listar
    '------------------------------------------------------------------------
    'Muestro todas las operaciones permitidas para un usuario determinado
    funListadoUsrAutorizadas = _
    "select * from sistema_perfiles,sistema_operaciones" & _
    " where sistema_perfiles.codOpr = sistema_operaciones.codOpr " & _
    " and sistema_perfiles.nomUsr = '" & nomUsr & "'"
    
End Function

Private Function funListadoUsrNoAutorizadas(nomUsr As String) As String
    '------------------------------------------------------------------------
    'Creo listado de operaciones NO autorizadas, para un usuario determinado.
    '------------------------------------------------------------------------
    'Prámetros.
    '   [nomUsr] nombre del usuario a listar
    '------------------------------------------------------------------------
    'selecciono todas las operaciones del sistema y les resto aquellas a las que el
    'usuario tiene acceso.
    funListadoUsrNoAutorizadas = _
    "select * from sistema_operaciones " & _
    "where sistema_operaciones.codOpr NOT IN ( " & _
    "select sistema_operaciones.codOpr from sistema_perfiles,sistema_operaciones" & _
    " where sistema_perfiles.codOpr = sistema_operaciones.codOpr " & _
    " and sistema_perfiles.nomUsr = '" & nomUsr & "')"
End Function

Private Function funListadoUsrTodas(nomUsr As String) As String
    '------------------------------------------------------------------------
    'Creo listado de todas las operaciones (autorizadas o no),
    'para un usuario determinado.
    '------------------------------------------------------------------------
    'Prámetros.
    '   [nomUsr] nombre del usuario a listar
    '------------------------------------------------------------------------
    
    'selecciono todas las operaciones a las que tiene acceso un usuario determinado
    'y las uno con las operaciones a las que no tiene acceso.
    'select sistema_operaciones.CodOpr,sistema_operaciones.perteneceA,sistema_operaciones.tipoOpr,sistema_operaciones.descOpr,sistema_perfiles.nomUsr from sistema_perfiles,sistema_operaciones UNION ALL
    'select sistema_operaciones.CodOpr,sistema_operaciones.perteneceA,sistema_operaciones.tipoOpr,sistema_operaciones.descOpr,''  from sistema_operaciones
    funListadoUsrTodas = _
    "select sistema_operaciones.CodOpr," & _
    "sistema_operaciones.perteneceA," & _
    "sistema_operaciones.descOpr," & _
    "sistema_operaciones.tipoOpr," & _
    "sistema_perfiles.nomUsr " & _
    "from sistema_perfiles,sistema_operaciones" & _
    " where sistema_perfiles.codOpr = sistema_operaciones.codOpr " & _
    " and sistema_perfiles.nomUsr = '" & nomUsr & "'" & _
    " UNION ALL " & _
    "select sistema_operaciones.CodOpr," & _
    "sistema_operaciones.perteneceA," & _
    "sistema_operaciones.descOpr," & _
    "sistema_operaciones.tipoOpr," & _
    "'' " & _
    " from sistema_operaciones " & _
    "where sistema_operaciones.codOpr NOT IN ( " & _
    "select sistema_operaciones.codOpr from sistema_perfiles,sistema_operaciones" & _
    " where sistema_perfiles.codOpr = sistema_operaciones.codOpr " & _
    " and sistema_perfiles.nomUsr = '" & nomUsr & "')"
End Function


