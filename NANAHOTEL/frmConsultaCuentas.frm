VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmConsultaCuentas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Resumen de cuentas"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton botImprimirTodo 
      Caption         =   "Imprimir &todo"
      Height          =   375
      Left            =   7440
      TabIndex        =   28
      Top             =   7200
      Width           =   1455
   End
   Begin Hotel_Nana.gaHOTELtitular gaHOTELtitular1 
      Height          =   1335
      Left            =   120
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   0
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   2355
      BackColor       =   -2147483633
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5655
      Left            =   120
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1440
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   9975
      _Version        =   327680
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Resumen de extras"
      TabPicture(0)   =   "frmConsultaCuentas.frx":0000
      Tab(0).ControlCount=   2
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "dbgrid1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      TabCaption(1)   =   "Resumen de alojamiento"
      TabPicture(1)   =   "frmConsultaCuentas.frx":001C
      Tab(1).ControlCount=   2
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame6"
      Tab(1).Control(1).Enabled=   0   'False
      Begin VB.Frame Frame6 
         BorderStyle     =   0  'None
         Height          =   4935
         Left            =   4440
         TabIndex        =   26
         Top             =   480
         Width           =   7095
         Begin VB.Data Data1 
            Caption         =   "Data1"
            Connect         =   ";PWD=manyacapo;"
            DatabaseName    =   "C:\NANAHOTEL\hotel.mdb"
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   375
            Left            =   120
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   2  'Snapshot
            RecordSource    =   "select * from cuentas_aloja,sistema_constantes"
            Top             =   4440
            Visible         =   0   'False
            Width           =   3135
         End
         Begin MSDBGrid.DBGrid DBGrid3 
            Bindings        =   "frmConsultaCuentas.frx":0038
            Height          =   4575
            Left            =   120
            OleObjectBlob   =   "frmConsultaCuentas.frx":0048
            TabIndex        =   3
            Top             =   360
            Width           =   6855
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Detalle de &alojamientos cargados"
            Height          =   240
            Left            =   120
            TabIndex        =   2
            Top             =   120
            Width           =   3045
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Total a pagar"
         Height          =   4935
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   4215
         Begin VB.Frame Frame5 
            Caption         =   "Totales convertidos "
            Height          =   2655
            Left            =   240
            TabIndex        =   13
            Top             =   840
            Width           =   3735
            Begin VB.TextBox txtTotGral 
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
               Height          =   375
               Left            =   1920
               Locked          =   -1  'True
               TabIndex        =   16
               Top             =   1320
               Width           =   1335
            End
            Begin VB.TextBox txtTotAloja 
               Alignment       =   1  'Right Justify
               Height          =   375
               Left            =   1920
               Locked          =   -1  'True
               TabIndex        =   15
               Top             =   825
               Width           =   1335
            End
            Begin VB.TextBox txtTotExtra 
               Alignment       =   1  'Right Justify
               Height          =   375
               Left            =   1920
               Locked          =   -1  'True
               TabIndex        =   14
               Top             =   345
               Width           =   1335
            End
            Begin VB.Label lblSigno 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "lblSigno"
               Height          =   240
               Index           =   2
               Left            =   1080
               TabIndex        =   34
               Top             =   1380
               Width           =   735
            End
            Begin VB.Label lblSigno 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "lblSigno"
               Height          =   240
               Index           =   1
               Left            =   1080
               TabIndex        =   33
               Top             =   885
               Width           =   735
            End
            Begin VB.Label lblSigno 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "lblSigno"
               Height          =   240
               Index           =   0
               Left            =   1080
               TabIndex        =   32
               Top             =   405
               Width           =   735
            End
            Begin VB.Label labcoti 
               AutoSize        =   -1  'True
               Caption         =   "labcoti"
               Height          =   240
               Left            =   1200
               TabIndex        =   21
               Top             =   2040
               Width           =   600
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "Cotización"
               Height          =   240
               Left            =   120
               TabIndex        =   20
               Top             =   2040
               Width           =   930
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               Caption         =   "Tot. General"
               Height          =   240
               Left            =   120
               TabIndex        =   19
               Top             =   1380
               Width           =   1110
            End
            Begin VB.Label Label12 
               Caption         =   "Alojamiento"
               Height          =   255
               Left            =   120
               TabIndex        =   18
               Top             =   885
               Width           =   1335
            End
            Begin VB.Label Label11 
               Caption         =   "Extras"
               Height          =   255
               Left            =   120
               TabIndex        =   17
               Top             =   405
               Width           =   975
            End
         End
         Begin VB.TextBox txtExtraDol 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   2640
            Locked          =   -1  'True
            TabIndex        =   12
            Top             =   4440
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.TextBox txtExtraMN 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   2640
            Locked          =   -1  'True
            TabIndex        =   11
            Top             =   4080
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.TextBox txtAloja 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   10
            Top             =   3720
            Visible         =   0   'False
            Width           =   1455
         End
         Begin ComctlLib.TabStrip tabTotales 
            Height          =   3255
            Left            =   120
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   360
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   5741
            MultiRow        =   -1  'True
            _Version        =   327682
            BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
               NumTabs         =   2
               BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
                  Caption         =   "&Dólares"
                  Key             =   "do"
                  Object.Tag             =   ""
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
                  Caption         =   "&M/N"
                  Key             =   "M"
                  Object.Tag             =   ""
                  ImageVarType    =   2
               EndProperty
            EndProperty
         End
         Begin VB.Label lblSignoMN 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "lblSignoMN"
            Height          =   240
            Index           =   0
            Left            =   1440
            TabIndex        =   31
            Top             =   4080
            Visible         =   0   'False
            Width           =   1050
         End
         Begin VB.Label lblSignoDol 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "lblSignoDol"
            Height          =   240
            Index           =   1
            Left            =   1440
            TabIndex        =   30
            Top             =   4500
            Visible         =   0   'False
            Width           =   1050
         End
         Begin VB.Label lblSignoDol 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "lblSignoDol"
            Height          =   240
            Index           =   0
            Left            =   1320
            TabIndex        =   29
            Top             =   3720
            Visible         =   0   'False
            Width           =   1050
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Extras"
            Height          =   240
            Left            =   120
            TabIndex        =   25
            Top             =   4200
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Extras"
            Height          =   240
            Left            =   240
            TabIndex        =   24
            Top             =   4500
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Alojamiento"
            Height          =   240
            Left            =   120
            TabIndex        =   23
            Top             =   3840
            Visible         =   0   'False
            Width           =   1065
         End
      End
      Begin MSFlexGridLib.MSFlexGrid dbgrid1 
         Height          =   4935
         Left            =   -74880
         TabIndex        =   1
         Top             =   600
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   8705
         _Version        =   393216
         Cols            =   12
         AllowBigSelection=   0   'False
         FocusRect       =   2
         HighLight       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   $"frmConsultaCuentas.frx":0F22
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "&Gastos extras"
         Height          =   240
         Left            =   -74880
         TabIndex        =   0
         Top             =   360
         Width           =   1230
      End
   End
   Begin VB.CommandButton botImprimir 
      Height          =   375
      Left            =   9120
      Picture         =   "frmConsultaCuentas.frx":0FC4
      Style           =   1  'Graphical
      TabIndex        =   4
      Tag             =   "Imprimir"
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton botSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   10560
      TabIndex        =   5
      Top             =   7200
      Width           =   1215
   End
   Begin Hotel_Nana.gaHOTELbarra gaHOTELbarra1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   7605
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   582
   End
   Begin Hotel_Nana.gaHOTELcli gaHOTELcli1 
      Height          =   735
      Left            =   4800
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   1296
      BackColor       =   -2147483633
   End
   Begin VB.Menu mnuFormulario 
      Caption         =   "&Formulario"
      Begin VB.Menu mnuFormularioAceptar 
         Caption         =   "Aceptar"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnudiv1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFormularioImprimir 
         Caption         =   "Imprimir"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuFormularioImprimirTodo 
         Caption         =   "Imprimir todo"
         Shortcut        =   ^T
      End
   End
   Begin VB.Menu mnuVer 
      Caption         =   "&Ver información de..."
      Begin VB.Menu mnuVerExtras 
         Caption         =   "Resumen de extras"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuVerAloja 
         Caption         =   "Resumen de alojamiento"
         Shortcut        =   {F6}
      End
   End
End
Attribute VB_Name = "frmConsultaCuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'utilizadas en corte para alojamiento
Private dias As Integer
Private total_dia As Double
Private fecha_ant As Date

Private primera_vez As Boolean

Private total_dia_dol As Double
Private total_dia_mnac As Double
Private total_hab_dol As Double
Private total_hab_mnac As Double
Private total_hab_dol_gral As Double
Private total_hab_mnac_gral As Double
Private total_alojamiento As Double
Private cotizacion As Double

Private titular_aloja As Long
Private titular_extra As Long

Private habCuenta As Long   'almacena la habitación con la que se esta trabajando
                            'se utiliza cuando tipo_accion_ConsultaCuentas=1
                            
Private cliCuenta As Long   'almacena el número de cliente con el que se esta trabajando
                            'se utiliza cuando tipo_accion_ConsultaCuentas = 2
                            
Private Sub inicializo_var()
    total_dia_dol = 0
    total_dia_mnac = 0
    total_hab_dol = 0
    total_hab_mnac = 0
    total_hab_dol_gral = 0
    total_hab_mnac_gral = 0
End Sub

Private Sub botSalir_Click()
    Unload Me
    Select Case tipo_accion_ConsultaCuentas
        Case 1 'por habitación
            frmIngHabitacion.Show 1
        Case 2  'por cliente
            frmIngPaxEmp.Show 1
    End Select
End Sub

Private Sub Form_Activate()
    'por defecto muestro el tabs de gastos extras
    Me.ssTab1.Tab = 0
End Sub

Private Sub Form_Load()
    'Apariencia formulario
    mSubConfiguroFuentesControlesSistema Me
    'obtengo cotización del archivo
    cotizacion = busco_cotiza
    'inicializo control data
    subInicializoControlData Me.Data1
        
    Select Case tipo_accion_ConsultaCuentas
        Case 1  'por habitacion
            habCuenta = Val(frmIngHabitacion.txtNroHab.Text)
            cabezal_formulario_habitacion habCuenta
            obtengo_titular_habitacion habCuenta
            
        Case 2  'por cliente
            cliCuenta = Val(frmIngPaxEmp.txtCodCli.Text)
            cabezal_formulario_cliente cliCuenta
            obtengo_titular_cliente cliCuenta
    End Select
    
    mSub_bloqueo_controles_formulario Me, True
    inicializo_var
    primera_vez = True
    genero_consulta_extras
    'configuro cabezal grilla extras
    subConfiguroCabezalExtras
    'la última columna de la grilla no la muestro
    DBGrid1.ColWidth(11) = 0
End Sub

Private Sub subConfiguroCabezalExtras()
    '-------------------------------------------------
    'Realizo el cabezal de la grilla de gastos extras
    '-------------------------------------------------
    Me.DBGrid1.FormatString = _
    "| Fecha    " & _
    "| Hab. " & _
    "| Bol.   " & _
    "| Cód. " & _
    "| Descripción                                           " & _
    "| Cant." & _
    "| P.uni. " & gblSignoMonedaNacional & "   " & _
    "| Total  " & gblSignoMonedaNacional & "   " & _
    "| P.uni. " & gblSignoDolares & "   " & _
    "| Total  " & gblSignoDolares & "    |"
End Sub

Private Sub cabezal_formulario_habitacion(hab_cuenta As Long)
    'No muestro cabezal de clientes
    Me.gaHOTELcli1.Visible = False
    
    Me.gaHOTELtitular1.CaminoBaseDeDatos = vardir
    Me.gaHOTELtitular1.NumeroHabitacion = hab_cuenta
    Me.gaHOTELtitular1.Width = 11655
End Sub

Private Sub obtengo_titular_habitacion(hab_cuenta As Long)
    'Obtengo el titular o los titulares de la habitción seleccionada.
    titular_extra = busco_titular_hab2(hab_cuenta, "extra")
    titular_aloja = busco_titular_hab2(hab_cuenta, "aloja")
End Sub

Private Sub cabezal_formulario_cliente(cli_cuenta As Long)
    'No muestro cabezal de habitación
    Me.gaHOTELtitular1.Visible = False
    
    Me.gaHOTELcli1.CaminoBaseDeDatos = vardir
    Me.gaHOTELcli1.CodigoCliente = cli_cuenta
    
    'Dibujo cabezal de clientes
    Me.gaHOTELcli1.Left = 120
    Me.gaHOTELcli1.Width = 11655
    Me.ssTab1.Top = 960
End Sub

Private Sub obtengo_titular_cliente(cli_cuenta As Long)
    'Cuando realizo una consulta cuenta por cliente, no es necesario
    'discriminar las cuentas, ya que tengo que mostrar ambos tipo de gastos
    '(alojamiento y extras) en caso de que el cliente los tenga.
    titular_extra = cli_cuenta
    titular_aloja = cli_cuenta
End Sub

Private Sub genero_consulta_extras()
    'Recorro los gastos de extras por el titular y realizo un corte por fecha
    Dim fecha_gasto_ant As Date
    
    tbCUENTAS.Index = "i_titular"
    tbCUENTAS.Seek ">=", 0, titular_extra, 0
    If Not tbCUENTAS.NoMatch Then   'si se posiciona
        Do While Not tbCUENTAS.EOF
            If tbCUENTAS("titular_cuenta") = titular_extra And _
             tbCUENTAS("facturado") = 0 Then
                fecha_gasto_ant = tbCUENTAS("fechagasto_cuenta")
                inicilizo_dia
                Do While Not tbCUENTAS.EOF
                    If tbCUENTAS("titular_cuenta") = titular_extra _
                    And tbCUENTAS("fechagasto_cuenta") = fecha_gasto_ant And _
                    tbCUENTAS("facturado") = 0 Then
                        muestro_grilla
                        'acumulo totales dia
                        total_dia_dol = total_dia_dol + tbCUENTAS("total_dolares_cuenta")
                        total_dia_mnac = total_dia_mnac + tbCUENTAS("total_mnacional_cuenta")
                    
                        tbCUENTAS.MoveNext
                    Else
                        Exit Do
                    End If
                Loop
                muestro_total_día
                totales_habitacion
            Else
                Exit Do
            End If
        Loop
        muestro_total_habitacion
    End If
    
End Sub

Private Sub subGeneroConsultaAlojamiento()
    '--------------------------------------------------------------------------
    'Muestro los gastos de alojamiento asignados a un titular determinado.
    '--------------------------------------------------------------------------
    Dim consultaSql As String
    
    'Muestro todos los alojamientos no facturados pertenecientes a un mismo titular
    consultaSql = _
    "select * from cuentas_aloja,sistema_constantes " & _
    "where cuentas_aloja.tipoAloja = sistema_constantes.codConst and " & _
    "sistema_constantes.tipoConst = 1 and " & _
    "cuentas_aloja.facturado = 0 and " & _
    "cuentas_aloja.titular_aloja = " & titular_aloja & _
    " order by fecha,tipoAloja ASC"
    
    Me.Data1.RecordSource = consultaSql
    Me.Data1.Refresh
    
    'calculo total de alojamiento
    total_alojamiento = 0
    If Data1.Recordset.RecordCount > 0 Then
        'existen alojamientos cargados
        'recorro los alojamientos
        Data1.Recordset.MoveFirst
        Do While Not Data1.Recordset.EOF
            total_alojamiento = total_alojamiento + Data1.Recordset.Fields("tarifa")
            Data1.Recordset.MoveNext
        Loop
    End If
End Sub

Private Sub totales_habitacion()
    total_hab_dol = total_hab_dol + total_dia_dol
    total_hab_mnac = total_hab_mnac + total_dia_mnac
End Sub

Private Sub totales_convertidos()
    'paso total moneda nacional a dolares
    total_hab_dol_gral = total_hab_mnac / cotizacion
    total_hab_dol_gral = total_hab_dol_gral + total_hab_dol
    'paso total dolares a pesos
    total_hab_mnac_gral = total_hab_dol * cotizacion
    total_hab_mnac_gral = total_hab_mnac_gral + total_hab_mnac
End Sub

Private Sub muestro_total_habitacion()
    Dim linea_total As String
    'agrego linea total
    linea_total = _
    Chr(9) & _
    Chr(9) & _
    Chr(9) & _
    Chr(9) & _
    Chr(9) & _
    Chr(9) & _
    Chr(9) & _
    Chr(9) & Format(total_hab_mnac, "####0.00;;#") & _
    Chr(9) & _
    Chr(9) & Format(total_hab_dol, "####0.00;;#")
    
    DBGrid1.AddItem Empty
    
    DBGrid1.AddItem linea_total
    marco_totales mSisColor_3TotalDeGastosTitular, 4
    cambio_fuente_totales mSisColor_3TotalDeGastosTitular
    
    'agrego linea total convertidos
    totales_convertidos
    
    'formacion de linea
    linea_total = _
    Chr(9) & _
    Chr(9) & _
    Chr(9) & _
    Chr(9) & _
    Chr(9) & "Totales ---> " & _
    Chr(9) & _
    Chr(9) & gblSignoMonedaNacional & _
    Chr(9) & Format(total_hab_mnac_gral, "####0.00;;#") & _
    Chr(9) & gblSignoDolares & _
    Chr(9) & Format(total_hab_dol_gral, "####0.00;;#")
    
    DBGrid1.AddItem linea_total
    marco_totales mSisColor_3TotalDeGastosTitular, 4
    cambio_fuente_totales mSisColor_3TotalDeGastosTitular
    
    formato_totales
End Sub

Private Sub formato_totales()
    'cambio fuente de celda totales
    DBGrid1.col = 5
    DBGrid1.CellFontWidth = 10
    DBGrid1.CellFontBold = True
    
    'signo m/nac
    DBGrid1.col = 7
    DBGrid1.CellFontBold = True
    DBGrid1.CellAlignment = 7
    
    'signo dolares
    DBGrid1.col = 9
    DBGrid1.CellFontBold = True
    DBGrid1.CellAlignment = 7
    
    'linea bacia para adorno
    DBGrid1.AddItem Empty
    marco_totales mSisColor_3TotalDeGastosTitular, 4
End Sub

Private Sub inicilizo_dia()
    total_dia_dol = 0
    total_dia_mnac = 0
End Sub

Private Sub muestro_total_día()
    Dim linea_grilla As String
    'formacion de linea
    linea_grilla = _
    Chr(9) & _
    Chr(9) & _
    Chr(9) & _
    Chr(9) & _
    Chr(9) & _
    Chr(9) & _
    Chr(9) & _
    Chr(9) & Format(total_dia_mnac, "####0.00;;#") & _
    Chr(9) & _
    Chr(9) & Format(total_dia_dol, "####0.00;;#")
    
    DBGrid1.AddItem linea_grilla
    marco_totales mSisColor_2TotalDeGastosDiarios, 7
    cambio_fuente_totales mSisColor_2TotalDeGastosDiarios
End Sub

Private Sub marco_totales(color As OLE_COLOR, inicio As Byte)
    'marca toda la fila de marron
    DBGrid1.row = DBGrid1.Rows - 1
    Dim i As Byte
    i = inicio
    Do While i < DBGrid1.Cols - 1
        DBGrid1.col = i
        DBGrid1.CellBackColor = color
        DBGrid1.ForeColor = &H80000012  'negro
        i = i + 1
    Loop
End Sub

Private Sub cambio_fuente_totales(color As OLE_COLOR)
    DBGrid1.col = 8
    'dbgrid1.CellFontBold = True
    DBGrid1.CellBackColor = color
    
    DBGrid1.col = 10
    'dbgrid1.CellFontBold = True

    DBGrid1.CellBackColor = color
    
    'cambio ancho de la fila
    DBGrid1.RowHeight(DBGrid1.row) = 300
End Sub

Private Sub muestro_grilla()
    Dim linea_grilla As String
    Dim descri_art As String

    If busco_articuloTF(Val(tbCUENTAS("articulo_cuenta"))) Then
        descri_art = tbARTICULOS("descriarticulo")
    End If
    
    'formacion de linea
    linea_grilla = _
    Chr(9) & tbCUENTAS("fechagasto_cuenta") & _
    Chr(9) & tbCUENTAS("habitacion_cuenta") & _
    Chr(9) & tbCUENTAS("boleta_cuenta") & _
    Chr(9) & tbCUENTAS("articulo_cuenta") & _
    Chr(9) & descri_art & _
    Chr(9) & tbCUENTAS("cantidad_cuenta") & _
    Chr(9) & Format(tbCUENTAS("importe_mnacional_cuenta"), "####0.00;;#") & _
    Chr(9) & Format(tbCUENTAS("total_mnacional_cuenta"), "####0.00;;#") & _
    Chr(9) & Format(tbCUENTAS("importe_dolares_cuenta"), "####0.00;;#") & _
    Chr(9) & Format(tbCUENTAS("total_dolares_cuenta"), "####0.00;;#") & _
    Chr(9) & tbCUENTAS("nrocorr_cuenta")
    
    DBGrid1.AddItem linea_grilla
    marco_totales mSisColor_1DetalleDeGastos, 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Descargo formulario de memoria
    Set frmConsultaCuentas = Nothing
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
   If ssTab1.Tab = 1 Then
        If primera_vez Then
            'inicializo etiquetas de símbolos de moneda
            Me.lblSignoDol(0).Caption = gblSignoDolares
            Me.lblSignoDol(1).Caption = gblSignoDolares
            Me.lblSignoMN(0).Caption = gblSignoMonedaNacional
            
            subGeneroConsultaAlojamiento
            muestro_datos
            tabTotales_Click
            primera_vez = False
        End If
    'realizo refresh
    Frame2.Refresh
    Frame6.Refresh
    Frame5.Refresh
    End If
End Sub

Private Sub muestro_datos()
    'totales
    txtExtraMN.Text = Format(total_hab_mnac, "####0.00;;#")
    txtExtraDol.Text = Format(total_hab_dol, "####0.00;;#")
    txtAloja.Text = Format(total_alojamiento, "####0.00;;#")
End Sub

Private Sub tabTotales_Click()
    Dim aloja As Double
    Dim extra As Double
    Dim gral As Double
    labcoti.Caption = gblSignoMonedaNacional & " " & Format(cotizacion, "####0.00")
    If tabTotales.SelectedItem.Key = "M" Then 'm/n
        extra = total_hab_mnac_gral
        aloja = total_alojamiento * cotizacion
        txtTotAloja.Text = Format(aloja, "####0.00")
        txtTotExtra.Text = Format(extra, "####0.00")
        lblSigno(0).Caption = gblSignoMonedaNacional
        lblSigno(1).Caption = gblSignoMonedaNacional
        lblSigno(2).Caption = gblSignoMonedaNacional
    Else
        extra = total_hab_dol_gral
        aloja = total_alojamiento
        txtTotExtra.Text = Format(extra, "####0.00")
        txtTotAloja.Text = Format(aloja, "####0.00")
        lblSigno(0).Caption = gblSignoDolares
        lblSigno(1).Caption = gblSignoDolares
        lblSigno(2).Caption = gblSignoDolares
    End If
    gral = extra + aloja
    txtTotGral.Text = Format(gral, "####0.00")
End Sub

Private Sub mnuFormularioAceptar_Click()
    'Equivale a presionar el boton de aceptar o la tecla F12
    botSalir_Click
End Sub

Private Sub mnuFormularioImprimir_Click()
    'Equivale a presionar el boton de imprimir o la tecla Ctrol+I
    botImprimir_Click
End Sub

Private Sub mnuVerAloja_Click()
    'Muestra la ficha de gastos alojamiento
    Me.ssTab1.Tab = 1
End Sub

Private Sub mnuVerExtras_Click()
    'Muestra la ficha de gastos extras
    Me.ssTab1.Tab = 0
End Sub

Private Sub mnuFormularioImprimirTodo_Click()
    'Imrimo reporte de gastos extras y alojamiento
    botImprimirTodo_Click
End Sub

'******************************************************
'*
'*  Impresión de reportes
'*
'******************************************************

Private Sub botImprimirTodo_Click()
    'Imprimo reporte de gastos extras y alojamiento
    If mfunAplicoConfImp(2, 7) = 1 Then
        'realizo reporte
        suArmoReporteAlojamientoExtras titular_aloja, titular_extra, 0, tipo_accion_ConsultaCuentas
    End If
End Sub

Private Sub botImprimir_Click()
    'Imprimo reportes
    If Me.ssTab1.Tab = 0 Then 'si estoy mostrando gastos extras
        'imprimo gastos extras
        If mfunAplicoConfImp(2, 5) = 1 Then 'listado de gastos extras
            'realizo listado
            subArmoReporteExtras titular_extra, tipo_accion_ConsultaCuentas
        End If
    Else
        'imprimo gastos alojamiento
        If mfunAplicoConfImp(2, 6) = 1 Then 'listado de gastos alojamientos
            'realizo listado
            subArmoReporteAlojamiento titular_aloja, tipo_accion_ConsultaCuentas
        End If
    End If
End Sub

Private Sub suArmoReporteAlojamientoExtras(titularAloja As Long, titularExtra As Long, _
                                            titularUnico As Long, tipoReporte As Byte)
    '----------------------------------------------------------------------------------
    'Configuro el control Crystal genérico,obtengo datos y emite el listado
    'de gastos extras y de alojamiento
    '-------------------------------------------------------------------------------
    'Parámetros.
    '   Entrada:    [titularAloja] número del titular de los gastos alojamiento.
    '               [titularExtra] número del titular de los gastos extras
    '               [titularUnico] número del titular único
    '               [tipoReporte]   determina si el origen del listado es
    '               a partir de una habitación o a partir de un cliente.
    '               1 = listado por habitación
    '               2 = listado por cliente
    '-------------------------------------------------------------------------------
    Dim nroCliTitAloja As Long
    Dim nacCliTitAloja As Integer
    
    'inicializo control data
     subInicializoControlData frmMAIN.Data1CrystalReport
    
    'NOTA: esta consulta debe de ser la misma que se utilizo para crear el reporte,
    'por ese motivo no debe de ser modificada en aspectos tales como campos a mostrar y
    'tablas utilizadas.

    'establesco consulta a utilizar
    'selecciono los gastos extras y alojamiento con el mismo criterio que lo hago para
    'los listado individuales. Después realizo una union de las dos consultas.

    frmMAIN.Data1CrystalReport.RecordSource = "select " & _
    "fechaGasto_cuenta as fecha,habitacion_cuenta as 'hab'," & _
    "articulos.descriArticulo as 'desc',total_mnacional_cuenta as 'totMN'," & _
    "total_dolares_cuenta as 'totDOL',cantidad_cuenta as 'cant'," & _
    "' ' as 'obs' " & _
    "from cuentas_extra,articulos " & _
    "where cuentas_extra.articulo_cuenta = articulos.nroArticulo and " & _
    "cuentas_extra.facturado = 0 and " & _
    "cuentas_extra.titular_cuenta = " & titularExtra & _
    " UNION ALL select " & _
    "fecha as 'fecha',habitacion_cuenta_aloja as 'hab'," & _
    "sistema_constantes.descConst as 'desc',' ' as 'totMN'," & _
    "tarifa as 'totDOL','' as 'cant',obsAloja as 'obs' " & _
    " from cuentas_aloja,sistema_constantes " & _
    "where cuentas_aloja.tipoAloja = sistema_constantes.codConst and " & _
    "sistema_constantes.tipoConst = 1 and " & _
    "cuentas_aloja.facturado = 0 and " & _
    "cuentas_aloja.titular_aloja = " & titularAloja '& _
    'ejecuto consulta control data
    frmMAIN.Data1CrystalReport.Refresh
    
    'como de aquí en más voy a trabajar directamente con el control data
    'me aseguro de que se allan encontrado gastos extras.
    If frmMAIN.Data1CrystalReport.Recordset.RecordCount > 0 Then
        'determino el nombre del reporte a utilizar
        frmMAIN.CrystalReport1.ReportFileName = vardir2 + "rptGastos3t.rpt"

        'inicializo fórmulas
        With frmMAIN.CrystalReport1
            .Formulas(0) = "cabVersion = '" & mFunImpVersionAplicacion & "'"        'nombre y versión aplicación
            .Formulas(1) = "cabRegistro = '" & mFunImpRegistro & "'"                'hotel propietario de la aplicación
            .Formulas(2) = "cabHoraImp = ' " & CStr(Time()) & "'"                   'hora actual
            Select Case tipoReporte
                Case 1  'por habitación
                    .Formulas(3) = "parte1Titulo = 'Habitación'"
                    .Formulas(4) = "parte1Hab = '" & habCuenta & " Suite " & busco_tipo_hab_descri(habCuenta) & "'"
                    .Formulas(5) = "parte2Titulo = 'Titular " & mFunBuscoDescripcionTipoTitular(habCuenta, "extra") & "'"
                    .Formulas(6) = "parte2NomTitular = '" & busco_titular_hab(habCuenta, "extra") & "'"
                    .Formulas(7) = "parte3Titulo = 'Titular " & mFunBuscoDescripcionTipoTitular(habCuenta, "aloja") & "'"
                    .Formulas(8) = "parte3NomTitular = '" & busco_titular_hab(habCuenta, "aloja") & "'"
                    .Formulas(9) = "TipoOrigenConsulta = 'Por habitación'"
                    'obtengo número cliente para poder obtener nacionalidad y después tipo de impuesto
                    nroCliTitAloja = busco_titular_hab2SinCambiarPunteroHab(habCuenta, "aloja")
                Case 2      'por cliente
                    .Formulas(3) = "parte1Titulo = ''"
                    .Formulas(4) = "parte1Hab = ''"
                    .Formulas(5) = "parte2Titulo = 'Titular'"
                    .Formulas(6) = "parte2NomTitular = '" & obtengo_nombre_pasajero(cliCuenta) & "'"
                    .Formulas(7) = "parte3Titulo = ''"
                    .Formulas(8) = "parte3NomTitular = ''"
                    .Formulas(9) = "TipoOrigenConsulta = 'Por cliente'"
                    'obtengo número cliente para poder obtener nacionalidad y después tipo de impuesto
                    nroCliTitAloja = cliCuenta
            End Select
            .Formulas(10) = "simboloMn = '" & gblSignoMonedaNacional & "'"
            .Formulas(11) = "simboloDol = '" & gblSignoDolares & "'"
            .Formulas(12) = "ValorCotizacion = '" & mFunObtengoUltimaCotizacion(1, 1, m_FechaSis) & "'"
            'obtengo nacionalidad
            nacCliTitAloja = Val(mfunObtengoDatosCli(4, nroCliTitAloja))
            .Formulas(13) = "porcentajeIvaAloja = '" & mFunTipoIvaALoja(nacCliTitAloja, 2) & "'"
        End With
        'genero listado
        frmMAIN.CrystalReport1.Action = 1
        'aviso de confirmación de impresión de gastos extras
        mSubMensaje 4, 136   'se imprimieron los gastos extras
        'inicializo fórmulas
        mSubInicializoFormulas 13
    Else
        'aviso de que no hay gastos para imprimir
        mSubMensaje 3, 9
    End If
End Sub
Private Sub subArmoReporteAlojamiento(titularAloja As Long, tipoReporte As Byte)
    '-------------------------------------------------------------------------------
    'Configuro el control Crystal genérico,obtengo datos y emite el listado
    'de gastos de alojamiento
    '-------------------------------------------------------------------------------
    'Parámetros.
    '   Entrada:    [titularAloja] número del titular de los gastos alojamiento.
    '               [tipoReporte]   determina si el origen del listado es
    '               a partir de una habitación o a partir de un cliente.
    '               1 = listado por habitación
    '               2 = listado por cliente
    '-------------------------------------------------------------------------------
    'inicializo control data
     subInicializoControlData frmMAIN.Data1CrystalReport
     
    'NOTA: esta consulta debe de ser la misma que se utilizo para crear el reporte,
    'por ese motivo no debe de ser modificada en aspectos tales como campos a mostrar y
    'tablas utilizadas.

    'establesco consulta a utilizar
    'Muestro todos los alojamientos no facturados pertenecientes a un mismo titular
    frmMAIN.Data1CrystalReport.RecordSource = _
    "select * from cuentas_aloja,sistema_constantes " & _
    "where cuentas_aloja.tipoAloja = sistema_constantes.codConst and " & _
    "sistema_constantes.tipoConst = 1 and " & _
    "cuentas_aloja.facturado = 0 and " & _
    "cuentas_aloja.titular_aloja = " & titularAloja & _
    " order by fecha,tipoAloja ASC"
    'ejecuto consulta control data
    frmMAIN.Data1CrystalReport.Refresh
    
    'como de aquí en más voy a trabajar directamente con el control data
    'me aseguro de que se allan encontrado gastos extras.
    If frmMAIN.Data1CrystalReport.Recordset.RecordCount > 0 Then
        'determino el nombre del reporte a utilizar
        frmMAIN.CrystalReport1.ReportFileName = vardir2 + "rptGastos2a.rpt"

        'inicializo fórmulas
        With frmMAIN.CrystalReport1
            .Formulas(0) = "cabVersion = '" & mFunImpVersionAplicacion & "'"        'nombre y versión aplicación
            .Formulas(1) = "cabRegistro = '" & mFunImpRegistro & "'"                'hotel propietario de la aplicación
            .Formulas(2) = "cabHoraImp = ' " & CStr(Time()) & "'"                   'hora actual
            Select Case tipoReporte
                Case 1  'por habitación
                    .Formulas(3) = "parte1Titulo = 'Habitación'"
                    'obtengo descripción del tipo de habitación
                    .Formulas(4) = "parte1Hab = '" & habCuenta & " Suite " & busco_tipo_hab_descri(habCuenta) & "'"
                    .Formulas(5) = "parte2Titulo = 'Titular " & mFunBuscoDescripcionTipoTitular(habCuenta, "aloja") & "'"
                    .Formulas(6) = "parte2NomTitular = '" & busco_titular_hab(habCuenta, "aloja") & "'"
                    .Formulas(7) = "TipoOrigenConsulta = 'Por habitación'"
                Case 2      'por cliente
                    .Formulas(3) = "parte1Titulo = ''"
                    .Formulas(4) = "parte1Hab = ''"
                    .Formulas(5) = "parte2Titulo = 'Titular'"
                    .Formulas(6) = "parte2NomTitular = '" & obtengo_nombre_pasajero(cliCuenta) & "'"
                    .Formulas(7) = "TipoOrigenConsulta = 'Por cliente'"
            End Select
            .Formulas(8) = "simboloMonedaNacional = '" & gblSignoMonedaNacional & "'"
            .Formulas(9) = "simboloDolares = '" & gblSignoDolares & "'"
            .Formulas(10) = "ValorCotizacion = '" & mFunObtengoUltimaCotizacion(1, 1, m_FechaSis) & "'"
            .Formulas(11) = "porcentajeIvaAloja = '" & mFunTipoIvaALoja(Val(mfunObtengoDatosCli(4, titularAloja)), 2) & "'"
        End With
        'genero listado
        frmMAIN.CrystalReport1.Action = 1
        'aviso de confirmación de impresión de gastos extras
        mSubMensaje 4, 135   'se imprimieron los gastos extras
        'inicializo fórmulas
        mSubInicializoFormulas 11
    Else
        'aviso de que no hay gastos para imprimir
        mSubMensaje 3, 9
    End If
End Sub

Private Sub subArmoReporteExtras(titularExtras As Long, tipoReporte As Byte)
    '-------------------------------------------------------------------------------
    'Configuro el control Crystal genérico,obtengo datos y emite el listado
    '-------------------------------------------------------------------------------
    'Parámetros.
    '   Entrada:    [titularExtras] número del titular de los gastos extras.
    '               [tipoReporte]   determina si el origen del listado es
    '               a partir de una habitación o a partir de un cliente.
    '               1 = listado por habitación
    '               2 = listado por cliente
    '-------------------------------------------------------------------------------
    'inicializo control data
     subInicializoControlData frmMAIN.Data1CrystalReport
     
    'establesco consulta a utilizar
    'NOTA: esta consulta debe de ser la misma que se utilizo para crear el reporte,
    'por ese motivo no debe de ser modificada en aspectos tales como campos a mostrar y
    'tablas utilizadas.

    frmMAIN.Data1CrystalReport.RecordSource = _
    "select * from cuentas_extra,articulos " & _
    "where cuentas_extra.articulo_cuenta = articulos.nroArticulo and " & _
    "cuentas_extra.facturado = 0 and " & _
    "cuentas_extra.titular_cuenta = " & titularExtras & _
    " order by fechaGasto_Cuenta ASC"

    'ejecuto consulta control data
    frmMAIN.Data1CrystalReport.Refresh
    
    'como de aquí en más voy a trabajar directamente con el control data
    'me aseguro de que se allan encontrado gastos extras.
    If frmMAIN.Data1CrystalReport.Recordset.RecordCount > 0 Then
        'determino el nombre del reporte a utilizar
        frmMAIN.CrystalReport1.ReportFileName = vardir2 + "rptGastos1e.rpt"

        'inicializo fórmulas
        With frmMAIN.CrystalReport1
            .Formulas(0) = "cabVersion = '" & mFunImpVersionAplicacion & "'"        'nombre y versión aplicación
            .Formulas(1) = "cabRegistro = '" & mFunImpRegistro & "'"                'hotel propietario de la aplicación
            .Formulas(2) = "cabHoraImp = ' " & CStr(Time()) & "'"                   'hora actual
            Select Case tipoReporte
                Case 1  'por habitación
                    .Formulas(3) = "parte1Titulo = 'Habitación'"
                    'obtengo descripción del tipo de habitación
                    .Formulas(4) = "parte1Hab = '" & habCuenta & " Suite " & busco_tipo_hab_descri(habCuenta) & "'"
                    .Formulas(5) = "parte2Titulo = 'Titular " & mFunBuscoDescripcionTipoTitular(habCuenta, "extra") & "'"
                    .Formulas(6) = "parte2NomTitular = '" & busco_titular_hab(habCuenta, "extra") & "'"
                    .Formulas(7) = "parte1TipoOrigenConsulta = 'Por habitación'"
                Case 2      'por cliente
                    .Formulas(3) = "parte1Titulo = ''"
                    .Formulas(4) = "parte1Hab = ''"
                    .Formulas(5) = "parte2Titulo = 'Titular'"
                    .Formulas(6) = "parte2NomTitular = '" & obtengo_nombre_pasajero(cliCuenta) & "'"
                    .Formulas(7) = "parte1TipoOrigenConsulta = 'Por cliente'"
            End Select
            .Formulas(8) = "simboloMn = '" & gblSignoMonedaNacional & "'"
            .Formulas(9) = "simboloDolares = '" & gblSignoDolares & "'"
            .Formulas(10) = "parte5ValorCotizacion = '" & mFunObtengoUltimaCotizacion(1, 1, m_FechaSis) & "'"
        End With
        
        'genero listado
        frmMAIN.CrystalReport1.Action = 1
        'aviso de confirmación de impresión de gastos extras
        mSubMensaje 4, 134   'se imprimieron los gastos extras
        'inicializo fórmulas
        mSubInicializoFormulas 24
    Else
        'aviso de que no hay gastos para imprimir
        mSubMensaje 3, 9
    End If
End Sub

'******************************************************
'*
'*  Asistencia a usuario
'*
'******************************************************

Private Sub botImprimirTodo_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 208
End Sub

Private Sub botImprimir_GotFocus()
    If Me.ssTab1.Tab = 0 Then 'si estoy mostrando gastos extras
        'muestro aviso de impresión de gastos extras
        mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 21
    Else
        'muestro aviso de impresión de gastos alojamiento
        mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 207
    End If
End Sub

Private Sub botSalir_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 1, 1
End Sub

Private Sub botImprimir_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub botSalir_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

