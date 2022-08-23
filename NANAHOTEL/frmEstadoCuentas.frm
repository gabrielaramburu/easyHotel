VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmEstadoCuentas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Estado de cuentas"
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   6255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   11033
      _Version        =   327680
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "&Resumen"
      TabPicture(0)   =   "frmEstadoCuentas.frx":0000
      Tab(0).ControlCount=   1
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      TabCaption(1)   =   "&Dolares"
      TabPicture(1)   =   "frmEstadoCuentas.frx":001C
      Tab(1).ControlCount=   1
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1(0)"
      Tab(1).Control(0).Enabled=   0   'False
      TabCaption(2)   =   "Moneda &Nacional"
      TabPicture(2)   =   "frmEstadoCuentas.frx":0038
      Tab(2).ControlCount=   1
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1(1)"
      Tab(2).Control(0).Enabled=   0   'False
      TabCaption(3)   =   "&Ambas Monedas"
      TabPicture(3)   =   "frmEstadoCuentas.frx":0054
      Tab(3).ControlCount=   1
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame1(2)"
      Tab(3).Control(0).Enabled=   0   'False
      Begin VB.Frame Frame1 
         Caption         =   "D&etalle de movimientos en "
         Height          =   5775
         Index           =   0
         Left            =   -74880
         TabIndex        =   10
         Top             =   360
         Width           =   11415
         Begin VB.PictureBox Picture4 
            Height          =   5350
            Left            =   5850
            ScaleHeight     =   5295
            ScaleWidth      =   75
            TabIndex        =   27
            Top             =   250
            Width           =   135
         End
         Begin MSFlexGridLib.MSFlexGrid gmsSaldos 
            Height          =   5415
            Index           =   0
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   11175
            _ExtentX        =   19711
            _ExtentY        =   9551
            _Version        =   393216
            Cols            =   9
            FocusRect       =   2
            HighLight       =   0
            SelectionMode   =   1
            FormatString    =   "| Fecha        | Moneda     | Tipo Documento              | Número      || Debe           | Haber          | Saldo          "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "D&etalle de movimientos en"
         Height          =   5775
         Index           =   1
         Left            =   -74880
         TabIndex        =   8
         Top             =   360
         Width           =   11415
         Begin VB.PictureBox Picture3 
            Height          =   5350
            Left            =   5850
            ScaleHeight     =   5295
            ScaleWidth      =   75
            TabIndex        =   26
            Top             =   250
            Width           =   135
         End
         Begin MSFlexGridLib.MSFlexGrid gmsSaldos 
            Height          =   5415
            Index           =   1
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   11175
            _ExtentX        =   19711
            _ExtentY        =   9551
            _Version        =   393216
            Cols            =   9
            FocusRect       =   2
            HighLight       =   0
            SelectionMode   =   1
            FormatString    =   "| Fecha        | Moneda     | Tipo Documento              | Número      || Debe           | Haber          | Saldo          "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Saldos"
         Height          =   5775
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   11415
         Begin VB.Label lblSignoMN 
            AutoSize        =   -1  'True
            Caption         =   "lblSignoMN"
            Height          =   240
            Index           =   2
            Left            =   1560
            TabIndex        =   29
            Top             =   3120
            Width           =   330
         End
         Begin VB.Label lblValorCotizacion 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblValorCotizacion"
            Height          =   375
            Left            =   2160
            TabIndex        =   28
            Top             =   3060
            Width           =   855
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Cotización a "
            Height          =   240
            Left            =   240
            TabIndex        =   23
            Top             =   3120
            Width           =   1140
         End
         Begin VB.Label lblSignoDol 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "lblSignoDol"
            Height          =   240
            Index           =   1
            Left            =   4800
            TabIndex        =   22
            Top             =   1320
            Width           =   1050
         End
         Begin VB.Label lblSignoMN 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "lblSignoMN"
            Height          =   240
            Index           =   1
            Left            =   4800
            TabIndex        =   21
            Top             =   720
            Width           =   1050
         End
         Begin VB.Label lblSignoDol 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "lblSignoDol"
            Height          =   240
            Index           =   0
            Left            =   4800
            TabIndex        =   20
            Top             =   3600
            Width           =   1050
         End
         Begin VB.Label lblSignoMN 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "lblSignoMN"
            Height          =   240
            Index           =   0
            Left            =   4800
            TabIndex        =   19
            Top             =   3120
            Width           =   1050
         End
         Begin VB.Label lblTotDeudaDol 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblTotDeudaDol"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6000
            TabIndex        =   18
            Top             =   3533
            Width           =   1455
         End
         Begin VB.Label lblTotDeudaMn 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblTotDeudaMn"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6000
            TabIndex        =   17
            Top             =   3053
            Width           =   1455
         End
         Begin VB.Line Line1 
            X1              =   240
            X2              =   7560
            Y1              =   2880
            Y2              =   2880
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Total adeudado a la fecha"
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
            Left            =   240
            TabIndex        =   16
            Top             =   2520
            Width           =   2760
         End
         Begin VB.Label lblTotSaldoDol 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblTotSaldoDol"
            Height          =   375
            Left            =   6000
            TabIndex        =   15
            Top             =   1253
            Width           =   1455
         End
         Begin VB.Label lblTotSaldoMn 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblTotSaldoMn"
            Height          =   375
            Left            =   6000
            TabIndex        =   14
            Top             =   653
            Width           =   1455
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Saldo total de movimeintos en moneda nacional"
            Height          =   240
            Left            =   240
            TabIndex        =   13
            Top             =   1320
            Width           =   4305
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Saldo total de movimientos en  dólares"
            Height          =   240
            Left            =   240
            TabIndex        =   12
            Top             =   720
            Width           =   3480
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "D&etalle completo de movimientos"
         Height          =   5775
         Index           =   2
         Left            =   -74880
         TabIndex        =   5
         Top             =   360
         Width           =   11415
         Begin VB.PictureBox Picture2 
            Height          =   5350
            Left            =   4740
            ScaleHeight     =   5295
            ScaleWidth      =   75
            TabIndex        =   25
            Top             =   250
            Width           =   135
         End
         Begin VB.PictureBox Picture1 
            Height          =   5350
            Left            =   7750
            ScaleHeight     =   5295
            ScaleWidth      =   75
            TabIndex        =   24
            Top             =   250
            Width           =   135
         End
         Begin MSFlexGridLib.MSFlexGrid gmsSaldos 
            Height          =   5415
            Index           =   2
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   11175
            _ExtentX        =   19711
            _ExtentY        =   9551
            _Version        =   393216
            Cols            =   13
            Redraw          =   -1  'True
            FocusRect       =   2
            HighLight       =   0
            SelectionMode   =   1
            FormatString    =   $"frmEstadoCuentas.frx":0070
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
   End
   Begin Hotel_Nana.gaHOTELbarra gaHOTELbarra1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   7605
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   582
   End
   Begin Hotel_Nana.gaHOTELcli gaHOTELcli1 
      Height          =   735
      Left            =   120
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   1296
      BackColor       =   -2147483633
   End
   Begin VB.CommandButton botImprimir 
      Height          =   375
      Left            =   9240
      Picture         =   "frmEstadoCuentas.frx":0105
      Style           =   1  'Graphical
      TabIndex        =   0
      Tag             =   "Imprimir"
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton botSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   10560
      TabIndex        =   1
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Menu mnuFormulario 
      Caption         =   "&Formulario"
      Begin VB.Menu mnuFormularioAceptar 
         Caption         =   "Salir"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuDiv1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuImprimirConsulta 
         Caption         =   "Imprimir"
         Shortcut        =   ^I
      End
   End
   Begin VB.Menu mnuVer 
      Caption         =   "&Ver información de ..."
      Begin VB.Menu mnuVerResumen 
         Caption         =   "Resumen"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuVerDolares 
         Caption         =   "Detalle movimientos en dólares"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuVerMonedaNacional 
         Caption         =   "Detalle movimientos en moneda nacional"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuVerDetalleAmbasMonedas 
         Caption         =   "Detalle de movimientos en ambas monedas"
         Shortcut        =   {F8}
      End
   End
   Begin VB.Menu mnuOrdenar 
      Caption         =   "&Ordenado por..."
      Visible         =   0   'False
      Begin VB.Menu mnuOrdenarPorFecha 
         Caption         =   "Fecha"
         Checked         =   -1  'True
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu mnuOrdenarPorMoneda 
         Caption         =   "Por moneda"
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu mnuOrdenadoPorTipoDocu 
         Caption         =   "Por tipo de documento"
         Shortcut        =   ^{F3}
      End
      Begin VB.Menu mnuOrdenadoPorNroDocu 
         Caption         =   "Por número de documento"
         Shortcut        =   ^{F4}
      End
   End
End
Attribute VB_Name = "frmEstadoCuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private nrocli As Long

Private totSaldoMn As Double    'almaceno el saldo final de movimientos en moneda nacional
Private totSaldoDol As Double   'almaceno el saldo final de movimientos en dólares
Private mostrarGrillasAmbos As Byte 'determino si muetro grilla de ambos movimientos,esta
                                    'grilla solo se muestra cuando tengo movimientos en
                                    'ambas monedas (mostrarGrillasAmbos = 2)
                                    
'NOTA: los procedimientos de ordenación de fecha no se usan más porque no se mantiene
'la coherencia con el cálculo de saldos.

Private Sub Form_Load()
    'Apariencia formulario
    mSubConfiguroFuentesControlesSistema Me
    
    nrocli = Val(frmIngPaxEmp.txtCodCli.Text)
    Me.gaHOTELcli1.CaminoBaseDeDatos = vardir
    Me.gaHOTELcli1.CodigoCliente = nrocli
    
    'bloque todos los frame del formulario
    subBloqueoFrame
    Frame2.Enabled = True
    
    'realizo cabezal grilla ambas monedas
    subCabezalGrillaAmbasMonedas
    'inicializo variables publicas de módulo
    mostrarGrillasAmbos = 0
    'proceso movimientos en moneda nacional
    subObtengoSaldos nrocli, 0
    'proceso movimientos en pesos
    subObtengoSaldos nrocli, 1
    'mostrar resumen
    subMuestroResumen
    'muestro o no, la ficha de ambos movimientos
    If mostrarGrillasAmbos <> 2 Then
        'no muestro tab ya que no tengo movimientos en más de una moneda.
        subNoPermitoTab 3
    Else
        'coloreo columna de saldos en grilla de ambos movs.
        subColoreoColumnaSaldos Me.gmsSaldos(2), 2
    End If
End Sub

Private Sub subMuestroResumen()
    '-------------------------------------------------------------------------
    'Inicializo ficha de resumen.
    '-------------------------------------------------------------------------
    Dim valorCot As Double
    Dim aux As Double
    
    'inicializo etiquetas de signo de monedas
    Me.lblSignoDol(0) = mFunObtengoSignoMoneda(1)
    Me.lblSignoDol(1) = mFunObtengoSignoMoneda(1)
    Me.lblSignoMN(0) = mFunObtengoSignoMoneda(0)
    Me.lblSignoMN(1) = mFunObtengoSignoMoneda(0)
    Me.lblSignoMN(2) = mFunObtengoSignoMoneda(0)
    'muestro saldos en cada cuenta
    lblTotSaldoMn = Format(totSaldoMn, "####0.00;;0")
    lblTotSaldoDol.Caption = Format(totSaldoDol, "####0.00;;0")
    'muestro valor cotización
    valorCot = mFunObtengoUltimaCotizacion(1, 1, m_FechaSis)
    Me.lblValorCotizacion = valorCot
    'calculo deuda total convertida a dólares
    aux = totSaldoMn + (totSaldoDol * valorCot)
    lblTotDeudaMn.Caption = Format(aux, "####0.00;;0")
    aux = totSaldoDol + (totSaldoMn / valorCot)
    lblTotDeudaDol.Caption = Format(aux, "####0.00;;0")
    
    lblTotDeudaDol.FontBold = True
    lblTotDeudaMn.FontBold = True
End Sub

Private Sub subBloqueoFrame()
    '------------------------------------------------------------------------
    'No permito trabajar con ningún frame del formulario. Mejora la interface.
    '-------------------------------------------------------------------------
    Frame2.Enabled = False
    Frame1(0).Enabled = False
    Frame1(1).Enabled = False
    Frame1(2).Enabled = False
End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)
    subBloqueoFrame
    Select Case ssTab1.Tab
        Case 0  'ficha resumen
            'no permito aplicar criterios de ordenación
            Me.mnuOrdenar.Enabled = False
            Frame2.Enabled = True
        Case 1  'ficha dolares
            'si permito aplicar criterios de ordenación
            Me.mnuOrdenar.Enabled = True
            Frame1(0).Enabled = True
            
            Frame1(0).Refresh
        Case 2  'ficha moneda nacional
            'si permito aplicar criterios de ordenación
            Me.mnuOrdenar.Enabled = True
            Frame1(1).Enabled = True
            
            Frame1(1).Refresh
        Case 3  'ficha ambos
            'si permito aplicar criterios de ordenación
            Me.mnuOrdenar.Enabled = True
            Frame1(2).Enabled = True
            'ordeno grilla por fecha+moneda+tipodocu+nrodocu
            gmsSaldos(2).col = 1
            gmsSaldos(2).ColSel = 4
            gmsSaldos(2).Sort = flexSortGenericAscending
    
            Frame1(2).Refresh
    End Select
    
End Sub

Private Sub subObtengoSaldos(nrocliente As Long, tipoMovs As Byte)
    '----------------------------------------------------------------------------------
    'Obtengo todos los movimientos en dólares y moneda nacional para un cliente determinado.
    'Los muestro en las grillas correspondientes.
    '-----------------------------------------------------------------------------------
    'Parámetros.
    '   Entrada [nroCliente] número de cliente que estoy procesando.
    '           [tipoMovs]  tipo de movimiento.
    '                       0 = mov. de moneda nacional
    '                       1 = mov. de dólares.
    '
    '-----------------------------------------------------------------------------------
    Dim rstMovs As Recordset
    Dim qdfMovs As QueryDef
    
    Dim indGrilla As Integer
    Dim nroTabs As Byte
    
    'determino que grilla debo de cargar y con que tabs estoy trabajando
    If tipoMovs = 0 Then
        indGrilla = 1
        nroTabs = 2
    End If
    
    If tipoMovs = 1 Then
        indGrilla = 0
        nroTabs = 1
    End If
    
    'Ejecuto consulta
    Set qdfMovs = bdHOTEL.CreateQueryDef("")
    qdfMovs.SQL = funConsultaMovimientos(nrocliente, tipoMovs)
    Set rstMovs = qdfMovs.OpenRecordset(dbOpenSnapshot)
    'verifico si existen movimientos
    If rstMovs.RecordCount > 0 Then
        'determino si muestro grilla de ambos movimientos
        mostrarGrillasAmbos = mostrarGrillasAmbos + 1
        'realizo cabezal grilla unica
        subCabezalGrillaUnicaMoneda gmsSaldos(indGrilla)
        'cargo registros procesados, a grilla de movimientos únicos
        subCargoGrillaUnica gmsSaldos(indGrilla), rstMovs
        'coloreo columna de saldos
        subColoreoColumnaSaldos Me.gmsSaldos(indGrilla), indGrilla
    Else
        'si no existen movimientos no permito trabajar con tab correspondiente.
        subNoPermitoTab nroTabs
    End If
    Set qdfMovs = Nothing
    Set rstMovs = Nothing
End Sub

Private Sub subColoreoColumnaSaldos(grilla As MSFlexGrid, tipoGrilla As Integer)
    '-------------------------------------------------------------------------------
    'Coloreo del color predeterminado la columan de saldos, en cada grilla.
    '-------------------------------------------------------------------------------
    'Parámetros.
    '   Entrada:    [grilla]        grilla que estoy trabajando
    '               [tipoGrilla]    tipo de grilla con la que estoy trabajando
    '                               0 = grilla dólares
    '                               1 = grilla moneda nacional
    '                               2 = grilla ambas monedas
    '--------------------------------------------------------------------------------
    If tipoGrilla = 0 Then
        marco_celdas_grilla grilla, 8, 8, 2, grilla.Rows - 1
        grilla.CellBackColor = mSisColor_7SaldoDolares
    End If
    If tipoGrilla = 1 Then
        marco_celdas_grilla grilla, 8, 8, 2, grilla.Rows - 1
        grilla.CellBackColor = mSisColor_6SaldoMonedaNacional
    End If
    If tipoGrilla = 2 Then
        marco_celdas_grilla grilla, 8, 8, 2, grilla.Rows - 1
        grilla.CellBackColor = mSisColor_7SaldoDolares
        marco_celdas_grilla grilla, 12, 12, 2, grilla.Rows - 1
        grilla.CellBackColor = mSisColor_6SaldoMonedaNacional
    End If
End Sub

Private Sub subCargoGrillaUnica(grilla As MSFlexGrid, movs As Recordset)
    '------------------------------------------------------------------------
    'Recorro recordset de movimientos y cargo en grilla correspondiente.
    '------------------------------------------------------------------------
    'Parámetros.
    '   Entrada [grilla]    grilla a cargar
    '
    '           [movs]      recordSet de movimientos
    '-------------------------------------------------------------------------
    movs.MoveFirst
    Do While Not movs.EOF
        'calculo saldos
        subCalculoSaldos movs("debe"), movs("haber"), movs("moneda")
        'cargo grilla de moneda única
        subRealizoLineaUnicaMoneda grilla, movs
        'cargo grilla de ambos monedas
        subRealizoLineaAmbasMonedas movs
        movs.MoveNext
    Loop
End Sub

Private Sub subCalculoSaldos(debe As Double, haber As Double, tipoMovs As Byte)
    '-------------------------------------------------------------------------------
    'Calcula el saldo total de movimientos, ya sea en dólares o pesos. Esta misma
    'cifra se incluye a medida que se van creando las líneas en la grilla, como saldos
    'parciales.
    '---------------------------------------------------------------------------------
    'Parámetros.
    '   Entrada [debe]  importe perteneciente al debe
    '           [haber] importe perteneciente al haber
    '           [tipoMovs]  tipo de movimiento.
    '                       0 = mov. de moneda nacional
    '                       1 = mov. de dólares.
    '-----------------------------------------------------------------------------------
    
    'determino tipo de movimiento
    If tipoMovs = 0 Then totSaldoMn = totSaldoMn + (debe - haber)
    If tipoMovs = 1 Then totSaldoDol = totSaldoDol + (debe - haber)
End Sub

Private Sub subRealizoLineaUnicaMoneda(grilla As MSFlexGrid, movs As Recordset)
    '------------------------------------------------------------------------------
    'Creo línea en grilla de movimientos únicos, tanto para moneda nacional como
    'para dólares.
    '------------------------------------------------------------------------------
    'Parámetros.
    '   Entrada [grilla]    grilla que estoy cargando
    '           [movs]      datos que voy a cargar en la línea
    '------------------------------------------------------------------------------
    Dim saldoParcial As Double
    'obtengo saldo parcial, calculado al momento de realizar la línea.
    If movs("moneda") = 0 Then saldoParcial = totSaldoMn
    If movs("moneda") = 1 Then saldoParcial = totSaldoDol
    
    Dim linea As String
    linea = _
    Chr(9) & _
    movs("fecha") & _
    Chr(9) & _
    funObtengoMoneda(movs("moneda")) & _
    Chr(9) & _
    funObtengoTipoDoc(movs("tipodoc")) & _
    Chr(9) & _
    movs("nrodoc") & _
    Chr(9) & _
    Chr(9) & _
    Format(movs("debe"), "####0.00;;#") & _
    Chr(9) & _
    Format(movs("haber"), "####0.00;;#") & _
    Chr(9) & _
    Format(saldoParcial, "####0.00;;#")
    
    grilla.AddItem linea
End Sub

Private Sub subRealizoLineaAmbasMonedas(movs As Recordset)
    Dim linea As String
    
    Dim debeMN As Double
    Dim haberMN As Double
    Dim saldoMN As Double
    
    Dim debeDOL As Double
    Dim haberDOL As Double
    Dim saldoDOL As Double
    
    'Determino el tipo de saldo a cargar
    If movs("moneda") = 0 Then  'monda nacional
        debeMN = movs("debe")
        haberMN = movs("haber")
        saldoMN = totSaldoMn
    Else
        debeDOL = movs("debe")
        haberDOL = movs("haber")
        saldoDOL = totSaldoDol
    End If
    
    linea = Chr(9) & _
    movs("fecha") & _
    Chr(9) & _
    funObtengoMoneda(movs("moneda")) & _
    Chr(9) & _
    funObtengoTipoDoc(movs("tipodoc")) & _
    Chr(9) & _
    movs("nrodoc") & _
    Chr(9) & _
    Chr(9) & _
    Format(debeMN, "####0.00;;#") & _
    Chr(9) & _
    Format(haberMN, "####0.00;;#") & _
    Chr(9) & _
    Format(saldoMN, "####0.00;;#") & _
    Chr(9) & _
    Chr(9) & _
    Format(debeDOL, "####0.00;;#") & _
    Chr(9) & _
    Format(haberDOL, "####0.00;;#") & _
    Chr(9) & _
    Format(saldoDOL, "####0.00;;#")

    gmsSaldos(2).AddItem linea
End Sub

Private Function funConsultaMovimientos(nrocli As Long, tipoMovs As Byte)
    '--------------------------------------------------------------------------------
    'Devuelve todos los movimientos pertenecientes a un cliente determinado, en una
    'moneda determinada.
    '--------------------------------------------------------------------------------
    'Parámetros.
    '   Entrada [nroCli]    cliente que estoy procesando
    '           [tipoMovs]  tipo de movimientos que tengo que obtener.
    '                       0 = moneda nacional
    '                       1 = dólares.
    '
    '   Salida  si [tipoMovs] = 0 movs. en moneda nacional de un cliente determinado.
    '           si [tipoMovs] = 1 movs. en dólares de un cliente determinado.
    '--------------------------------------------------------------------------------
    Dim consulta As String
    consulta = _
    "Select fecha, " & _
            "tipodoc, " & _
            "nrodoc, " & _
            "debe, " & _
            "haber," & _
            "moneda " & _
    "From estado_cuentas " & _
    " where nrocli = " & Str(nrocli) & _
    " and moneda = " & tipoMovs & _
    " order by fecha,moneda,tipodoc,nrodoc "
    funConsultaMovimientos = consulta
End Function

Private Sub subCabezalGrillaUnicaMoneda(grilla As MSFlexGrid)
    '---------------------------------------------------------------
    'Creo cabezal de grilla de única moneda moneda.
    '---------------------------------------------------------------
    'Parámetros.
    '   Entrada    [grilla] grilla a la que quiero relizar cabezal
    '----------------------------------------------------------------
    grilla.FormatString = _
    " | Fecha        | " & _
    "Moneda     | " & _
    "Tipo Documento              | " & _
    "Número      |" & _
    "| " & _
    "Debe           | " & _
    "Haber          | " & _
    "Saldo          "
End Sub

Private Sub subCabezalGrillaAmbasMonedas()
    '---------------------------------------------------
    'Creo cabezal de grilla de ambas monedas
    '---------------------------------------------------
    gmsSaldos(2).FormatString = _
        " | Fecha       | " & _
        "Moneda     | " & _
        "Tipo Doc.        | " & _
        "Nro.         |" & _
        " |" & _
        "Debe       |" & _
        "Haber      |" & _
        "Saldo         |" & _
        " | " & _
        "Debe       |" & _
        "Haber      |" & _
        "Saldo         "
End Sub

Private Function funObtengoTipoDoc(tipo As Byte)
    '--------------------------------------------------------------------------
    'Devuelve la descripción del tipo de documento mostrado.
    '---------------------------------------------------------------------------
    'Parámetros.
    '   Entrada:    [tipo] tipo de documento
    '   Salida      string correspondienyte a la descripción del tipo de docu.
    '----------------------------------------------------------------------------
    Dim signoMn As String
    Dim signoDol As String
    signoMn = mFunObtengoSignoMoneda(0)
    signoDol = mFunObtengoSignoMoneda(1)
    Select Case tipo
        Case 3  'factura moneda nacional
            funObtengoTipoDoc = "Factura " & signoMn
        Case 4  'factura dólares
            funObtengoTipoDoc = "Factura " & signoDol
        Case 7  'dev. crédito moneda nacional
            funObtengoTipoDoc = "Dev. Crédito " & signoMn
        Case 8  'dev. crédito dólares
            funObtengoTipoDoc = "Dev. Crédito " & signoDol
        Case 9  'recivo moneda nacional
            funObtengoTipoDoc = "Recibo " & signoMn
        Case 10 'recivo dólares
            funObtengoTipoDoc = "Recibo " & signoDol
    End Select
End Function

Private Function funObtengoMoneda(moneda As Byte) As String
    'Obtengo la moneda del documento
    Select Case moneda
        Case 0  'moneda nacional
            funObtengoMoneda = mFunObtengoSignoMoneda(0)
        Case 1  'dólares
            funObtengoMoneda = mFunObtengoSignoMoneda(1)
    End Select
End Function

Private Sub subNoPermitoTab(nroTab As Byte)
    '------------------------------------------------------------------------------
    'No permito trabajar con el tab que se pasa como parámetro, ya que no existen
    'datos para mostrar.
    'Tampoco permito opción desde el menú "Ver información de ..."
    '-------------------------------------------------------------------------------
    'Parámetros.
    '   Entrada [nroTab]    tab con el cual estoy trabajando.
    '                       0 = Resumen (siempre se muestra)
    '                       1 = Dolares
    '                       2 = Moneda nacional
    '                       3 = Ambas
    '---------------------------------------------------------------------------------
    Me.ssTab1.TabEnabled(nroTab) = False
    Select Case nroTab
        Case 1
            Me.mnuVerDolares.Enabled = False
        Case 2
            Me.mnuVerMonedaNacional.Enabled = False
        Case 3
            Me.mnuVerDetalleAmbasMonedas.Enabled = False
    End Select
End Sub

Private Sub botSalir_Click()
    Unload Me
    frmIngPaxEmp.Show 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Descargo formulario de memoria
    Set frmEstadoCuentas = Nothing
End Sub

'****************************************************
'* Impresión de estados de cuenta
'****************************************************
Private Sub botImprimir_Click()
    'Imprimo información
    If mfunAplicoConfImp(2, 10) = 1 Then
        'verifico que tipo de filtro estoy aplicando
        Select Case Me.ssTab1.Tab
            Case 0  'imprimo resumen
                subImprimoReporte 0, nrocli
            Case 1  'imprimo detalle dólares
                subImprimoReporte 1, nrocli
            Case 2  'imprimo detalle moneda nacional
                subImprimoReporte 2, nrocli
            Case 3  'imprimo ambas monedas
                subImprimoReporte 3, nrocli
        End Select
    End If
End Sub

Private Sub subImprimoReporte(tipoReporte As Byte, nrocli As Long)
    '----------------------------------------------------------------------------------
    'Configuro el control Crystal genérico,obtiene datos y emite el listado
    'correspondiente al tab seleccionado.
    '-------------------------------------------------------------------------------
    'Parámetros.
    '   Entrada:    [tipoReporte] determina que tipo de información muestro.
    '               0 = resumen
    '               1 = detalle movimientos dólares
    '               2 = detalle movimientos moneda nacional
    '               3 = ambos movimientos
    '               [nroCli]    cliente que voy a procesar
    '-------------------------------------------------------------------------------
    
    
    Dim nomReporte As String
    Dim tituloReporte As String
    Dim nroMensajeFin As Integer
    Dim consulta As String
    Dim totResumen As String
    
    'inicializo control data
     subInicializoControlData frmMAIN.Data1CrystalReport
      
    Select Case tipoReporte
        Case 0
            nroMensajeFin = 146
            nomReporte = "rptest3.rpt"
            'realizo cualquier consulta, lo importante es obtener datos del cliente.
            'realizando la consulta con tipoMovs = 3 me aseguro de siempre encontrar
            'registros.
            consulta = funConsultaMovimientosImp(nrocli, 0)
            tituloReporte = "Estado de cuenta: resumen saldos"
        
        Case 1
            nroMensajeFin = 143
            nomReporte = "rptest1.rpt"
            tituloReporte = "Estado de cuenta: dólares"
            totResumen = Me.lblTotSaldoDol.Caption
            consulta = funConsultaMovimientosImp(nrocli, 1)
            
        Case 2
            nroMensajeFin = 144
            nomReporte = "rptest1.rpt"
            tituloReporte = "Estado de cuenta: moneda nacional"
            totResumen = Me.lblTotSaldoMn.Caption
            consulta = funConsultaMovimientosImp(nrocli, 0)
            
        Case 3
            nroMensajeFin = 145
            nomReporte = "rptest2.rpt"
            tituloReporte = "Estado de cuenta: ambas monedas"
            consulta = funConsultaMovimientosImp(nrocli, 3)
    End Select
        
    frmMAIN.Data1CrystalReport.RecordSource = consulta
    'ejecuto consulta control data
    frmMAIN.Data1CrystalReport.Refresh
    
    'como de aquí en más voy a trabajar directamente con el control data
    'me aseguro de que se allan encontrado reservas
    If frmMAIN.Data1CrystalReport.Recordset.RecordCount > 0 Then
        'determino el nombre del reporte a utilizar
        frmMAIN.CrystalReport1.ReportFileName = vardir2 + nomReporte
        
        'inicializo fórmulas
        With frmMAIN.CrystalReport1
            .Formulas(0) = "cabVersion = '" & mFunImpVersionAplicacion & "'"        'nombre y versión aplicación
            .Formulas(1) = "cabRegistro = '" & mFunImpRegistro & "'"                'hotel propietario de la aplicación
            .Formulas(2) = "cabHoraImp = ' " & CStr(Time()) & "'"                   'hora actual
            .Formulas(3) = "cabTitulo = '" & tituloReporte & "'"
            .Formulas(4) = "signoMn = '" & gblSignoMonedaNacional & "'"
            .Formulas(5) = "signoDol = '" & gblSignoDolares & "'"
            'muestro resuen
            If tipoReporte = 1 Or tipoReporte = 2 Then
                .Formulas(6) = "ResumenTotal =' " & totResumen & "'"
            Else
                .Formulas(6) = "totalSaldoMn = '" & Me.lblTotSaldoMn & "'"
                .Formulas(7) = "totalSaldoDol = '" & Me.lblTotSaldoDol & "'"
                .Formulas(8) = "totalDeudaMn = '" & Me.lblTotDeudaMn & "'"
                .Formulas(9) = "totalDeudaDol = '" & Me.lblTotDeudaDol & "'"
                .Formulas(10) = "valorCotizacion = '" & mFunObtengoUltimaCotizacion(1, 1, m_FechaSis) & "'"
            End If
        End With
        'genero listado
        frmMAIN.CrystalReport1.Action = 1
        'aviso de confirmación de impresión de reporte
        mSubMensaje 4, nroMensajeFin  'se imprimieron los egresos previstos
        'inicializo fórmulas
        mSubInicializoFormulas 9
        'inicializo campos de ordenación del informe
        'mSubInicializoCamposOrden 1
    Else
        'aviso de que no hay datos para imprimir
        mSubMensaje 3, 9
    End If
End Sub

Private Function funConsultaMovimientosImp(nrocli As Long, tipoMovs As Byte)
                                            
    '--------------------------------------------------------------------------------
    'Devuelve todos los movimientos pertenecientes a un cliente determinado, en una
    'moneda determinada, para impresión.
    'Basicamente la diferencia entre este procedimiento y el utilizado para
    'mostrar los datos por pantalla, es la cantidad de campos que se manejan.
    '--------------------------------------------------------------------------------
    'Parámetros.
    '   Entrada [nroCli]    cliente que estoy procesando
    '           [tipoMovs]  tipo de movimientos que tengo que obtener.
    '                       0 = moneda nacional
    '                       1 = dólares.
    '                       3 = no filtro por moneda (impresión de ambas cuentas)
    '
    '   Salida  si [tipoMovs] = 0 movs. en moneda nacional de un cliente determinado.
    '           si [tipoMovs] = 1 movs. en dólares de un cliente determinado.
    '--------------------------------------------------------------------------------
    Dim consulta As String
    Dim filtroMoneda As String
    
    'determino si filtro por moneda
    If tipoMovs <> 3 Then
        filtroMoneda = " and moneda = " & tipoMovs
    End If
    consulta = _
    "Select * " & _
    " from estado_cuentas,clientes " & _
    " where estado_cuentas.nrocli = clientes.nrocorr and " & _
    " nrocli = " & Str(nrocli) & _
    filtroMoneda & _
    " order by fecha,moneda,tipodoc,nrodoc "
    funConsultaMovimientosImp = consulta
End Function

'****************************************************
'* Menú flotante
'****************************************************
Private Sub mnuVerResumen_Click()
    Me.ssTab1.Tab = 0
End Sub

Private Sub mnuVerDolares_Click()
    Me.ssTab1.Tab = 1
End Sub

Private Sub mnuVerMonedaNacional_Click()
    Me.ssTab1.Tab = 2
End Sub

Private Sub mnuVerDetalleAmbasMonedas_Click()
    Me.ssTab1.Tab = 3
End Sub

Private Sub mnuFormularioAceptar_Click()
    'Equivale a preionar la tecla F12 o el boton de aceptar
    botSalir_Click
End Sub

Private Sub mnuImprimirConsulta_Click()
    'Equivale a presionar el boton de imprimir o la tecla Ctrol+I
    botImprimir_Click
End Sub

Private Sub mnuOrdenarPorFecha_Click()
    'Ordena por fecha
    subMarcoOpcionesMenu
    mnuOrdenarPorFecha.Checked = True
    subOrdenoGrilla Me.gmsSaldos(funIndGrillaAOrdenar), 1
End Sub

Private Sub mnuOrdenarPorMoneda_Click()
    'Ordena por moneda
    subMarcoOpcionesMenu
    mnuOrdenarPorMoneda.Checked = True
    subOrdenoGrilla Me.gmsSaldos(funIndGrillaAOrdenar), 2
End Sub

Private Sub mnuOrdenadoPorNroDocu_Click()
    'Ordena por número de documento
    subMarcoOpcionesMenu
    mnuOrdenadoPorNroDocu.Checked = True
    subOrdenoGrilla Me.gmsSaldos(funIndGrillaAOrdenar), 4
End Sub

Private Sub mnuOrdenadoPorTipoDocu_Click()
    'Ordena por tipo de documento
    subMarcoOpcionesMenu
    mnuOrdenadoPorTipoDocu.Checked = True
    subOrdenoGrilla Me.gmsSaldos(funIndGrillaAOrdenar), 3
End Sub

Private Function funIndGrillaAOrdenar() As Integer
    '-------------------------------------------------------------------------------------
    'Determino el índice de la matriz de controles grilla, que estoy mostrando actualmente.
    'El valor devuelto se utiliza para aplicar criterio de ordenación a la grilla visible.
    '-------------------------------------------------------------------------------------

    funIndGrillaAOrdenar = Me.ssTab1.Tab - 1
End Function

Private Sub subMarcoOpcionesMenu()
    'Inicializa la propiedad cheked de las opciones de menú a false.
    Me.mnuOrdenadoPorNroDocu.Checked = False
    Me.mnuOrdenadoPorTipoDocu.Checked = False
    Me.mnuOrdenarPorFecha.Checked = False
    Me.mnuOrdenarPorMoneda.Checked = False
End Sub

Private Sub subOrdenoGrilla(grilla As MSFlexGrid, criterio As Long)
    '---------------------------------------------------------------------
    'Ordena la grilla según un criterio.
    '---------------------------------------------------------------------
    'Parámetros.
    '   Entrada [criterio]      1 pordena por fecha
    '                           2  por moneda
    '                           3  por tipo de documento
    '                           4  por número de documento
    '---------------------------------------------------------------------
    'establese la columna por la cual se ordena
    grilla.col = criterio
    'asume que quiero ordenar todas las filas no fijas de la grilla
    grilla.ColSel = grilla.col
    grilla.Sort = flexSortGenericAscending
    'muestro ícono en grilla
    mSubMuestroIcono grilla, criterio
    grilla.SetFocus
End Sub

'************************************************************
'*
'*  Asistencia a usuarios
'*
'************************************************************

Private Sub botImprimir_GotFocus()
    Select Case Me.ssTab1.Tab
        Case 0  'imprimo resumen
            mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 217
        Case 1  'imprimo dólares
            mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 31
        Case 2  'imprimo moneda nacional
            mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 215
        Case 3  'imprimo ambas
            mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 216
    End Select
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

