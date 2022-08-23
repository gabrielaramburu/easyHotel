VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmDevolucion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Devolución"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   11520
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   11520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton botSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5880
      TabIndex        =   30
      Top             =   6360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame5 
      Caption         =   "Cabezal"
      Height          =   2175
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   11295
      Begin VB.TextBox txtNroCli 
         Height          =   285
         Index           =   0
         Left            =   4200
         TabIndex        =   14
         Text            =   "0"
         Top             =   360
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtRuc 
         Alignment       =   1  'Right Justify
         Height          =   375
         Index           =   0
         Left            =   8280
         TabIndex        =   13
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox txtCP 
         Alignment       =   1  'Right Justify
         Height          =   375
         Index           =   0
         Left            =   8280
         TabIndex        =   12
         Top             =   1200
         Width           =   1215
      End
      Begin VB.ComboBox cboPais 
         Height          =   360
         Index           =   0
         Left            =   5520
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox txtLoc 
         Height          =   375
         Index           =   0
         Left            =   2040
         TabIndex        =   10
         Top             =   1680
         Width           =   2055
      End
      Begin VB.TextBox txtDir 
         Height          =   375
         Index           =   0
         Left            =   2040
         TabIndex        =   9
         Top             =   1200
         Width           =   5055
      End
      Begin VB.TextBox fechaemi 
         Height          =   375
         Index           =   0
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtNom 
         Height          =   375
         Index           =   0
         Left            =   2040
         TabIndex        =   7
         Top             =   720
         Width           =   5055
      End
      Begin VB.Label Label9 
         Caption         =   "R.U.C"
         Height          =   255
         Index           =   0
         Left            =   7560
         TabIndex        =   21
         Top             =   780
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "C.P."
         Height          =   255
         Index           =   0
         Left            =   7560
         TabIndex        =   20
         Top             =   1260
         Width           =   375
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Pais"
         Height          =   240
         Index           =   0
         Left            =   4320
         TabIndex        =   19
         Top             =   1740
         Width           =   405
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Localidad"
         Height          =   240
         Index           =   0
         Left            =   240
         TabIndex        =   18
         Top             =   1740
         Width           =   900
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Dirección"
         Height          =   240
         Index           =   0
         Left            =   240
         TabIndex        =   17
         Top             =   1260
         Width           =   855
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nombre completo"
         Height          =   240
         Index           =   0
         Left            =   240
         TabIndex        =   16
         Top             =   780
         Width           =   1620
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         Height          =   240
         Index           =   0
         Left            =   240
         TabIndex        =   15
         Top             =   300
         Width           =   570
      End
   End
   Begin VB.CommandButton botImprimir 
      Height          =   375
      Index           =   0
      Left            =   8760
      Picture         =   "frmDevolucion.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Tag             =   "Imprimir"
      Top             =   6360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin Hotel_Nana.gaHOTELbarra gaHOTELbarra1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   6840
      Width           =   11520
      _ExtentX        =   20320
      _ExtentY        =   582
   End
   Begin VB.CommandButton botAnular 
      Caption         =   "&Anular"
      Height          =   375
      Left            =   7320
      TabIndex        =   3
      Top             =   6360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton botCancelar 
      Height          =   375
      Index           =   0
      Left            =   10200
      Picture         =   "frmDevolucion.frx":0942
      Style           =   1  'Graphical
      TabIndex        =   2
      Tag             =   "Cancelar"
      Top             =   6360
      Width           =   1215
   End
   Begin Hotel_Nana.gaHOTELcli gaHOTELcli1 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   1296
      BackColor       =   -2147483633
   End
   Begin MSFlexGridLib.MSFlexGrid dbgrid1 
      Height          =   2490
      Index           =   0
      Left            =   120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3120
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   4392
      _Version        =   393216
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      ScrollBars      =   2
   End
   Begin MSFlexGridLib.MSFlexGrid msfgTotales 
      Height          =   615
      Index           =   0
      Left            =   4680
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   5640
      Width           =   6750
      _ExtentX        =   11906
      _ExtentY        =   1085
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      ScrollBars      =   0
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
   Begin VB.Label lblSignoMon 
      Caption         =   "lblSignoMon(0)"
      Height          =   255
      Index           =   0
      Left            =   2520
      TabIndex        =   31
      Top             =   6480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblNroDocu 
      Caption         =   "0"
      Height          =   375
      Index           =   0
      Left            =   3960
      TabIndex        =   29
      Top             =   6000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblImpMinimo 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   27
      Top             =   6120
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label lblImpBasico 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Index           =   0
      Left            =   720
      TabIndex        =   26
      Top             =   6120
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label lblImpExento 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Index           =   0
      Left            =   1320
      TabIndex        =   25
      Top             =   6120
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label lblIVAm 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Index           =   0
      Left            =   1920
      TabIndex        =   24
      Top             =   6120
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label lblTotalGral 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Index           =   0
      Left            =   3000
      TabIndex        =   23
      Top             =   6120
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label lblIVAb 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Index           =   0
      Left            =   2520
      TabIndex        =   22
      Top             =   6120
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Menu mnuFormulario 
      Caption         =   "&Formulario"
      Begin VB.Menu mnuFormularioImprimir 
         Caption         =   "Imprimir"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuFormularioCancelar 
         Caption         =   "Cancelar"
      End
      Begin VB.Menu mnuFormularioAnular 
         Caption         =   "Anular          F12"
      End
      Begin VB.Menu mnuFormularioCancelarAnulacion 
         Caption         =   "Cancelar"
      End
      Begin VB.Menu mnuFormularioSalir 
         Caption         =   "Salir          F12"
      End
   End
End
Attribute VB_Name = "frmDevolucion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_Nro_Dev As Long
Private m_Tipo_Dev As Byte
Private m_Nro_Docu As Long
Private m_Tipo_Docu As Byte

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'Intercepto la tecla F12
    If KeyCode = vbKeyF12 Then
        Form_KeyPress (KeyCode)
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyF12 Then
        If Me.botAnular.Visible = True Then
            'anulo la devolución
            'botAnular_Click
        Else
            If Me.botSalir.Visible = True Then
                'cierro el formulario
                botSalir_Click
            End If
        End If
    End If
    
    If KeyAscii = vbKeyEscape Then
        Unload Me
        frmTipoDocumento.Show 1
    End If
End Sub

Private Sub Form_Load()
    'Apariencia formulario
    mSubConfiguroFuentesControlesSistema Me
    
    'configuro menú de opciones
    subConfiguroMenuOpciones tipo_accion_devo
    
    'configuro cabezal de grilla
    subCabezalGrilla Me.DBGrid1(0)
    subOcultoColumnasNoVisibles
    
    Select Case tipo_accion_devo
        Case 1  'nueva devolución
            obtengo_datos_docu
            inicializo_form_para_nueva_dev
            '
            mSub_muestro_lineas_documento m_Tipo_Docu, m_Nro_Docu, frmDevolucion
            'muestro documento al cual pertenece la devolución
            Me.Frame5(0).Caption = frmTipoDocumento.LstTipoDoc.Text & " " & frmTipoDocumento.txtNroDoc
            
        Case 3  'consulto
            obtengo_datos_dev
            Me.lblNroDocu(0).Caption = m_Nro_Dev
            inicializo_form
            mSub_muestro_lineas_documento m_Tipo_Dev, m_Nro_Dev, frmDevolucion
            'muestro documento al cual pertenece la devolución
            Me.Frame5(0).Caption = frmTipoDocumento.LstTipoDoc.Text & " " & frmTipoDocumento.txtNroDoc & " Factura: " & tbCABEZAL("nro_fact_docu")
    
    End Select
    botones
    
    'Muestro grilla de totales
    mSubMuestro_Totales msfgTotales(0), _
        Me.lblImpExento(0).Caption, _
        Me.lblImpMinimo(0).Caption, _
        Me.lblIVAm(0).Caption, _
        Me.lblImpBasico(0).Caption, _
        Me.lblIVAb(0).Caption, _
        Me.lblTotalGral(0).Caption
End Sub

Private Sub subOcultoColumnasNoVisibles()
    DBGrid1(0).ColWidth(10) = 0 'habitación del gasto
    DBGrid1(0).ColWidth(11) = 0 'nrocorr del gasto
    DBGrid1(0).ColWidth(12) = 0 'tipo
End Sub

Private Sub subConfiguroMenuOpciones(tipo_accion As Byte)
    'Determino que opción semuestra enel menú de opciones
    Select Case tipo_accion
        Case 1  'nueva devolución
            Me.mnuFormularioAnular.Visible = False
            Me.mnuFormularioCancelarAnulacion.Visible = False
            Me.mnuFormularioSalir.Visible = False
        Case 2  'anulación
            Me.mnuFormularioSalir.Visible = False
            Me.mnuFormularioImprimir.Visible = False
            Me.mnuFormularioCancelar.Visible = False
        Case 3  'consulta
            Me.mnuFormularioAnular.Visible = False
            Me.mnuFormularioCancelarAnulacion.Visible = False
            Me.mnuFormularioImprimir.Visible = False
            Me.mnuFormularioCancelar.Visible = False
    End Select
End Sub

Private Sub obtengo_datos_docu()
    m_Nro_Docu = frmTipoDocumento.txtNroDoc.Text
    m_Tipo_Docu = frmTipoDocumento.LstTipoDoc.ItemData(frmTipoDocumento.LstTipoDoc.ListIndex)
    m_Tipo_Dev = obtengo_tipo_dev
End Sub

Private Sub obtengo_datos_dev()
    m_Nro_Dev = frmTipoDocumento.txtNroDoc.Text
    m_Tipo_Dev = frmTipoDocumento.LstTipoDoc.ItemData(frmTipoDocumento.LstTipoDoc.ListIndex)
End Sub

Private Sub inicializo_form()
    'Inicializo formulario para Consulta o Anulación
    
    mSub_cambio_cabezal False, 0, Me    'No permito modificar cabezal.
    carga_tipo_pais frmDevolucion.cboPais(0)
    'Muestro cabezal devolución.
    mSub_cargo_cabezal_desde_documento m_Tipo_Dev, m_Nro_Dev, frmDevolucion, 0
    
    'Muestro datos cliente en cabezal formulario.
    Me.gaHOTELcli1.CaminoBaseDeDatos = vardir
    Me.gaHOTELcli1.CodigoCliente = txtNroCli(0).Text
End Sub

Private Sub inicializo_form_para_nueva_dev()
    'Inicializo formulario para nueva devolución
    mSub_cambio_cabezal False, 0, Me    'No permito modificar cabezal.
    carga_tipo_pais frmDevolucion.cboPais(0)
    
    'Muestro cabezal devolución desde factura.
    mSub_cargo_cabezal_desde_documento m_Tipo_Docu, m_Nro_Docu, frmDevolucion, 0
    
    'Muestro datos cliente en cabezal formulario.
    Me.gaHOTELcli1.CaminoBaseDeDatos = vardir
    Me.gaHOTELcli1.CodigoCliente = txtNroCli(0).Text
    
End Sub

Private Sub botones()
    Select Case tipo_accion_devo
        Case 1  'nueva
            Me.botImprimir(0).Visible = True
            Me.botSalir.Visible = False
        Case 2  'anulo
            'muestro boton de anulación y saco de impresión y de salir
            botImprimir(0).Visible = False
            Me.botSalir.Visible = False
            botAnular.Visible = True
            'posiciono boton de anular
            botAnular.Top = 6360
            botAnular.Left = 8400
            
        Case 3  'consulto
            'oculto boton anular e imprimir y cancelar
            Me.botImprimir(0).Visible = False
            Me.botAnular.Visible = False
            Me.botCancelar(0).Visible = False
            'muestro boton de salir
            Me.botSalir.Visible = True
            Me.botSalir.Top = 6360
            Me.botSalir.Left = 10200
            
    End Select
End Sub

'Private Sub botAnular_Click()
'    'Esto no se usa. gabriel!!!
'    'Elimino gastos devueltos
'
'
'    'Elimino documento anulado
'    mSub_Elimino_Documento m_Tipo_Dev, m_Nro_Dev
'
'    'Elimino documento del estado de cuentas, solo si es crédito.
'    mSub_Elimino_Documento_EstadoCuenta m_Tipo_Documento, m_Nro_Documento
'
'    'cuando termino muestro mensaje y me voy
'    MsgBox "Anulación de devolución salio bien (de puro culo)", vbExclamation
'    Unload Me
'    frmTipoDocumento.Show 1
'End Sub

Private Sub botImprimir_Click(Index As Integer)
    'aviso confirmación al usuario
    'aviso de confirmación de impresión
    If mfunAplicoConfImp(1, 8) = 1 Then
        Me.lblNroDocu(0).Caption = mFun_obtengo_proximo_documento(m_Tipo_Dev)
        'Grabo lineas de la devolución
        mFun_realizo_lineas m_Tipo_Dev, Me, 0
        'Grabo cabezal de la devolución
        mSub_grabo_cabezal_documento m_Tipo_Dev, _
                                    Val(Me.lblNroDocu(0).Caption), _
                                    m_FechaSis, _
                                    Me.txtNom(0).Text, _
                                    Me.txtDir(0).Text, _
                                    Me.txtLoc(0).Text, _
                                    Me.txtRuc(0).Text, _
                                    Me.txtCP(0).Text, _
                                    Me.cboPais(0).ItemData(cboPais(Index).ListIndex), _
                                    Me.txtNroCli(0).Text, _
                                    Me.lblTotalGral(0).Caption, _
                                    m_Nro_Docu, _
                                    0, _
                                    lblIVAb(0).Caption, _
                                    lblIVAm(0).Caption, _
                                    lblImpExento(0).Caption, _
                                    lblImpBasico(0).Caption, _
                                    lblImpMinimo(0).Caption, 0, 0, 0, "", "", "", ""
                
    
        'Creo nuevamente los gastos en la cuenta del titular de la devolución
        'Este procedimiento toma los datos desde el archivo de documentos, por
        'ese motivo es necesario que se ejecute después de haber creado la devolución
        
        mSub_creo_gastos_nuevamente m_Tipo_Dev, Val(Me.lblNroDocu(0).Caption), Me.txtNroCli(0).Text
        
        'actualizo estado de cuentas
        mSub_grabo_estado_cuentas _
                            m_Tipo_Dev, _
                            Val(lblNroDocu(0).Caption), _
                            Val(txtNroCli(0).Text), _
                            Val(lblTotalGral(0).Caption), _
                            m_FechaSis
        mSubArmoReporteFactura m_Tipo_Dev, Val(lblNroDocu(0).Caption)
        'aviso de confirmación de operación realizada correctamente
        mSubMensaje 4, 51, CStr(Me.lblNroDocu(0).Caption)
        
        'Grabo bitacora
        GraboBitacora "Docu. " & lblNroDocu(0).Caption
        Unload Me
        frmTipoDocumento.Show 1
    End If
End Sub

Private Function obtengo_tipo_dev()
    'Dependiendo del tipo de documento que seleccione, dependerá el tipo de
    'de devolución que voy a realizar.
    Select Case frmTipoDocumento.LstTipoDoc.ItemData(frmTipoDocumento.LstTipoDoc.ListIndex) 'tipo documento
        Case 1
            obtengo_tipo_dev = 5
        Case 2
            obtengo_tipo_dev = 6
        Case 3
            obtengo_tipo_dev = 7
        Case 4
            obtengo_tipo_dev = 8
    End Select
End Function

Private Sub botCancelar_Click(Index As Integer)
    Unload Me
    frmTipoDocumento.Show 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Descargo formulario de memoria
    Set frmDevolucion = Nothing
End Sub

Private Sub mnuFormularioAnular_Click()
    'Equivale a presionar el boton de anular
    'botAnular_Click
End Sub

Private Sub mnuFormularioCancelar_Click()
    'Equivale a presionar elbotón de cancelar
    botCancelar_Click 0
End Sub

Private Sub mnuFormularioCancelarAnulacion_Click()
    'Equivale a presionar el boton de cancelar
    botCancelar_Click 0
End Sub

Private Sub mnuFormularioImprimir_Click()
    'Equivale a presionar el botón de imprimir
    botImprimir_Click 0
End Sub

Private Sub mnuFormularioSalir_Click()
    'Equivale a presionar el botón de salir
    botSalir_Click
End Sub

Private Sub botSalir_Click()
    Unload Me
    frmTipoDocumento.Show 1
End Sub

'*****************************************************
'*
'*  Asistencia al usuario
'*
'*****************************************************

Private Sub botSalir_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 1, 1
End Sub

Private Sub botImprimir_GotFocus(Index As Integer)
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 131
End Sub

Private Sub botCancelar_GotFocus(Index As Integer)
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 1, 3
End Sub

Private Sub botCancelar_LostFocus(Index As Integer)
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub botImprimir_LostFocus(Index As Integer)
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub botSalir_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub


