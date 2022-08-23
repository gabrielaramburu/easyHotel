VERSION 5.00
Object = "{EB8C3860-8603-11D0-A35D-7062E4000000}#1.1#0"; "VCBND.OCX"
Begin VB.Form frmRecivo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Recibos"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Hotel_Nana.gaHOTELbarra gaHOTELbarra1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   5355
      Width           =   9510
      _ExtentX        =   16775
      _ExtentY        =   582
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del recivo"
      Height          =   5295
      Left            =   120
      TabIndex        =   14
      Top             =   0
      Width           =   9255
      Begin VB.TextBox txtObsRecivo 
         Height          =   975
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   3600
         Width           =   7575
      End
      Begin VB.CheckBox chkAgenciaEmpresa 
         Caption         =   "&Agencia/Empresa"
         Height          =   255
         Left            =   5280
         TabIndex        =   8
         Top             =   2280
         Width           =   2055
      End
      Begin VB.CommandButton botSalir 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   3960
         TabIndex        =   20
         Top             =   4800
         Width           =   1215
      End
      Begin VB.TextBox txtNroRecivo 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtNroCli 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   1680
         MaxLength       =   5
         TabIndex        =   7
         Top             =   2640
         Width           =   1335
      End
      Begin VB.TextBox txtNomCli 
         Height          =   375
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   2640
         Width           =   4095
      End
      Begin VB.TextBox txtImporte 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   1680
         TabIndex        =   5
         Top             =   2040
         Width           =   1335
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
         Left            =   7320
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   2640
         Width           =   495
      End
      Begin VB.CommandButton botCancelar 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   7920
         Picture         =   "frmRecivo.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         Tag             =   "Cancelar"
         Top             =   4800
         Width           =   1215
      End
      Begin VB.CommandButton botConfirmar 
         Height          =   375
         Left            =   6600
         Picture         =   "frmRecivo.frx":08C2
         Style           =   1  'Graphical
         TabIndex        =   12
         Tag             =   "Aceptar"
         Top             =   4800
         Width           =   1215
      End
      Begin VB.CommandButton botBorrar 
         Caption         =   "&Borrar"
         Height          =   375
         Left            =   5280
         TabIndex        =   11
         Top             =   4800
         Width           =   1215
      End
      Begin VcBndCtl.VcCalCombo fFechaRecivo 
         Height          =   375
         Left            =   1680
         TabIndex        =   3
         Top             =   1440
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _0              =   $"frmRecivo.frx":1178
         _1              =   $"frmRecivo.frx":1581
         _2              =   $"frmRecivo.frx":198A
         _3              =   ")@f-@@@E@@@A@@@@@@@@@'@@@E@@@A@@@@@@@@@)h@@@E@@@A@@@@@@@@@,467D"
         _count          =   4
         _ver            =   2
      End
      Begin VB.Image Image1 
         Height          =   105
         Left            =   240
         Picture         =   "frmRecivo.frx":1D93
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   8850
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "&Concepto"
         Height          =   240
         Left            =   240
         TabIndex        =   9
         Top             =   3240
         Width           =   870
      End
      Begin VB.Label lblSignoMoneda 
         Alignment       =   1  'Right Justify
         Caption         =   "sig"
         Height          =   240
         Left            =   1080
         TabIndex        =   21
         Top             =   2100
         Width           =   525
      End
      Begin VB.Label Label3 
         Caption         =   "Moneda:"
         Height          =   255
         Left            =   4920
         TabIndex        =   18
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "&Número recivo"
         Height          =   255
         Left            =   240
         TabIndex        =   0
         Top             =   540
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "F&echa"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   1500
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "&Importe"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   2100
         Width           =   735
      End
      Begin VB.Label lblMoneda 
         Caption         =   "lblMoneda"
         Height          =   375
         Left            =   5880
         TabIndex        =   17
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "&Realizado a:"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   2700
         Width           =   1215
      End
   End
   Begin VB.Menu mnuFormulario 
      Caption         =   "&Formulario"
      Begin VB.Menu mnuFormularioBorrar 
         Caption         =   "Borrar          F12"
      End
      Begin VB.Menu mnuFormularioAceptar 
         Caption         =   "Aceptar          F12"
      End
      Begin VB.Menu mnuFormularioCancelar 
         Caption         =   "Cancelar"
      End
      Begin VB.Menu mnuFormularioSalir 
         Caption         =   "Salir          F12"
      End
   End
   Begin VB.Menu mnuBuscar 
      Caption         =   "Buscar..."
      Begin VB.Menu mnuBuscarClientes 
         Caption         =   "Clientes..."
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmRecivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private nro_recivo As Long
Private tipo_recivo As Byte
Private moneda_recivo As Byte

Private Sub Form_Load()
    'Apariencia formulario
    mSubConfiguroFuentesControlesSistema Me
    
    Me.txtNroRecivo.BackColor = mSisColor_18ControlesNoHabilitados
    Me.txtNomCli.BackColor = mSisColor_18ControlesNoHabilitados
    
    'configuro menú de opciones
    subConfiguroMenuOpciones tipo_accion_recivo
    
    Select Case tipo_accion_recivo
        Case 1  'ingreso automático de recivo
            'obtengo número próximo recivo a realizar
            
            nro_recivo = obtengo_prox_recivo
            tipo_recivo = 1 'automático
            determino_moneda
        
        Case 2  'ingreso recivo manual
            nro_recivo = Val(frmTipoDocumento.txtNroDoc.Text)
            tipo_recivo = 2
            determino_moneda
            
        Case 3  'consulto recivo automatico
            nro_recivo = Val(frmTipoDocumento.txtNroDoc.Text)
            tipo_recivo = 1
            muestro_datos_recivo
            desmarco_controles
            
        Case 4  'consulto recivo manual
            nro_recivo = Val(frmTipoDocumento.txtNroDoc.Text)
            tipo_recivo = 2
            muestro_datos_recivo
            desmarco_controles
            
        Case 5  'borro registro automatico
            nro_recivo = Val(frmTipoDocumento.txtNroDoc.Text)
            tipo_recivo = 1
            muestro_datos_recivo
            desmarco_controles
                
        Case 6  'borro registro manual
            nro_recivo = Val(frmTipoDocumento.txtNroDoc.Text)
            tipo_recivo = 2
            muestro_datos_recivo
            desmarco_controles
    End Select
    titulo_formulario
    inicializo_formulario
End Sub

Private Sub botBorrar_Click()
    Dim tipoRecivo As Byte
    If tbRECIVOS("moneda_recivo") = 0 Then
        tipoRecivo = 9 'm/n
    Else
        tipoRecivo = 10    'dol
    End If
    
    'aviso de confirmación de borrado de recivo
    If mFunMensaje(4, 31) Then
        'borro estado cuentas
        tbESTADO_CUENTAS.Index = "pi_estado_cuentas"
        tbESTADO_CUENTAS.Seek "=", tipoRecivo, nro_recivo
        If Not tbESTADO_CUENTAS.NoMatch Then    'si existe
            tbESTADO_CUENTAS.Delete
        End If
        
        'borro recivo
        tbRECIVOS.Delete
                
        'aviso de confirmación de borrado
        mSubMensaje 4, 34
        
        'grabo bitacora
        GraboBitacora "Res. " & nro_recivo
        Unload Me
        frmTipoDocumento.Show 1
    End If
End Sub

Private Sub subConfiguroMenuOpciones(tipo_accion As Byte)
    'Configuro las opciones del menú dependiendo de la tarea que realiza el formulario
    Select Case tipo_accion
        Case 1 To 2 'ingreso recivo automático y manual
            Me.mnuFormularioBorrar.Visible = False
            Me.mnuFormularioSalir.Visible = False
            
        Case 3 To 4 'consultar recivo automático y manual
            Me.mnuBuscar.Visible = False
            Me.mnuFormularioAceptar.Visible = False
            Me.mnuFormularioBorrar.Visible = False
            Me.mnuFormularioCancelar.Visible = False
            
        Case 5 To 6 'borrar recivo automático y manual
            Me.mnuBuscar.Visible = False
            Me.mnuFormularioAceptar.Visible = False
            Me.mnuFormularioSalir.Visible = False
    End Select
End Sub

Private Sub titulo_formulario()
    Select Case tipo_accion_recivo
        Case 1
            Me.Caption = "Nuevo recibo"
        Case 2
            Me.Caption = "Nuevo recibo manual"
        Case 3
            Me.Caption = "Consulto recibo automático"
        Case 4
            Me.Caption = "Consulto recibo manual"
        Case 5
            Me.Caption = "Borro recibo automático"
        Case 6
            Me.Caption = "Borro recibo manual"
    End Select
End Sub

Private Sub inicializo_formulario()
    Select Case tipo_accion_recivo
        Case 1  'nuevo recivo automático
            txtNroRecivo.Text = nro_recivo
            fFechaRecivo.Value = m_FechaSis
            botBorrar.Visible = False
            botSalir.Visible = False
            
        Case 2  'nuevo recivo manual
            txtNroRecivo.Text = nro_recivo
            fFechaRecivo.Value = m_FechaSis
            botBorrar.Visible = False
            botSalir.Visible = False
            
        Case 3  'consulto recivo automático
            botConfirmar.Visible = False
            botBorrar.Visible = False
            botCancelar.Visible = False
            botSalir.Left = 7920
            
        Case 4  'consulto recivo manual
            botConfirmar.Visible = False
            botBorrar.Visible = False
            botCancelar.Visible = False
            botSalir.Left = 7920
            
        Case 5  'elimino recivo automático
            botConfirmar.Visible = False
            botSalir.Visible = False
            botBorrar.Left = 6600
            
        Case 6  'elimino recivo manual
            botConfirmar.Visible = False
            botSalir.Visible = False
            botBorrar.Left = 6600
    End Select
End Sub

Private Sub determino_moneda()
   'Determino la moneda del nuevo recivo
    Select Case frmTipoDocumento.LstTipoDoc.ItemData(frmTipoDocumento.LstTipoDoc.ListIndex)
        Case 9  'M/N
            Me.lblMoneda.Caption = gblSignoMonedaNacional
            Me.lblSignoMoneda.Caption = gblSignoMonedaNacional
            moneda_recivo = 0
        Case 10  'Dol
            Me.lblMoneda.Caption = gblSignoDolares
            Me.lblSignoMoneda.Caption = gblSignoDolares
            moneda_recivo = 1
    End Select
End Sub

Private Sub muestro_datos_recivo()
    If busco_recivoTF(tipo_recivo, nro_recivo) Then
        fFechaRecivo.Text = tbRECIVOS("fecha_recivo")
        txtImporte.Text = tbRECIVOS("importe_recivo")
        txtNroCli.Text = tbRECIVOS("cliente_recivo")
        txtNomCli.Text = tbRECIVOS("nomcli_recivo")
        txtNroRecivo = tbRECIVOS("nro_recivo")
        chkAgenciaEmpresa.Value = tbRECIVOS("tipoCliente_recivo")
        If Not IsNull(tbRECIVOS("obs_recivo")) Then txtObsRecivo = tbRECIVOS("obs_recivo")
        If tbRECIVOS("moneda_recivo") = 0 Then  'm/n
            lblMoneda.Caption = gblSignoMonedaNacional
            lblSignoMoneda.Caption = gblSignoMonedaNacional
        Else                                    'dol
            lblMoneda.Caption = gblSignoDolares
            lblSignoMoneda.Caption = gblSignoDolares
        End If
    End If
End Sub

Private Sub chkAgenciaEmpresa_Click()
    'Inicializo controles de datos de cliente
    Me.txtNomCli.Text = Empty
    Me.txtNroCli.Text = Empty
End Sub

Private Sub BotAyuda_Click()
    Dim nro_corr_aux As String
    If Me.chkAgenciaEmpresa.Value = 1 Then
        'el cliente es una empresa
        nro_corr_aux = mFunBusqueda(3)
        If Val(nro_corr_aux) <> 0 Then
            txtNroCli.Text = nro_corr_aux
            cargo_nombre_emp
        End If
    Else
        'el cliente es un pasajeros
        nro_corr_aux = mFunBusqueda(1)
        If Val(nro_corr_aux) <> 0 Then
            txtNroCli.Text = nro_corr_aux
            cargo_nombre_cli
        End If
    End If
    
End Sub

Private Sub botConfirmar_Click()
    'aviso de confirmación de impresión de recivo
    If funValidoDatos Then
        Select Case tipo_accion_recivo
            Case 1  'nuevo recivo automático
                genero_proximo_nro_recivo
                actualizo_estado_cuentas
                grabo_nuevo_recivo
                imprimo_recivo
                'grabobitacora
                GraboBitacora "Res. " & nro_recivo
                Unload Me
                frmTipoDocumento.Show 1
    
            Case 2  'nuevo recivo manual
                If mFunMensaje(4, 147) Then
                    actualizo_estado_cuentas
                    grabo_nuevo_recivo
                    'grabobitacora
                    GraboBitacora "Res. " & nro_recivo
                    Unload Me
                    frmTipoDocumento.Show 1
                End If
        End Select
    End If
End Sub

Private Function funValidoDatos() As Boolean
    '----------------------------------------------------------------------
    'Detemino si los datos ingresados para realizar el recivo son correctos
    'y estan completos.
    '----------------------------------------------------------------------
    'Parámetros.
    '   Salida  True = los datos estan bien y completos.
    '           False = falta algún dato o los valores ingresados
    '                   no son correctos.
    '-----------------------------------------------------------------------
    
    'verifico que se ingrese fecha
    If IsDate(Me.fFechaRecivo.Text) Then
        'verifico importe
        If Val(Me.txtImporte.Text) > 0 Then
            'verifico se ingrese y exista pasajero
            If mfunExisteCliente(Me.chkAgenciaEmpresa.Value, Val(Me.txtNroCli)) Then
                funValidoDatos = True
            Else
                funValidoDatos = False
                'aviso de cliente inexistente
                mSubMensaje 4, 149
                txtNroCli.SetFocus
                Exit Function
            End If
        Else
            funValidoDatos = False
            'aviso de importe incorrecto
            mSubMensaje 4, 148
            txtImporte.SetFocus
            Exit Function
        End If
    Else
        funValidoDatos = False
        'aviso de formato de fecha incorrecto
        mSubMensaje 3, 1
        fFechaRecivo.SetFocus
        Exit Function
    End If
End Function
Private Sub grabo_nuevo_recivo()
    tbRECIVOS.AddNew
        tbRECIVOS("tipo_recivo") = tipo_recivo
        tbRECIVOS("nro_recivo") = nro_recivo
        tbRECIVOS("fecha_recivo") = fFechaRecivo.Value
        tbRECIVOS("cliente_recivo") = Val(txtNroCli.Text)
        tbRECIVOS("nomcli_recivo") = txtNomCli.Text
        tbRECIVOS("importe_recivo") = Val(txtImporte.Text)
        tbRECIVOS("moneda_recivo") = moneda_recivo
        tbRECIVOS("tipoCliente_recivo") = chkAgenciaEmpresa.Value
        If Not IsNull(Me.txtObsRecivo) Then tbRECIVOS("obs_recivo") = Me.txtObsRecivo.Text
    tbRECIVOS.Update
End Sub

Private Sub genero_proximo_nro_recivo()
    If tbPARAMETROS("prox_recivo") = nro_recivo Then
        tbPARAMETROS.Edit
        tbPARAMETROS("prox_recivo") = nro_recivo + 1
        tbPARAMETROS.Update
    Else
        nro_recivo = tbPARAMETROS("prox_recivo")
        tbPARAMETROS.Edit
        tbPARAMETROS("prox_recivo") = nro_recivo + 1
        tbPARAMETROS.Update
        'Si se cambia el número de recivo aviso al usuario.
        'atención se cambió el número de recivo a " & (nuevo recivo)
        mSubMensaje 4, 32, CStr(nro_recivo)
    End If
End Sub

Private Sub actualizo_estado_cuentas()
    tbESTADO_CUENTAS.AddNew
        tbESTADO_CUENTAS("tipodoc") = obtengo_tipo_recivo
        tbESTADO_CUENTAS("nrodoc") = nro_recivo
        tbESTADO_CUENTAS("nrocli") = Val(txtNroCli.Text)
        tbESTADO_CUENTAS("debe") = 0
        tbESTADO_CUENTAS("haber") = Val(txtImporte.Text)
        tbESTADO_CUENTAS("fecha") = fFechaRecivo.Value
        tbESTADO_CUENTAS("moneda") = moneda_recivo
    tbESTADO_CUENTAS.Update
End Sub

Private Sub imprimo_recivo()
    'confirmación de impresión de recivo
    If mfunAplicoConfImp(1, 9) = 1 Then
        subArmoImpresionRecivo tipo_recivo, nro_recivo
    End If
End Sub

Private Sub subArmoImpresionRecivo(tipoRecivo As Byte, nroRecivo As Long)
    '----------------------------------------------------------------------------------
    'Configuro el control Crystal genérico,obtiene datos y emite el recivo
    'automático ingresado.
    '-------------------------------------------------------------------------------
    'Parámetros.
    '   Entrada:    [tipoRecivo] tipo del recivo que voy a imprimir.
    '                   1 = recivo automático
    '                   2 = recivo manual
    '               [nroRecivo] número de recivo que voy a imprimir.
    '-------------------------------------------------------------------------------
    
    Dim consulta As String
        
    'inicializo control data
     subInicializoControlData frmMAIN.Data1CrystalReport
    
    'NOTA: esta consulta debe de ser la misma que se utilizo para crear el reporte,
    'por ese motivo no debe de ser modificada en aspectos tales como campos a mostrar y
    'tablas utilizadas.
    
    consulta = "select * from recivos " & _
    " where recivos.tipo_recivo = " & tipoRecivo & _
    " and recivos.nro_recivo = " & nroRecivo
    
    frmMAIN.Data1CrystalReport.RecordSource = consulta
    'ejecuto consulta control data
    frmMAIN.Data1CrystalReport.Refresh
    
    'como de aquí en más voy a trabajar directamente con el control data
    'me aseguro de que se allan encontrado reservas
    If frmMAIN.Data1CrystalReport.Recordset.RecordCount > 0 Then
        'determino el nombre del reporte a utilizar
        frmMAIN.CrystalReport1.ReportFileName = vardir2 + "rptrecibo.rpt"
        
        'inicializo fórmulas
        With frmMAIN.CrystalReport1
            .Formulas(0) = "cabVersion = '" & mFunImpVersionAplicacion & "'"        'nombre y versión aplicación
            .Formulas(1) = "cabRegistro = '" & mFunImpRegistro & "'"                'hotel propietario de la aplicación
            .Formulas(2) = "cabHoraImp = ' " & CStr(Time()) & "'"                   'hora actual
            .Formulas(3) = "cabTitulo = '" & "tituloReporte" & "'"
            .Formulas(4) = "signoMn = '" & gblSignoMonedaNacional & "'"
            .Formulas(5) = "signoDol = '" & gblSignoDolares & "'"
            .Formulas(6) = "descImporte = '" & Numlet(Me.txtImporte) & "'"
            .Formulas(7) = "cabTitulo = ' " & "Recivo " & "'"
        End With
        'genero listado
        frmMAIN.CrystalReport1.Action = 1
        'aviso de confirmación de impresión de reporte
        mSubMensaje 4, 33, CStr(nroRecivo)  'se imprimió el recivo
        'inicializo fórmulas
        mSubInicializoFormulas 6
    Else
        'aviso de que no hay datos para imprimir
        mSubMensaje 3, 9
    End If
End Sub

Private Function obtengo_tipo_recivo()
    If moneda_recivo = 0 Then   'm/n
        obtengo_tipo_recivo = 9
    Else                        'dol
        obtengo_tipo_recivo = 10
    End If
End Function

Private Function obtengo_prox_recivo()
    obtengo_prox_recivo = tbPARAMETROS("prox_recivo")
End Function

Private Function cargo_nombre_cli()
    If busco_clienteTF(Val(txtNroCli.Text)) Then
        txtNomCli.Text = tbCLIENTES("nombre_completo_titular")
    Else
        txtNroCli.Text = ""
        txtNomCli.Text = ""
    End If
End Function

Private Function cargo_nombre_emp()
    If busco_empTF(Val(txtNroCli.Text)) Then
        txtNomCli.Text = tbEMPRESAS("nomEmp")
    Else
        txtNroCli.Text = ""
        txtNomCli.Text = ""
    End If
End Function

Private Sub Form_Unload(Cancel As Integer)
    'Descargo formulario de memoria
    Set frmRecivo = Nothing
End Sub

Private Sub txtImporte_KeyPress(KeyAscii As Integer)
    CapturoEnter KeyAscii
    ValidoNum KeyAscii, False, True
End Sub

Private Sub txtNroCli_KeyPress(KeyAscii As Integer)
    CapturoEnter KeyAscii
    ValidoNum KeyAscii, False, False
End Sub

Private Sub txtNroCli_LostFocus()
    Dim nroEmp As Long
    If Me.chkAgenciaEmpresa.Value = 1 Then
        'cliente es una agencia empresa
        cargo_nombre_emp
    Else
        'cliente es un pax
        cargo_nombre_cli
    End If
    
    'asistencia a usuarios
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub botCancelar_Click()
    Unload Me
    frmTipoDocumento.Show 1
End Sub

Private Sub desmarco_controles()
    mSub_bloqueo_controles_formulario Me, True
    fFechaRecivo.Enabled = False
    botAyuda.Enabled = False
    chkAgenciaEmpresa.Enabled = False
End Sub

Private Sub mnuFormularioAceptar_Click()
    'Equivale a presionar el boton aceptar o la tecla F12
    botConfirmar_Click
End Sub

Private Sub mnuFormularioCancelar_Click()
    'Equivale a presionar el boton cancelar
    botCancelar_Click
End Sub

Private Sub mnuFormularioBorrar_Click()
    'Equivale a presionar el boton de borrar
    botBorrar_Click
End Sub

Private Sub mnuFormularioSalir_Click()
    'Equivale a presionar el boton de salir
    botSalir_Click
End Sub

Private Sub mnuBuscarClientes_Click()
    'Equivale a presionar el botón de ayuda
    BotAyuda_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 Then
        If botSalir.Visible = True Then
            'cierro el formulario
            botSalir_Click
        Else
            If botBorrar.Visible = True Then
                'borro el recivo
                botBorrar_Click
            Else
                If botConfirmar.Visible = True Then
                    'creo un nuevo recivo
                    botConfirmar_Click
                End If
            End If
        End If
    End If
    
    If KeyCode = vbKeyEscape Then
        botSalir_Click
    End If
End Sub

Private Sub botSalir_Click()
    Unload Me
    frmTipoDocumento.Show 1
End Sub

'**************************************************
'*
'*  Asistencia de usuarios
'*
'**************************************************

Private Sub txtNroRecivo_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 114
End Sub

Private Sub fFechaRecivo_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 115
End Sub

Private Sub txtImporte_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 116
End Sub

Private Sub txtNroCli_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 117
End Sub

Private Sub botSalir_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 1, 1
End Sub

Private Sub botBorrar_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 119
End Sub

Private Sub botConfirmar_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 118
End Sub

Private Sub botCancelar_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 1, 3
End Sub

Private Sub chkAgenciaEmpresa_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 219
End Sub

Private Sub txtObsRecivo_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 218
End Sub

Private Sub botCancelar_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub botConfirmar_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub botBorrar_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub botSalir_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub fFechaRecivo_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub txtImporte_LostFocus()
    'obtengo solamente los dos valores decimales, para que la función de conversión
    'de números a letras funcione correctamnete, ya que la misma, tiene problemas
    'para convertir números con más de dos cifras deciamles.
    txtImporte.Text = Format(txtImporte.Text, "#########0.00;;#")
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub txtNroRecivo_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub txtObsRecivo_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub chkAgenciaEmpresa_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

