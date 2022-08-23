VERSION 5.00
Object = "{EB8C3860-8603-11D0-A35D-7062E4000000}#1.1#0"; "VCBND.OCX"
Begin VB.Form frmCargaAlojaManual 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Alojamiento Manual"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   9510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Hotel_Nana.gaHOTELtitular gaHOTELtitular1 
      Height          =   1335
      Left            =   120
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   0
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   2355
      BackColor       =   -2147483633
   End
   Begin VB.Frame Frame1 
      Caption         =   "Fecha del alojamiento a modificar"
      Height          =   1215
      Left            =   120
      TabIndex        =   11
      Top             =   1440
      Width           =   9255
      Begin VB.CommandButton botContinuar 
         Caption         =   "&Procesar"
         Height          =   375
         Left            =   2880
         TabIndex        =   2
         Tag             =   "Procesar"
         Top             =   480
         Width           =   1215
      End
      Begin VcBndCtl.VcCalCombo fecha 
         Height          =   375
         Left            =   1200
         TabIndex        =   1
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _0              =   $"frmCargaAlojaManual.frx":0000
         _1              =   $"frmCargaAlojaManual.frx":0409
         _2              =   $"frmCargaAlojaManual.frx":0812
         _3              =   "-@@@B@@@@@%@@@C@@@@@@@D@@@A@@@)f@@@E@@@A@@@@@@@@@'@@@E@@@A@@@@@@@@@)h@@@E@@@A@@@@@@@@@,1E1B"
         _count          =   4
         _ver            =   2
      End
      Begin VB.Label Label2 
         Caption         =   "F&echa:"
         Height          =   255
         Left            =   240
         TabIndex        =   0
         Top             =   540
         Width           =   615
      End
   End
   Begin Hotel_Nana.gaHOTELbarra gaHOTELbarra1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   6495
      Width           =   9510
      _ExtentX        =   16775
      _ExtentY        =   582
   End
   Begin VB.Frame Frame2 
      Caption         =   "Nuevo importe del alojamiento"
      Height          =   3615
      Left            =   120
      TabIndex        =   12
      Top             =   2760
      Width           =   9255
      Begin VB.TextBox txtObsMotivoIngreso 
         Height          =   375
         Left            =   5040
         MaxLength       =   25
         TabIndex        =   8
         Top             =   1320
         Width           =   3495
      End
      Begin VB.ComboBox cboMotivoIngreso 
         Height          =   360
         Left            =   5040
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   600
         Width           =   3495
      End
      Begin VB.CommandButton botCancelar 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   7920
         Picture         =   "frmCargaAlojaManual.frx":0C1B
         Style           =   1  'Graphical
         TabIndex        =   10
         Tag             =   "Cancelar"
         Top             =   3120
         Width           =   1215
      End
      Begin VB.CommandButton botConfirmar 
         Enabled         =   0   'False
         Height          =   375
         Left            =   6600
         Picture         =   "frmCargaAlojaManual.frx":14DD
         Style           =   1  'Graphical
         TabIndex        =   9
         Tag             =   "Aceptar"
         Top             =   3120
         Width           =   1215
      End
      Begin VB.TextBox txtImporteNue 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   2040
         Width           =   1455
      End
      Begin VB.TextBox txtTarifaNueva 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   2880
         MaxLength       =   9
         TabIndex        =   4
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox txtTarifaAnt 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "&Observaciones"
         Height          =   255
         Left            =   5040
         TabIndex        =   7
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "&Motivo de ingreso "
         Height          =   255
         Left            =   5040
         TabIndex        =   5
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label lblSignoDol 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "lblSignoDol"
         Height          =   240
         Index           =   2
         Left            =   1680
         TabIndex        =   22
         Top             =   2107
         Width           =   1050
      End
      Begin VB.Label lblSignoDol 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "lblSignoDol"
         Height          =   240
         Index           =   1
         Left            =   1680
         TabIndex        =   21
         Top             =   1380
         Width           =   1050
      End
      Begin VB.Label lblSignoDol 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "lblSignoDol"
         Height          =   240
         Index           =   0
         Left            =   1680
         TabIndex        =   20
         Top             =   660
         Width           =   1050
      End
      Begin VB.Label LblAlojamiento 
         AutoSize        =   -1  'True
         Caption         =   "LblAlojamiento"
         Height          =   240
         Left            =   5040
         TabIndex        =   18
         Top             =   2107
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label labTarifaAnt 
         AutoSize        =   -1  'True
         Caption         =   "Importe actual"
         Height          =   240
         Left            =   240
         TabIndex        =   17
         Top             =   660
         Width           =   1260
      End
      Begin VB.Label Label1 
         Caption         =   "Nuevo importe del alojamiento"
         Height          =   495
         Left            =   240
         TabIndex        =   14
         Top             =   2100
         Width           =   2415
      End
      Begin VB.Label labTarifaNue 
         AutoSize        =   -1  'True
         Caption         =   "&Importe a modificar"
         Height          =   240
         Left            =   240
         TabIndex        =   3
         Top             =   1380
         Width           =   1710
      End
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
      Begin VB.Menu mnuProcesar 
         Caption         =   "Procesar"
         Shortcut        =   {F9}
      End
   End
End
Attribute VB_Name = "frmCargaAlojaManual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private hab_cuenta As Long
Private tarifa As Double
Private importe_nue As Double

Private Sub Form_Load()
    'Apariencia formulario
    mSubConfiguroFuentesControlesSistema Me
    'Inicializo etiquetas de signo moneda
    Me.lblSignoDol(0).Caption = gblSignoDolares
    Me.lblSignoDol(1).Caption = gblSignoDolares
    Me.lblSignoDol(2).Caption = gblSignoDolares
    'cargo combo motivo de alojamiento manual
    mSubCargoComboMotivoAlojaManual Me.cboMotivoIngreso
    
    'obtengo habitacion
    hab_cuenta = Val(frmIngHabitacion.txtNroHab.Text)
    cabezal_formulario
    
    mSub_bloqueo_controles_formulario Me, True
    Me.botConfirmar.Enabled = False
    
    importe_nue = 0
End Sub

Private Sub botConfirmar_Click()
    Dim prox_gasto As Integer
    'verifico si es posible ingresar el alojamiento
    If funValidoDatos Then
        prox_gasto = obtengo_ultimo_corr_aloja(fecha.Value)
        tbCUENTAS_ALOJA.AddNew
            tbCUENTAS_ALOJA("habitacion_cuenta_aloja") = hab_cuenta
            tbCUENTAS_ALOJA("nrocorr_cuenta_aloja") = prox_gasto
            tbCUENTAS_ALOJA("tarifa") = Val(txtTarifaNueva.Text)
            tbCUENTAS_ALOJA("fecha") = fecha.Value
            tbCUENTAS_ALOJA("titular_aloja") = busco_titular_hab2(hab_cuenta, "aloja")
            tbCUENTAS_ALOJA("tipoAloja") = Me.cboMotivoIngreso.ItemData(Me.cboMotivoIngreso.ListIndex)
            tbCUENTAS_ALOJA("obsAloja") = Me.txtObsMotivoIngreso.Text
        tbCUENTAS_ALOJA.Update
        'aviso de confirmación de proceso
        mSubMensaje 4, 8
        'grabo bitacora
        GraboBitacora "Día " & fecha.Text
        Unload Me
        frmIngHabitacion.Show 1
    End If
End Sub

Private Function funValidoDatos() As Boolean
    '-----------------------------------------------------------------------
    'Determina si es posible ingresar el alojamiento manual.
    '-----------------------------------------------------------------------
    'Parámetros.
    '   Salida  True si se seleccionó en el combo un motivo de ingreso
    '           False no se seleccionó motivo de ingreso en el combo
    '-----------------------------------------------------------------------
    If Me.cboMotivoIngreso.ListIndex >= 0 Then
        'se seleccionó un motivo de ingreso
        funValidoDatos = True
    Else
        'no se seleccionó un motivo de ingreso
        funValidoDatos = False
        'muestro aviso de no ingreso de motivo alojamiento
        mSubMensaje 4, 133
        'le doy el focus al combo
        Me.cboMotivoIngreso.SetFocus
    End If
End Function

Private Sub botContinuar_Click()
    If IsDate(fecha.Text) Then
        modifico_formulario
        If busco_alojamiento Then
            LblAlojamiento.Caption = "Modificación de alojamiento"
            txtTarifaAnt.Text = Format(tarifa, "####0.00")
        Else
            LblAlojamiento.Caption = "Ingreso nuevo alojamiento"
            txtTarifaAnt.Text = "0,00"
        End If
        txtTarifaNueva.SetFocus
    Else
        'formato de fecha incorrecto
        mSubMensaje 3, 1
        fecha.SetFocus
    End If
End Sub

Private Sub modifico_formulario()
    Me.botConfirmar.Enabled = True
    LblAlojamiento.Visible = True
    botContinuar.Enabled = False
    mSubBloqueoControlFormulario Me.fecha, True
    mSubBloqueoControlFormulario Me.txtTarifaNueva, False
    mSubBloqueoControlFormulario Me.cboMotivoIngreso, False
    mSubBloqueoControlFormulario Me.txtObsMotivoIngreso, False
End Sub

Private Function busco_alojamiento()
    Dim titular As Long
    Dim fecha As Date
    
    'Recorro los gastos de alojamiento que tiene el titular en un día dado,
    'siempre y cuando no estén facturados.
    'Estos gástos pueden pertenecer a distintas habitaciones, por ejemplo
    'si dos habitaciones tienen titutal alojamiento compartido
    'aparecerán ámbas habitaciones en su estado de cuentas.
    'Se puede dar el caso también que el titular se cambie de habitación,
    'dejando la habitación origen libre, los gastos de su nueva habitación
    'aparecerán en el estado de cuenta también.
    'cuando se realize un alojamiento manual en una fecha determinada, para
    'modificar el importe de un alojamiento anterior, es posible que la habitación
    'a la cual se desea modificar el importe no pertenesca ya a el titular (caso explicado
    'anteriormente), por ese motivo no se podrá modificar el importe del alojamiento,
    'sino que en su defecto se creará otro importe que corresponderá a la nueva habitación
    'el cual modificará el importe total del alojamiento a pagar por el cliente.
    
    
    titular = busco_titular_hab2(hab_cuenta, "aloja")
    
    tarifa = 0
    busco_alojamiento = False
    fecha = Me.fecha.Value
    
    tbCUENTAS_ALOJA.Index = "i_titular"
    tbCUENTAS_ALOJA.Seek ">=", 0, titular, fecha, hab_cuenta
    If Not tbCUENTAS_ALOJA.NoMatch Then
        Do While Not tbCUENTAS_ALOJA.EOF
            If tbCUENTAS_ALOJA("titular_aloja") = titular And _
                tbCUENTAS_ALOJA("fecha") = fecha And _
                tbCUENTAS_ALOJA("habitacion_cuenta_aloja") = hab_cuenta And _
                tbCUENTAS_ALOJA("facturado") = 0 Then
                    tarifa = tarifa + tbCUENTAS_ALOJA("tarifa")
                    busco_alojamiento = True
                tbCUENTAS_ALOJA.MoveNext
            Else
                Exit Do
            End If
        Loop
    End If
End Function

Private Sub cabezal_formulario()
    Me.gaHOTELtitular1.CaminoBaseDeDatos = vardir
    Me.gaHOTELtitular1.NumeroHabitacion = hab_cuenta
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Descargo formulario de memoria
    Set frmCargaAlojaManual = Nothing
End Sub

Private Sub txtTarifaNueva_Change()
    Dim ting As Double
    ting = Val(txtTarifaNueva.Text)
    importe_nue = ting + tarifa
    txtImporteNue.Text = Format(importe_nue, "####0.00")
End Sub

Private Sub txtTarifaNueva_KeyPress(KeyAscii As Integer)
    ValidoNum KeyAscii, True, True
End Sub

Private Sub botCancelar_Click()
    Unload Me
    frmIngHabitacion.Show 1
End Sub

Private Sub mnuProcesar_Click()
    'Esta opción del menú es lo mismo que presionar la tecla F9
    'o el boton de procesar
    If Me.botContinuar.Enabled = True Then
        botContinuar_Click
    End If
End Sub

Private Sub mnuFormularioAceptar_Click()
    'Esta opción es lo mismo que apretar el boton de aceptar o la tecla F12
    If Me.botConfirmar.Enabled = True Then
        botConfirmar_Click
    End If
End Sub

Private Sub mnuFormularioCancelar_Click()
    'Esta opción del menu es lo mismo que presionar el botón de cancelar o la tecla Esc
    botCancelar_Click
End Sub

'********************************************************************
'*
'* Asistencia para usuario
'*
'********************************************************************

Private Sub txtObsMotivoIngreso_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 210
End Sub

Private Sub cboMotivoIngreso_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 209
End Sub

Private Sub fecha_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 15
End Sub

Private Sub botContinuar_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 1, 4
End Sub

Private Sub botConfirmar_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 1, 2
End Sub

Private Sub botCancelar_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 1, 3
End Sub

Private Sub txtTarifaNueva_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 16
End Sub

Private Sub fecha_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub botContinuar_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub botConfirmar_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub botCancelar_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub txtTarifaNueva_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub txtObsMotivoIngreso_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub cboMotivoIngreso_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

