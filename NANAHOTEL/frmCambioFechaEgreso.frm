VERSION 5.00
Object = "{EB8C3860-8603-11D0-A35D-7062E4000000}#1.1#0"; "VCBND.OCX"
Begin VB.Form frmCambioFechaEgreso 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cambio de fecha de egreso"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   9510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Hotel_Nana.gaHOTELtitular gaHOTELtitular1 
      Height          =   1335
      Left            =   120
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   0
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   2355
      BackColor       =   -2147483633
   End
   Begin Hotel_Nana.gaHOTELbarra gaHOTELbarra1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3975
      Width           =   9510
      _ExtentX        =   16775
      _ExtentY        =   582
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cambio de fecha egreso"
      Height          =   2415
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   9255
      Begin VB.CommandButton botCancelar 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   7920
         Picture         =   "frmCambioFechaEgreso.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         Tag             =   "Cancelar"
         Top             =   1920
         Width           =   1215
      End
      Begin VB.CommandButton botAceptar 
         Height          =   375
         Left            =   6600
         Picture         =   "frmCambioFechaEgreso.frx":08C2
         Style           =   1  'Graphical
         TabIndex        =   8
         Tag             =   "Aceptar"
         Top             =   1920
         Width           =   1215
      End
      Begin VcBndCtl.VcCalCombo fEgresoNueva 
         Height          =   375
         Left            =   7080
         TabIndex        =   1
         Top             =   495
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _0              =   $"frmCambioFechaEgreso.frx":1178
         _1              =   $"frmCambioFechaEgreso.frx":1581
         _2              =   $"frmCambioFechaEgreso.frx":198A
         _3              =   "-@@@A@@@)f@@@E@@@A@@@@@@@@@'@@@E@@@A@@@@@@@@@)h@@@E@@@A@@@@@@@@@,456D"
         _count          =   4
         _ver            =   2
      End
      Begin VB.TextBox txtFingreso 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         Height          =   405
         Left            =   1920
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txtFegreso 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         Height          =   405
         Left            =   1920
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1245
         Width           =   1455
      End
      Begin VB.Label lblAtencion 
         Caption         =   "lblAtencion"
         Height          =   735
         Left            =   4800
         TabIndex        =   11
         Top             =   960
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   4200
         Picture         =   "frmCambioFechaEgreso.frx":1D93
         Top             =   960
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "&Nueva fecha de egreso"
         Height          =   240
         Left            =   4800
         TabIndex        =   0
         Top             =   562
         Width           =   2115
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de egreso"
         Height          =   240
         Left            =   240
         TabIndex        =   4
         Top             =   1320
         Width           =   1545
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de ingreso"
         Height          =   240
         Left            =   240
         TabIndex        =   3
         Top             =   562
         Width           =   1575
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
   End
End
Attribute VB_Name = "frmCambioFechaEgreso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private hab_cuenta As Long

Private Sub Form_Load()
    'Apariencia formulario
    mSubConfiguroFuentesControlesSistema Me
    
    'obtengo habitacion
    hab_cuenta = Val(frmIngHabitacion.txtNroHab.Text)
    
    If Not mFunDeterminoOcupacionValida(hab_cuenta) Then
        Me.lblAtencion.Visible = True
        lblAtencion.Caption = "Atención:" & Chr(10) & "Período de ocupación fuera del establecido."
        Me.Image1.Visible = True
    End If
    
    cabezal_formulario
    muestro_fechas_ingreso_egreso
End Sub

Private Sub muestro_fechas_ingreso_egreso()
    'busco en chekin la habitación (primer pasajero hospedado)
    'y obtengo las fechas actuales de ingreso y egreso
    If busco_habita_checkin(hab_cuenta) Then
        txtFingreso.Text = tbCHECKIN("fcheckdes")
        txtFegreso.Text = tbCHECKIN("fcheckhas")
    End If
End Sub

Private Sub cabezal_formulario()
    Me.gaHOTELtitular1.CaminoBaseDeDatos = vardir
    Me.gaHOTELtitular1.NumeroHabitacion = hab_cuenta
End Sub

Private Sub botAceptar_Click()
    Dim fd As Date
    Dim fh As Date
    fd = m_FechaSis
    fh = fEgresoNueva.Value
    If valido_fecha_nueva Then
        'verifico no esté reservada
        If Not habitacion_reservada(hab_cuenta, fd, fh) Then
            'verifico no este bloqueada
            If Not habitacion_bloqueada(hab_cuenta, fd, fh) Then
                cambio_fecha_egreso
                'aviso de confirmación de cambio de fecha de egreso
                mSubMensaje 4, 6
                'grabobitacora
                GraboBitacora "Hab. " & hab_cuenta & "N.F.E. " & fEgresoNueva.Text
                Unload Me
                frmIngHabitacion.Show 1
            Else
                'en el nuevo período seleccionado la habitación se encuntra bloqueada
                mSubMensaje 4, 3
                fEgresoNueva.SetFocus
            End If
        Else
            'en el nuevo período seleccionado la habitación se encuentra reservada
            mSubMensaje 4, 2
            fEgresoNueva.SetFocus
        End If
    End If
End Sub

Private Function valido_fecha_nueva()
    'Valído que la nueva fecha de egreso sea una fecha válida.
    valido_fecha_nueva = True
    If Not IsDate(fEgresoNueva.Text) Then
        'el formato de la fecha no es correcto
        mSubMensaje 3, 1
        fEgresoNueva.SetFocus
        valido_fecha_nueva = False
        Exit Function
    End If
    
    If fEgresoNueva.Value < m_FechaSis Then
        'la nueva fecha de egreso no puede ser menor a la fecha de hoy
        mSubMensaje 4, 5
        fEgresoNueva.SetFocus
        valido_fecha_nueva = False
        Exit Function
    End If

    If fEgresoNueva.Value = CDate(Me.txtFingreso.Text) Then
        'la nueva fecha de egreso no puede ser igual a la fecha de ingreso
        mSubMensaje 4, 4
        fEgresoNueva.SetFocus
        valido_fecha_nueva = False
        Exit Function
    End If
End Function

Private Sub cambio_fecha_egreso()
    'Recorro todos los pasajeros hospedados de la habitación
    'y cambio fecha de egreso por la nueva fecha.
    Dim consulta As String
    consulta = "UPDATE checkin " & _
    "SET fcheckhas = " & fechaSQL(fEgresoNueva.Text) & _
    " WHERE nrohab= " & Str(hab_cuenta)
    bdHOTEL.Execute consulta
End Sub

Private Sub botCancelar_Click()
    Unload Me
    frmIngHabitacion.Show 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Descargo formulario de memoria
    Set frmCambioFechaEgreso = Nothing
End Sub

Private Sub mnuFormularioAceptar_Click()
    'Presionar la tecla F12 es lo mismo que el boton aceptar
    botAceptar_Click
End Sub

Private Sub mnuFormularioCancelar_Click()
    'Desde esta opción del menú cancelo la operación
    'Es lo mismo que presionar la teccla ESC
    botCancelar_Click
End Sub

'******************************************************************************
'*
'*  Asistencia al usuario
'*
'******************************************************************************

Private Sub botAceptar_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 1, 2
End Sub

Private Sub fEgresoNueva_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 8
End Sub

Private Sub botCancelar_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 1, 3
End Sub

Private Sub botAceptar_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub botCancelar_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub fEgresoNueva_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

