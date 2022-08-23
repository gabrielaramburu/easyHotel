VERSION 5.00
Begin VB.Form frmIngHabitacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Habitación"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   4920
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   4920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Hotel_Nana.gaHOTELbarra gaHOTELbarra1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   5205
      Width           =   4920
      _ExtentX        =   8678
      _ExtentY        =   582
   End
   Begin VB.Frame Frame1 
      Caption         =   "&Número de habitación "
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.CommandButton botAyudaHab 
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   1560
         TabIndex        =   5
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton botSalir 
         Height          =   370
         Left            =   3360
         Picture         =   "frmIngHabitacion.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Tag             =   "Cancelar"
         Top             =   4680
         Width           =   1200
      End
      Begin VB.TextBox txtNroHab 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   480
         Left            =   240
         MaxLength       =   4
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton botConfirmar 
         Height          =   370
         Left            =   2040
         Picture         =   "frmIngHabitacion.frx":08C2
         Style           =   1  'Graphical
         TabIndex        =   2
         Tag             =   "Aceptar"
         Top             =   4680
         Width           =   1200
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
   Begin VB.Menu mnuBuscar 
      Caption         =   "&Buscar"
      Begin VB.Menu mnuBuscarHabitaciones 
         Caption         =   "Buscar habitaciones"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmIngHabitacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub botAyudaHab_Click()
    'Muestro todas las habitaciones ocupadas del hotel.
    Me.txtNroHab.Text = mFunBusqueda(9)
End Sub

'Este formulario se utiliza solo con los procesos que
'requieren que la habitación este ocupada.
Private Sub Form_Load()
    'Apariencia formulario
    mSubConfiguroFuentesControlesSistema Me
    configuracion_apariencia
    
    Select Case tipo_accion_inghabitacion
        Case 1
            frmIngHabitacion.Caption = "Ingreso de extras"
        Case 2
            frmIngHabitacion.Caption = "Resumen de cuentas"
        Case 3
            frmIngHabitacion.Caption = "Cambio de tarifa"
        Case 4
            frmIngHabitacion.Caption = "Alojamiento manual"
        Case 5
            frmIngHabitacion.Caption = "Cambio de titular"
        Case 6
            frmIngHabitacion.Caption = "Facturación"
        Case 7
            frmIngHabitacion.Caption = "Pasajeros por habitación"
        Case 8
            frmIngHabitacion.Caption = "Check-Out"
        Case 9
            frmIngHabitacion.Caption = "Cambio fecha de egreso"
    End Select
End Sub

Private Sub botConfirmar_Click()
    'Valido que exista la habitación y busco si está ocupada
    Dim consulta As String
    If busco_habitaTF(Val(txtNroHab.Text)) Then
        If busco_habita_checkin(Val(txtNroHab.Text)) Then
            Me.Hide
            'si la habitación esta ocupada tengo que validar
            'que el período de ocupación se válido
            If Not mFunDeterminoOcupacionValida(Val(txtNroHab.Text)) Then
                'si el período de ocupación es incorrecto muestro mensaje
                mSubMensaje 4, 129, _
                "Se aconseja realizar el Check-Out o cambiar la fecha de egreso."
            End If
            llamo_formulario
        Else
            'no hay pasajeros hospedados en esa habitación
            mSubMensaje 4, 16
            limpio_txthab
        End If
    Else
        'no existe la habitación
        mSubMensaje 4, 17
        limpio_txthab
    End If
End Sub

Private Sub llamo_formulario()
    Select Case tipo_accion_inghabitacion
        Case 1
            frmIngExtras.Show 1
        Case 2
            frmConsultaCuentas.Show 1
        Case 3
            frmTarifas.Show 1
        Case 4
            frmCargaAlojaManual.Show 1
        Case 5
            'iniciaizo propiedades del formulario
            frmTitularesHabitacion.propTipoAccionFormularioTitular = 4  'cambio de titular
            frmTitularesHabitacion.propHabCuenta = CLng(Me.txtNroHab.Text)
            frmTitularesHabitacion.Show 1
        Case 6
            frmFacturacion.Show 1
        Case 7
            frmPasajerosHabitacion.Show 1
        Case 8
            frmCheck_Out.Show 1
        Case 9
            frmCambioFechaEgreso.Show 1
    End Select
End Sub

Private Sub limpio_txthab()
    txtNroHab.Text = Empty
    txtNroHab.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Descargo formulario de memoria
    Set frmIngHabitacion = Nothing
End Sub

Private Sub txtNroHab_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        botConfirmar.Value = True
    End If
    ValidoNum KeyAscii, False, False
End Sub

Private Sub configuracion_apariencia()
    'Determina la apariencia del los elemento configurables del formulario
End Sub

Private Sub botSalir_Click()
    Unload Me
End Sub

Private Sub mnuBuscarHabitaciones_Click()
    'Equivale a presionar F1 o el boton de ayuda
    botAyudaHab_Click
End Sub

Private Sub mnuFormularioAceptar_Click()
    'Equivale a digitar F12 o a el boton aceptar
    botConfirmar_Click
End Sub

Private Sub mnuFormularioCancelar_Click()
    'Equivale a digitar Esc o el boton de cancelar
    botSalir_Click
End Sub

