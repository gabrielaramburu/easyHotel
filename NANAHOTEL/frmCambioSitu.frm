VERSION 5.00
Begin VB.Form frmCambioSitu 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cambio de situación"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   8160
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   8160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Hotel_Nana.gaHOTELbarra gaHOTELbarra1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   4980
      Width           =   8160
      _ExtentX        =   14393
      _ExtentY        =   582
   End
   Begin VB.Frame Frame2 
      Caption         =   "Habitación"
      Height          =   855
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   7935
      Begin Hotel_Nana.gaHOTELtipo gaHOTELtipo1 
         Height          =   300
         Left            =   120
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   360
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   529
         BackColor       =   -2147483633
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Estado/Situacion "
      Height          =   3975
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   7935
      Begin VB.CommandButton botSalir 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   6600
         Picture         =   "frmCambioSitu.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         Tag             =   "Cancelar"
         Top             =   3480
         Width           =   1215
      End
      Begin VB.CommandButton botConfirmar 
         Height          =   375
         Left            =   5280
         Picture         =   "frmCambioSitu.frx":08C2
         Style           =   1  'Graphical
         TabIndex        =   10
         Tag             =   "Aceptar"
         Top             =   3480
         Width           =   1215
      End
      Begin VB.TextBox txtEstadoActual 
         Height          =   375
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   420
         Width           =   3615
      End
      Begin VB.TextBox txtfecha 
         Height          =   375
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1440
         Width           =   1575
      End
      Begin VB.ComboBox cboNueva 
         Height          =   315
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   2520
         Width           =   3615
      End
      Begin VB.ComboBox cboActual 
         Height          =   315
         Left            =   2280
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   960
         Width           =   3615
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha en la que se estableció la situación"
         Height          =   855
         Left            =   240
         TabIndex        =   12
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Estado actual"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "&Situación nueva"
         Height          =   375
         Left            =   240
         TabIndex        =   0
         Top             =   2520
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Situación actual"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   1815
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
Attribute VB_Name = "frmCambioSitu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private hab_cuenta As Long

Private Sub botConfirmar_Click()
    actualizo_situ_habitacion
    grabo_historico_situ
    'aviso de confirmación de cambio de situación
    mSubMensaje 4, 7
    'grabo bitacora
    GraboBitacora "Hab. " & hab_cuenta
    Unload Me
    frmIngHabitacion2.Show 1
End Sub

Private Sub actualizo_situ_habitacion()
    'Cambio la situación actual de la habitación
    tbHABITACIONES.Edit
        tbHABITACIONES("situacionhab") = cboNueva.ItemData(cboNueva.ListIndex)
        tbHABITACIONES("fechasituacionhab") = m_FechaSis
    tbHABITACIONES.Update
End Sub

Private Sub grabo_historico_situ()
    Dim prox As Long
    prox = corr_situ(hab_cuenta)
    tbSITUACION_HIS.AddNew
        tbSITUACION_HIS("nrohab_situ") = hab_cuenta
        tbSITUACION_HIS("corr_situ") = prox
        tbSITUACION_HIS("fechacambio_situ") = m_FechaSis
        tbSITUACION_HIS("situacion_situ") = cboNueva.ItemData(cboNueva.ListIndex)
    tbSITUACION_HIS.Update
End Sub

Private Sub Form_Load()
    'Apariencia formulario
    mSubConfiguroFuentesControlesSistema Me
    
    'obtengo habitacion
    hab_cuenta = Val(frmIngHabitacion2.txtNroHab.Text)
        
    BloqueoControles
    
    cabezal_formulario
    'cargo combos
    busco_estado_habitacion
    carga_tipo_estado_hab cboActual, 2
    carga_tipo_estado_hab cboNueva, 2
    obtengo_situacion_actual
End Sub

Private Sub cabezal_formulario()
    Me.gaHOTELtipo1.CaminoBaseDeDatos = vardir
    Me.gaHOTELtipo1.NumeroHabitacion = hab_cuenta
End Sub

Private Sub obtengo_situacion_actual()
    posiciono_combo cboActual, tbHABITACIONES("situacionhab")
    txtfecha.Text = tbHABITACIONES("fechasituacionhab")
End Sub

Private Sub busco_estado_habitacion()
    If busco_habita_checkin(hab_cuenta) Then
        Me.txtEstadoActual.Text = "Ocupada"
    Else    'si la habitacion esta libre
        If habitacion_bloqueada(hab_cuenta, m_FechaSis, m_FechaSis) Then
            Me.txtEstadoActual.Text = "Bloqueada"
        Else
            Me.txtEstadoActual.Text = "Libre"
        End If
    End If
End Sub

Private Sub BloqueoControles()
    mSub_bloqueo_controles_formulario Me, True
    cboNueva.BackColor = mConstSisColor_Blanco
    cboNueva.Locked = False
    cboNueva.TabStop = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Descargo formulario de memoria
    Set frmCambioSitu = Nothing
End Sub

Private Sub mnuFormularioAceptar_Click()
    'Esta opción del menú equivale a presionar la tecla F12 o
    'a el boton aceptar
    botConfirmar_Click
End Sub

Private Sub mnuFormularioCancelar_Click()
    'Esta opción del menú equivale a preionar la tecla Esc o
    'a el boton cancealr
    botSalir_Click
End Sub

Private Sub botSalir_Click()
    Unload Me
    frmIngHabitacion2.Show 1
End Sub

'******************************************************************************
'*
'*  Asistencia al usuario
'*
'******************************************************************************

Private Sub cboNueva_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 14
End Sub

Private Sub cboNueva_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub botConfirmar_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 1, 2
End Sub

Private Sub botConfirmar_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub botSalir_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 1, 3
End Sub

Private Sub botSalir_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

