VERSION 5.00
Begin VB.Form frmIngHabitacion2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ingreso habitación"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   4920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   4920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Hotel_Nana.gaHOTELbarra gaHOTELbarra1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   5115
      Width           =   4920
      _ExtentX        =   8678
      _ExtentY        =   582
   End
   Begin VB.Frame Frame1 
      Caption         =   "&Número de habitación "
      Height          =   5055
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
      Begin VB.CommandButton botCancelar 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   3360
         Picture         =   "frmIngHabitacion2.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Tag             =   "Cancelar"
         Top             =   4560
         Width           =   1215
      End
      Begin VB.CommandButton botConfirmar 
         Height          =   375
         Left            =   2040
         Picture         =   "frmIngHabitacion2.frx":08C2
         Style           =   1  'Graphical
         TabIndex        =   2
         Tag             =   "Aceptar"
         Top             =   4560
         Width           =   1215
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
Attribute VB_Name = "frmIngHabitacion2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Este formulario se utiliza con todos los procesos que
'requieren ingresar habitación.
'No es necesario que la misma este ocupada; solo que exista.

Private Sub Form_Load()
    'Apariencia formulario
    mSubConfiguroFuentesControlesSistema Me
    
    Select Case tipo_accion_inghabitacion2
        Case 1
            Me.Caption = "Cambio de situación"
        Case 2
            Me.Caption = "Consulta de cambios de situación"
        Case 3
            Me.Caption = "Bloqueo de habitaciones"
    End Select
End Sub

Private Sub botConfirmar_Click()
    'Valido que exista la habitación
    If busco_habitaTF(Val(txtNroHab.Text)) Then
        Me.Hide
        llamo_formulario
    Else
        'no existe la habitación
        mSubMensaje 4, 17
        limpio_txthab
    End If
End Sub

Private Sub botAyudaHab_Click()
    'Llamo ayuda de habitaciones para mostrar
    'todas las habitaciones del hotel
    Me.txtNroHab.Text = mFunBusqueda(8)
End Sub

Private Sub llamo_formulario()
    Select Case tipo_accion_inghabitacion2
        Case 1
            frmCambioSitu.Show 1
        Case 3
            frmBloquearHab.Show 1
    End Select
End Sub

Private Sub limpio_txthab()
    txtNroHab.Text = ""
    txtNroHab.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Descargo formulario de memoria
    Set frmIngHabitacion2 = Nothing
End Sub

Private Sub txtNroHab_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        botConfirmar.Value = True
    End If
    ValidoNum KeyAscii, False, False
End Sub

Private Sub botCancelar_Click()
    Unload Me
End Sub

Private Sub mnuBuscarHabitaciones_Click()
    'Equivale a presionar la tecla F1 o el boton de ayuda
    botAyudaHab_Click
End Sub

Private Sub mnuFormularioAceptar_Click()
    'Equivale e preionar el boton de acepta o la tecla F12
    botConfirmar_Click
End Sub

Private Sub mnuFormularioCancelar_Click()
    'Equivale a presionar la tecla Esc o el boton de cancelar
    botCancelar_Click
End Sub

