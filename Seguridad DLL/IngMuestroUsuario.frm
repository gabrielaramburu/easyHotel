VERSION 5.00
Begin VB.Form IngMuestroUsuario 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control de acceso"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4920
   ControlBox      =   0   'False
   Icon            =   "IngMuestroUsuario.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   4920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstContraseña 
      Height          =   300
      Left            =   2640
      TabIndex        =   7
      Top             =   3120
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox txtContraseña 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   120
      MaxLength       =   50
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2640
      Width           =   4695
   End
   Begin VB.CommandButton botSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton botAceptar 
      Default         =   -1  'True
      Height          =   375
      Left            =   2280
      Picture         =   "IngMuestroUsuario.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton botCancelar 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3600
      Picture         =   "IngMuestroUsuario.frx":0CF8
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3480
      Width           =   1215
   End
   Begin VB.ListBox lstUsuarios 
      Height          =   1740
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4695
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "&Contraseña"
      Height          =   240
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   1035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Seleccionar nombre de usuario"
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2805
   End
End
Attribute VB_Name = "IngMuestroUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Event NotificoCliente(usuario As String, boton As Byte)
'Este es un evento del objeto IngMuestroUsuario
'el cual lo desencadeno cuando se hace click sobre un boton

Private UltimoUsuario As Byte

Private Sub botAceptar_Click()
    'comparo contraseñas
    'comparo la contraseña de la caja de texto
    'con el valor de la lista de contraseñas en la posición
    'del usuario seleccionado.
    
    'si no hay usuarios seleccionados no proceso
    If lstUsuarios.ListIndex <> -1 Then
        If Me.lstContraseña.List(Me.lstUsuarios.ListIndex) = _
        Me.txtContraseña Then
            Me.Visible = False
            UltimoUsuario = lstUsuarios.ListIndex
            RaiseEvent NotificoCliente(Me.lstUsuarios.Text, 1)   'boton aceptar
        Else
            MsgBox "La contraseña es incorrecta", vbExclamation
            'me posiciono en contraseña y marco el texto
            txtContraseña.SetFocus
            txtContraseña.SelStart = 0
            txtContraseña.SelLength = Len(txtContraseña)
        End If
    End If
End Sub

Private Sub botCancelar_Click()
    Me.Visible = False
    RaiseEvent NotificoCliente(Empty, 2)   'boton cancelar
End Sub

Private Sub botSalir_Click()
    Me.Visible = False
    RaiseEvent NotificoCliente(Empty, 3)  'boton salir
End Sub

Private Sub Form_Paint()
    If lstUsuarios.ListCount > 0 Then
      lstUsuarios.ListIndex = UltimoUsuario
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'si se hace click en el boton cerrar, oculta el
    'cuadro de diálogo en lugar de descargarlo
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
        Me.Visible = False
    End If
End Sub
