VERSION 5.00
Begin VB.Form IngSinUsuario 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control de acceso"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4890
   Icon            =   "IngSinUsuario.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   4890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstContraseña 
      Height          =   780
      Left            =   1080
      TabIndex        =   7
      Top             =   2160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ListBox lstUsuarios 
      Height          =   780
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtUsuario 
      Height          =   375
      Left            =   240
      MaxLength       =   50
      TabIndex        =   1
      Top             =   600
      Width           =   3015
   End
   Begin VB.TextBox txtContraseña 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   240
      MaxLength       =   50
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1320
      Width           =   3015
   End
   Begin VB.CommandButton botAceptar 
      Default         =   -1  'True
      Height          =   375
      Left            =   2280
      Picture         =   "IngSinUsuario.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton botCancelar 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3600
      Picture         =   "IngSinUsuario.frx":0CF8
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Nombre de usuario"
      Height          =   240
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1725
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "&Contraseña"
      Height          =   240
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   1035
   End
End
Attribute VB_Name = "IngSinUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Event NotificoCliente(usuario As String, boton As Byte)

Private Sub botAceptar_Click()
    'busco usuario
    If existe_usuario Then
        If Me.lstContraseña.List(Me.lstUsuarios.ListIndex) = _
        Me.txtContraseña Then
            Me.Visible = False
            RaiseEvent NotificoCliente(Me.txtUsuario.Text, 1)    'boton aceptar
        Else
            MsgBox "La contraseña no es correcta", vbExclamation
            txtContraseña.SetFocus
            txtContraseña.SelStart = 0
            txtContraseña.SelLength = Len(txtContraseña.Text)
        End If
    Else
        MsgBox "El usuario no existe", vbExclamation
        txtUsuario.SetFocus
        txtUsuario.SelStart = 0
        txtUsuario.SelLength = Len(txtUsuario.Text)
    End If
End Sub

Private Function existe_usuario()
    'busco el usuario ingresado en la lista de usuarios
    existe_usuario = False
    lstUsuarios.Text = Me.txtUsuario.Text
    If lstUsuarios.ListIndex <> -1 Then 'encontre
        existe_usuario = True
    End If
End Function

Private Sub botCancelar_Click()
    Me.Visible = False
    RaiseEvent NotificoCliente(Empty, 2)   'boton cancelar
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'si se hace click en el boton cerrar, oculta el
    'cuadro de diálogo en lugar de descargarlo
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
        Me.Visible = False
    End If
End Sub


