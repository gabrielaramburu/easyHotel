VERSION 5.00
Begin VB.Form IngContra 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Contraseña de usuarios"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   4920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton botEliminar 
      Caption         =   "Eliminar"
      Height          =   375
      Left            =   2280
      TabIndex        =   9
      Top             =   2640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox lstUsuarios 
      Height          =   540
      Left            =   240
      TabIndex        =   8
      Top             =   2520
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtContraseñaConfirm 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   240
      MaxLength       =   50
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   1920
      Width           =   3015
   End
   Begin VB.TextBox txtContraseñaNueva 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   240
      MaxLength       =   50
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1200
      Width           =   3015
   End
   Begin VB.TextBox txtUsuario 
      Height          =   375
      Left            =   240
      MaxLength       =   50
      TabIndex        =   1
      Top             =   480
      Width           =   3015
   End
   Begin VB.CommandButton botCancelar 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3600
      Picture         =   "IngContra.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton botAceptar 
      Default         =   -1  'True
      Height          =   375
      Left            =   2280
      Picture         =   "IngContra.frx":08C2
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Usuario"
      Height          =   240
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   705
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "&Confirmar contraseña"
      Height          =   240
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   1905
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "&Nueva contraseña"
      Height          =   240
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   1650
   End
End
Attribute VB_Name = "IngContra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public TipoAccionContra As Byte
Event NotificoCliente(boton As Byte)

Private Sub botEliminar_Click()
    'valido que se halla ingresado usuario
    If Len(txtUsuario.Text) > 0 Then
        'valido que el usuario exista
        If Not existe_usuario Then
            MsgBox "El usuario " & txtUsuario.Text & " no existe", vbExclamation
            CorrijoUsuario
        Else
            Me.Visible = False
            RaiseEvent NotificoCliente(3)   'boton eliminar
        End If
    Else
        MsgBox "Debe de ingresar usuario", vbExclamation
        txtUsuario.SetFocus
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

Private Sub botCancelar_Click()
    Me.Visible = False
    RaiseEvent NotificoCliente(2)   'boton cancelar
End Sub

Private Sub botAceptar_Click()
    Dim UsuarioOk As Boolean
    UsuarioOk = True
    
    'valido que se halla ingresado usuario
    If Len(txtUsuario.Text) > 0 Then
        If TipoAccionContra = 1 Then 'nuevo usuario
            'valido que el usuario no exista
            If existe_usuario Then
                UsuarioOk = False
                MsgBox "El usuario " & txtUsuario.Text & " ya existe", vbExclamation
                CorrijoUsuario
            End If
        End If
        If TipoAccionContra = 2 Then    'modifico
            'valido que el usuario exista
            If Not existe_usuario Then
                UsuarioOk = False
                MsgBox "El usuario " & txtUsuario.Text & " no existe", vbExclamation
                CorrijoUsuario
            End If
        End If
    Else
        MsgBox "Debe de ingresar usuario", vbExclamation
        txtUsuario.SetFocus
        UsuarioOk = False
    End If

    If UsuarioOk Then
        'valido que la nueva contraseña sea válida
        If Len(Me.txtContraseñaNueva.Text) > 0 Then
            'valido que la contraseña nueva sea igual a la confirmación
            If txtContraseñaNueva.Text = txtContraseñaConfirm.Text Then
                Me.Visible = False
                RaiseEvent NotificoCliente(1)   'boton aceptar
            Else
                MsgBox "La confirmación no es correcta", vbExclamation
                txtContraseñaConfirm.SetFocus
                txtContraseñaConfirm.SelStart = 0
                txtContraseñaConfirm.SelLength = (Len(txtContraseñaConfirm.Text))
            End If
        Else
            MsgBox "La contraseña debe de tener por lo menos 1 dígito", vbExclamation
            txtContraseñaNueva.SetFocus
        End If
    End If
End Sub

Private Sub CorrijoUsuario()
    'limpio contraseñas
    txtContraseñaNueva.Text = Empty
    txtContraseñaConfirm.Text = Empty
    'posiciono y marco usuario erroneo
    txtUsuario.SetFocus
    txtUsuario.SelStart = 0
    txtUsuario.SelLength = (Len(txtUsuario.Text))
End Sub

Private Function existe_usuario()
    'busco el usuario ingresado en la lista de usuarios
    existe_usuario = False
    lstUsuarios.Text = Me.txtUsuario.Text
    If lstUsuarios.ListIndex <> -1 Then 'encontre
        existe_usuario = True
    End If
End Function

