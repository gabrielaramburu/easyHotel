VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmIngPaxEmp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Estadados de cuenta"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7905
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   7905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Hotel_Nana.gaHOTELbarra gaHOTELbarra1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   6
      Top             =   3600
      Width           =   7905
      _ExtentX        =   13944
      _ExtentY        =   582
   End
   Begin VB.CommandButton botAyuda 
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
      Height          =   435
      Left            =   7080
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   840
      Width           =   615
   End
   Begin VB.CommandButton botConfirmar 
      Height          =   375
      Left            =   5160
      Picture         =   "frmIngPaxEmp.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Tag             =   "Aceptar"
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox txtNomCli 
      BackColor       =   &H80000016&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   840
      Width           =   4935
   End
   Begin VB.TextBox txtCodCli 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   420
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton botSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   6480
      Picture         =   "frmIngPaxEmp.frx":08B6
      Style           =   1  'Graphical
      TabIndex        =   3
      Tag             =   "Cancelar"
      Top             =   3000
      Width           =   1215
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   3375
      Left            =   120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   5953
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Pax"
            Key             =   ""
            Object.Tag             =   "cli"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Empresa o Agencia"
            Key             =   ""
            Object.Tag             =   "emp"
            ImageVarType    =   2
         EndProperty
      EndProperty
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
   Begin VB.Menu mnuVer 
      Caption         =   "&Ver información de..."
      Begin VB.Menu mnuVerPax 
         Caption         =   "Pax"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuVerEmpAgencia 
         Caption         =   "Empresa o agencia"
         Shortcut        =   {F6}
      End
   End
   Begin VB.Menu mnuBuscar 
      Caption         =   "&Buscar..."
      Begin VB.Menu mnuFormularioClientes 
         Caption         =   "Clientes..."
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmIngPaxEmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private existe As Boolean

Private Sub BotAyuda_Click()
    If TabStrip1.SelectedItem.Tag = "cli" Then
        txtCodCli.Text = mFunBusqueda(1)    'todos los cliente
        If Val(txtCodCli.Text) <> 0 Then
            busco_cli
        End If
    Else
        txtCodCli.Text = mFunBusqueda(3)   'empresas
        If Val(txtCodCli.Text) <> 0 Then
            busco_emp
        End If
    End If
End Sub

Private Sub botConfirmar_Click()
    If existe Then
        Me.Hide
        Select Case tipo_accion_IngEstadoCuenta
            Case 1
                frmEstadoCuentas.Show 1
            Case 2
                frmConsultaCuentas.Show 1
        End Select
    End If
End Sub

Private Sub botSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    'Apariencia formulario
    mSubConfiguroFuentesControlesSistema Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Descargo formulario de memoria
    Set frmIngPaxEmp = Nothing
End Sub

Private Sub TabStrip1_Click()
    txtCodCli.Text = ""
    txtNomCli.Text = ""
    existe = False
End Sub

Private Sub txtCodCli_Change()
    existe = False
    If TabStrip1.SelectedItem.Tag = "cli" Then
        busco_cli
    Else
        busco_emp
    End If
End Sub

Private Sub txtCodCli_KeyPress(KeyAscii As Integer)
    ValidoNum KeyAscii, False, False
End Sub

Private Sub txtCodCli_LostFocus()
    existe = False
    If TabStrip1.SelectedItem.Tag = "cli" Then
        busco_cli
    Else
        busco_emp
    End If
    'asistencia a usuario
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub busco_cli()
    If busco_clienteTF(Val(txtCodCli.Text)) Then
        existe = True
        txtNomCli.Text = tbCLIENTES("nombre_completo_titular")
    Else
        txtNomCli.Text = Empty
    End If
End Sub

Private Sub busco_emp()
    If busco_empTF(Val(txtCodCli.Text)) Then
        existe = True
        txtNomCli.Text = tbEMPRESAS("nomemp")
    Else
        txtNomCli.Text = Empty
    End If
End Sub

Private Sub mnuFormularioAceptar_Click()
    'Equivale a presionar el botón de aceptar o la tecla F12
    botConfirmar_Click
End Sub

Private Sub mnuFormularioCancelar_Click()
    'Equivale a presionar el boton de cancelar o la tecla Esc
    botSalir_Click
End Sub

Private Sub mnuFormularioClientes_Click()
    'Equivale a presionar el botón de ayuda
    BotAyuda_Click
End Sub

Private Sub mnuVerPax_Click()
    'Selecciono la primer ficha de la controls tab
    TabStrip1.Tabs(1).Selected = True
End Sub

Private Sub mnuVerEmpAgencia_Click()
    'Selecciono la segunda ficha
    TabStrip1.Tabs(2).Selected = True
End Sub

'************************************************************
'*
'*  Asistencia a usuarios
'*
'************************************************************

Private Sub botConfirmar_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 32
End Sub

Private Sub botSalir_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 1, 1
End Sub

Private Sub txtCodCli_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 33
End Sub

Private Sub botConfirmar_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub botSalir_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

