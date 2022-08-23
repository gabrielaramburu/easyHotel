VERSION 5.00
Begin VB.Form frmTarifas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento de tarifas"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   6915
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   6915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Hotel_Nana.gaHOTELtitular gaHOTELtitular1 
      Height          =   1335
      Left            =   120
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   2355
      BackColor       =   12632256
   End
   Begin Hotel_Nana.gaHOTELbarra gaHOTELbarra1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5355
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   582
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cambio de tarifa "
      Height          =   3855
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   6615
      Begin VB.CommandButton botCancelar 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   5280
         Picture         =   "frmTarifas.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         Tag             =   "Cancelar"
         Top             =   3360
         Width           =   1215
      End
      Begin VB.CommandButton botConfirmar 
         Height          =   375
         Left            =   3960
         Picture         =   "frmTarifas.frx":08C2
         Style           =   1  'Graphical
         TabIndex        =   2
         Tag             =   "Aceptar"
         Top             =   3360
         Width           =   1215
      End
      Begin VB.TextBox txttarifa_nueva 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   1
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox txttarifa_ant 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "&Tarifa nueva"
         Height          =   240
         Left            =   240
         TabIndex        =   0
         Top             =   1260
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tarifa anteriror"
         Height          =   240
         Left            =   240
         TabIndex        =   5
         Top             =   540
         Width           =   1305
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
Attribute VB_Name = "frmTarifas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private nrohab As Long

Private Sub botCancelar_Click()
    Unload Me
    frmIngHabitacion.Show 1
End Sub

Private Sub botConfirmar_Click()
    'valido que se ingrese algo en el cuadro de tarifas
    If Trim(Me.txttarifa_nueva.Text) = Empty Then
        'debe de ingresar nueva tarifa
        mSubMensaje 4, 39
    Else
        tbHABITACIONES.Edit
            tbHABITACIONES("tarifa") = Val(txttarifa_nueva.Text)
        tbHABITACIONES.Update
        'grabo bitacora
        'aviso de cambio de tarifa correctamente
        mSubMensaje 4, 38
        GraboBitacora "hab. " & nrohab
        Unload Me
        frmIngHabitacion.Show 1
    End If
End Sub

Private Sub Form_Load()
    'Apariencia formulario
    mSubConfiguroFuentesControlesSistema Me
    
    Me.txttarifa_ant.BackColor = mSisColor_18ControlesNoHabilitados
    Me.txttarifa_ant.TabStop = False
    
    'obtengo tarifa anterior
    nrohab = Val(frmIngHabitacion.txtNroHab.Text)
    If busco_habitaTF(nrohab) Then
        txttarifa_ant.Text = tbHABITACIONES("tarifa")
    End If
    
    cabezal_formulario
End Sub

Private Sub cabezal_formulario()
    Me.gaHOTELtitular1.CaminoBaseDeDatos = vardir
    Me.gaHOTELtitular1.NumeroHabitacion = nrohab
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Descargo formulario de memoria
    Set frmTarifas = Nothing
End Sub

Private Sub txttarifa_nueva_KeyPress(KeyAscii As Integer)
    ValidoNum KeyAscii, True, True
End Sub

Private Sub mnuFormularioAceptar_Click()
    'Equivale a presionar el boton de aceptar o la tecla F12
    botConfirmar_Click
End Sub

Private Sub mnuFormularioCancelar_Click()
    'Equivale a presionar el boton de cancelar o la tecla Esc
    botCancelar_Click
End Sub

'****************************************************
'*
'*  Asisencia a usuarios
'*
'****************************************************

Private Sub txttarifa_nueva_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 52
End Sub

Private Sub botConfirmar_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 1, 2
End Sub

Private Sub botCancelar_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 1, 1
End Sub

Private Sub botConfirmar_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub botCancelar_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub txttarifa_nueva_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

