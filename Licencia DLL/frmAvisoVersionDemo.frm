VERSION 5.00
Begin VB.Form frmAvisoVersionDemo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7320
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   480
      Top             =   6600
   End
   Begin VB.Frame Frame1 
      Height          =   6375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7095
      Begin VB.Image Image2 
         Height          =   210
         Left            =   240
         Picture         =   "frmAvisoVersionDemo.frx":0000
         Stretch         =   -1  'True
         Top             =   5040
         Width           =   6660
      End
      Begin VB.Label lblVersionAplicación 
         Caption         =   "lblVersionAplicación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   4455
      End
      Begin VB.Label lbl1Aviso 
         Caption         =   "lbl1Aviso"
         Height          =   1095
         Left            =   120
         TabIndex        =   7
         Top             =   1680
         Width           =   4455
      End
      Begin VB.Label lblPeriodoDeUso 
         Alignment       =   2  'Center
         Caption         =   "lblPeriodoDeUso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   2880
         Width           =   4455
      End
      Begin VB.Label lbl2Aviso 
         Caption         =   "lbl2Aviso"
         Height          =   1335
         Left            =   120
         TabIndex        =   5
         Top             =   3360
         Width           =   4455
      End
      Begin VB.Label lblDerechos 
         Caption         =   "lblDerechos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   4
         Top             =   5280
         Width           =   6735
      End
      Begin VB.Label lblNomAplicacion 
         Caption         =   "lblNomAplicacion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   22.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   4095
      End
      Begin VB.Label lblSistemaAplicacion 
         Caption         =   "lblSistemaAplicacion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   4095
      End
      Begin VB.Image Image1 
         Height          =   4530
         Left            =   4560
         Picture         =   "frmAvisoVersionDemo.frx":4BBA
         Stretch         =   -1  'True
         Top             =   240
         Width           =   2445
      End
   End
   Begin VB.CommandButton botAceptar 
      Caption         =   "&Aceptar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6000
      TabIndex        =   0
      Top             =   6600
      Width           =   1215
   End
End
Attribute VB_Name = "frmAvisoVersionDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Declaración de constantes
Private Const cTiempoEspera As Integer = 4000   'Tiempo de espera antes de permitir
                                                'cerrar el formulario.

'Declaración de propiedades del formulario
Public propDiasVersionDemo As String    'cantidad de días de uso que se autoriza a la versión demo

Private Sub Form_Activate()
    'cuando muestro el formulario empiezo a contar por un lapso de tiempo
    'determinado antes de desencadenar el evento timer
    Timer1.Enabled = True
    Timer1.Interval = cTiempoEspera
End Sub

Private Sub Timer1_Timer()
    'Cuando se desencadena el evento permito cerrar el formulario.
    Timer1.Enabled = False
    botAceptar.Enabled = True
End Sub

'Este formulario, como todos los formularios, no se invocan
'directamene desde la aplicación cliente porque los mismos son clases privadas.
'Los clientes no pueden crear instancias de clases privadas.

Private Sub Form_Load()
    'Cargo el texto que se muestra en las ventanas de avisos,
    'el cual no se establece por propiedades del objeto.
    lbl1Aviso.Caption = "Usted puede evaluar gratuitamente este programa por un período de " _
                        & propDiasVersionDemo & " días." & Chr(10) & _
                        "Después debe de REGISTRAR el producto o dejar de usar el mismo."
    
    lbl2Aviso.Caption = "Registrándose usted tiene acceso a todas las funciones de la última" & _
                        " versión del programa. No tendrá recordatorios ni límite de tiempo. " & _
                        "Usted también tiene derecho a versiones futuras a precios REDUCIDOS."
End Sub

Private Sub botAceptar_Click()
    Me.Visible = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Evita que se cierre el formulario sin el conocimento del objeto
    'AvisoVersionDemo.
    'Si frmAvisoVersionDemo se muestra como un cuadro de diálogo modal, ocultar el
    'cuadro de diálogo en lugar de descargarlo permite que un método de la clase
    'AvisoVersionDemo recupere algún valor de este formulario.
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
        Me.Visible = False
    End If
End Sub

