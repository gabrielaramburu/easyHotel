VERSION 5.00
Begin VB.Form frmInicializarAplicacion 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6885
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Left            =   360
      Top             =   1080
   End
   Begin VB.Label lbl1Aviso 
      Caption         =   "lbl1Aviso"
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   480
      Width           =   5655
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "frmInicializarAplicacion.frx":0000
      Top             =   480
      Width           =   480
   End
End
Attribute VB_Name = "frmInicializarAplicacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Declaración de constantes
Private Const cTiempoEspera As Integer = 2000   'Milisegundos que se muestra el formulario

Private Sub Form_Load()
    'Inicializo etiquta
    lbl1Aviso.Caption = "Inicializando aplicación." & Chr(10) & _
                        "Por favor espere unos segundos."
    'inicializo control timer
    Timer1.Interval = cTiempoEspera
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    'Finalizó el tiempo de espera, en el cual se muestra el formulario.
    Me.Visible = False
    Timer1.Enabled = False
End Sub
