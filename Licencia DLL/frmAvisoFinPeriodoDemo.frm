VERSION 5.00
Begin VB.Form frmAvisoFinPeriodoDemo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   6780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   6600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton botCerrar 
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   5280
      TabIndex        =   4
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      Begin VB.Image Image1 
         Height          =   210
         Left            =   120
         Picture         =   "frmAvisoFinPeriodoDemo.frx":0000
         Stretch         =   -1  'True
         Top             =   4800
         Width           =   6060
      End
      Begin VB.Image Image2 
         Height          =   210
         Left            =   0
         Picture         =   "frmAvisoFinPeriodoDemo.frx":4BBA
         Stretch         =   -1  'True
         Top             =   0
         Width           =   6660
      End
      Begin VB.Label lbl2Aviso 
         Caption         =   "lbl2Aviso"
         Height          =   735
         Left            =   120
         TabIndex        =   8
         Top             =   3960
         Width           =   5895
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
         TabIndex        =   7
         Top             =   5040
         Width           =   6135
      End
      Begin VB.Label lbl1Aviso 
         Caption         =   "lbl1Aviso"
         Height          =   2175
         Left            =   120
         TabIndex        =   6
         Top             =   1800
         Width           =   5895
      End
      Begin VB.Label lblPeriodoTerminado 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "lblPeriodoTerminado"
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
         TabIndex        =   5
         Top             =   1440
         Width           =   5895
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
         TabIndex        =   3
         Top             =   720
         Width           =   4095
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
         TabIndex        =   2
         Top             =   240
         Width           =   4095
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
         TabIndex        =   1
         Top             =   1080
         Width           =   4455
      End
   End
End
Attribute VB_Name = "frmAvisoFinPeriodoDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Declaración de propiedades del formulario
Public propMuestroAvisoExtension As Boolean     'establece si se muestra el aviso de solicitud de
                                                'de extensión del período de evaluación.


Private Sub Form_Load()
    'Cargo el texto que se muestra en las ventanas de avisos,
    'el cual no se establece por propiedades del objeto.
    lbl1Aviso.Caption = "El período de evaluación a TERMINADO, " & _
                        "por lo que no podrá seguir utilizando esta versión de evaluación." & Chr(10) & Chr(10) & _
                        "Si usted quedó conforme con las características del programa, " & _
                        "deberá adquirir la LICENCIA del mismo para poder utilizarlo sin límites de tiempo." & Chr(10) & _
                        "Además podrá acceder a futuras versiones a presios reducidos, entre otros veneficios."
                       
    If propMuestroAvisoExtension Then
        'muestro aviso de extensión del período de evaluación
        lbl2Aviso.Caption = "Si usted concidera que el período de evaluación no le fué suficiente, " & _
                        "póngase en contancto con nosotros para obtener una EXTENSIÓN " & _
                        "del período de evaluación."
    Else
        lbl2Aviso.Visible = False
    End If
End Sub

Private Sub botCerrar_Click()
    Me.Visible = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Evita que se cierre el formulario sin el conocimento del objeto
    'AvisoFinPeriodoDemo.
    'Si AvisoFinPeriodoDemo se muestra como un cuadro de diálogo modal, ocultar el
    'cuadro de diálogo en lugar de descargarlo permite que un método de la clase
    'AvisoFinPeriodoDemo recupere algún valor de este formulario.
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
        Me.Visible = False
    End If
End Sub

