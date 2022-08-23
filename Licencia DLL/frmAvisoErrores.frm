VERSION 5.00
Begin VB.Form frmAvisoErrores 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Imposible ejecutar aplicación."
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton botAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   5040
      TabIndex        =   0
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label lblContactarse 
      Caption         =   "lblContactarse"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   6135
   End
   Begin VB.Image Image2 
      Height          =   210
      Left            =   120
      Picture         =   "frmAvisoErrores.frx":0000
      Stretch         =   -1  'True
      Top             =   960
      Width           =   6060
   End
   Begin VB.Label lblMsg 
      Caption         =   "lblMsg"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   960
      TabIndex        =   2
      Top             =   240
      Width           =   5175
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "frmAvisoErrores.frx":4BBA
      Top             =   360
      Width           =   480
   End
   Begin VB.Label lblDesMsg 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblDesMsg"
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   6135
   End
End
Attribute VB_Name = "frmAvisoErrores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub botAceptar_Click()
    Me.Visible = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Evita que se cierre el formulario sin el conocimento del objeto
    'AvisoErrores
    'Si frmAvisoErrores se muestra como un cuadro de diálogo modal, ocultar el
    'cuadro de diálogo en lugar de descargarlo permite que un método de la clase
    'AvisoErrores recupere algún valor de este formulario.
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
        Me.Visible = False
    End If
End Sub


