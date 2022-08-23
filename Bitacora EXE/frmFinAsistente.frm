VERSION 5.00
Begin VB.Form frmFinAsistente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fin del asistente de creación de listados."
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   6585
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton botAceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   5280
      TabIndex        =   1
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label lblAviso 
      Caption         =   "lblAviso"
      Height          =   855
      Left            =   1920
      TabIndex        =   0
      Top             =   240
      Width           =   3735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1380
      Left            =   240
      Picture         =   "frmFinAsistente.frx":0000
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "frmFinAsistente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub botAceptar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.lblAviso.Caption = _
    "Uds. ha creado con éxito un nuevo listado," & _
    "para ejecutarlo diríjase al cuadro de listados."
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Descargo formulario de memoria
    Set frmFinAsistente = Nothing
End Sub
