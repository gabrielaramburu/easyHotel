VERSION 5.00
Begin VB.Form msgAccesoDenegado 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Acceso denegado"
   ClientHeight    =   1635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   5730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton botAceptar 
      Default         =   -1  'True
      Height          =   375
      Left            =   4440
      Picture         =   "msgAccesoDenegado.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "msgAccesoDenegado.frx":08B6
      Top             =   240
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Acceso denegado"
      Height          =   735
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   4695
   End
End
Attribute VB_Name = "msgAccesoDenegado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Event NotificoCliente(boton As Byte)

Private Sub botAceptar_Click()
    Me.Visible = False
    RaiseEvent NotificoCliente(1)   'boton aceptar
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'si se hace click en el boton cerrar, oculta el
    'cuadro de diálogo en lugar de descargarlo
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
        Me.Visible = False
    End If
End Sub


