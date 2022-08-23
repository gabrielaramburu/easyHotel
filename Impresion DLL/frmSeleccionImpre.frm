VERSION 5.00
Begin VB.Form frmSeleccionImpre 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Selección de impresora"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5910
   Icon            =   "frmSeleccionImpre.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Seleccione impresora a utilizar"
      Height          =   2535
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   5655
      Begin VB.ComboBox cboImpreSistema 
         Height          =   360
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   720
         Width           =   4335
      End
      Begin VB.Label lblAviso 
         Caption         =   "lblAviso"
         Height          =   735
         Left            =   360
         TabIndex        =   5
         Top             =   1440
         Visible         =   0   'False
         Width           =   5055
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   4800
         Picture         =   "frmSeleccionImpre.frx":27A2
         Top             =   660
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&Impresoras del sistema"
         Height          =   240
         Left            =   240
         TabIndex        =   0
         Top             =   480
         Width           =   2100
      End
   End
   Begin VB.CommandButton botCancelar 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   4560
      Picture         =   "frmSeleccionImpre.frx":4F44
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton botAceptar 
      Default         =   -1  'True
      Height          =   375
      Left            =   3240
      Picture         =   "frmSeleccionImpre.frx":5806
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2760
      Width           =   1215
   End
End
Attribute VB_Name = "frmSeleccionImpre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Declaro eventos del formulario
Event notificoCliente(boton As Byte)    'Este evento se desencadena cuando se presiona
                                        'un boton del formulario
                                        '0 = boton cancelar
                                        '1 = boton aceptar

Private Sub botAceptar_Click()
    'oculto formulario
    Me.Visible = False
    'desencadeno evento aceptar
    RaiseEvent notificoCliente(1)
End Sub

Private Sub botCancelar_Click()
    'oculto formulario
    Me.Visible = False
    'desencadeno evento cancelar
    RaiseEvent notificoCliente(0)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Si se hace click en el boton cerrar, oculta el
    'cuadro de diálogo en lugar de descargarlo
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
        Me.Visible = False
    End If
End Sub
