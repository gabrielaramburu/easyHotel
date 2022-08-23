VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "DLL impresión"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   6615
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton botSeleccionImpresoras 
      Caption         =   "&Selección de impresoras del sistema"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   4455
   End
   Begin VB.CommandButton botAceptar 
      Cancel          =   -1  'True
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   5040
      TabIndex        =   0
      Top             =   2760
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Creo variables para acceder a las bibliotecas
Private biblioImpresion As SeleccionImpre

Private Sub botAceptar_Click()
    Unalod Me
End Sub

Private Sub botSeleccionImpresoras_Click()
    Dim impresora As String
    Set biblioImpresion = New SeleccionImpre
        impresora = biblioImpresion.mFunSeleccionoImpresora(Printer.DeviceName)
        If impresora = "" Then
            MsgBox "No se seleccionó ninguna impresora"
        Else
            MsgBox "Se seleccionó la impresora " & impresora, vbExclamation
        End If
    Set biblioImpresion = Nothing
End Sub
