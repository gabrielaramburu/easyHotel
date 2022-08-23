VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4995
   LinkTopic       =   "Form1"
   ScaleHeight     =   3255
   ScaleWidth      =   4995
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Muestro inicial"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents Inicial As ConfigurOinicial
Attribute Inicial.VB_VarHelpID = -1

Private Sub Command1_Click()
    If Inicial Is Nothing Then
        Set Inicial = New ConfigurOinicial
        Inicial.MostrarPantallaConfigurar _
                    "PerfilesUsuarios.txt", _
                    App.Path, _
                    "APLICACIÓN DE PRUEBA"

    End If
    Set Inicial = Nothing
End Sub
