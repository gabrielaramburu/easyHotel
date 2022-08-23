VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AvisoErrores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Declaración de propiedades
Public propMsgError As String
Public propDescMsgError As String
Public propContactarse As String

'Declaración de variavles
Private mFrmAvisoErrores As frmAvisoErrores

Private Sub Class_Initialize()
    'creo objeto formulario
    Set mFrmAvisoErrores = New frmAvisoErrores
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    'Descargo formulario
    Unload frmAvisoErrores
    Set frmAvisoErrores = Nothing
End Sub

Public Sub MostrarMensaje()
    'inicializo propiedades
    mFrmAvisoErrores.lblMsg = propMsgError
    mFrmAvisoErrores.lblDesMsg = propDescMsgError
    mFrmAvisoErrores.lblContactarse = propContactarse
    'muestro formulario con mensaje de error.
    mFrmAvisoErrores.Show 1
End Sub

