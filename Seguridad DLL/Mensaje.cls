VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Mensaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private WithEvents msgAcceso As msgAccesoDenegado
Attribute msgAcceso.VB_VarHelpID = -1
Event NotificoCliente(boton As Byte)

Public Sub MensajeAccesoDenegado(Usr As String)
    msgAcceso.Label1.Caption = _
    "Acceso denegado." & Chr(10) & _
    "El usuario " & Usr & " no está autorizado a ejecutar esta opción." & Chr(10) & _
    "Consulte con el administrador del sistema."
    msgAcceso.Show 1
End Sub

Private Sub Class_Initialize()
    Set msgAcceso = New msgAccesoDenegado
End Sub

Private Sub Class_Terminate()
    'Cuando se termina el objeto descargo el formulario
    'y  establesco su referencia a nothing para
    'evitar ocupar espacios de memoria huerfanos
    On Error Resume Next
    Unload msgAccesoDenegado
    Set msgAccesoDenegado = Nothing
End Sub

Private Sub msgAcceso_NotificoCliente(boton As Byte)
    RaiseEvent NotificoCliente(boton)
End Sub
