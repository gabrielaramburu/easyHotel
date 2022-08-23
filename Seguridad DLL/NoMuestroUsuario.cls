VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NoMuestroUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private WithEvents dlgNoMuestro As IngSinUsuario
Attribute dlgNoMuestro.VB_VarHelpID = -1
Event NotificoCliente(usuario As String, boton As Byte)

Public Sub MuestroSinUsuario(tabla As Recordset)
    'limpio lista usuarios
    dlgNoMuestro.lstUsuarios.Clear
    'limpio lista contraseñas
    dlgNoMuestro.lstContraseña.Clear
    'limpio contraseña
    dlgNoMuestro.txtContraseña.Text = Empty
    If tabla.RecordCount > 0 Then
        tabla.MoveFirst
        Do While Not tabla.EOF
            dlgNoMuestro.lstUsuarios.AddItem tabla(0)
            dlgNoMuestro.lstContraseña.AddItem tabla(1)
            tabla.MoveNext
        Loop
    End If
    dlgNoMuestro.Show 1
End Sub

Public Sub MuestroAdmin(contraAdmin As String)
    'Solo permito ingresar al usuario administrador
    
    dlgNoMuestro.txtUsuario.Text = "Admin"
    
    'no permito ingresar usuario
    dlgNoMuestro.txtUsuario.Locked = True
    dlgNoMuestro.txtUsuario.BackColor = &H80000000
    'por defecto quedo pidiendo contraseña
    dlgNoMuestro.txtContraseña.TabIndex = 0
    
    dlgNoMuestro.lstUsuarios.AddItem "Admin"
    dlgNoMuestro.lstContraseña.AddItem contraAdmin
    dlgNoMuestro.Show 1
End Sub

Private Sub dlgNoMuestro_NotificoCliente(usuario As String, boton As Byte)
    RaiseEvent NotificoCliente(usuario, boton)
End Sub

Private Sub Class_Initialize()
    Set dlgNoMuestro = New IngSinUsuario
End Sub

Private Sub Class_Terminate()
    'Cuando se termina el objeto descargo el formulario
    'y  establesco su referencia a nothing para
    'evitar ocupar espacios de memoria huerfanos
    On Error Resume Next
    Unload IngSinUsuario
    Set IngSinUsuario = Nothing
End Sub

