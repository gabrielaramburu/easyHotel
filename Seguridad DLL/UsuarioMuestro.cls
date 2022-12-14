VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UsuarioMuestro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private WithEvents dlgIngMuestroUsuario As IngMuestroUsuario
Attribute dlgIngMuestroUsuario.VB_VarHelpID = -1
Event NotificoCliente(usuario As String, boton As Byte)

Public Sub MuestroUsuario(tabla As Recordset)
    'Utilizado para cambiar de usuario dentro de la aplicaci?n.
    
    'limpio lista usuarios
    dlgIngMuestroUsuario.lstUsuarios.Clear
    'limpio lista contrase?as
    dlgIngMuestroUsuario.lstContrase?a.Clear
    'limpio contrase?a
    dlgIngMuestroUsuario.txtContrase?a.Text = Empty
    If tabla.RecordCount > 0 Then
        tabla.MoveFirst
        Do While Not tabla.EOF
            dlgIngMuestroUsuario.lstUsuarios.AddItem tabla(0)
            dlgIngMuestroUsuario.lstContrase?a.AddItem tabla(1)
            tabla.MoveNext
        Loop
    End If
    dlgIngMuestroUsuario.botSalir.Visible = False
    dlgIngMuestroUsuario.Show 1
End Sub

Public Sub MuestroUsuarioStandBy(tabla As Recordset)
    'Brinda adem?s la posibilidad de dejar la aplicaci?n
    'en caso de no ser un usuario del sistema.
    
    'limpio lista usuarios
    dlgIngMuestroUsuario.lstUsuarios.Clear
    'limpio lista contrase?as
    dlgIngMuestroUsuario.lstContrase?a.Clear
    'limpio contrase?a
    dlgIngMuestroUsuario.txtContrase?a.Text = Empty
    If tabla.RecordCount > 0 Then
        tabla.MoveFirst
        Do While Not tabla.EOF
            dlgIngMuestroUsuario.lstUsuarios.AddItem tabla(0)
            dlgIngMuestroUsuario.lstContrase?a.AddItem tabla(1)
            tabla.MoveNext
        Loop
    End If
    dlgIngMuestroUsuario.botCancelar.Visible = False
    dlgIngMuestroUsuario.botAceptar.Left = 3600
    
    dlgIngMuestroUsuario.botSalir.Visible = True
    
    'cambio t?tulo formulario
    dlgIngMuestroUsuario.Caption = "En espera de usuario"
    dlgIngMuestroUsuario.Show 1
End Sub

Private Sub dlgIngMuestroUsuario_NotificoCliente(usuario As String, boton As Byte)
    RaiseEvent NotificoCliente(usuario, boton)
End Sub

Private Sub Class_Initialize()
    Set dlgIngMuestroUsuario = New IngMuestroUsuario
End Sub

Private Sub Class_Terminate()
    'Cuando se termina el objeto descargo el formulario
    'y  establesco su referencia a nothing para
    'evitar ocupar espacios de memoria huerfanos
    On Error Resume Next
    Unload IngMuestroUsuario
    Set IngMuestroUsuario = Nothing
End Sub

