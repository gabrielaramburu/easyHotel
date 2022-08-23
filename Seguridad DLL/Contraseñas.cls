VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Contraseñas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mtabla As Recordset
Private mtabla2 As Recordset

Private WithEvents dlgContra As IngContra
Attribute dlgContra.VB_VarHelpID = -1
Event NotificoCliente(boton As Byte)

Public Sub NuevoUsuarioAdmin(tabla As Recordset)
    'cambio titulo del formulario
    dlgContra.Caption = "Nuevo administrador"
    
    'cargo lista de usuarios
    Set mtabla = tabla
    With dlgContra
        .txtUsuario.Text = "Admin"  'El usuario es el administrador siempres se llama Admin        .txtUsuario.Locked = True   'bloqueo para que no se pueda modificar
        .txtUsuario.TabStop = False 'no permito desplazarme con tab
        .txtUsuario.BackColor = &H80000000  'Modifico el fondo del control
        .txtContraseñaNueva.TabIndex = 0
        .txtContraseñaConfirm.TabIndex = 1
    End With
    
    dlgContra.TipoAccionContra = 3
    dlgContra.Show 1
End Sub

Public Sub NuevoUsuario(tabla As Recordset)
    'Creo nuevo usuario
    
    'cambio titulo del formulario
    dlgContra.Caption = "Nuevo usuario"
    
    'cargo lista de usuarios
    Set mtabla = tabla
    
    'verifico si existen registros
    If tabla.RecordCount > 0 Then
        'existen registros
        tabla.MoveFirst
        'cargo lista de usuariso existentes
        Do While Not tabla.EOF
            dlgContra.lstUsuarios.AddItem tabla(0)
            tabla.MoveNext
        Loop
    End If
    
    dlgContra.TipoAccionContra = 1
    dlgContra.Show 1
    Set mtabla = Nothing
End Sub

Public Sub ModificoUsuario(tabla As Recordset, Optional Usr As String)
    'Modifico contraseña
    
    'cambio titulo del formulario
    dlgContra.Caption = "Modificar contraseña"
    
    'cargo lista de usuarios
    Set mtabla = tabla
    If tabla.RecordCount > 0 Then
        tabla.MoveFirst
        Do While Not tabla.EOF
            dlgContra.lstUsuarios.AddItem tabla(0)
            tabla.MoveNext
        Loop
    End If
    
    'antes de mostrar cargo usuario pasado como parámetro
    dlgContra.txtUsuario.Text = Usr
    
    dlgContra.TipoAccionContra = 2
    dlgContra.Show 1
End Sub

Public Sub EliminoUsuario(tablaUsr As Recordset, _
                        Optional tablaPer As Recordset, _
                        Optional Usr As String)
    'Elimino un usuario
    
    'cambio titulo del formulario
    dlgContra.Caption = "Eliminar usuario"
    
    'cargo lista de usuarios
    Set mtabla = tablaUsr
    Set mtabla2 = tablaPer
    If tablaUsr.RecordCount > 0 Then
        tablaUsr.MoveFirst
        Do While Not tablaUsr.EOF
            dlgContra.lstUsuarios.AddItem tablaUsr(0)
            tablaUsr.MoveNext
        Loop
    End If
    subMuestroControlesEliminar
    
    'antes de mostrar cargo usuario pasado como parámetro
    dlgContra.txtUsuario.Text = Usr
    dlgContra.Show 1
End Sub

Private Sub subMuestroControlesEliminar()
    With dlgContra
        'muestro y posiciono boton de eliminar
        .botEliminar.Visible = True
        .botEliminar.Top = 3120
        .botEliminar.Default = True
        
        'oculto resto de controles que no se usan
        .botAceptar.Visible = False
        .txtContraseñaConfirm.Visible = False
        .txtContraseñaNueva.Visible = False
        .Label2.Visible = False
        .Label3.Visible = False
    End With
End Sub

Private Sub dlgContra_NotificoCliente(boton As Byte)
    If boton = 1 Then   'aceptar
        If dlgContra.TipoAccionContra = 2 Then   'modifico
            'modifico contraseña
            mtabla.Index = "iclaves"
            mtabla.Seek "=", dlgContra.txtUsuario.Text
            If Not mtabla.NoMatch Then  'existe
                mtabla.Edit
                    mtabla(1) = dlgContra.txtContraseñaNueva.Text
                mtabla.Update
            End If
        End If
        If dlgContra.TipoAccionContra = 1 Then   'nuevo usuario
            'nuevo usuario
            mtabla.Index = "iclaves"
            mtabla.Seek "=", dlgContra.txtUsuario.Text
            If mtabla.NoMatch Then  'no existe
                mtabla.AddNew
                    mtabla(0) = dlgContra.txtUsuario.Text
                    mtabla(1) = dlgContra.txtContraseñaNueva.Text
                mtabla.Update
            End If
        End If
        If dlgContra.TipoAccionContra = 3 Then  'nuevo usuario admin
            mtabla.Edit
                mtabla("SisAdmin") = dlgContra.txtContraseñaNueva.Text
                'El campo mtabla("SisAdminTF") se inicializa a 1 en la aplicación
                'perfiles de usuario
            mtabla.Update
        End If
    End If
    
    If boton = 3 Then 'eliminar
        mtabla.Index = "iclaves"
        mtabla.Seek "=", dlgContra.txtUsuario.Text
        If Not mtabla.NoMatch Then  'existe
            'borro permisos
            subBorroPermisos
            mtabla.Delete
        End If
    End If
    RaiseEvent NotificoCliente(boton)
End Sub

Private Sub subBorroPermisos()
    'Borro los permisos del usuarios borrado
    mtabla2.Index = "i_NomUsr"
    mtabla2.Seek "=", dlgContra.txtUsuario.Text
    If Not mtabla2.NoMatch Then 'existe
        Do While Not mtabla2.EOF
            If mtabla2("NomUsr") = dlgContra.txtUsuario.Text Then
                mtabla2.Delete
                mtabla2.MoveNext
            Else
                Exit Do
            End If
        Loop
    End If
End Sub

Private Sub Class_Initialize()
    Set dlgContra = New IngContra
End Sub

Private Sub Class_Terminate()
    'Cuando se termina el objeto descargo el formulario
    'y  establesco su referencia a nothing para
    'evitar ocupar espacios de memoria huerfanos
    On Error Resume Next
    Unload IngContra
    Set IngContra = Nothing
    Set mtabla = Nothing
End Sub

