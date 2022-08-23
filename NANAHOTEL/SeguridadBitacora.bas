Attribute VB_Name = "SeguridadBitacora"
Option Explicit

'Esta variable es utilizada para ejecutar el procedimiento
'que graba las operaciones ejecutadas y se encuntra en bitacora.dll
Public ControlOperaciones As GraboOperacion

'Se iniciliazan cada vez que empiezo y termino de ejecutar
'una operación de la cual se lleve control en bitácora
Public HoraIni As String

'Se iniciliaza con el número de operación que se ejecuta
'de la cual se lleva control de bitacora.
Public OprEjecutada As Integer

Private MsgNoAcceso As mensaje

Public Function funUsuarioAutorizoSinMensaje(Usr As String, Opr As Integer)
    'Esta función se encarga de determinar si un usuario
    'está autorizado a ejecutra un determinada opción de dentro del
    'programa.
    'La diferencia con funUsuarioAutorizo es que esta no muestra mensaje de acceso denegado.
    If tbPARAMETROS("SisAdminTF") = 0 Then
        'Nunca definí perfiles de usuario, por ese motivo
        'no pido contraseña ninguna.
        funUsuarioAutorizoSinMensaje = True
    Else
        tbSISTEMA_PERFILES.Index = "pk_perfiles"
        tbSISTEMA_PERFILES.Seek "=", Opr, Usr
        If Not tbSISTEMA_PERFILES.NoMatch Then 'existe
            funUsuarioAutorizoSinMensaje = True
        Else
            funUsuarioAutorizoSinMensaje = False
        End If
    End If
End Function

Public Function funUsuarioAutorizo(Usr As String, Opr As Integer)
    'Esta función se encarga de determinar si un usuario
    'está autorizado a ejecutra un determinada opción de dentro del
    'programa.
    
    If tbPARAMETROS("SisAdminTF") = 0 Then
        'Nunca definí perfiles de usuario, por ese motivo
        'no pido contraseña ninguna.
        funUsuarioAutorizo = True
    Else
        tbSISTEMA_PERFILES.Index = "pk_perfiles"
        tbSISTEMA_PERFILES.Seek "=", Opr, Usr
        If Not tbSISTEMA_PERFILES.NoMatch Then 'existe
            funUsuarioAutorizo = True
        Else
           'If NoMostrarMsg <> True Then
                If MsgNoAcceso Is Nothing Then
                    Set MsgNoAcceso = New mensaje
                End If
                'Muestro mensaje de acceso denegado
                MsgNoAcceso.MensajeAccesoDenegado m_UsuarioSisNom
                funUsuarioAutorizo = False
            'End If
        End If
    End If
End Function

Public Sub GraboBitacora(Obs As String)
    'Es llamado después de realizar alguna de las operaciones
    'del sistema.
    'Los parámetros que se le pasan a la functión que se encuntra
    'implementada en bitacora.dll se inicializan antes de ejecutar la operación
    'epsecífica.
    'Después de ejecutada la misma, se llama a este procedimiento.
                'grabo en bitacora
                
    ControlOperaciones.GraboOperacionEnBaseDeDatos _
            m_FechaSis, _
            m_UsuarioSisNom, _
            OprEjecutada, _
            HoraIni, _
            Time, _
            Obs, _
            tbSISTEMA_BITACORA
End Sub
