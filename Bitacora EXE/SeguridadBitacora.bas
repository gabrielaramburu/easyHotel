Attribute VB_Name = "SeguridadBitacora"
Option Explicit
Private MsgNoAcceso As Mensaje

Public Function mfunUsuarioAutorizo(Usr As String, Opr As Integer)
    'Esta función se encarga de determinar si un usuario
    'está autorizado a ejecutra un determinada opción de dentro del
    'programa.
    
    If tbSISTEMA_PARAMETROS("SisAdminTF") = 0 Then
        'Nunca definí perfiles de usuario, por ese motivo
        'no pido contraseña ninguna.
        mfunUsuarioAutorizo = True
    Else
        tbSISTEMA_PERFILES.Index = "pk_perfiles"
        tbSISTEMA_PERFILES.Seek "=", Opr, Usr
        If Not tbSISTEMA_PERFILES.NoMatch Then 'existe
            mfunUsuarioAutorizo = True
        Else
            If MsgNoAcceso Is Nothing Then
                Set MsgNoAcceso = New Mensaje
            End If
            'Muestro mensaje de acceso denegado
            MsgNoAcceso.MensajeAccesoDenegado m_UsuarioSisNom
            mfunUsuarioAutorizo = False
        End If
    End If
End Function

