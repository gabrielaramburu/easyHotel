Attribute VB_Name = "SeguridadBitacora"
Option Explicit
Private MsgNoAcceso As Mensaje

Public Function mfunUsuarioAutorizo(Usr As String, Opr As Integer)
    'Esta funci�n se encarga de determinar si un usuario
    'est� autorizado a ejecutra un determinada opci�n de dentro del
    'programa.
    
    If tbSISTEMA_PARAMETROS("SisAdminTF") = 0 Then
        'Nunca defin� perfiles de usuario, por ese motivo
        'no pido contrase�a ninguna.
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

