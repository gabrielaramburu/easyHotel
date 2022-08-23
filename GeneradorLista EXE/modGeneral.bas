Attribute VB_Name = "modGeneral"
Public Sub mSubMensaje(tipoMsg As Byte, codMsg As Integer, Optional descAux As String)
    'Muestro un cuadro de d�alogo al usuario.
    'tipoMsg y codMsg: con estos datos se accede a un registro de la tabla SISTEMA_MENSAJES
    'deacuerdo a los valores de ese registro se muestra un determinado cuadro de di�logo.
    'descAux es utilizada para mensajes que tienen que mostrar datos extras en el mensaje,
    'como por ejemplo un n�mero de recivo, etc.
    
    'Los mensajes de tipoMsg = 3 son los generales para todas las aplicaciones
    '                tipoMsg = 4 son los particulares de esta aplicaci�n.
    
    '0 solo boton de aceptar
    '1 aceptar y cancelar
    
    '16 icono cr�tico
    '32 pregunta de advertencia
    '48 mensaje de advertencia
    '64 mensaje de informaci�n
    
    tbSISTEMA_MENSAJES.Index = "pk_msg"
    tbSISTEMA_MENSAJES.Seek "=", tipoMsg, codMsg
    If Not tbSISTEMA_MENSAJES.NoMatch Then
        'si existe el mensaje, muestro un cuadro de di�logo.
        MsgBox tbSISTEMA_MENSAJES("descMsg") & " " & descAux & " ", _
                tbSISTEMA_MENSAJES("estiloMsg"), _
                tbSISTEMA_MENSAJES("tituloMsg")
    End If
End Sub

Public Function mFunMensaje(tipoMsg As Byte, codMsg As Integer) As Boolean
    'Muestro un cuadro de d�alogo al usuario.
    'tipoMsg y codMsg: con estos datos se accede a un registro de la tabla SISTEMA_MENSAJES
    'deacuerdo a los valores de ese registro se muestra un determinado cuadro de di�logo.
    'Retorno true si el usuario presiona el boton de aceptar y
    'false si presiona el boton de cncelar.
    tbSISTEMA_MENSAJES.Index = "pk_msg"
    tbSISTEMA_MENSAJES.Seek "=", tipoMsg, codMsg
    If Not tbSISTEMA_MENSAJES.NoMatch Then
        'si existe el mensaje, muestro un cuadro de di�logo.
        If MsgBox(tbSISTEMA_MENSAJES("descMsg"), _
                tbSISTEMA_MENSAJES("estiloMsg"), _
                tbSISTEMA_MENSAJES("tituloMsg")) = vbOK Then
                'se presiono el boton de aceptar
                mFunMensaje = True
        Else
            'se presiono el boton de cancelar
            mFunMensaje = False
        End If
    End If
End Function


