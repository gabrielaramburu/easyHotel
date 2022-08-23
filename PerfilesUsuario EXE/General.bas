Attribute VB_Name = "General"
Option Explicit

Public Function funExisteOtraInstancia() As Boolean
    'Determino si ya hay una instancia de la aplicación ejecutándose.
    Dim msg As String
    If App.PrevInstance Then
        msg = App.EXEName & ".EXE" & " ya está en ejecución"
        MsgBox msg, 16, "Aplicación."
        funExisteOtraInstancia = True
    Else
        'no existe ninguna instancia
        funExisteOtraInstancia = False
    End If
End Function

Public Sub mSubCargoComboUsr(cboUsr As comboBox, anexoTodas As Boolean)
    '------------------------------------------------------
    'Recorro archivo de usuarios y cargo en un combo
    '------------------------------------------------------
    'Parámetros.
    '   [cboUsr] combo de usr. que quiero inicializar
    '   [anexoTodas] si es True = anexo desc (Todos)
    '                si es Fasle = no anexo nada
    '-------------------------------------------------------
    tbSISTEMA_USUARIOS.Index = "iclaves"
    tbSISTEMA_USUARIOS.Seek ">=", ""
    If Not tbSISTEMA_USUARIOS.NoMatch Then
        'verifico si anexo desc.
        If anexoTodas Then
            cboUsr.AddItem "(Todos)"
        End If
        Do While Not tbSISTEMA_USUARIOS.EOF
            cboUsr.AddItem tbSISTEMA_USUARIOS("NomUsr")
            tbSISTEMA_USUARIOS.MoveNext
        Loop
    End If
    'por defecto posiciono en el primer elemento del combo
    If cboUsr.ListCount > 0 Then
        cboUsr.ListIndex = 0
    End If
End Sub

Public Sub mSubCargoComboOpr(comboOpr As comboBox, anexoTodas As Boolean, nivel As Byte)
    '-------------------------------------------------------
    'Recorro archivo de operaciones y las cargo en un combo
    '-------------------------------------------------------
    'Parámetros.
    '   [cboOpr] combo de opr. que quiero inicializar
    '   [anexoTodas] si es True = anexo desc (Todas)
    '                si es Fasle = no anexo nada
    '   [nivel]      1 operaciones de primer nivel
    '                2 operaciones de segundo nivel
    '-------------------------------------------------------
    tbSISTEMA_OPERACIONES.Index = "i_DescOpr"
    tbSISTEMA_OPERACIONES.Seek ">=", ""
    If Not tbSISTEMA_OPERACIONES.NoMatch Then
        'verifico si anexo desc.
        If anexoTodas Then
            comboOpr.AddItem "(Todas)"
        End If
        Do While Not tbSISTEMA_OPERACIONES.EOF
            'verifico nivel de la operación
            If tbSISTEMA_OPERACIONES("TipoOpr") = nivel Then
                comboOpr.AddItem tbSISTEMA_OPERACIONES("DescOpr")
                comboOpr.ItemData(comboOpr.NewIndex) = _
                tbSISTEMA_OPERACIONES("CodOpr")
            End If
            tbSISTEMA_OPERACIONES.MoveNext
        Loop
    End If
    'por defecto posiciono en el primer elemento del combo
    If comboOpr.ListCount > 0 Then
        comboOpr.ListIndex = 0
    End If
End Sub

Public Sub mSubMensaje(tipoMsg As Byte, codMsg As Integer, Optional descAux As String)
    'Muestro un cuadro de díalogo al usuario.
    'tipoMsg y codMsg: con estos datos se accede a un registro de la tabla SISTEMA_MENSAJES
    'deacuerdo a los valores de ese registro se muestra un determinado cuadro de diálogo.
    'descAux es utilizada para mensajes que tienen que mostrar datos extras en el mensaje,
    'como por ejemplo un número de recivo, etc.
    
    'Los mensajes de tipoMsg = 3 son los generales para todas las aplicaciones
    '                tipoMsg = 4 son los particulares de esta aplicación.
    
    '0 solo boton de aceptar
    '1 aceptar y cancelar
    
    '16 icono crítico
    '32 pregunta de advertencia
    '48 mensaje de advertencia
    '64 mensaje de información
    
    tbSISTEMA_MENSAJES.Index = "pk_msg"
    tbSISTEMA_MENSAJES.Seek "=", tipoMsg, codMsg
    If Not tbSISTEMA_MENSAJES.NoMatch Then
        'si existe el mensaje, muestro un cuadro de diálogo.
        MsgBox tbSISTEMA_MENSAJES("descMsg") & " " & descAux & " ", _
                tbSISTEMA_MENSAJES("estiloMsg"), _
                tbSISTEMA_MENSAJES("tituloMsg")
    End If
End Sub

Public Function mFunMensaje(tipoMsg As Byte, codMsg As Integer) As Boolean
    'Muestro un cuadro de díalogo al usuario.
    'tipoMsg y codMsg: con estos datos se accede a un registro de la tabla SISTEMA_MENSAJES
    'deacuerdo a los valores de ese registro se muestra un determinado cuadro de diálogo.
    'Retorno true si el usuario presiona el boton de aceptar y
    'false si presiona el boton de cncelar.
    tbSISTEMA_MENSAJES.Index = "pk_msg"
    tbSISTEMA_MENSAJES.Seek "=", tipoMsg, codMsg
    If Not tbSISTEMA_MENSAJES.NoMatch Then
        'si existe el mensaje, muestro un cuadro de diálogo.
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


