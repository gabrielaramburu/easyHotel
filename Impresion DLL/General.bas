Attribute VB_Name = "General"
Option Explicit

'NOTA: estos procedimientos se implementaron con el itento de pasar los procedimientos y
'funciones implementados en el módulo Impresión de la apliación EasyHotel a esta biblioteca,
'con la idea de poderlos reutilizar en las apliaciones perfiles y bitácora, como en cualquier
'otra futura aplicación a desarrollar. El problema fue que muchos de éstos procedimientos
'utilizan componentes de la aplicación que no puedo pasar como parámetros, como ser
'control crystal, comboBox, tipo App, etc.
'Por lo tanto este procedimiento y función no se utilizan en esta biblioteca.

Public Sub mSubMensaje(codMsg As Integer)
    '-----------------------------------------------------------------------------
    'Muestro un cuadro de díalogo al usuario.
    '-----------------------------------------------------------------------------
    'Parámetros.
    '       Entrada [codMsg]    código del mensaje a mostrar
    '
    'NOTA: este procedimiento se utiliza también en las aplicaciones generales, con
    'la diferencia de que en las mismas se trabaja con el archivo SISTEMA_MENSAJES
    '-----------------------------------------------------------------------------
    
    '16 icono crítico
    '32 pregunta de advertencia
    '48 mensaje de advertencia
    '64 mensaje de información
    Dim existeMsg As Boolean
    
    Dim tituloMsg As String
    Dim tipoSignoMsg As Byte
    Dim textoMsg As String
    
    'por defecto asumo que existe
    existeMsg = True
    Select Case codMsg
        Case 1
            tituloMsg = "No hay impresora instalada en el sistema. No se puede  continuar con la impresión."
            tipoSignoMsg = 48
            tituloMsg = "Error de impresión."
        Case Else
        existeMsg = False
    End Select
    'verifico si exuste mensaje
    If existeMsg Then
        MsgBox textoMsg & " ", tipoSignoMsg, tituloMsg
    End If
End Sub

Public Function mFunMensaje(codMsg As Integer) As Boolean
    '----------------------------------------------------------------------------------
    'Muestro un cuadro de díalogo al usuario.
    '----------------------------------------------------------------------------------
    'Parámetros.
    '         Entrada [codMsg]    código del mensaje a mostrar
    '         Salida True = el usuario presiona el boton de aceptar
    '                False =  si presiona el boton de cancelar.
    '-----------------------------------------------------------------------------------
    
    '16 icono crítico
    '32 pregunta de advertencia
    '48 mensaje de advertencia
    '64 mensaje de información
    Dim existeMsg As Boolean
    
    Dim tituloMsg As String
    Dim tipoSignoMsg As Byte
    Dim textoMsg As String
    
    'por defecto asumo que existe
    existeMsg = True
    Select Case codMsg
        Case 1
            tituloMsg = "¿Confirma la impresión del reporte?."
            tipoSignoMsg = 33
            tituloMsg = "¿Confirma impresión?"
        Case Else
        existeMsg = False
    End Select
    
    If existeMsg Then
        'si existe el mensaje, muestro un cuadro de diálogo.
        If MsgBox(textoMsg, tipoSignoMsg, tituloMsg) = vbOK Then
            'se presiono el boton de aceptar
            mFunMensaje = True
        Else
            'se presiono el boton de cancelar
            mFunMensaje = False
        End If
    End If
End Function


