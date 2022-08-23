Attribute VB_Name = "Formularios"
Option Explicit

Public Sub mSubConfiguroFuentesControlesSistema(f As Form)
    'Cambia la fuente de las caja de texto y etiquetas especiales
    'a el tipo y tamaño configurado por el usuario SISTEMA_FUENTES
    Dim objcontrol As Variant

    For Each objcontrol In f.Controls
        If TypeOf objcontrol Is TextBox Then
            objcontrol.Font.Name = mSisFuente_1GeneralTipo
            objcontrol.Font.Size = msisFuente_1GeneralTam
            'Si el tamaño ed letra es de 14 cambio el ancho de los formularios a 405
            If msisFuente_1GeneralTam = 14 Then
                'scrollBars = 2 son caja de observaciones y no se cambian
                If objcontrol.ScrollBars = 0 Then   'textbox comunes
                    objcontrol.Height = 405
                End If
            End If
        Else
            If TypeOf objcontrol Is ComboBox Then
                objcontrol.Font.Name = mSisFuente_1GeneralTipo
                objcontrol.Font.Size = msisFuente_1GeneralTam
            Else
                If TypeOf objcontrol Is Label Then
                    If objcontrol.BorderStyle = 1 Then  'solo lbl con borde
                        objcontrol.Font.Name = mSisFuente_1GeneralTipo
                        objcontrol.Font.Size = msisFuente_1GeneralTam
                        If msisFuente_1GeneralTam = 14 Then
                            objcontrol.Height = 405
                        End If
                    End If
                End If
            End If
        End If
    Next
End Sub

Public Sub mSubEtiquetasInicializo(f As Form)
    'Recorre las etiquetas de un formulario y modifica la propiedad
    'backcolor. Solo las etiquetas que simulan caja de texto.
    'Tambien borra el contenido de la etiqueta
    Dim objcontrol As Variant
    For Each objcontrol In f.Controls
        If TypeOf objcontrol Is Label Then
            If objcontrol.BorderStyle = 1 Then 'con borde
                objcontrol.BackColor = mSisColor_18ControlesNoHabilitados
                objcontrol.Caption = ""
            End If
        End If
    Next
End Sub

Public Sub mSub_bloqueo_controles_formulario(f As Form, x As Boolean)
    'Recorro todos los controles del formulario
    'y los bloqueo (Loked = true) o desbloqueo (Loked = false)
    'dependiendo del valor de x.
    'Tambien cambio el color de los mismos.
    'Establesco la propiedad tabStop a false para que no se puedan acceder mediante
    'el uso de tabs.
    
    On Error Resume Next
    Dim x2 As Boolean
    'true = bloqueo controles
    'false=desbloqueo controles
    Dim objcontrol As Variant
    Dim color As OLE_COLOR
    If x Then
        color = mSisColor_18ControlesNoHabilitados
        x2 = False
    Else
        color = mConstSisColor_Blanco
        x2 = True
    End If
        
    For Each objcontrol In f.Controls
        If TypeOf objcontrol Is TextBox Or _
        TypeOf objcontrol Is ComboBox Then
            objcontrol.Locked = x
            objcontrol.BackColor = color
            objcontrol.TabStop = x2
        End If
    Next
End Sub

Public Sub mSubBloqueoControlFormulario(c As control, x As Boolean)
    '---------------------------------------------------------------------------------
    'Cambia las propiedades del control que se pasa como parámetros, habilitando o
    'no el funcionamiento del control para el usuario.
    '---------------------------------------------------------------------------------
    'Parámetros.
    '   Entrada [c] control que deseo bloquear o desbloquear
    '           [x] tipo de acción a realizar
    '               true = bloqueo el control
    '               false = desbloqueo el control
    'NOTA:
    '   Se cambia la propiedad Loked
    '   Se cambio el color del fondo del control.
    '   Se cambio la propiedad TabStop
    '---------------------------------------------------------------------------------
    On Error Resume Next

    Dim color As OLE_COLOR
    Dim xAux As Boolean
    If x Then
        'bloqueo control
        color = mSisColor_18ControlesNoHabilitados
        xAux = False
    Else
        'desbloqueo control
        color = mConstSisColor_Blanco
        xAux = True
    End If
    
    c.Locked = x
    c.BackColor = color
    c.TabStop = xAux
End Sub

Public Sub mSub_limpio_controles_formulario(f As Form)
    'Limpio los controles de los formularios
    Dim objcontrol As Variant
    For Each objcontrol In f.Controls
        If TypeOf objcontrol Is TextBox Then
            objcontrol.Text = ""
        End If
        If TypeOf objcontrol Is ComboBox Then
            objcontrol.ListIndex = -1
        End If
        If TypeOf objcontrol Is Label Then
            If objcontrol.BorderStyle = 1 Then
                objcontrol.Caption = ""
            End If
        End If
    Next
End Sub

Public Sub ValidoNum(KeyAscii As Integer, Menos As Boolean, Punto As Boolean)
    Select Case KeyAscii
        Case 48 To 57       ' Permite los dígitos
        Case 8              ' Permite el carácter de retroceso
        Case 46 And Punto   ' Permite el punto
        Case 45 And Menos
        Case Else
            KeyAscii = 0
            Beep
    End Select
End Sub

Public Sub CapturoEnter(KeyAscii As Integer)
    'rutina que pasa el foco al control siguiente del que la llamo
    'en caso de haber apretado el ENTER
    
    If KeyAscii = 13 Then
        SendKeys (Chr(9))
    End If
End Sub

Public Sub mSubMuestroInformacionEnLinea(barra As gaHOTELbarra, tipo As Byte, codMsg As Integer)
    'Muestra información en la barra de tareas de los formularios.
    barra.Leyenda (mFunBuscoDescMsg(tipo, codMsg))
End Sub

Public Sub mSubLimpioInformacionEnLinea(barra As gaHOTELbarra)
    'No muestro nada en la barra de estado.
    'Cuando le doy el focus a un control muestro una descripción en la barra de tareas,
    'la cual sirve como asistencia para el usuario.
    'Cuando dejo el control (eveno lost focus) muestro una línea vacía para asegurarme
    'que no quede información incorrecta en la barra de tares.
    barra.Leyenda ("")
End Sub

