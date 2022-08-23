Attribute VB_Name = "General"
Option Explicit
'Estas variables son utilizadas por las funciones de conversión a
'letras.
Dim Unidades$(9), Decenas$(9), Oncenas$(9)
Dim Veintes$(9), Centenas$(9)

Sub Main()
    'Procedimiento de inicio.
    'Fue creado con la idea de mostrar un formulario de presentación.
    frmInicioSesion.Show
    'si no realizo esta espera no se muestra el formulario de inicio
    mSubEspera 0.1
    
    terminarEjecucion = False
    'cargo el formulario
    Load frmMAIN
    'evealúo si continúo con la ejecución
    If Not terminarEjecucion Then
        'oculto formulario de inicio sesión
        Unload frmInicioSesion
        'muestro el formulario
        frmMAIN.Show
        
    Else
        'descatgo el formulario
        Unload frmMAIN
    End If
End Sub

Public Function comparo_fecha_mmyyyy_con_actual(mm As Integer, yyyy As Integer)
    Dim anio_actual As Integer
    Dim mes_actual As Integer
    
    comparo_fecha_mmyyyy_con_actual = True
    anio_actual = Year(m_FechaSis)
    mes_actual = Month(m_FechaSis)
       
    'si estan bien los parametros
    If mm > 12 Or mm < 1 Then
        comparo_fecha_mmyyyy_con_actual = False
        Exit Function
    End If
    If yyyy < 1 Then
        comparo_fecha_mmyyyy_con_actual = False
        Exit Function
    End If
    
    If yyyy < anio_actual Then
        comparo_fecha_mmyyyy_con_actual = False
    Else
        If yyyy = anio_actual Then
            If mm < mes_actual Then
                comparo_fecha_mmyyyy_con_actual = False
            End If
        End If
    End If
End Function

Public Function formo_fecha(fecha As String)
    Dim fecha_aux As String
    Dim dia As String, caracter As String
    Dim cursor As Byte
    If Val(fecha) <> 0 Then
        cursor = 1
        Do While cursor <= Len(fecha)
            caracter = Mid(fecha, cursor, 1)
            If caracter <> "/" Then
                fecha_aux = fecha_aux + caracter
            End If
            cursor = cursor + 1
        Loop
        fecha_aux = Mid(fecha_aux, 1, 2) & "/" & Mid(fecha_aux, 3, 2) & "/" & Mid(fecha_aux, 5, 4)
        formo_fecha = fecha_aux
    End If
End Function

Public Function corto_palabras(frase As String)
    'recorre una cadena y se queda con la primera palabra comenzado de la izquierda
    Dim aux As String, cursor As String, caracter As String
    cursor = 1
    frase = LTrim(frase)
    Do While cursor <= Len(frase)
        caracter = Mid(frase, cursor, 1)
        If caracter <> " " Then
            aux = aux & caracter
        Else
            Exit Do
        End If
        cursor = cursor + 1
    Loop
    corto_palabras = aux
End Function
Public Function corto_strMedio(frase As String, car As String)
    'Dado un string obtengo el string que aparece después del primer espacio en blanco
    Dim aux As String, cursor As String, caracter As String
    Dim proceso As Byte
    proceso = 0
    cursor = 1
    Do While cursor <= Len(frase)
        caracter = Mid(frase, cursor, 1)
        If caracter = car Then
            proceso = proceso + 1
        End If
        If proceso = 1 Then
            aux = aux & caracter
        End If
        cursor = cursor + 1
    Loop
    corto_strMedio = aux
End Function

Public Function corto_strDer(frase As String, car As String)
    Dim aux As String, cursor As Integer, caracter As String
    Dim proceso As Boolean
    proceso = False
    cursor = 1
    Do While cursor <= Len(frase)
        caracter = Mid(frase, cursor, 1)
        If proceso Then
            aux = aux & caracter
        End If
        If caracter = car Then
            proceso = True
        End If
        cursor = cursor + 1
    Loop
    corto_strDer = aux
End Function

Public Function corto_strIzq(frase As String, car As String)
    Dim aux As String, cursor As Integer, caracter As String
    Dim proceso As Boolean
    cursor = 1
    Do While cursor <= Len(frase)
        caracter = Mid(frase, cursor, 1)
        If caracter <> car Then
            aux = aux & caracter
        Else
            Exit Do
        End If
        cursor = cursor + 1
    Loop
    corto_strIzq = aux
End Function

Public Function NroResFormato(nroRes As Variant)
    'Da formato al número de reserva
    'Si se envía con guión se lo saca
    'Si se envía sin guión se lo pone
    If Len(nroRes) > 9 Then                                'saco guión
        NroResFormato = Mid(nroRes, 1, 5) & Mid(nroRes, 7, 10)
    Else                        'pongo guión
        NroResFormato = Mid(nroRes, 1, 4) & "-" & Mid(nroRes, 5, 10)
    End If
End Function

Public Function busco_titular_hab(hab As Long, tipo As String)
    'Devuelve el nombre de titular de una habitación
    Dim titular As Long
    busco_titular_hab = ""
    If busco_habitaTF(hab) Then
        Select Case tipo
            Case "aloja"
                If tbHABITACIONES("titular_aloja") <> 0 Then
                    titular = tbHABITACIONES("titular_aloja")
                Else
                    titular = tbHABITACIONES("titular_unica")
                End If
            Case "extra"
                If tbHABITACIONES("titular_extra") <> 0 Then
                    titular = tbHABITACIONES("titular_extra")
                Else
                    titular = tbHABITACIONES("titular_unica")
                End If
            Case "unica"
                    titular = tbHABITACIONES("titular_unica")
        End Select
        If busco_clienteTF(titular) Then
            busco_titular_hab = tbCLIENTES("nombre_completo_titular")
        End If
    End If
End Function

Public Function busco_titular_hab2(hab As Long, tipo As String)
    'Devuelve el número de titular de una habitación.
    'tipo: determina el tipo de titular que quiero obtener.
    
    busco_titular_hab2 = 0
    If busco_habitaTF(hab) Then
        Select Case tipo
            Case "aloja"
                If tbHABITACIONES("titular_aloja") <> 0 Then
                    busco_titular_hab2 = tbHABITACIONES("titular_aloja")
                Else
                    busco_titular_hab2 = tbHABITACIONES("titular_unica")
                End If
            Case "extra"
                If tbHABITACIONES("titular_extra") <> 0 Then
                    busco_titular_hab2 = tbHABITACIONES("titular_extra")
                Else
                    busco_titular_hab2 = tbHABITACIONES("titular_unica")
                End If
            Case "unica"
                    busco_titular_hab2 = tbHABITACIONES("titular_unica")
        End Select
    End If
End Function

Public Function busco_titular_hab2SinCambiarPunteroHab(hab As Long, tipo As String)
    '----------------------------------------------------------------------------
    'Obtengo número de titular de una habitación, dependiendo del tipo del mismo.
    '
    'Estoy teniendo problemas con el trabajo de variables de tipo tabla de nivel
    'general. Un caso concreto: detecté que la función busco_titular_hab2
    'se posiciona en un registro de forma directa lo que provoca que el procedimiento
    'proceso_alojamiento que recorre el mismo archivo pero en forma secuencial pierda
    'la referencia al registro de trabajo luego de la llamada a esta función.
    'Para solucionarlo creo esta nueva función.
    '-------------------------------------------------------------------------------
    'Paraámetros.
    '-------------------------------------------------------------------------------
    '   Entrada :
    '               [hab]    habitación de la cual quiero obtener número titular
    '               [tipo]   tipo de titular que quiero obtener
    '
    '   Salida:     número de titular.
    '--------------------------------------------------------------------------------
    'declaro variables para trabajar con archivo de habitaciones
    Dim tbHab As Recordset
    Set tbHab = tbHABITACIONES
    'busco habitación
    tbHab.Index = "inrohab"
    tbHab.Seek "=", hab
    If Not tbHab.NoMatch Then
        busco_titular_hab2SinCambiarPunteroHab = 0
        Select Case tipo
            Case "aloja"
                If tbHab("titular_aloja") <> 0 Then
                    busco_titular_hab2SinCambiarPunteroHab = tbHab("titular_aloja")
                Else
                    busco_titular_hab2SinCambiarPunteroHab = tbHab("titular_unica")
                End If
            Case "extra"
                If tbHab("titular_extra") <> 0 Then
                    busco_titular_hab2SinCambiarPunteroHab = tbHab("titular_extra")
                Else
                    busco_titular_hab2SinCambiarPunteroHab = tbHab("titular_unica")
                End If
            Case "unica"
                    busco_titular_hab2SinCambiarPunteroHab = tbHab("titular_unica")
        End Select
    End If
    Set tbHab = Nothing
End Function

Public Function mFunBuscoDescripcionTipoTitular(habTit As Long, TipoTit As String) As String
    '------------------------------------------------------------------------------------
    'Esta función es encarga de obtener el tipo de titular que tiene una habitación,
    'discriminado por tipo de cuenta.
    'A diferencia de las funciones anteriores, ésta no devuelve ni el número, ni el nombre
    'sinó que devuelve un string indicando si el titular es de tipo único o solo de gastos
    'extras o gastos alojamiento.
    '------------------------------------------------------------------------------------
    'Parámetros.
    '   Entrada [habTit]   habitación de la cual nos interesa saber el titular
    '           [tipoTit]   indica el tipo de cuenta con la que se esta trabajando
    '               "aloja" cuenta de gastos de alojamiento
    '               "extra" cuenta de gastos extras
    '
    'NOTA: no se trabaja con tipoTit "unica" ya que no tiene sentido.
    '------------------------------------------------------------------------------------
    'declaro variables para trabajar con tabla de habitaciones
    Dim tbHab As Recordset
    Set tbHab = tbHABITACIONES
    
    'busco habitación
    tbHab.Index = "inrohab"
    tbHab.Seek "=", habTit
    If Not tbHab.NoMatch Then
        'existe habitación
        Select Case TipoTit
            Case "aloja"
                If tbHab("titular_aloja") <> 0 Then
                    mFunBuscoDescripcionTipoTitular = "gastos alojamiento"
                Else
                    mFunBuscoDescripcionTipoTitular = "unico"
                End If
            Case "extra"
                If tbHab("titular_extra") <> 0 Then
                    mFunBuscoDescripcionTipoTitular = "gastos extras"
                Else
                    mFunBuscoDescripcionTipoTitular = "unico"
                End If
        End Select
    End If
    Set tbHab = Nothing
End Function

Public Sub cambio_situacion(hab As Long, situ As Byte)
    If busco_habitaTF(hab) Then
        tbHABITACIONES.Edit
            tbHABITACIONES("situacionhab") = situ
            tbHABITACIONES("fechasituacionhab") = m_FechaSis
        tbHABITACIONES.Update
    End If
End Sub

Public Sub inicializo_habitacion(hab As Long)
    If busco_habitaTF(hab) Then
        tbHABITACIONES.Edit
            tbHABITACIONES("tipocuenta_unica") = 0
            tbHABITACIONES("tipocuenta_aloja") = 0
            tbHABITACIONES("tipocuenta_extra") = 0
            tbHABITACIONES("titular_unica") = 0
            tbHABITACIONES("titular_aloja") = 0
            tbHABITACIONES("titular_extra") = 0
            tbHABITACIONES("tarifa") = 0
        tbHABITACIONES.Update
    End If
End Sub

Public Function paso_moneda_a_codigo(desc_moneda As String)
    If Trim(desc_moneda) = gblSignoMonedaNacional Then paso_moneda_a_codigo = 0
    If Trim(desc_moneda) = gblSignoDolares Then paso_moneda_a_codigo = 1
End Function

Public Function paso_moneda_a_desc(cod_moneda As Byte)
    If cod_moneda = 0 Then paso_moneda_a_desc = gblSignoMonedaNacional
    If cod_moneda = 1 Then paso_moneda_a_desc = gblSignoDolares
End Function

Public Sub mSub_Cargo_Fecha_Sistema()
    'Cargo fecha del sistema
    'Llamado desde cierre diario
    
    m_FechaSis = tbPARAMETROS("fecha_ultimo_cierre_realizado")
End Sub

Public Function mFunObtengoFechaSistema() As Boolean
    'Es llamado al iniciar la ejecución de la aplicación (Main)
    '------------------------------------------------------------------------------------------------
    'Parámetos.
    '   Salida: True, existe un valor fecha en el archivo parámetros
    '           True, no existe un valor fecha en el archivo parámetros ya que es la primera
    '           vez que ejecuto la aplicación, pero el usuario confirma la nueva fecha del sistema
    '
    '           False,no existe un valor fecha en el archivo parámetros ya que es la primera
    '           vez que ejecuto la aplicación, y el usuario No confirma la nueva fecha del sistema
    '-------------------------------------------------------------------------------------------------
    If IsDate(tbPARAMETROS("fecha_ultimo_cierre_realizado")) Then
        m_FechaSis = tbPARAMETROS("fecha_ultimo_cierre_realizado")
        'tengo fecha del sistema
        mFunObtengoFechaSistema = True
    Else
        'si no es un valor fecha es porque el campo todavía no esta inicializado
        'es decir es la primera vez que se ejecuta la aplicación.
        
        'inicializo valor del campo en tabla parámetros
        If mFunInicializoFechaSistema Then
            'elusuario confirmó la nueva fecha
            m_FechaSis = tbPARAMETROS("fecha_ultimo_cierre_realizado")
            'tengo fecha del sistema
            mFunObtengoFechaSistema = True
        Else
            'el usuario NO confirmó la nueva fecha
            
            'NO tengo fecha del sistema
            mFunObtengoFechaSistema = False
        End If
    End If
End Function

Public Sub mSub_Inicializo_fuentes_sistema()
    'Recorro el archivo de fuentes y asigno el valor correspondiente
    'a cada constante de fuente establecida
    
    Dim tipo As String
    Dim tam As Byte
    tbSIS_FUENTES.MoveFirst
    Do While Not tbSIS_FUENTES.EOF
        tipo = tbSIS_FUENTES("tipoapafuente")
        tam = tbSIS_FUENTES("tamapafuente")
        Select Case tbSIS_FUENTES("codapafuente")
            Case 1
                mSisFuente_1GeneralTipo = tipo
                msisFuente_1GeneralTam = tam
        End Select
        tbSIS_FUENTES.MoveNext
    Loop
End Sub

Public Sub mSub_Inicialixo_colores_sistema()
    'Recorro el archivo de colores y le asigno el valor correspondiente
    'a cada constante de color
    
    Dim color As OLE_COLOR
    tbSIS_COLORES.MoveFirst
    Do While Not tbSIS_COLORES.EOF
        color = tbSIS_COLORES("colorapa")
        Select Case tbSIS_COLORES("codapa")
            Case 1
                mSisColor_1DetalleDeGastos = color
            Case 2
                mSisColor_2TotalDeGastosDiarios = color
            Case 3
                mSisColor_3TotalDeGastosTitular = color
            Case 6
                mSisColor_6SaldoMonedaNacional = color
            Case 7
                mSisColor_7SaldoDolares = color
            Case 10
                mSisColor_10CheckinSeleccionHab = color
            Case 11
                mSisColor_11SeleccionHabLibre = color
            Case 12
                mSisColor_12SeleccionHabOcupada = color
            Case 15
                mSisColor_15FilaSeleccionada = color
            Case 18
                mSisColor_18ControlesNoHabilitados = color
            Case 19
                mSisColor_19FilaSeleccionadaTexto = color
        End Select
        tbSIS_COLORES.MoveNext
    Loop
End Sub

Private Function mSubMuestro_leyenda_barra(b As StatusBar, codigo As Integer)
    'Debuelve el texto que se mostrará en el primer panel de la barra de estado
    Dim Leyenda As String
    Select Case codigo
        Case 1
            Leyenda = "Hola como te va"
        Case 2
        Case 3
        Case 4
    End Select
    b.Panels(1).Text = Leyenda
End Function

Public Function mFunObtengoEstadoHab(hab As Long)
    'Devuelve el estado de una habitación
    'Es muy importante el orden en que llamo a estas funciones.
    'Ejemplo: si llamo primero a la función de control de habitaciones reservadas
    'la consulta no funciona bien, ya que este estado no es determinante, es decir,
    'la habitación puede estar ocupada, ya que la reserva se hizo efectiva.
    'En conclusión podemos decir que existen estados determinnates y no determinanates.
    'Estados determinantes:
    '   Ocupada
    '   Bloqueada
    '   Libre
    'Estado no determinante:
    '   Reservada
    
    If busco_habita_checkin(hab) Then
        mFunObtengoEstadoHab = "Ocupada"
    Else
        If habitacion_reservada(hab, m_FechaSis, m_FechaSis) Then
            mFunObtengoEstadoHab = "Reservada"
        Else
            If habitacion_bloqueada(hab, m_FechaSis, m_FechaSis) Then
                 mFunObtengoEstadoHab = "Bloqueada"
            Else
                mFunObtengoEstadoHab = "Libre"
            End If
        End If
    End If
End Function

Public Function mFunDeterminoOcupacionValida(habOcupada As Long) As Boolean
    'Determina si la ocupación de una habitación este dentro del período establecido
    '--------------------------------------------------------------------------------
    'Parámetros.
    '   Entrada [habOcupada] Habitación del hotel actualmente ocupada
    '
    '   Salida  True, si el período de ocupación esta dentro de lo previsto
    '           False, el período de ocupación llegó a su fin y no se ha realizado
    '                checkout a la habitación
    '----------------------------------------------------------------------------------
    'por defecto asumo que la ocupación es correcta
    mFunDeterminoOcupacionValida = True
    'busco datos del alojamiento
    If busco_habita_checkin(habOcupada) Then
        'verifico si esta dentro del período de ocupación
        If tbCHECKIN("fCheckHas") < m_FechaSis Then
            'la habitación devió de ser dejada libre (checkout)
            mFunDeterminoOcupacionValida = False
        End If
    End If
End Function

Public Sub msubPosicionoListasAlPrincipio(lista As ListBox)
    'Doy el focus al primer elemento de un listbox
    If lista.ListCount >= 1 Then
        lista.ListIndex = 0
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

Public Function funExisteOtraInstancia() As Boolean
    'Determino si ya hay una instancia de la aplicación ejecutándose.
    Dim msg As String
    If App.PrevInstance Then
        msg = App.EXEName & ".EXE" & " ya está en ejecución"
        MsgBox msg, 16, "Aplicación."
        funExisteOtraInstancia = True
        funExisteOtraInstancia = False
    Else
        'no existe ninguna instancia
        funExisteOtraInstancia = False
    End If
End Function

Public Sub mSubEspera(Segundos As Single)
    'Produce una pausa
    Dim ComienzoSeg As Single
    Dim FinSeg As Single
    ComienzoSeg = Timer
    FinSeg = ComienzoSeg + Segundos
    Do While FinSeg > Timer
        DoEvents
        If ComienzoSeg > Timer Then
            FinSeg = FinSeg - 24 * 60 * 60
        End If
    Loop
End Sub

Public Function mFunObtengoTotHabHotel() As Integer
    'Devuelve el total de habitaciones del hotel
    '----------------------------------------------------------
    'Parámetros.
    '   Salida: total de registros del archivo tbHABITACIONES
    '----------------------------------------------------------
    Dim consulta As String
    Dim rstHab As Recordset
    Dim qdfHab As QueryDef
    consulta = "select * from habitaciones"
    'ejecuto consulta
    Set qdfHab = bdHOTEL.CreateQueryDef("")
    qdfHab.SQL = consulta
    Set rstHab = qdfHab.OpenRecordset(dbOpenSnapshot)
    'me muevo al último registro para poder utilizar la propiedad recordcount
    rstHab.MoveLast
    mFunObtengoTotHabHotel = rstHab.RecordCount
    
    Set qdfHab = Nothing
    Set rstHab = Nothing
End Function

Public Function mFunObtengoUltimaCotizacion(tipoSalida As Byte, codMoneda As Byte, fechaSis As Date) As Variant
    'Devuelve la última cotización (la más nueva), para una moneda determinada
    '----------------------------------------------------------------------------------
    'Parámetros.
    '   Entrada:
    '               [tipoSalida]    1= debuelve el valor de la última cotización
    '                               2= debuelve la fecha de la última cotización
    '
    '               [codMoneda] código de la moneda de la cual se desea obtener
    '                           la cotización
    '               [fechaSis]  Fecha actual del sistema, (m_FechaSis)
    '
    '   Salida:     0= si no se encontró ningún registro en el archivo de Cotizaciones
    '               para la moneda correspondiente
    '
    '               Si el tipoSalida = 1,debuelve el valor de la última cotización
    '               Si el tipoSalida = 2,debuelve la fecha de la última cotización
    '-----------------------------------------------------------------------------------
    'por defecto asumo que no existe cotización para la moneda
    mFunObtengoUltimaCotizacion = 0
    'busco si se encuentra definida una cotización para la fecha actual
    tbCOTIZACIONES.Index = "pkCotizaciones"
    tbCOTIZACIONES.Seek "=", codMoneda, fechaSis
    If Not tbCOTIZACIONES.NoMatch Then
        'existe cotización definida para la fecha
        If tipoSalida = 1 Then
            'devuelvo el valor de la cotización
            mFunObtengoUltimaCotizacion = tbCOTIZACIONES("valorCot")
        Else
            If tipoSalida = 2 Then
                'devuelvo la fecha de la cotización
                mFunObtengoUltimaCotizacion = tbCOTIZACIONES("fechaCot")
            End If
        End If
    Else
        'no existe cotización para la fecha
        tbCOTIZACIONES.Seek ">=", codMoneda, 0
        If Not tbCOTIZACIONES.NoMatch Then
            'me posiciono en el primer registro para la moneda y recorro el archvio
            'hasta encontrar el último registro para la moneda
            Do While Not tbCOTIZACIONES.EOF
                If tbCOTIZACIONES("codMoneda") = codMoneda Then
                    If tipoSalida = 1 Then
                        'devuelvo el valor de la cotización
                        mFunObtengoUltimaCotizacion = tbCOTIZACIONES("valorCot")
                    Else
                        If tipoSalida = 2 Then
                            'devuelvo la fecha de la cotización
                            mFunObtengoUltimaCotizacion = tbCOTIZACIONES("fechaCot")
                        End If
                    End If
                Else
                    'no tengo más registro para la moneda
                    Exit Do
                End If
                tbCOTIZACIONES.MoveNext
            Loop
        End If
    End If
End Function

Public Function mFunMuestroNroReserva(nroRes As Long) As String
    '-----------------------------------------------------------------------
    'Convierte el número de reserva a un formato más legible para el usuario
    '-----------------------------------------------------------------------
    'Parámetros:
    '   Entrada: número de reservaen formato long.
    '   Salida: número de reserva en formato string.
    '           ejemplo: entra 200300029 y sale 2003-00029
    '                   (paret año + 5 dígitos correlativos)
    '------------------------------------------------------------------------
    On Error Resume Next
    mFunMuestroNroReserva = Mid(Str(nroRes), 1, 5) + "-" + Mid(Str(nroRes), 6, 10)
End Function

Public Function mFunFormatoNombre(nom As String) As String
    '--------------------------------------------------------------------------
    'Para una mayor prolijidad en la presentación de la información se establece
    'un formato determinado para mostrar los nombre.
    'Esta función se encarga de convertir el nombre ingresado por el usuario al
    'formato correspondiente.
    '---------------------------------------------------------------------------
    'Parámetros.
    '   Entrada = nombre sin formato
    '   Salida  = nombre con formato, el formato es igual a primer letra en mayúsculas
    '             y el resto del nombre en minúsculas.
    '             ejm.: gabriel, Gabriel.
    '------------------------------------------------------------------------------
    On Error Resume Next
    
    Dim nomAux As String
    nomAux = StrConv(nom, 2)    'convuerto todo el string a minúsculas
    nomAux = StrConv(nomAux, 3)    'convierto la primer letra a mayúsculas
    mFunFormatoNombre = nomAux
End Function

Public Function mFunFormatoApellido(Ape As String) As String
    '--------------------------------------------------------------------------
    'Para una mayor prolijidad en la presentación de la información se establece
    'un formato determinado para mostrar los apellidos.
    'Esta función se encarga de convertir el apellido ingresado por el usuario al
    'formato correspondiente.
    '---------------------------------------------------------------------------
    'Parámetros.
    '   Entrada = apellido sin formato
    '   Salida  = apellido con formato, el formato es todas las letras en mayúsculas
    '             ejm.: aramburu, ARAMBURU
    '------------------------------------------------------------------------------
    On Error Resume Next
    
    Dim ApeAux As String
    ApeAux = StrConv(Ape, 1)    'convierto todo el string a mayúsculas
    mFunFormatoApellido = ApeAux
End Function

Public Function mFunObtengoSignoMoneda(tipoMoneda As Byte) As String
    '----------------------------------------------------------------------
    'Devuelve el signo utilizado para los tipos de monedas utilizados en el
    'sistema.
    '-----------------------------------------------------------------------
    'Parámetros.
    '   Entrada [tipoMoneda] 0 = moneda nacional
    '                        1 = dólares
    '   Salida  signo almacenado en el archivo parámetros.
    '-----------------------------------------------------------------------
    On Error Resume Next
    Select Case tipoMoneda
        Case 0  'm/n
            mFunObtengoSignoMoneda = tbPARAMETROS("simboloMonedaNacional")
        Case 1  'dólares
            mFunObtengoSignoMoneda = tbPARAMETROS("simboloDolares")
    End Select
End Function

'-------------------------------------------------------------------------------------------------------
'Esta función, convierte un número en su correspondiente trascripción a letras. Funciona bien con
'números enteros y con hasta 2 decimales, pero más de 2 decimales se pierde y no "sabe" lo que dice.
'
'Debes introducir este código en un módulo (por ejemplo) y realizar la llamada con el número que
'deseas convertir. Por Ejemplo: Label1 = Numlet(CCur(Text1))
'-------------------------------------------------------------------------------------------------------


Public Function Numlet$(NUM#)
    Dim DEC$, MILM$, MILL$, MILE$, UNID$
    ReDim SALI$(11)
    Dim var$, i%, aux$
    'NUM# = Round(NUM#, 2)
    var$ = Trim$(Str$(NUM#))
        If InStr(var$, ".") = 0 Then
            var$ = var$ + ".00"
        End If
       
        If InStr(var$, ".") = Len(var$) - 1 Then
            var$ = var$ + "0"
        End If
    var$ = String$(15 - Len(LTrim$(var$)), "0") + LTrim$(var$)
    DEC$ = Mid$(var$, 14, 2)
    MILM$ = Mid$(var$, 1, 3)
    MILL$ = Mid$(var$, 4, 3)
    MILE$ = Mid$(var$, 7, 3)
    UNID$ = Mid$(var$, 10, 3)
    For i% = 1 To 11: SALI$(i%) = " ": Next i%
    i% = 0
    Unidades$(1) = "UN "
    Unidades$(2) = "DOS "
    Unidades$(3) = "TRES "
    Unidades$(4) = "CUATRO "
    Unidades$(5) = "CINCO "
    Unidades$(6) = "SEIS "
    Unidades$(7) = "SIETE "
    Unidades$(8) = "OCHO "
    Unidades$(9) = "NUEVE "

    Decenas$(1) = "DIEZ "
    Decenas$(2) = "VEINTE "
    Decenas$(3) = "TREINTA "
    Decenas$(4) = "CUARENTA "
    Decenas$(5) = "CINCUENTA "
    Decenas$(6) = "SESENTA "
    Decenas$(7) = "SETENTA "
    Decenas$(8) = "OCHENTA "
    Decenas$(9) = "NOVENTA "

    Oncenas$(1) = "ONCE "
    Oncenas$(2) = "DOCE "
    Oncenas$(3) = "TRECE "
    Oncenas$(4) = "CATORCE "
    Oncenas$(5) = "QUINCE "
    Oncenas$(6) = "DIECISEIS "
    Oncenas$(7) = "DIECISIETE "
    Oncenas$(8) = "DIECIOCHO "
    Oncenas$(9) = "DIECINUEVE "

    Veintes$(1) = "VEINTIUN "
    Veintes$(2) = "VEINTIDOS "
    Veintes$(3) = "VEINTITRES "
    Veintes$(4) = "VEINTICUATRO "
    Veintes$(5) = "VEINTICINCO "
    Veintes$(6) = "VEINTISEIS "
    Veintes$(7) = "VEINTISIETE "
    Veintes$(8) = "VEINTIOCHO "
    Veintes$(9) = "VEINTINUEVE "

    Centenas$(1) = " CIENTO "
    Centenas$(2) = " DOSCIENTOS "
    Centenas$(3) = " TRESCIENTOS "
    Centenas$(4) = "CUATROCIENTOS "
    Centenas$(5) = " QUINIENTOS "
    Centenas$(6) = " SEISCIENTOS "
    Centenas$(7) = " SETECIENTOS "
    Centenas$(8) = " OCHOCIENTOS "
    Centenas$(9) = " NOVECIENTOS "

    If NUM# > 999999999999.99 Then Numlet$ = " ": Exit Function
        If Val(MILM$) >= 1 Then
            SALI$(2) = " MIL ": '** MILES DE MILLONES
            SALI$(4) = " MILLONES "
                If Val(MILM$) <> 1 Then
                    Unidades$(1) = "UN "
                    Veintes$(1) = "VEINTIUN "
                    SALI$(1) = Descifrar$(Val(MILM$))
                End If
        End If
        If Val(MILL$) >= 1 Then
            If Val(MILL$) < 2 Then
                SALI$(3) = "UN ": '*** UN MILLON
                    If Trim$(SALI$(4)) <> "MILLONES" Then
                        SALI$(4) = " MILLON "
                    End If
                Else
                    SALI$(4) = " MILLONES ": '*** VARIOS MILLONES
                    Unidades$(1) = "UN "
                    Veintes$(1) = "VEINTIUN "
                    SALI$(3) = Descifrar$(Val(MILL$))
                End If
        End If

    For i% = 2 To 9
        Centenas$(i%) = Mid$(Centenas(i%), 1, 11) '+ "AS" no son pesetas son pesos (por ahora!!)
    Next i%
        If Val(MILE$) > 0 Then
            SALI$(6) = " MIL ": '*** MILES
                If Val(MILE$) <> 1 Then
                    SALI$(5) = Descifrar$(Val(MILE$))
                End If
      End If
        Unidades$(1) = "UN "
        Veintes$(1) = "VEINTIUN "
            If Val(UNID$) >= 1 Then
                SALI$(7) = Descifrar$(Val(UNID$)): '*** CIENTOS
                    If Val(DEC$) >= 10 Then
                        SALI$(8) = " CON ": '*** DECIMALES
                        'SALI$(10) = Descifrar$(Val(DEC$))
                        SALI$(10) = Val(DEC$) & "/100"
                    End If
            End If
            If Val(MILM$) = 0 And Val(MILL$) = 0 And Val(MILE$) = 0 And Val(UNID$) = 0 Then SALI$(7) = " CERO "
            aux$ = ""
                For i% = 1 To 11
                    aux$ = aux$ + SALI$(i%)
                Next i%
       Numlet$ = Trim$(aux$)
  End Function

Function Descifrar$(numero%)
Static SAL$(4)
Dim i%, CT As Double, DC As Double, DU As Double, UD As Double
Dim VARIABLE$

    For i% = 1 To 4: SAL$(i%) = " ": Next i%
        VARIABLE$ = String$(3 - Len(Trim$(Str$(numero%))), "0") + Trim$(Str$(numero%))
        CT = Val(Mid$(VARIABLE$, 1, 1)): '*** CENTENA
        DC = Val(Mid$(VARIABLE$, 2, 1)): '*** DECENA
        DU = Val(Mid$(VARIABLE$, 2, 2)): '*** DECENA + UNIDAD
        UD = Val(Mid$(VARIABLE$, 3, 1)): '*** UNIDAD
        If numero% = 100 Then
            SAL$(1) = "CIEN "
        Else
            If CT <> 0 Then SAL$(1) = Centenas$(CT)
                If DC <> 0 Then
                    If DU <> 10 And DU <> 20 Then
                        If DC = 1 Then SAL$(2) = Oncenas$(UD): Descifrar$ = Trim$(SAL$(1) + " " + SAL$(2)):  Exit Function
                                If DC = 2 Then SAL$(2) = Veintes$(UD): Descifrar$ = Trim$(SAL$(1) + " " + SAL$(2)):  Exit Function
                                End If
                            SAL$(2) = " " + Decenas$(DC)
                                If UD <> 0 Then SAL$(3) = "Y "
                        End If
                            If UD <> 0 Then SAL$(4) = Unidades$(UD)
                    End If
                        Descifrar = Trim$(SAL$(1) + SAL$(2) + SAL$(3) + SAL$(4))
            End Function

