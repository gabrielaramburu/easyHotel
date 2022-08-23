Attribute VB_Name = "Correlativos"
'En este m�dulo estan los procedimientos encargados de obtener
'n�meros correlativos de las diferentes tablas que utilizan este m�todo.

Option Explicit

Public Function corr_situ(hab_cuenta As Long)
    'Obtengo corelativo de hist�rico de situaci�nes
    corr_situ = 1
    tbSITUACION_HIS.Index = "i_situacion"
    tbSITUACION_HIS.Seek ">=", hab_cuenta, 1
    If Not tbSITUACION_HIS.NoMatch Then
        Do While Not tbSITUACION_HIS.EOF
            If tbSITUACION_HIS("nrohab_situ") = hab_cuenta Then
                corr_situ = tbSITUACION_HIS("corr_situ") + 1
                tbSITUACION_HIS.MoveNext
            Else
                Exit Do
            End If
        Loop
    End If
End Function

Public Function proxima_reserva()
    'Obtengo el pr�ximo nro. de reserva a utilizar
    Dim anioSis As String
    Dim parteCifra As String
    
    proxima_reserva = tbPARAMETROS("nroreserva")

    If proxima_reserva = 0 Then
        'todav�a no se ha realizado ning�n walkin en el sistema y es
        'la primer reserva que estoy realizando
        mSubInicializoNroReserva
        proxima_reserva = tbPARAMETROS("nroreserva")
    End If
End Function

Public Sub sumo_corr_reserva()
    'Genera un nuevo n�mero de reserva que ser� utilizado por
    'la pr�xima reserva a realizar.

    If tbPARAMETROS("nroreserva") = nro_reserva Then
        tbPARAMETROS.Edit
        tbPARAMETROS("nroreserva") = nro_reserva + 1
        tbPARAMETROS.Update
    Else
        corr_reserva
    End If
End Sub

Public Sub corr_reserva()
    'En el caso que se halla realizado otra reserva (o m�s de una) en
    'el interin, debo de asignar un nuevo n�mero libre a la reserva actual
    'y calcular el pr�ximo para las siguientes (reservas).
    
    nro_reserva = tbPARAMETROS("nroreserva")
    tbPARAMETROS.Edit
    tbPARAMETROS("nroreserva") = nro_reserva + 1
    tbPARAMETROS.Update
End Sub

Public Sub mSubCorrReservaWalkin()
    'Es utilizado tamb�en cuando se realiza un Walkin.
    'Cada vez que se cancela un Walkinqueda un n�mero de reserva libre.
    'La diferencia con el procedimiento anterior es este se utiliza para walkin
    'y no quise incluir el control de nro_reserva = 0 en el procedimiento corr_reserva.
    
    nro_reserva = tbPARAMETROS("nroreserva")
    If nro_reserva = 0 Then
        'todav�a no se ha realizado ninguna reserva en el sistema
        'y estoy haciendo el primer walkin
        mSubInicializoNroReserva
    End If
    nro_reserva = tbPARAMETROS("nroreserva")
    tbPARAMETROS.Edit
    tbPARAMETROS("nroreserva") = nro_reserva + 1
    tbPARAMETROS.Update
End Sub

Public Function obtengo_proximo_gasto(fecha As Date)
    'Llamado desde frmIngExtras
    
    'Recorro todos los gastos realizados en una fecha determinada
    'y obtengo el �ltimo +1

    obtengo_proximo_gasto = 1
    
    tbCUENTAS.Index = "i_cuentas"
    tbCUENTAS.Seek ">=", fecha, 1
    If Not tbCUENTAS.NoMatch Then   'si se posiciona
        Do While Not tbCUENTAS.EOF
            If tbCUENTAS("fechagasto_cuenta") = fecha Then
                obtengo_proximo_gasto = tbCUENTAS("nrocorr_cuenta") + 1
                tbCUENTAS.MoveNext
            Else
                Exit Do
            End If
        Loop
    End If
End Function

Public Function obtengo_ultimo_corr_aloja(fecha As Date)
    obtengo_ultimo_corr_aloja = 1
    tbCUENTAS_ALOJA.Index = "pi_cuentas_aloja"
    tbCUENTAS_ALOJA.Seek ">=", fecha, 1
    If Not tbCUENTAS_ALOJA.NoMatch Then  'existe
        Do While Not tbCUENTAS_ALOJA.EOF
            If tbCUENTAS_ALOJA("fecha") = fecha Then
                obtengo_ultimo_corr_aloja = tbCUENTAS_ALOJA("nrocorr_cuenta_aloja") + 1
            Else
                Exit Do
            End If
            tbCUENTAS_ALOJA.MoveNext
        Loop
    End If
End Function

Public Function obtengo_ultimo_corr_cierre(fecha As Date)
    'Utilizado para crear un registro en el archivo que registra el resulta de
    'las ejecuciones del cierre diario.
    obtengo_ultimo_corr_cierre = 1
    tbCIERRE_DIARIO.Index = "pk_cierre"
    tbCIERRE_DIARIO.Seek ">=", fecha, 1
    If Not tbCIERRE_DIARIO.NoMatch Then 'existe
        Do While Not tbCIERRE_DIARIO.EOF
            If tbCIERRE_DIARIO("fecha_cierre") = fecha Then
                obtengo_ultimo_corr_cierre = tbCIERRE_DIARIO("nrocorr_cierre") + 1
            Else
                Exit Do
            End If
            tbCIERRE_DIARIO.MoveNext
        Loop
    End If
End Function

Public Function mFunObtengoCorrListadoPoblacionFlotante(fecha As Date) As Integer
    '--------------------------------------------------------------------
    'Recorro el archivo de listado de poblaci�n f. y obtengo el pr�ximo
    'n�mero correlativo.
    '
    'Par�metros:
    '   Entrada: [fecha] fecha del d�a que estoy cerrando
    '   Salida:  proximo correlativo
    '---------------------------------------------------------------------
    Dim tablaListado As Recordset
    
    Set tablaListado = tbPOBLACION_FLOTANTE
    
    mFunObtengoCorrListadoPoblacionFlotante = 1
    tablaListado.Index = "pk_listado"
    tablaListado.Seek ">=", fecha, 0
    If Not tablaListado.NoMatch Then 'existe
        Do While Not tablaListado.EOF
            If tablaListado("fechaListado") = fecha Then
                mFunObtengoCorrListadoPoblacionFlotante = _
                tablaListado("nroLineaListado") + 1
            Else
                Exit Do
            End If
            tablaListado.MoveNext
        Loop
    End If
    Set tablaListado = Nothing
End Function

Public Function mFun_obtengo_nrocorr_bloqueo(hab As Long)
    'Utilizado para crear obtener un nuevo n�mero de bloqueo
    'que ser� asignado al nuevo bloqueo de una habitaci�n dada.
    'frmBloquearHab
    mFun_obtengo_nrocorr_bloqueo = 1
    tbBLOQUEO_HAB.Index = "pk_bloqueo_hab"
    tbBLOQUEO_HAB.Seek ">=", hab, 0
    If Not tbBLOQUEO_HAB.NoMatch Then
        Do While Not tbBLOQUEO_HAB.EOF
            If tbBLOQUEO_HAB("hab_bloq") = hab Then
                mFun_obtengo_nrocorr_bloqueo = tbBLOQUEO_HAB("nrocorr_bloq") + 1
            Else
                Exit Do
            End If
            tbBLOQUEO_HAB.MoveNext
        Loop
    End If
End Function

Public Function mFunObtengoNroCorrTipoHabitaciones() As Integer
    'Devuelve el pr�ximo n�mero libre del archivo de Tipos de Habitaciones
    '----------------------------------------------------------------------
    'Par�metros.
    '   Salida:     el valor correspondiente, a el primer n�mero libre
    '               de la clave de tipos de habitaciones.
    '
    '               el valor es 1 si el archivo esta vac�o
    '-----------------------------------------------------------------------
    'por defecto asumo que el archivo esta vac�o
    mFunObtengoNroCorrTipoHabitaciones = 1
    tbTIPO_HABITACIONES.Index = "i_tipo_hab"
    tbTIPO_HABITACIONES.Seek ">=", 1
    'comienzo en el primer registro del archivo
    If Not tbTIPO_HABITACIONES.NoMatch Then
        'recorro todos los registros hasta el final
        Do While Not tbTIPO_HABITACIONES.EOF
            'almaceno el valor del �ltimo registro leido
            mFunObtengoNroCorrTipoHabitaciones = _
                            tbTIPO_HABITACIONES("tipoHab") + 1
            tbTIPO_HABITACIONES.MoveNext
        Loop
    End If
End Function

Public Function mFunObtengoProxTipoEstado(tipoReg As Byte) As Integer
    'Devuelve el pr�ximo n�mero libre del archivo de TIPO_ESTADO_HAB
    'para registros de un tipo determinado.
    '----------------------------------------------------------------------
    'Par�metros.
    '   Entrada:    [tipoReg]   valor del primer campo de la clave del archivo
    '                           por el cual realizo la b�squeda
    '
    '   Salida:     el valor correspondiente, a el primer n�mero libre
    '               de la clave de tipos de habitaciones.
    '
    '               el valor es 1 si el archivo esta vac�o
    '---------------------------------------------------------------------
    'por defecto asumo que el archivo esta vac�o
    mFunObtengoProxTipoEstado = 1
    tbTIPO_ESTADO_HAB.Index = "i_estado"
    tbTIPO_ESTADO_HAB.Seek ">=", tipoReg, 1
    If Not tbTIPO_ESTADO_HAB.NoMatch Then
        'recorro todos los registros cuyo tipo sea igual al tipo del par�metro
        Do While Not tbTIPO_ESTADO_HAB.EOF
            'valido tipo del registro
            If tbTIPO_ESTADO_HAB("tipo_cod") = tipoReg Then
                mFunObtengoProxTipoEstado = tbTIPO_ESTADO_HAB("cod") + 1
            Else
                Exit Do
            End If
            tbTIPO_ESTADO_HAB.MoveNext
        Loop
    End If
End Function
