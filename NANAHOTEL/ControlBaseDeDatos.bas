Attribute VB_Name = "ControlBaseDeDatos"
Option Explicit
'Al ejecutar un formulario, se asume que existen datos en la base de datos choerentes,
'que permitir�n la ejecuci�n correcta del c�digo de dichos formularios o m�dulos.

'Sin embargo, cuando se instala la aplicaci�n, las tablas de la base de datos estan vac�as,
'o sin inicializar. Esto puede originar que ciertos procesos no se puedan ejecutar, ya que
'es impresindible que los mismos cuenten con informaci�n b�sica.

'En este m�dulo se realiza el control de existencia de datos m�nimos, es decir, antes de
'abrir un determinado formulario, se valida que existan los datos m�nimos necesarios
'en la base de datos para que el mismo se pueda ejecutar.


Public Function mFunControlDeBaseDeDatos(formulario As String) As Boolean
    'Determino si existe informaci�n para mostrar determinados formularios
    '------------------------------------------------------------------------
    'Par�metros.
    '   Entrada:    [formulario] Formulario que voy a ejecutar
    '
    '   Salida:     True, existe informaci�n m�nima
    '               False, no existe informaci�n m�nima
    '-------------------------------------------------------------------------
    
    'por defecto asumo que no existe informaci�n m�nima
    mFunControlDeBaseDeDatos = False
    Select Case formulario
        Case "frmCargaReserva"
            'valido que existan habitaciones
            'valido que existan tipos de habitaci�n
            If funExistenRegistros(tbHABITACIONES) And _
                funExistenRegistros(tbTIPO_HABITACIONES) Then
                    mFunControlDeBaseDeDatos = True
            End If
            
        Case "frmBloquearHab"
            'valido que existan motivos de bloqueo
            If funExistenRegistrosTipoEstado(1) Then
                mFunControlDeBaseDeDatos = True
            End If
            
        Case "frmCambioSitu"
            'valido que exista situaciones de habitaciones
            If funExistenRegistrosTipoEstado(2) Then
                mFunControlDeBaseDeDatos = True
            End If
            
        Case "frmConsultaCompleta"
            'valido que existan habitaciones
            If funExistenRegistros(tbHABITACIONES) Then
                mFunControlDeBaseDeDatos = True
            End If
            
        Case "frmConsultaTitular"
            'valido que existan habitaciones
            If funExistenRegistros(tbHABITACIONES) Then
                mFunControlDeBaseDeDatos = True
            End If
        
        Case "frmIngExtras"
            'valido que existan art�culos
            If funExistenRegistros(tbARTICULOS) Then
                mFunControlDeBaseDeDatos = True
            End If
        
        Case "frmListadoIngresos"
            'valido que existan tipos de habitaciones
            If funExistenRegistros(tbTIPO_HABITACIONES) Then
                mFunControlDeBaseDeDatos = True
            End If
            
        Case "frmListadoEgresos"
            'valido que existan tipos de habitaciones
            If funExistenRegistros(tbTIPO_HABITACIONES) Then
                mFunControlDeBaseDeDatos = True
            End If
            
        Case "frmVerDisponibilidad"
            'valido que existan tipos de habitaciones
            If funExistenRegistros(tbTIPO_HABITACIONES) Then
                mFunControlDeBaseDeDatos = True
            End If
                    
        Case "frmCuadroHab"
            'valido que existan tipos de habitaciones
            If funExistenRegistros(tbHABITACIONES) Then
                mFunControlDeBaseDeDatos = True
            End If
                    
        Case "frmCierreDiario"
            'valido que existan habitaciones
            If funExistenRegistros(tbHABITACIONES) Then
                mFunControlDeBaseDeDatos = True
            End If
            
        Case "frmEstadoCuentas"
            'valido que exista por lo menos un registro definido de cotizaciones
            'de la moneda dolar. Si no es as� los procedimientos de conversi�n
            'cancelan al dividir por 0.
            If funExistenRegistrosCotizaciones(1) Then
                mFunControlDeBaseDeDatos = True
            End If
        
        Case "frmConsultaCuentas"
            'valido que exista por lo menos un registro definido de cotizaciones
            'de la moneda dolar. Si no es as� los procedimientos de conversi�n
            'cancelan al dividir por 0.
            
            'Este argumento se utiliza para la consulta cuentas por habitaci�n y por cliente .
            If funExistenRegistrosCotizaciones(1) Then
                mFunControlDeBaseDeDatos = True
            End If
            
        Case "frmFacturacion"
            'valido que exista por lo menos un registro definido de cotizaciones
            'de la moneda dolar. Si no es as� los procedimientos de conversi�n
            'cancelan al dividir por 0.
            
            'Este argumento se utiliza al realizar nuevo documento y nueva devoluci�n
            If funExistenRegistrosCotizaciones(1) Then
                mFunControlDeBaseDeDatos = True
            End If
        
    End Select
End Function

Private Function funExistenRegistrosCotizaciones(tipoMoneda As Byte) As Boolean
    'Determina si existe por lo menos 1 registro de cotizaciones de la moneda determinada.
    '------------------------------------------------------------------------------------
    'Par�metros.
    '   Entrada:    [tipoMoneda]    c�digo que identifica a la moneda
    '                               0 M/N   1 D�lares
    '
    '   Salida:     True, si existe definida por lo menos una cotizaci�n para la moneda
    '               False, no existe ning�n valor definido.
    '-------------------------------------------------------------------------------------
    'por defecto asumo que no existen registros
    funExistenRegistrosCotizaciones = False
    tbCOTIZACIONES.Index = "pkCotizaciones"
    tbCOTIZACIONES.Seek ">=", tipoMoneda, 0
    If Not tbCOTIZACIONES.NoMatch Then
        If tbCOTIZACIONES("codMoneda") = tipoMoneda Then
            'existe por lo menos 1 valor definido para la moneda
            funExistenRegistrosCotizaciones = True
        End If
    End If
End Function

Private Function funExistenRegistros(tabla As Recordset) As Boolean
    'Determina si existe 1 o m�s registros en una tabla determinada
    '----------------------------------------------------------------------
    'Par�metros.
    '   Entrada:    [tabla] tabla a la cual se quiere verificar la
    '                       existencia de registros
    '
    '   Salida:     True, si la tabla contiene 1 o m�s registros
    '               False, si la tabla esta vac�a
    '
    '   NOTA:   utilizar la funci�n RecordCount, elentecer�a la aplicaci�n
    '           por lo que utilizo el metodo MoveFirst, conjuntamente con
    '           con la intersepci�n de errores
    '-----------------------------------------------------------------------
    On Error GoTo error
    
    Dim tablaAux As Recordset
    
    'por defecto asumo que existen registros
    funExistenRegistros = True
    
    'utilizo una variable auxiliar para no trabajar directamente con la tabla,
    'lo que producir�a que se modificara el �ndice activo, al ejecutar MoveFirst
    Set tablaAux = tabla
    tablaAux.MoveFirst
    Set tablaAux = Nothing
Exit Function
error:
    funExistenRegistros = False
End Function

Private Function funExistenRegistrosTipoEstado(tipoReg As Integer) As Boolean
    'Determina si existen registros de un tipo determinado en la tabla
    'de tbTIPO_ESTADO_HAB
    '----------------------------------------------------------------------
    'Par�metros.
    '   Entrada:    [tipoReg] valor del primer campo de la clave de la tabla.
    '
    '   Salida:     True, si la tabla contiene 1 o m�s registros de un tipo det.
    '               False, si la tabla esta vac�a o no contiene registros del tipo det.
    '
    '-----------------------------------------------------------------------
    On Error Resume Next
    'por defecto asumo que no existen registros de ese tipo
    funExistenRegistrosTipoEstado = False
    'busco registro
    tbTIPO_ESTADO_HAB.Index = "i_estado"
    tbTIPO_ESTADO_HAB.Seek ">=", tipoReg, 0
    If Not tbTIPO_ESTADO_HAB.NoMatch Then
        If tbTIPO_ESTADO_HAB(0) = tipoReg Then
            'encontre un registro del tipo determinado
            funExistenRegistrosTipoEstado = True
        End If
    End If
End Function

'Archivo de par�metros.
'En este archivo existe informaci�n referente a distintos aspectos de la aplicaci�n,
'como ser, fechas de ejecuci�n, pr�ximos n�meros correlativos, etc.

'A medida que la aplicaci�n ejecuta determinados procesos, estos valores se van actualizando

'El problemas es que la primera vez que un proceso quiera utilizar algunos de estos valores,
'los mismos estar�n sin inicializar. A continuaci�n se implementan los procesos de inicializaci�n
'para determinados campos de esta tabla.

'Esto procedimientos de inicializaci�n son llamados por los procesos que requieren informaci�n de ciertos campos
'del archivo y no encuntran la misma.

'Esto deber� ocurrir la primera vez que se ejecuta cada proceso.

'Estos procesos de validaci�n son �nicos para cada campo del archivo.

'----------------------------------------------------------------------------------------
'NOTA: no se pueden modificar la posici�n de los campos actuales del archivo par�metros
'ya que algunos procesos acceden a los mismos por su �ndice.
'Esta precauci�n debe de ser tomada tambi�n para todas las demas tablas de la base de datos.

'CAMPOS A INICIALIZAR por medio de procedimientos
'   NroReserva
'   Fecha_ultimo_cierre_realizado

'CAMPOS QUE NO SE UTILIZAN
'   anioSis
'   fecha_aloja_auto
'   tot_habitaciones

'Es necesario inicializar los campos del archivo correspondiente al d�gito de cada n�mero de
'documento.
'El valor correspondiente debe de ser el siguiente:
'campo  valor
'7      0
'10     1
'13     2
'16     3
'19     4
'22     5
'25     6
'28     7
'Tambi�n debo de inicializar los campos que almacenan el pr�ximo n�mero de documento a utilizar.
'con el valor 1.

Public Sub mSubInicializoNroReserva()
    'Este procedimento es llamado cuando el campo nro_reserva del archivo par�metros = 0
    'Esto ocurre cuando:
    'Se realiza la primer reserva del sistema sin haber realizado un walkin.
    'Se realiza el primer walkin del sistema sin haber realizado una reserva.
    
    Dim anioSis As String
    Dim parteCifra As String
    Dim nroResAux As Long
    
    'obtengo el a�o del sistema
    anioSis = Year(m_FechaSis)
    parteCifra = "00001"
    nroResAux = Val(anioSis & parteCifra)
    'grabo archivo par�metros
    tbPARAMETROS.Edit
        tbPARAMETROS("nroreserva") = nroResAux
    tbPARAMETROS.Update
End Sub

Public Function mFunInicializoFechaSistema() As Boolean
    'Inicializo el campo "fecha_ultimo_cierre_realizado" con la fecha del sistema.
    'Este procedimiento es llamado cuando se inicia la aplicaci�n y el valor de este
    'campo no es una fecha, es decir es nulo.
    '-------------------------------------------------------------------------------
    'Par�metros.
    '   Salida: True el usuario confirma la nueva fecha del sistema.
    '           False el usuario no confirma la nueva fecha
    '-------------------------------------------------------------------------------
    
    Dim mensaje As String
    'por defecto asumo que la fecha del sistema no es correcta
    mFunInicializoFechaSistema = False
    mensaje = "Aviso al usuario: 801 " & Chr(10) & _
                "No hay fecha de sistema establecida." & Chr(10) & _
                "Esto se debe a que es la primera vez que ejecuta la aplicaci�n." & Chr(10) & _
                "Se establecer� la fecha del sistema a " & Format(Date, "dddd, d mmm yyyy") & Chr(10) & _
                "Si la fecha es correcta presione ACEPTAR para continuar, si no lo es presione CANCELAR para corregirla."
    If MsgBox(mensaje, vbInformation + vbOKCancel, "Confirmar fecha del sistema.") = vbOK Then
        tbPARAMETROS.Edit
            tbPARAMETROS("fecha_ultimo_cierre_realizado") = Date
        tbPARAMETROS.Update
        'la fecha del sistema es correcta
        mFunInicializoFechaSistema = True
    End If
End Function
