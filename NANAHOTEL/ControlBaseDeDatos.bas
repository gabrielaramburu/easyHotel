Attribute VB_Name = "ControlBaseDeDatos"
Option Explicit
'Al ejecutar un formulario, se asume que existen datos en la base de datos choerentes,
'que permitirán la ejecución correcta del código de dichos formularios o módulos.

'Sin embargo, cuando se instala la aplicación, las tablas de la base de datos estan vacías,
'o sin inicializar. Esto puede originar que ciertos procesos no se puedan ejecutar, ya que
'es impresindible que los mismos cuenten con información básica.

'En este módulo se realiza el control de existencia de datos mínimos, es decir, antes de
'abrir un determinado formulario, se valida que existan los datos mínimos necesarios
'en la base de datos para que el mismo se pueda ejecutar.


Public Function mFunControlDeBaseDeDatos(formulario As String) As Boolean
    'Determino si existe información para mostrar determinados formularios
    '------------------------------------------------------------------------
    'Parámetros.
    '   Entrada:    [formulario] Formulario que voy a ejecutar
    '
    '   Salida:     True, existe información mínima
    '               False, no existe información mínima
    '-------------------------------------------------------------------------
    
    'por defecto asumo que no existe información mínima
    mFunControlDeBaseDeDatos = False
    Select Case formulario
        Case "frmCargaReserva"
            'valido que existan habitaciones
            'valido que existan tipos de habitación
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
            'valido que existan artículos
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
            'de la moneda dolar. Si no es así los procedimientos de conversión
            'cancelan al dividir por 0.
            If funExistenRegistrosCotizaciones(1) Then
                mFunControlDeBaseDeDatos = True
            End If
        
        Case "frmConsultaCuentas"
            'valido que exista por lo menos un registro definido de cotizaciones
            'de la moneda dolar. Si no es así los procedimientos de conversión
            'cancelan al dividir por 0.
            
            'Este argumento se utiliza para la consulta cuentas por habitación y por cliente .
            If funExistenRegistrosCotizaciones(1) Then
                mFunControlDeBaseDeDatos = True
            End If
            
        Case "frmFacturacion"
            'valido que exista por lo menos un registro definido de cotizaciones
            'de la moneda dolar. Si no es así los procedimientos de conversión
            'cancelan al dividir por 0.
            
            'Este argumento se utiliza al realizar nuevo documento y nueva devolución
            If funExistenRegistrosCotizaciones(1) Then
                mFunControlDeBaseDeDatos = True
            End If
        
    End Select
End Function

Private Function funExistenRegistrosCotizaciones(tipoMoneda As Byte) As Boolean
    'Determina si existe por lo menos 1 registro de cotizaciones de la moneda determinada.
    '------------------------------------------------------------------------------------
    'Parámetros.
    '   Entrada:    [tipoMoneda]    código que identifica a la moneda
    '                               0 M/N   1 Dólares
    '
    '   Salida:     True, si existe definida por lo menos una cotización para la moneda
    '               False, no existe ningún valor definido.
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
    'Determina si existe 1 o más registros en una tabla determinada
    '----------------------------------------------------------------------
    'Parámetros.
    '   Entrada:    [tabla] tabla a la cual se quiere verificar la
    '                       existencia de registros
    '
    '   Salida:     True, si la tabla contiene 1 o más registros
    '               False, si la tabla esta vacía
    '
    '   NOTA:   utilizar la función RecordCount, elentecería la aplicación
    '           por lo que utilizo el metodo MoveFirst, conjuntamente con
    '           con la intersepción de errores
    '-----------------------------------------------------------------------
    On Error GoTo error
    
    Dim tablaAux As Recordset
    
    'por defecto asumo que existen registros
    funExistenRegistros = True
    
    'utilizo una variable auxiliar para no trabajar directamente con la tabla,
    'lo que produciría que se modificara el índice activo, al ejecutar MoveFirst
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
    'Parámetros.
    '   Entrada:    [tipoReg] valor del primer campo de la clave de la tabla.
    '
    '   Salida:     True, si la tabla contiene 1 o más registros de un tipo det.
    '               False, si la tabla esta vacía o no contiene registros del tipo det.
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

'Archivo de parámetros.
'En este archivo existe información referente a distintos aspectos de la aplicación,
'como ser, fechas de ejecución, próximos números correlativos, etc.

'A medida que la aplicación ejecuta determinados procesos, estos valores se van actualizando

'El problemas es que la primera vez que un proceso quiera utilizar algunos de estos valores,
'los mismos estarán sin inicializar. A continuación se implementan los procesos de inicialización
'para determinados campos de esta tabla.

'Esto procedimientos de inicialización son llamados por los procesos que requieren información de ciertos campos
'del archivo y no encuntran la misma.

'Esto deberá ocurrir la primera vez que se ejecuta cada proceso.

'Estos procesos de validación son únicos para cada campo del archivo.

'----------------------------------------------------------------------------------------
'NOTA: no se pueden modificar la posición de los campos actuales del archivo parámetros
'ya que algunos procesos acceden a los mismos por su índice.
'Esta precaución debe de ser tomada también para todas las demas tablas de la base de datos.

'CAMPOS A INICIALIZAR por medio de procedimientos
'   NroReserva
'   Fecha_ultimo_cierre_realizado

'CAMPOS QUE NO SE UTILIZAN
'   anioSis
'   fecha_aloja_auto
'   tot_habitaciones

'Es necesario inicializar los campos del archivo correspondiente al dígito de cada número de
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
'También debo de inicializar los campos que almacenan el próximo número de documento a utilizar.
'con el valor 1.

Public Sub mSubInicializoNroReserva()
    'Este procedimento es llamado cuando el campo nro_reserva del archivo parámetros = 0
    'Esto ocurre cuando:
    'Se realiza la primer reserva del sistema sin haber realizado un walkin.
    'Se realiza el primer walkin del sistema sin haber realizado una reserva.
    
    Dim anioSis As String
    Dim parteCifra As String
    Dim nroResAux As Long
    
    'obtengo el año del sistema
    anioSis = Year(m_FechaSis)
    parteCifra = "00001"
    nroResAux = Val(anioSis & parteCifra)
    'grabo archivo parámetros
    tbPARAMETROS.Edit
        tbPARAMETROS("nroreserva") = nroResAux
    tbPARAMETROS.Update
End Sub

Public Function mFunInicializoFechaSistema() As Boolean
    'Inicializo el campo "fecha_ultimo_cierre_realizado" con la fecha del sistema.
    'Este procedimiento es llamado cuando se inicia la aplicación y el valor de este
    'campo no es una fecha, es decir es nulo.
    '-------------------------------------------------------------------------------
    'Parámetros.
    '   Salida: True el usuario confirma la nueva fecha del sistema.
    '           False el usuario no confirma la nueva fecha
    '-------------------------------------------------------------------------------
    
    Dim mensaje As String
    'por defecto asumo que la fecha del sistema no es correcta
    mFunInicializoFechaSistema = False
    mensaje = "Aviso al usuario: 801 " & Chr(10) & _
                "No hay fecha de sistema establecida." & Chr(10) & _
                "Esto se debe a que es la primera vez que ejecuta la aplicación." & Chr(10) & _
                "Se establecerá la fecha del sistema a " & Format(Date, "dddd, d mmm yyyy") & Chr(10) & _
                "Si la fecha es correcta presione ACEPTAR para continuar, si no lo es presione CANCELAR para corregirla."
    If MsgBox(mensaje, vbInformation + vbOKCancel, "Confirmar fecha del sistema.") = vbOK Then
        tbPARAMETROS.Edit
            tbPARAMETROS("fecha_ultimo_cierre_realizado") = Date
        tbPARAMETROS.Update
        'la fecha del sistema es correcta
        mFunInicializoFechaSistema = True
    End If
End Function
