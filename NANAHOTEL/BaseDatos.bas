Attribute VB_Name = "BaseDatos"
'En este módulo están todo los procedimientos y funciones involucrados
'en el manejo de la base de datos
Option Explicit
Public bdHOTEL As Database, bdWK As Workspace
Public tbRESERVAS As Recordset
Public tbPARAMETROS As Recordset
Public tbHAB_RESERVAS As Recordset
Public tbCLAVES_ACCESO As Recordset
Public tbHABITACIONES As Recordset
Public tbHAB_RESERVAS_AUX As Recordset
Public tbNACIONALIDADES As Recordset
Public tbPAISES As Recordset
Public tbCOTIZACIONES As Recordset
Public tbCLIENTES As Recordset
Public tbCHECKIN As Recordset
Public tbCHECKOUT As Recordset
Public tbCUENTAS As Recordset
Public tbARTICULOS As Recordset
Public tbPUNTO_VENTA As Recordset
Public tbTIPO_HABITACIONES As Recordset
Public tbCUENTAS_ALOJA As Recordset
Public tbCABEZAL As Recordset
Public tbLINEAS As Recordset
Public tbEMPRESAS As Recordset
Public tbTIPO_ESTADO_HAB As Recordset
Public tbBLOQUEO_HAB As Recordset
Public tbSITUACION_HIS As Recordset
Public tbANULADAS As Recordset
Public tbHAB_ANULADAS As Recordset
Public tbESTADO_CUENTAS As Recordset
Public tbRECIVOS As Recordset
Public tbCIERRE_DIARIO As Recordset
Public tbSIS_COLORES As Recordset
Public tbSIS_FUENTES As Recordset
Public tbIVA As Recordset
Public tbSISTEMA_USUARIOS As Recordset
Public tbSISTEMA_PERFILES As Recordset
Public tbSISTEMA_BITACORA As Recordset
Public tbSISTEMA_OPERACIONES As Recordset
Public tbSISTEMA_CONF_FORMULARIOS As Recordset
Public tbSISTEMA_MENSAJES As Recordset
Public tbSISTEMA_LICENCIA As Recordset
Public tbMONEDAS As Recordset
Public tbSEXO As Recordset
Public tbESTADO_CIVIL As Recordset
Public tbTARJETAS As Recordset
Public tbPOBLACION_FLOTANTE As Recordset
Public tbSISTEMA_LISTADOS As Recordset
Public tbSISTEMA_CONSTANTES As Recordset

'Declaración de constante de contraseña
Public Const cContraseñaBD As String = ";PWD=manyacapo;"

'*******************************************************************************************
'   NOTA:       Asignación de campos a variables
'
'   Los campos de la base de datos de tipo String y Fecha, se cargan con valores nulos
'   Los campos de tipo Integer se cargan con 0.
'   Esto ocurre al crear un nuevo registro. Si se accede algun campo de tipo string o fecha
'   y el mismo contiene valor nulo, cancela la sentencia de asignación de dicho campo a
'   una variable.
'   Para evitar esto hay que utilizar la funcion IsNull.
'*******************************************************************************************

'*******************************************************************************************
'   NOTA:       Declaración de variables de archivo (RecordSet) públicas.       09/12/02
'
'   Si bien todos los archivos utilizados por la aplicación, se usan mediante la utilización
'   de las variables públicas tbNOMBRE_ARCHIVO, es una buena práctica empezar a implementar
'   declaraciones privadas de dichas variables en los procedimientos o funciones que
'   accedan a archivos. Esto evita que el puntero,quede incorrectamente posicionado ya que
'   su posición es compartida por todos los procedimientos que utilizan la declaración
'   pública de las variables de archvio tbNOMBRE_ARCHIVO.
'   A medida que implementan nuevos procedimientos y funciones se empezará a utilizar
'   este nuevo creiterio.
'*********************************************************************************************

Public Sub mSubAbroBaseDeDatos()
    'asigna espacio trabajo
    Set bdWK = DBEngine.Workspaces(0)
    
    'obtengo el directorio de ejecución del exe.
    vardir = BaseDeDatosAplicacion           'directorio para y BD
    vardir2 = App.Path & "\"                 'directorio para reportes
        
    'abre base de datos
    Set bdHOTEL = bdWK.OpenDatabase(vardir, False, False, cContraseñaBD)
    
    'abre tablas
    Set tbRESERVAS = bdHOTEL.OpenRecordset("RESERVAS", dbOpenTable)
    Set tbPARAMETROS = bdHOTEL.OpenRecordset("SISTEMA_PARAMETROS", dbOpenTable)
    Set tbHAB_RESERVAS = bdHOTEL.OpenRecordset("HAB_RESERVA", dbOpenTable)
    Set tbHABITACIONES = bdHOTEL.OpenRecordset("HABITACIONES", dbOpenTable)
    Set tbHAB_RESERVAS_AUX = bdHOTEL.OpenRecordset("HAB_RESERVA_AUX", dbOpenTable)
    Set tbPAISES = bdHOTEL.OpenRecordset("PAISES", dbOpenTable)
    Set tbNACIONALIDADES = bdHOTEL.OpenRecordset("NACIONALIDADES", dbOpenTable)
    Set tbCOTIZACIONES = bdHOTEL.OpenRecordset("COTIZACIONES", dbOpenTable)
    Set tbCLIENTES = bdHOTEL.OpenRecordset("CLIENTES", dbOpenTable)
    Set tbCHECKIN = bdHOTEL.OpenRecordset("CHECKIN", dbOpenTable)
    Set tbCHECKOUT = bdHOTEL.OpenRecordset("CHECKOUT", dbOpenTable)
    Set tbCUENTAS = bdHOTEL.OpenRecordset("CUENTAS_EXTRA", dbOpenTable)
    Set tbARTICULOS = bdHOTEL.OpenRecordset("ARTICULOS", dbOpenTable)
    Set tbPUNTO_VENTA = bdHOTEL.OpenRecordset("PUNTO_VENTA", dbOpenTable)
    Set tbTIPO_HABITACIONES = bdHOTEL.OpenRecordset("TIPO_HABITACIONES", dbOpenTable)
    Set tbCUENTAS_ALOJA = bdHOTEL.OpenRecordset("CUENTAS_ALOJA", dbOpenTable)
    Set tbCABEZAL = bdHOTEL.OpenRecordset("FAC_CABEZAL", dbOpenTable)
    Set tbLINEAS = bdHOTEL.OpenRecordset("FAC_LINEAS", dbOpenTable)
    Set tbEMPRESAS = bdHOTEL.OpenRecordset("EMPRESAS", dbOpenTable)
    Set tbTIPO_ESTADO_HAB = bdHOTEL.OpenRecordset("TIPO_ESTADO_HAB", dbOpenTable)
    Set tbBLOQUEO_HAB = bdHOTEL.OpenRecordset("BLOQUEO_HAB", dbOpenTable)
    Set tbSITUACION_HIS = bdHOTEL.OpenRecordset("SITUACION_HIS", dbOpenTable)
    Set tbANULADAS = bdHOTEL.OpenRecordset("ANULADAS", dbOpenTable)
    Set tbHAB_ANULADAS = bdHOTEL.OpenRecordset("HAB_ANULADAS", dbOpenTable)
    Set tbESTADO_CUENTAS = bdHOTEL.OpenRecordset("ESTADO_CUENTAS", dbOpenTable)
    Set tbRECIVOS = bdHOTEL.OpenRecordset("RECIVOS", dbOpenTable)
    Set tbCIERRE_DIARIO = bdHOTEL.OpenRecordset("CIERRE_DIARIO", dbOpenTable)
    Set tbSIS_COLORES = bdHOTEL.OpenRecordset("SISTEMA_COLORES", dbOpenTable)
    Set tbSIS_FUENTES = bdHOTEL.OpenRecordset("SISTEMA_FUENTES", dbOpenTable)
    Set tbSISTEMA_USUARIOS = bdHOTEL.OpenRecordset("SISTEMA_USUARIOS", dbOpenTable)
    Set tbSISTEMA_PERFILES = bdHOTEL.OpenRecordset("SISTEMA_PERFILES", dbOpenTable)
    Set tbSISTEMA_BITACORA = bdHOTEL.OpenRecordset("SISTEMA_BITACORA", dbOpenTable)
    Set tbSISTEMA_OPERACIONES = bdHOTEL.OpenRecordset("SISTEMA_OPERACIONES", dbOpenTable)
    Set tbSISTEMA_MENSAJES = bdHOTEL.OpenRecordset("SISTEMA_MENSAJES", dbOpenTable)
    Set tbIVA = bdHOTEL.OpenRecordset("IVA", dbOpenTable)
    Set tbSISTEMA_CONF_FORMULARIOS = bdHOTEL.OpenRecordset("SISTEMA_CONF_FORMULARIOS", dbOpenTable)
    Set tbMONEDAS = bdHOTEL.OpenRecordset("MONEDAS", dbOpenTable)
    Set tbSEXO = bdHOTEL.OpenRecordset("SEXO", dbOpenTable)
    Set tbESTADO_CIVIL = bdHOTEL.OpenRecordset("ESTADO_CIVIL", dbOpenTable)
    Set tbTARJETAS = bdHOTEL.OpenRecordset("TARJETAS_CREDITO", dbOpenTable)
    Set tbSISTEMA_LICENCIA = bdHOTEL.OpenRecordset("SISTEMA_LICENCIA", dbOpenTable)
    Set tbPOBLACION_FLOTANTE = bdHOTEL.OpenRecordset("POBLACION_FLOTANTE", dbOpenTable)
    Set tbSISTEMA_LISTADOS = bdHOTEL.OpenRecordset("SISTEMA_LISTADOS", dbOpenTable)
    Set tbSISTEMA_CONSTANTES = bdHOTEL.OpenRecordset("SISTEMA_CONSTANTES", dbOpenTable)
End Sub

Public Sub mSubInicioAplicacion()
    'cargo colores desde archivo colores
    mSub_Inicialixo_colores_sistema
    'cargo fuentes desde archivo de fuentes
    mSub_Inicializo_fuentes_sistema
    
    'Inicializo los diferentes vectores utilizados en el programa
    'con los valores correspondientes
    mSubcargo_combos_vectores
    'inicializo variables globales de signo de moneda
    gblSignoMonedaNacional = mFunObtengoSignoMoneda(0)
    gblSignoDolares = mFunObtengoSignoMoneda(1)
    
End Sub

Public Function busco_reservaCheckinTF(res As Long)
    'Determino si una reserva determinada (una habitación) ya fue ocupada
    'Utilizado para distinguir reservas que ingresan hoy libres entre
    'reserva que ingresan hoy ocupadas.
    
    busco_reservaCheckinTF = False
    tbCHECKIN.Index = "i_checkin_rh"
    tbCHECKIN.Seek ">=", res, 0
    If Not tbCHECKIN.NoMatch Then
        If tbCHECKIN("nroreserva") = res Then
            busco_reservaCheckinTF = True
        End If
    End If
End Function

Public Function busco_ReservaHabita_checkin(res As Long, hab As Long)
    'Busca si la habitación correspondiente a una reserva ya ingresó al hotel
    busco_ReservaHabita_checkin = False
    tbCHECKIN.Index = "i_checkin_rh"
    tbCHECKIN.Seek "=", res, hab
    If Not tbCHECKIN.NoMatch Then
        busco_ReservaHabita_checkin = True
    End If
End Function

Public Function mFun_BuscoIvaTF(CodIva As Byte)
    mFun_BuscoIvaTF = False
    tbIVA.Index = "pk_iva"
    tbIVA.Seek "=", CodIva
    If Not tbIVA.NoMatch Then
        mFun_BuscoIvaTF = True
    End If
End Function

Public Function mFunObtengoPorcentajeIva(tipoIva As Byte) As Single
    '------------------------------------------------------------------------
    'Devuelvo el porcentaje de iva asociado al código de IVA que paso como
    'parámetro.
    '------------------------------------------------------------------------
    'Parámetros.
    '   Entrada.    [tipoIva]   tipo de iva
    '   Salida      porcentaje asociado a el tipo de iva
    '               0 si no encuentro tipo de iva en archivo de IVA
    '--------------------------------------------------------------------------
    'declaro variable para acceder a la tabla de IVAS
    Dim tbTablaIva As Recordset
    Set tbTablaIva = tbIVA
    tbTablaIva.Index = "pk_iva"
    tbTablaIva.Seek "=", tipoIva
    If Not tbTablaIva.NoMatch Then
        mFunObtengoPorcentajeIva = tbTablaIva("valorIva")
    Else
        mFunObtengoPorcentajeIva = 0
    End If
    Set tbTablaIva = Nothing
End Function

Public Function busco_recivoTF(tipo As Byte, nro As Long)
    busco_recivoTF = False
    tbRECIVOS.Index = "pk_recivo"
    tbRECIVOS.Seek "=", tipo, nro
    If Not tbRECIVOS.NoMatch Then
        busco_recivoTF = True
    End If
End Function

Public Function busco_reserva_anuladaTF(res_anu As Long)
    busco_reserva_anuladaTF = False
    tbANULADAS.Index = "i_reservas"
    tbANULADAS.Seek "=", res_anu
    If Not tbANULADAS.NoMatch Then
        busco_reserva_anuladaTF = True
    End If
End Function

Public Function busco_reservaTF(res As Long)
    busco_reservaTF = False
    tbRESERVAS.Index = "i_reservas"
    tbRESERVAS.Seek "=", res
    If Not tbRESERVAS.NoMatch Then
        busco_reservaTF = True
    End If
End Function

Public Function busco_clienteTF(cli As Long)
    busco_clienteTF = False
    tbCLIENTES.Index = "iclie_nrocorr"
    tbCLIENTES.Seek "=", cli
    If Not tbCLIENTES.NoMatch Then
        busco_clienteTF = True
    End If
End Function

Public Function busco_articuloTF(art As Long)
    busco_articuloTF = False
    tbARTICULOS.Index = "i_articulo"
    tbARTICULOS.Seek "=", art
    If Not tbARTICULOS.NoMatch Then
        busco_articuloTF = True
    End If
End Function

Public Function busco_habitaTF(hab As Long)
    busco_habitaTF = False
    tbHABITACIONES.Index = "inrohab"
    tbHABITACIONES.Seek "=", hab
    If Not tbHABITACIONES.NoMatch Then
        busco_habitaTF = True
    End If
End Function

Public Function busco_habita_checkin(hab As Long)
    'Busca si una habitación esta ocupada (tiene pasajeros hospedados)
    'Me posiciono en el cliente con menor número correlativo alojado en la habitacion.
    busco_habita_checkin = False
    tbCHECKIN.Index = "i_habitacion"
    tbCHECKIN.Seek "=", hab
    If Not tbCHECKIN.NoMatch Then   'existe
        busco_habita_checkin = True
    End If
End Function

Public Function busco_titular_checkinTF(hab As Long, tit As Long)
    'Busco un pasajero en una habitación.
    busco_titular_checkinTF = False
    tbCHECKIN.Index = "i_checkin"
    tbCHECKIN.Seek "=", hab, tit
    If Not tbCHECKIN.NoMatch Then
        busco_titular_checkinTF = True
    End If
End Function

Public Function busco_empTF(cod As Long)
    busco_empTF = False
    tbEMPRESAS.Index = "i_empresa"
    tbEMPRESAS.Seek "=", cod
    If Not tbEMPRESAS.NoMatch Then
        busco_empTF = True
    End If
End Function

Public Function busco_tipo_habTF(tipo As Long)
    busco_tipo_habTF = False
    tbTIPO_HABITACIONES.Index = "i_tipo_hab"
    tbTIPO_HABITACIONES.Seek "=", tipo
    If Not tbTIPO_HABITACIONES.NoMatch Then
        busco_tipo_habTF = True
    End If
End Function

Public Function mFunObtengoSituacionHab(situ As Long)
    mFunObtengoSituacionHab = ""
    If busco_estado_habTF(2, situ) Then
        mFunObtengoSituacionHab = tbTIPO_ESTADO_HAB("descri")
    End If
End Function

Public Function mFunObtengoTipoHab(hab As Long) As Integer
    'busco habitación
    If busco_habitaTF(hab) Then
        'devuelvo el tipo de la habitación
        mFunObtengoTipoHab = tbHABITACIONES("tipohab")
    End If
End Function

Public Function mFun_BuscoDescriTipoHab(tipo As Long)
    'Devuelve la descripción de un tipo de habitación
    mFun_BuscoDescriTipoHab = ""
    If busco_tipo_habTF(tipo) Then
        mFun_BuscoDescriTipoHab = tbTIPO_HABITACIONES("descripcion")
    End If
End Function

Public Function busco_tipo_hab_descri(hab As Long)
    'Devuelve la descripción del tipo de habitación de una habitación
    If busco_habitaTF(hab) Then
        busco_tipo_hab_descri = mFun_BuscoDescriTipoHab(tbHABITACIONES("tipohab"))
    End If
End Function

Public Function mfunBuscoReservaNoAsignada(res As Long, corr As Long)
    'Busca una reserva no asignada
    mfunBuscoReservaNoAsignada = False
    tbHAB_RESERVAS.Index = "ihab_reserva"
    tbHAB_RESERVAS.Seek "=", res, corr
    If Not tbHAB_RESERVAS.NoMatch Then  'existe
        mfunBuscoReservaNoAsignada = True
    End If
End Function

Public Function busco_puntoventaTF(codigo As Long)
    busco_puntoventaTF = False
    tbPUNTO_VENTA.Index = "i_punto_venta"
    tbPUNTO_VENTA.Seek "=", codigo
    If Not tbPUNTO_VENTA.NoMatch Then
        busco_puntoventaTF = True
    End If
End Function

Public Function busco_estado_habTF(tipo As Byte, codigo As Long)
    busco_estado_habTF = False
    tbTIPO_ESTADO_HAB.Index = "i_estado"
    tbTIPO_ESTADO_HAB.Seek "=", tipo, codigo
    If Not tbTIPO_ESTADO_HAB.NoMatch Then
        busco_estado_habTF = True
    End If
End Function

Public Function busco_cotiza()
    'Obtiene la cotización del dolar (tipo moneda = 1), para la fecha de la aplicación
    busco_cotiza = mFunObtengoUltimaCotizacion(1, 1, m_FechaSis)
End Function

Public Function busco_pasajero(cli As Long)
    busco_pasajero = 0
    tbCHECKIN.Index = "i_checkin_cli"
    tbCHECKIN.Seek "=", cli
    If Not tbCHECKIN.NoMatch Then
        If tbCHECKIN("nrocorrcli") = cli Then
            busco_pasajero = tbCHECKIN("nrohab")
        End If
    End If
End Function

Public Function busco_operacion(Opr As Integer)
    'Busco una operacion del sistema
    busco_operacion = False
    tbSISTEMA_OPERACIONES.Index = "pk_Opr"
    tbSISTEMA_OPERACIONES.Seek "=", Opr
    If Not tbSISTEMA_OPERACIONES.NoMatch Then
        busco_operacion = True
    End If
End Function

Public Function busco_documentoTF(tipoDoc As Byte, NroDoc As Long)
    'Busca cualquier tipo de documento, desde facturas (1 al 4) hasta devoluciones (5 al 8)
    busco_documentoTF = False
    tbCABEZAL.Index = "i_cabezal"
    tbCABEZAL.Seek "=", tipoDoc, NroDoc
    If Not tbCABEZAL.NoMatch Then   'existe
        busco_documentoTF = True
    End If
End Function

Public Function funBuscoBloqueoTF(hab As Long, nroCorrBloq As Long)
    'Busco un bloqueo determinado
    funBuscoBloqueoTF = False
    tbBLOQUEO_HAB.Index = "pk_bloqueo_hab"
    tbBLOQUEO_HAB.Seek "=", hab, nroCorrBloq
    If Not tbBLOQUEO_HAB.NoMatch Then   'existe
        funBuscoBloqueoTF = True
    End If
End Function

Public Function mFunNombreTitularReserva(res As Long)
    '-----------------------------------------------------------------
    'Devuelve nombre completo del titular de una reserva. (no anulada)
    '-----------------------------------------------------------------
    mFunNombreTitularReserva = ""
    If busco_reservaTF(res) Then
        mFunNombreTitularReserva = _
        tbRESERVAS("primer_ape_titular") & " " & _
        tbRESERVAS("segundo_ape_titular") & " " & _
        tbRESERVAS("primer_nom_titular") & " " & _
        tbRESERVAS("segundo_nom_titular")
    End If
End Function

Public Function obtengo_nombre_pasajero(cli As Long)
    obtengo_nombre_pasajero = ""
    If busco_clienteTF(cli) Then
        obtengo_nombre_pasajero = tbCLIENTES("nombre_completo_titular")
    End If
End Function

Public Function mFunPosicionoParaGrabar(formulario As Byte, conf As Byte)
    'Se utiliza para grabar los valores configurables, de
    'los formularios que utilizan este servicio.
    mFunPosicionoParaGrabar = False
    tbSISTEMA_CONF_FORMULARIOS.Index = "pk_CodFormulario"
    tbSISTEMA_CONF_FORMULARIOS.Seek "=", formulario, conf
    If Not tbSISTEMA_CONF_FORMULARIOS.NoMatch Then  'existe
        mFunPosicionoParaGrabar = True
    End If
End Function

Public Function mFunBuscoDescMsg(tipoMsg As Byte, codMsg As Integer) As String
    'Busco un mensaje en la tabla de mensajes y devulevo la descripción
    tbSISTEMA_MENSAJES.Index = "pk_msg"
    tbSISTEMA_MENSAJES.Seek "=", tipoMsg, codMsg
    If Not tbSISTEMA_MENSAJES.NoMatch Then
        mFunBuscoDescMsg = tbSISTEMA_MENSAJES("descMsg")
    End If
End Function

Public Function mFunBuscoDescPais(codPais As Integer) As String
    'Devulevo el nombre de un país
    mFunBuscoDescPais = ""
    tbPAISES.Index = "i_pais"
    tbPAISES.Seek "=", codPais
    If Not tbPAISES.NoMatch Then
        mFunBuscoDescPais = tbPAISES("descri_pais")
    End If
End Function

Public Function mFunBuscoDescEstadoCivil(codEstadoCivil As Byte) As String
    'Devuelvo la descripción del estado civil
    mFunBuscoDescEstadoCivil = ""
    tbESTADO_CIVIL.Index = "pk_estadoCivil"
    tbESTADO_CIVIL.Seek "=", codEstadoCivil
    If Not tbESTADO_CIVIL.NoMatch Then
        mFunBuscoDescEstadoCivil = tbESTADO_CIVIL("descEstadoCivil")
    End If
End Function

Public Function mFunBuscoTarifaHab(tipo As Long) As Double
    'Obtengo la tarifa correspondiente del tipo de habitación pasado como parametro.
    
    'busco tipo de habitación
    If busco_tipo_habTF(tipo) Then
        'obtengo tarifa
        mFunBuscoTarifaHab = tbTIPO_HABITACIONES("tarifa")
    End If
End Function

Public Sub subInicializoControlData(controlData As Object)
    'Inicializa el control data pasado como parámetro
    'con la base de datos que utiliza la aplicación
    
    controlData.Connect = cContraseñaBD 'establece la contraseña de la base de datos
    controlData.DatabaseName = vardir
End Sub

Public Function mfunObtengoDatosCli(tipoDato As Byte, cli As Long) As String
    '----------------------------------------------------------------------------
    'Parámetros.
    '   Entrada [tipoDato]  1 = fecha nac
    '                       2 = documento
    '                       3 = estadocivil
    '                       4 = nacionalidad
    '                       5 = tipo de documento
    '           [cli]       cliente
    '-----------------------------------------------------------------------------
    Dim tablaCli As Recordset
    Dim tablaEstadoCivil As Recordset
    Dim tablaNacionalidad As Recordset
    Dim tablaConstantes As Recordset
    
    Set tablaCli = tbCLIENTES
    
    'por defecto devuelvo empty
    mfunObtengoDatosCli = ""
    tablaCli.Index = "iclie_nrocorr"
    tablaCli.Seek "=", cli
    If Not tablaCli.NoMatch Then
        'determino el dato a devolver
        Select Case tipoDato
            Case 1  'fecha nac
                 If Not IsNull(tablaCli("fecha_nac_titular")) Then mfunObtengoDatosCli = tablaCli("fecha_nac_titular")
            Case 2  'documento
                If Not IsNull(tablaCli("documento_titular")) Then mfunObtengoDatosCli = tablaCli("documento_titular")
            Case 3  'estadocivil
                Set tablaEstadoCivil = tbESTADO_CIVIL
                    tablaEstadoCivil.Index = "pk_estadoCivil"
                    tablaEstadoCivil.Seek "=", tablaCli("estado_civil_titular")
                    If Not tablaEstadoCivil.NoMatch Then
                        mfunObtengoDatosCli = tablaEstadoCivil("descEstadoCivil")
                    End If
                Set tablaEstadoCivil = Nothing
                
            Case 4  'nacionalidad
                Set tablaNacionalidad = tbNACIONALIDADES
                    tablaNacionalidad.Index = "i_nacionalidad"
                    tablaNacionalidad.Seek "=", tablaCli("nacionalidad_titular")
                    If Not tablaNacionalidad.NoMatch Then
                        mfunObtengoDatosCli = tablaNacionalidad("descri_nacionalidad")
                    End If
                Set tablaNacionalidad = Nothing
            Case 5  'tipo de documento
                Set tablaConstantes = tbSISTEMA_CONSTANTES
                    tablaConstantes.Index = "pkConst"
                    tablaConstantes.Seek "=", tablaCli("tipoDocu_titular")
                    If Not tablaConstantes.NoMatch Then
                        mfunObtengoDatosCli = tablaConstantes("descConst")
                    End If
                Set tablaConstantes = Nothing
        End Select
    End If
        
    Set tablaCli = Nothing
End Function

Public Function mFunBuscoNombreEmpresa(nroEmp As Long) As String
    '-------------------------------------------------------------------------
    'Devuelve el nombre de una empresa existente.
    'Parámetros:
    '   Entrada:    [nroEmp]    número de empresa de la cual se quiere buscar
    '                           el nombre.
    '   Salida:     nombre de la empresa
    '--------------------------------------------------------------------------
    'declaro variables de archivo
    Dim tbEmp As Recordset
    Set tbEmp = tbEMPRESAS
    
    'busco empresa
    tbEmp.Index = "i_empresa"
    tbEmp.Seek "=", nroEmp
    If Not tbEmp.NoMatch Then
        'la empresa existe
        mFunBuscoNombreEmpresa = tbEmp("NomEmp")
    Else
        'la empresa no existe
        mFunBuscoNombreEmpresa = Empty
    End If
    
    Set tbEmp = Nothing
End Function

Public Function mFunBuscoNombreTarjetaCredito(nroTar As Integer) As String
    '-------------------------------------------------------------------------
    'Devuelve el nombre de una tarjeta de crédito existente.
    'Parámetros:
    '   Entrada:    [nroTar]    número de la tarjeta de crédito que se quiere buscar
    '                           el nombre
    '   Salida:     nombre de la la tarjeta de crédito
    '--------------------------------------------------------------------------
    'declaro variables de archivo
    Dim tbTar As Recordset
    Set tbTar = tbTARJETAS
    
    'busco empresa
    tbTar.Index = "pkTarjetas"
    tbTar.Seek "=", nroTar
    If Not tbTar.NoMatch Then
        'la empresa existe
        mFunBuscoNombreTarjetaCredito = tbTar("descTarjeta")
    Else
        'la empresa no existe
        mFunBuscoNombreTarjetaCredito = Empty
    End If
    
    Set tbTar = Nothing
End Function

Public Function mfunExisteCliente(tipoCli As Byte, nrocli As Long) As Boolean
    '----------------------------------------------------------------------------
    'Los clientes del hotel pueden ser de dos tipo: pax o empresas-agencias.
    'Esta función determina si existe el cliente, dependiendo del tipo del mismo.
    '-----------------------------------------------------------------------------
    'Parámetros.
    '   Entrada:    [tipoCli]   determina el tipo del cliente
    '                   0 = pax
    '                   1 = empresa
    '               [nroCli]    número del cliente a buscar
    '
    '   Salida:     True, si existe el cliente (en EMPRESAS o CLIENTES)
    '               False, si el cliente no existe
    '-------------------------------------------------------------------------------
    'declaro variables para utilizar tablas
    Dim tbCli As Recordset
    Dim tbEmp As Recordset
    
    Set tbCli = tbCLIENTES
    Set tbEmp = tbEMPRESAS
    'por defecto asumo que no existe el cliente
    mfunExisteCliente = False
    
    'determino que tipo de cliente estoy buscando
    Select Case tipoCli
        Case 0
            tbCli.Index = "iclie_nrocorr"
            tbCli.Seek "=", nrocli
            If Not tbCli.NoMatch Then
                'existe el cliente en archivo clientes
                mfunExisteCliente = True
            End If
        Case 1
            tbEmp.Index = "i_empresa"
            tbEmp.Seek "=", nrocli
            If Not tbEmp.NoMatch Then
                'existe el cliente en archivo empresas
                mfunExisteCliente = True
            End If
    End Select
    
    Set tbCli = Nothing
    Set tbEmp = Nothing
End Function

Public Function funObtengoDescSisConstantes(tipoConst As Byte, nroConst As Integer) As String
    '--------------------------------------------------------------------------------
    'Devuelvo el valor de una constante,almacenada en el archivo SISTEMA_CONSTANTES
    '--------------------------------------------------------------------------------
    'Parámetros.
    '   Entrada [tipoConst] tipo de constantes a buscar
    '           [nroConst]  número de constante a buscar
    '
    '   Salida  descripción de la constantes.
    '--------------------------------------------------------------------------------
    'declaro variables para utilizar archivo de constantes.
    Dim tbConst As Recordset
    Set tbConst = tbSISTEMA_CONSTANTES
    tbConst.Index = "pkConst"
    tbConst.Seek "=", tipoConst, nroConst
    If Not tbConst.NoMatch Then
        funObtengoDescSisConstantes = tbConst("descConst")
    Else
        funObtengoDescSisConstantes = Empty
    End If
    Set tbConst = Nothing
End Function

Public Function mFunObtengoFechaAlojaHab(hab As Long, tipoFecha As Byte) As Date
    '------------------------------------------------------------------------------
    'Devuelve la fecha de entrada o salida de una habitación alojada en el hotel
    '------------------------------------------------------------------------------
    'Parámetros.
    '   Entrada.    [hab] número de habitación con la cual estoy trabajando
    '               [tipoFecha] tipo de fecha que deseo devolver
    '               0 = fecha de entrada
    '               1 = fecha de salida
    '
    '   Salida      si [tipoFechaEntrada] = 0 entonces devuelvo tbCHECKIN("fCheckDes")
    '               si [tipoFechaEntrada] = 1 entonces devuelvo tbCHECKIN("fCheckHas")
    '-------------------------------------------------------------------------------
    'declaro variables para acceder a tabla
    Dim tbCheck As Recordset
    Set tbCheck = tbCHECKIN
    tbCheck.Index = "i_checkin"
    tbCheck.Seek ">=", hab, 0
    If Not tbCheck.NoMatch Then
        If tbCheck("nroHab") = hab Then
            'accedí a la habitación
            'determino que dato devuelvo
            Select Case tipoFecha
                Case 0
                    mFunObtengoFechaAlojaHab = tbCheck("fCheckDes")
                Case 1
                    mFunObtengoFechaAlojaHab = tbCheck("fCheckHas")
            End Select
        End If
    End If
    Set tbCheck = Nothing
End Function

Public Function mFunObtengoTotPaxAlojadosHab(hab As Long) As Integer
    '------------------------------------------------------------------------------
    'Cuenta la cantidad de pasajeros alojados en una habitación actualmente.
    '------------------------------------------------------------------------------
    'Parámetros.
    '   Entrada:    [hab]   nro de habitación
    '   Salida      total   de pasajeros alojados en una habitación del hotel
    '                       actualmente.
    '-------------------------------------------------------------------------------
    Dim rstHab As Recordset
    Dim qdfHab As QueryDef
    Dim consulta  As String
    consulta = "select * from checkin where nrohab= " & hab
    'ejecuto consulta
    Set qdfHab = bdHOTEL.CreateQueryDef("")
    qdfHab.SQL = consulta
    Set rstHab = qdfHab.OpenRecordset(dbOpenSnapshot)
    'me muevo al último registro para poder utilizar la propiedad recordcount
    If rstHab.RecordCount > 0 Then
        rstHab.MoveLast
        mFunObtengoTotPaxAlojadosHab = rstHab.RecordCount
    Else
        mFunObtengoTotPaxAlojadosHab = 0
    End If

    Set qdfHab = Nothing
    Set rstHab = Nothing
End Function
