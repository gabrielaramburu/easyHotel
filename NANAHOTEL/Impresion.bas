Attribute VB_Name = "Impresion"
Option Explicit

'***************************************************************************
'
'   Contiene procedimientos y funciones relativos a la impresión de informes
'
'***************************************************************************
Public Sub mSubCargoImpresorasInstaladas(comboImp As ComboBox)
    '-------------------------------------------------------------------------------------
    'Carga en el combo que se pasa como parámetros las impresoras instaladas del sistema
    '-------------------------------------------------------------------------------------
    'Parámetros.
    '       Entrada [comboImp] combo en el cual se cargan las impresoras
    '
    '-------------------------------------------------------------------------------------
    Dim impre As Printer
    For Each impre In Printers
        comboImp.AddItem impre.DeviceName
    Next
End Sub

Public Function mFunImpRegistro() As String
    '---------------------------------------------------------------------
    'Devuelve el nombre del hotel al cual pertenece la aplicación.
    '---------------------------------------------------------------------
    'Parámetros:
    '   Salida: si es una versión registrada, devuelve el nombre del hotel
    '           si es una versión demo devulve string correspondiente
    '---------------------------------------------------------------------
    On Error Resume Next
    'declaración de variables para poder utilizar los métodos
    'de la biblioteca Licencia.DLL, los cuales proveen información acerca de la
    'aplicación.
    
    Dim biblioLicenciaInfApli As InformacionApli
    Set biblioLicenciaInfApli = New InformacionApli
    
    mFunImpRegistro = biblioLicenciaInfApli.mFunObtenerLicenciaApli(idApli, tbSISTEMA_LICENCIA, 2)
    'tipoInf = 2 corresponde al dato a devolver, en este caso el nombre del hotel.
    
    Set biblioLicenciaInfApli = Nothing
End Function

Public Sub mSubInicializoCamposOrden(ultimoCampo As Byte)
    '-------------------------------------------------------------------------------
    'Inicializo los campo de ordenación utilizados en el reporte.
    'Es necesario ejecutar este procedimiento después de cada ejecución de algún
    'listado ya que las aplicacines trabaja con un solo control Crystal, por lo que es
    'necesario inicializarlo para poder realizar un nuevo listado.
    '--------------------------------------------------------------------------------
    'Parámetros.
    '   Entrada: [ultimoCampo] número del último campo utilizada en el último
    '                               reporte ejecutado.
    '--------------------------------------------------------------------------------
    On Error Resume Next
    Dim i As Byte
    i = 0
    Do While i <= ultimoCampo
        frmMAIN.CrystalReport1.SortFields(i) = ""
        i = i + 1
    Loop
End Sub

Public Sub mSubInicializoFormulas(ultimaFor As Integer)
    '------------------------------------------------------------------------------
    'Inicializo las formulas utilizadas por el último reporte ejecutado.
    'Es necesario ejecutar este procedimiento después de cada ejecución de algún
    'listado ya que las aplicaciones trabajan con un solo control Crystal, por lo que es
    'necesario inicializarlo para poder realizar un nuevo listado.
    '--------------------------------------------------------------------------------
    'Parámetros.
    '   Entrada:    [ultimaFormula] número de la última formula utilizada
    '               en el último reporte ejecutado.
    '--------------------------------------------------------------------------------
    On Error Resume Next
    Dim i As Byte
    i = 0
    Do While i <= ultimaFor
        frmMAIN.CrystalReport1.Formulas(i) = ""
        i = i + 1
    Loop
End Sub

Public Function mFunFormatoFecha(fecha As Date, tipoFormato As Byte) As String
    '---------------------------------------------------------------------------
    'Devuelve un string, correspondiente a una fecha en un formato determinado.
    '---------------------------------------------------------------------------
    'Parámetros.
    '   Entrada:    [fecha] fecha que se quiere convertir
    '               [tipoFormato] formato que se le quiere dar a la fecha
    '                   1 = dd/mm/aaaa
    '----------------------------------------------------------------------------
    On Error Resume Next
    Select Case tipoFormato
        Case 1  'dd/mm/aaaa
            mFunFormatoFecha = Format(fecha, "dd/mm/yyyy")
    End Select
End Function

Public Function mFunImpVersionAplicacion() As String
    '-----------------------------------------------------------------------------
    'Devuelve el nombre del programa y la versión que se está utilizando
    'Parámetros:
    '   Salida: titulo programa + versión
    '           El titulo del programa se configura dentro del menú proyecto.
    '           La versión esta formada por tres componentes.
    '------------------------------------------------------------------------------
    On Error Resume Next
    mFunImpVersionAplicacion = App.Title & " " & _
                            App.Major & "." & _
                            App.Minor & "." & _
                            App.Revision & "."
End Function


Public Function mfunAplicoConfImp(tipoLis As Byte, codLis As Integer) As Byte
    '----------------------------------------------------------------------
    'Establece la confgiguración del listado y realiza los pasos previos
    'necesario de acuerdo a dicha configuración.
    '----------------------------------------------------------------------
    'Parámetros.
    '           [tipoLis]       tipo del listado que se va a imprimir
    '           [codLis]        código del listado
    '
    '
    '   Salida  si se confirma la impresión se devuelve 1
    '           si se cancela la impresión se devuelve 0
    '-----------------------------------------------------------------------
    
    'declaración de variable para utilizar biblioteca de impresión
    Dim biblioSeleccion As SeleccionImpre
    Dim biblioImpGral As ImpresionGeneral

    Set biblioSeleccion = New SeleccionImpre
    Set biblioImpGral = New ImpresionGeneral
    
    Dim ImpDelLis As String
    Dim permitirSeleccionarImp As Byte
    Dim mostrarVistaPrevia As Byte
    Dim mostrarMsgConfi As Byte
    
    Dim impAUtilizar As String  'impresora que se utiliza finalmente para emitir el listado
                                'puede ser o no la impresora preestablecida para el listado.
                                'dependerá de:
                                '   si la misma se cambia en el cuadro de selección de impresoras
                                '   si la impresora preestablecida sigue estando instalada en el sistema
                                
    
    'por defecto asumo que se emite el listado
    mfunAplicoConfImp = 1
    
    'obtengo datos del reporte
    ImpDelLis = mFunObtengoDatosListados(tipoLis, codLis, 1)
    mostrarVistaPrevia = mFunObtengoDatosListados(tipoLis, codLis, 2)
    permitirSeleccionarImp = mFunObtengoDatosListados(tipoLis, codLis, 3)
    mostrarMsgConfi = mFunObtengoDatosListados(tipoLis, codLis, 4)
    
    'verifico si tengo que mostrar cuadro de seleción de impresora
    If permitirSeleccionarImp = 1 Then
        'muestro cuadro de seleccion de impresora, por defecto muestro la impresora del listado.
        impAUtilizar = biblioSeleccion.mFunSeleccionoImpresora(ImpDelLis)
        If impAUtilizar <> "" Then
            'se seleccionó una impresora (boton aceptar)
        Else
            'no se seleccionó una impresora (boton cancelar)
            'por lo tanto no continúo con el proceso de impresión
            mfunAplicoConfImp = 0   'cancelo impresión
        End If
    Else
        'no tengo que mostrar cuadro de selección de impresoras
        'verifico si la impresora del listado existe
        If Not biblioImpGral.mFunExisteImpresoraInstalada(ImpDelLis) Then
            'si no existe
            'verifico si hay impresoras instaladas
            If biblioImpGral.mFunCantidadImpresorasInstaladas > 0 Then
                'establesco com impresora del listado a la impresora predeterminada del sistema
                impAUtilizar = Printer.DeviceName
            Else
                'no se puede emitir el listado ya que no hay impresoras instaladas
                mSubMensaje 3, 7
                mfunAplicoConfImp = 0   'cancelo impresión
            End If
        End If
    End If
    
    'verifico si continúo con la impresión después de seleccionar impre
    If mfunAplicoConfImp = 1 Then
        'determino si tengo que mostrar mensaje de confirmación para seguir
        If mostrarMsgConfi = 1 Then
            'muestro mensaje de confirmación de impresión
            If Not mFunMensaje(3, 8) Then
                'no confirma impresión
                mfunAplicoConfImp = 0 'cancelo impresión
            End If
        End If
   End If
   
   'verifico si continúo con la impresión después del mensaje de confirmación
   If mfunAplicoConfImp = 1 Then
        'configuro otros aspectos del listado
        'determino si muestro vista previa
        Select Case mostrarVistaPrevia
            Case 1  'muestro vista prvia
                frmMAIN.CrystalReport1.Destination = crptToWindow
                frmMAIN.CrystalReport1.WindowState = crptMaximized  'ventana maximizada
            Case 0  'directo a la impresora
                frmMAIN.CrystalReport1.Destination = crptToPrinter

        End Select
    End If
    
    Set biblioSeleccion = Nothing
    Set biblioImpGral = Nothing
End Function

Public Function mFunObtengoDatosListados(tipoLis As Byte, _
                                        codLis As Integer, tipoDatoDev As Byte) As Variant
    '-------------------------------------------------------------------------------------
    'Dado un listado, obtiene información específica del mismo, la cual se encuntra en el
    'archivo SISTEMA_LISTADOS
    '-------------------------------------------------------------------------------------
    'Parámetros:
    '   Entrada
    '           [tipoLis]       tipo del listado del cual se quiere obtener información
    '                           1 = facturas
    '                           2 = varios
    '                           3 = perfiles
    '                           4 = nocrystal
    '           [codLis]        código del listado
    '           [tipoDatoDev]   tipo del dato a devolver
    '                           1 = impresora a utilizar
    '                           2 = mostrar vista previa
    '                           3 = permitir seleccionar impresora
    '                           4 = mostrar mensaje de confirmación
    '
    '   Salida  si [tipoDatoDev] = 1 tipo string
    '                            = 2 tipo byte (1 permite, 0 no permite)
    '                            = 3 tipo byte (1 permite, 0 no permite)
    '                            = 4 tipo byte (1 permite, 0 no permite)
    '
    '           si no existe listado devuleve tipo byte (2)
    '           si no se paso [tipoDatoDev] correcto devulve tipo byte (3)
    '---------------------------------------------------------------------------------------
       
    'declaro variables pra utilizar la tabla SISTEMA_LISTADOS
    Dim tablaSisLis As Recordset
    Set tablaSisLis = tbSISTEMA_LISTADOS
    
    'busco listado
    tablaSisLis.Index = "pk_listados"
    tablaSisLis.Seek "=", tipoLis, codLis
    If Not tablaSisLis.NoMatch Then
        'si existe listado
        'determino que valor devuelvo
        Select Case tipoDatoDev
            Case 1  'impresora
                If IsNull(tablaSisLis("impreLis")) Then
                    mFunObtengoDatosListados = ""
                Else
                    mFunObtengoDatosListados = tablaSisLis("impreLis")
                End If
            Case 2  'vista previa
                mFunObtengoDatosListados = tablaSisLis("mostrarVistaPrevia")
            Case 3  'permitir seleccionar impresora
                mFunObtengoDatosListados = tablaSisLis("seleccionarImpLis")
            Case 4  'mostrar mensaje confirmación
                mFunObtengoDatosListados = tablaSisLis("mensajeConfLis")
            Case Else
                'no existe el dato a devolver
                mFunObtengoDatosListados = 3
        End Select
    Else
        'no existe listado
        mFunObtengoDatosListados = 2
    End If
    Set tablaSisLis = Nothing
End Function


