Attribute VB_Name = "ControlDeLicencia"
Option Explicit

'En éste módulo se encuntra la función que determina si la aplicación
'es una aplicación válida.

'Esta función utiliza componentes implementados en Licencia.dll

'Una aplicación válida pude ser de dos tipos distintos:
'a) una versión demo, la cual posee un periódo de evaluación
'   dentro del cual se puede ejecutar
'b) una versión registrada, la cual se ejecuta sin límites de tiempo
'   pero solo en la máquina para la cual se adquirió la licencia.

'Si la aplicación es una versión demo, se valida que:
'   -   el período de evaluación no halla finalizado
'   -   la fecha del sistema sea choerente, es decir no se halla retrocedido

'Si la aplicación es una versión registrada, se valida que:
'   -   se esté ejecutando en la máquina para la cual se adquirió la licencia.

'Además de éstos controles específicos para cada tipo de versión, se valida simepre:
'   -   existencia y contendio del archivo de identificación id

'La primera vez que se ejecuta la aplicación, se inicializa la misma, es decir,
'se crea una registro en la tabla licencia, con información referente a la aplicación.


'Declaración de constantes
Private Const cNomArchivoId As String = "PerfilesUsuarios.id.txt"  'nombre del archivo donde se encuentra el
                                                        'Id de la aplicación
                            
'Declaración de variable para utilizar componente Licencia.dll
Private bibloControlVersionDemo As ControlVersionDemo
Private bibloAvisoErrores As AvisoErrores
Private funInicializarAplicacion As InicializarAplicacion
Private avisoFin As AvisoFinPeriodoDemo
Private avisoDemo As AvisoVersionDemo

Public Function mFunAplicacionValida() As Boolean
    'Determina si la aplicación es una aplicación válida.
    '-----------------------------------------------------------------------------
    '   Salida:
    '       True    si la aplicación es válida y se puede ejecutar.
    '
    '       False   si la aplicación no es válida o no están todas condiciones
    '               establecidas para que se pueda ejecutar.
    '------------------------------------------------------------------------------
    Dim codigoLicencia As Integer       'número de Id de la palicación
    Dim serieDisco As String            'serie del disco duro de la máquina donde se
                                        'está ejecutando la aplicación
    Dim resIniApli As Integer           'bandera de control
    
    'por defecto asumo que NO se pude ejecutar la aplicación
    mFunAplicacionValida = False
    
    'Creo nuevas instancias de las clases que contienen los componentes de código
    'que se utilizan para validar la aplicación.
    Set bibloControlVersionDemo = New ControlVersionDemo
    Set bibloAvisoErrores = New AvisoErrores
    
    'obtengo serie del disco duro
    serieDisco = bibloControlVersionDemo.funObtengoSerieDisco("C:\")
    'obtengo código de Id aplicación
    idApli = bibloControlVersionDemo.funObtengoIdAplicacion(App.Path & "\" & cNomArchivoId)

    'verifico tipo de licencia
    codigoLicencia = bibloControlVersionDemo.funControloLicenciaAplicacion(App.Path & "\" & cNomArchivoId, _
                                                                            tbSISTEMA_LICENCIA, _
                                                                              Date, _
                                                                            serieDisco)

    'evalúo el resutado que devuelve la función de control de licencia.
    Select Case codigoLicencia
        Case 621    'es una versión demo
            'muestro aviso de versión demo
            subMuestroAvisoVersionDemo
            'esta varible es utilizada cuando salgo de la aplicación para determinar
            'si se muestra o no, nuevamente el mensaje de versión demo.
            gEsUnaVersionDemo = True
          
            'es necesario actualizar la fecha de ejecución
            If funActualizoFechaEjecucion Then
                'se puede ejecutar la aplicación
                mFunAplicacionValida = True
            End If
            
        Case 622    'es una versión registrada
            'se puede ejecutar la aplicación
            mFunAplicacionValida = True
            
        Case 514    'finalizó el período de evaluación
            'muestro aviso de finalización del período
            subMuestroAvisoFinEvaluacion
            'desactivo control de usuarios
            subDesactivoControlDeUsuarios
            
        Case 513    'primera vez que ejecuto la aplicación o error
            'No se encontró un registro con la clave primaria
            'igual al número de identificación de la aplicación que estoy ejecutando.
            
            'ejecuto inicialización de la aplicación
            Set funInicializarAplicacion = New InicializarAplicacion
                resIniApli = funInicializarAplicacion.funInicializarAplicacion(tbSISTEMA_LICENCIA, _
                                                                            idApli, _
                                                                            Date)
                'evalúo el resultado de inicializar la aplicación
                Select Case resIniApli
                    Case 0
                        'después de inicializar la aplicación muestro aviso de versión demo.
                        subMuestroAvisoVersionDemo
                        
                        'esta varible es utilizada cuando salgo de la aplicación para determinar
                        'si se muestra o no, nuevamente el mensaje de versión demo.
                        gEsUnaVersionDemo = True
          
                        'se puede ejecutar la aplicación
                        mFunAplicacionValida = True
                        
                    Case Else
                        'problemas al inicalizar la aplicación
                        bibloAvisoErrores.propMsgError = "Error: " & resIniApli & Chr(10) & _
                                                            "Problemas al inicializar la aplicación."
                        bibloAvisoErrores.propDescMsgError = "El proceso de inicialización de la aplicación, " & _
                                                            "no se pudo realizar correctamente. Si es la primera vez que " & _
                                                            "ejecuta esta aplicación, es posible que el problema se origine " & _
                                                            "por una instalación incorrecta. Si por el contrario ustes ya la ha ejecutado " & _
                                                            "con éxito anteriormente, es muy problable que halla problemas con la base " & _
                                                            "de datos de la aplicación o con algún componente del hardware de su equipo."
                        bibloAvisoErrores.propContactarse = "Para solucionar el problema comuníquese con nosotros."
                        bibloAvisoErrores.MostrarMensaje
                End Select
            'destruyo instancia utilizada para inicializar la aplicación
            Set funInicializarAplicacion = Nothing
            
        Case 515    'error
            'la fecha del sistema fue retrocedida.
            bibloAvisoErrores.propMsgError = "Error: " & codigoLicencia & Chr(10) & _
                                                "Problemas con la fecha del sistema."
            bibloAvisoErrores.propDescMsgError = "La fecha del sistema es menor a la fecha en la cual se " & _
                                                "ejecutó la aplicación por última vez. Como usted está ejecutando, " & _
                                                "una versión de evaluación no registrada, esta distorsión en la fecha, " & _
                                                "impide ejecutar la aplicación."
            bibloAvisoErrores.propContactarse = "Para solucionar el problema comuníquese con nosotros."
            bibloAvisoErrores.MostrarMensaje
            
        Case 516    'error
            'no se puede identificar el tipo de licencia de la aplición
            bibloAvisoErrores.propMsgError = "Error: " & codigoLicencia & Chr(10) & _
                                                "No se puede identificar el tipo de licencia de la aplicación."
            bibloAvisoErrores.propDescMsgError = "La información que se posee sobre la licencia de la aplicación es " & _
                                                    "inchoerente. Esto puede deberse a problemas en la base de datos, " & _
                                                    "originados por una instalación incorrecta o por problemas con algún componente " & _
                                                    "del  hardware de su equipo. " & _
                                                    "La aplicación no posee información suficiente para poder ejecutarse."
            bibloAvisoErrores.propContactarse = "Para solucionar el problema comuníquese con nosotros."
            bibloAvisoErrores.MostrarMensaje
            
        Case 517    'error
            'la aplicación se esta ejecutando en un disco duro distinto al que fue instalada.
            bibloAvisoErrores.propMsgError = "Error: " & codigoLicencia & Chr(10) & _
                                             "Versión registrada: incorrecta."
            bibloAvisoErrores.propDescMsgError = "La aplicación no cumple con las condiciones necesarias para ejecutarse como versión registrada. " & _
                                                "La aplicación se esta ejecutando en un disco duro distinto al que fue instalada y para el cual se " & _
                                                "adquirió la licencia. Si esto no es así, el error se debe a que ha ocurrido un fallo inesperado " & _
                                                "en las rutinas de seguridad."
            bibloAvisoErrores.propContactarse = "Para solucionar el problema comuníquese con nosotros."
            bibloAvisoErrores.MostrarMensaje
                                                
        Case 518    'error
            'no coinciden los tipos
            bibloAvisoErrores.propMsgError = "Error: " & codigoLicencia & Chr(10) & _
                                             "No coinciden los tipos."
            bibloAvisoErrores.propDescMsgError = "La información en el archivo Id de la aplicación es erronea. " & _
                                                "Posiblemente el mismo fue modificado, alterando sus valores originales. " & _
                                                "La aplicación no puede obtener información correcta para poder ejecutarse."
            bibloAvisoErrores.propContactarse = "Para solucionar el problema comuníquese con nosotros."
            bibloAvisoErrores.MostrarMensaje
            
        Case 519    'error
            'no existe el archivo
            bibloAvisoErrores.propMsgError = "Error: " & codigoLicencia & Chr(10) & _
                                             "No existe archivo Id de la aplicación."
            bibloAvisoErrores.propDescMsgError = "El archivo Id de la aplicación no se pudo localizar. " & _
                                                "El mismo fue cambiado de lugar dentro de su disco duro o eliminado. " & _
                                                "Si es la primera vez que ejecuta la aplicación, entonces el problema se puede " & _
                                                "originar por una instalación incorrecta."
            bibloAvisoErrores.propContactarse = "Para solucionar el problema comuníquese con nosotros."
            bibloAvisoErrores.MostrarMensaje
            
        Case 520    'error
            'error en ejecución de algunas de las rutinas de la biblioteca
            bibloAvisoErrores.propMsgError = "Error: " & codigoLicencia & Chr(10) & _
                                             "Error no determinado."
            bibloAvisoErrores.propDescMsgError = "Se produjo un error inesperado al ejecutar las rutinas de seguridad. " & _
                                                "Imposible continuar con la ejecución de la aplicación."
            bibloAvisoErrores.propContactarse = "Para solucionar el problema comuníquese con nosotros."
            bibloAvisoErrores.MostrarMensaje
        
    End Select

    'Destruyo las instancias creadas.
    Set bibloControlVersionDemo = Nothing
    Set bibloAvisoErrores = Nothing
End Function

Private Function funActualizoFechaEjecucion() As Boolean
    'Actualiza la fecha de la última ejecución de una versión demo.
    '--------------------------------------------------------------------
    'Parámetros.
    '
    '   Salida: True    se pudo actualizar la última fecha de ejecución
    '           False   no se pudo actualizar
    '--------------------------------------------------------------------
    Set bibloControlVersionDemo = New ControlVersionDemo
        If bibloControlVersionDemo.funActualizarUltimoDiaEjecuciónVD(tbSISTEMA_LICENCIA, idApli, Date) Then
            'la fecha se actualizó correctamente
            funActualizoFechaEjecucion = True
        Else
            'se produjo un error al actualizar la fecha del sistema
            funActualizoFechaEjecucion = False
            Set bibloAvisoErrores = New AvisoErrores
            bibloAvisoErrores.propMsgError = "Error: " & 712 & Chr(10) & _
                                             "No se puede actualizar fecha de última ejecución."
            bibloAvisoErrores.propDescMsgError = "El procedo de actualización de la última fecha de ejecución " & _
                                                "de la aplicación, no se pudo realizar correctamente." & _
                                                "La aplicación no podrá seguir ejecutándose."
            bibloAvisoErrores.propContactarse = "Para solucionar el problema comuníquese con nosotros."
            bibloAvisoErrores.MostrarMensaje
            Set bibloAvisoErrores = Nothing
        End If
    Set bibloControlVersionDemo = Nothing
End Function

Public Sub subMuestroAvisoVersionDemo()
    'Muestra un aviso que indica que la aplicación es una versión demo.
    'NOTA: Este procedimiento se declara Public ya que es utilizado por el formulario frmMain
    
    'creo instancias
    Set avisoDemo = New AvisoVersionDemo
    Set bibloControlVersionDemo = New ControlVersionDemo
    
    'establesco propiedades del forumlario de versión demo
    avisoDemo.AvisoVersionDemoPropTituloForm = "Versión de evaluación(no registrada)."
    avisoDemo.AvisoVersionDemoPropNomAplicacion = "Perfiles"
    avisoDemo.AvisoVersionDemoPropSistemaAplicacion = "Para Window 9x/2000/XP/NT"
    avisoDemo.AvisoVersionDemoPropVersionAplicacion = "Version de evaluación 1.0"
    avisoDemo.AvisoVersionDemoPropDiasDemos = bibloControlVersionDemo.funObtenerCantDiasAutorizadosVD(tbSISTEMA_LICENCIA, idApli)
    avisoDemo.AvisoVersionDemoPropPeriodoDeUso = "Días utilizado: " & _
                                                    bibloControlVersionDemo.funObtenerCantDiasUtilizadosVD(tbSISTEMA_LICENCIA, Date, idApli) & _
                                                    " de su período de " & avisoDemo.AvisoVersionDemoPropDiasDemos & " días."
    avisoDemo.AvisoVersionDemoPropDerechos = "Copyright(c) 2000-2002" & Chr(10) & _
                                        "All Rights Reserved." & Chr(10) & _
                                        "Maldonado - Uruguay" & Chr(10) '& _
                                        '"www.chupacabrasventanita.com"
    avisoDemo.MostrarAvisoVersionDemo

    'destruyo instancias creadas
    Set avisoDemo = Nothing
    Set bibloControlVersionDemo = Nothing
End Sub

Private Sub subMuestroAvisoFinEvaluacion()
    'Muestra un aviso que indica que la instalción no se puede ejecutar
    'porque el período de evaluación ha terminado.
    
    'creo instancias
    Set avisoFin = New AvisoFinPeriodoDemo
    Set bibloControlVersionDemo = New ControlVersionDemo
    
    'establesco todas las propiedades del objeto creado AvisoFinPeriodoDemo
    avisoFin.AvisoFinPeriodoDemoPropNomAplicacion = "Perfiles"
    avisoFin.AvisoFinPeriodoDemoPropSistemaAplicacion = "Para Window 9x/2000/XP/NT"
    avisoFin.AvisoFinPeriodoDemoPropVersionAplicacion = "Version de evaluación 1.0"
    avisoFin.AvisoFinPeriodoDemoPropPeriodoTerminado = "Días utilizado: " & _
                                                        bibloControlVersionDemo.funObtenerCantDiasUtilizadosVD(tbSISTEMA_LICENCIA, Date, idApli) & _
                                                        " de su período de " & _
                                                        bibloControlVersionDemo.funObtenerCantDiasAutorizadosVD(tbSISTEMA_LICENCIA, idApli) & _
                                                        " días."
    avisoFin.AvisoFinPeriodoDemoPropTituloForm = "Fin del período de evaluación."
    avisoFin.AvisoFinPeriodoDemoPropExtension = True
    
    avisoFin.AvisoFinPeriodoDemoPropDerechos = "Copyright(c) 2000-2002" & Chr(10) & _
                                        "All Rights Reserved." & Chr(10) & _
                                        "Maldonado - Uruguay" & Chr(10) ' & _
                                        '"www.chupacabrasventanita.com"
    avisoFin.MostrarAvisoFinPeriodoDemo
    'destruyo instancias creadas
    Set avisoFin = Nothing
    Set bibloControlVersionDemo = Nothing
End Sub

Private Sub subDesactivoControlDeUsuarios()
    'Esta aplicación tiene la característica que luego que al misma se instala,
    'modfica una campo de la tabla parámetros, el cual tiene com objetivo, indicarle
    'a la aplicación principal, que esta activado el control de usuarios.
    'Si está activado el control de usuarios, la aplicación principal requiere contraseñas y
    'autorización para ejecutar las diferentes opciones del sistema, según lo establecido
    'por el administrador.
    'Cuando se deja de utilizar la versión demo, debido que finalizó el período de evaluación,
    'es necesario desactivar este control, para que la aplicación no se siga utiliznado
    'el sistema de control establecido.
    On Error Resume Next
    'modifico tabla parámetros
    tbSISTEMA_PARAMETROS.Edit
        'desactivo control
        tbSISTEMA_PARAMETROS("SisAdminTF") = 0
    tbSISTEMA_PARAMETROS.Update
End Sub
