Attribute VB_Name = "ControlDeLicencia"
Option Explicit

'En �ste m�dulo se encuntra la funci�n que determina si la aplicaci�n
'es una aplicaci�n v�lida.

'Esta funci�n utiliza componentes implementados en Licencia.dll

'Una aplicaci�n v�lida pude ser de dos tipos distintos:
'a) una versi�n demo, la cual posee un peri�do de evaluaci�n
'   dentro del cual se puede ejecutar
'b) una versi�n registrada, la cual se ejecuta sin l�mites de tiempo
'   pero solo en la m�quina para la cual se adquiri� la licencia.

'Si la aplicaci�n es una versi�n demo, se valida que:
'   -   el per�odo de evaluaci�n no halla finalizado
'   -   la fecha del sistema sea choerente, es decir no se halla retrocedido

'Si la aplicaci�n es una versi�n registrada, se valida que:
'   -   se est� ejecutando en la m�quina para la cual se adquiri� la licencia.

'Adem�s de �stos controles espec�ficos para cada tipo de versi�n, se valida simepre:
'   -   existencia y contendio del archivo de identificaci�n id

'La primera vez que se ejecuta la aplicaci�n, se inicializa la misma, es decir,
'se crea una registro en la tabla licencia, con informaci�n referente a la aplicaci�n.


'Declaraci�n de constantes
Private Const cNomArchivoId As String = "PerfilesUsuarios.id.txt"  'nombre del archivo donde se encuentra el
                                                        'Id de la aplicaci�n
                            
'Declaraci�n de variable para utilizar componente Licencia.dll
Private bibloControlVersionDemo As ControlVersionDemo
Private bibloAvisoErrores As AvisoErrores
Private funInicializarAplicacion As InicializarAplicacion
Private avisoFin As AvisoFinPeriodoDemo
Private avisoDemo As AvisoVersionDemo

Public Function mFunAplicacionValida() As Boolean
    'Determina si la aplicaci�n es una aplicaci�n v�lida.
    '-----------------------------------------------------------------------------
    '   Salida:
    '       True    si la aplicaci�n es v�lida y se puede ejecutar.
    '
    '       False   si la aplicaci�n no es v�lida o no est�n todas condiciones
    '               establecidas para que se pueda ejecutar.
    '------------------------------------------------------------------------------
    Dim codigoLicencia As Integer       'n�mero de Id de la palicaci�n
    Dim serieDisco As String            'serie del disco duro de la m�quina donde se
                                        'est� ejecutando la aplicaci�n
    Dim resIniApli As Integer           'bandera de control
    
    'por defecto asumo que NO se pude ejecutar la aplicaci�n
    mFunAplicacionValida = False
    
    'Creo nuevas instancias de las clases que contienen los componentes de c�digo
    'que se utilizan para validar la aplicaci�n.
    Set bibloControlVersionDemo = New ControlVersionDemo
    Set bibloAvisoErrores = New AvisoErrores
    
    'obtengo serie del disco duro
    serieDisco = bibloControlVersionDemo.funObtengoSerieDisco("C:\")
    'obtengo c�digo de Id aplicaci�n
    idApli = bibloControlVersionDemo.funObtengoIdAplicacion(App.Path & "\" & cNomArchivoId)

    'verifico tipo de licencia
    codigoLicencia = bibloControlVersionDemo.funControloLicenciaAplicacion(App.Path & "\" & cNomArchivoId, _
                                                                            tbSISTEMA_LICENCIA, _
                                                                              Date, _
                                                                            serieDisco)

    'eval�o el resutado que devuelve la funci�n de control de licencia.
    Select Case codigoLicencia
        Case 621    'es una versi�n demo
            'muestro aviso de versi�n demo
            subMuestroAvisoVersionDemo
            'esta varible es utilizada cuando salgo de la aplicaci�n para determinar
            'si se muestra o no, nuevamente el mensaje de versi�n demo.
            gEsUnaVersionDemo = True
          
            'es necesario actualizar la fecha de ejecuci�n
            If funActualizoFechaEjecucion Then
                'se puede ejecutar la aplicaci�n
                mFunAplicacionValida = True
            End If
            
        Case 622    'es una versi�n registrada
            'se puede ejecutar la aplicaci�n
            mFunAplicacionValida = True
            
        Case 514    'finaliz� el per�odo de evaluaci�n
            'muestro aviso de finalizaci�n del per�odo
            subMuestroAvisoFinEvaluacion
            'desactivo control de usuarios
            subDesactivoControlDeUsuarios
            
        Case 513    'primera vez que ejecuto la aplicaci�n o error
            'No se encontr� un registro con la clave primaria
            'igual al n�mero de identificaci�n de la aplicaci�n que estoy ejecutando.
            
            'ejecuto inicializaci�n de la aplicaci�n
            Set funInicializarAplicacion = New InicializarAplicacion
                resIniApli = funInicializarAplicacion.funInicializarAplicacion(tbSISTEMA_LICENCIA, _
                                                                            idApli, _
                                                                            Date)
                'eval�o el resultado de inicializar la aplicaci�n
                Select Case resIniApli
                    Case 0
                        'despu�s de inicializar la aplicaci�n muestro aviso de versi�n demo.
                        subMuestroAvisoVersionDemo
                        
                        'esta varible es utilizada cuando salgo de la aplicaci�n para determinar
                        'si se muestra o no, nuevamente el mensaje de versi�n demo.
                        gEsUnaVersionDemo = True
          
                        'se puede ejecutar la aplicaci�n
                        mFunAplicacionValida = True
                        
                    Case Else
                        'problemas al inicalizar la aplicaci�n
                        bibloAvisoErrores.propMsgError = "Error: " & resIniApli & Chr(10) & _
                                                            "Problemas al inicializar la aplicaci�n."
                        bibloAvisoErrores.propDescMsgError = "El proceso de inicializaci�n de la aplicaci�n, " & _
                                                            "no se pudo realizar correctamente. Si es la primera vez que " & _
                                                            "ejecuta esta aplicaci�n, es posible que el problema se origine " & _
                                                            "por una instalaci�n incorrecta. Si por el contrario ustes ya la ha ejecutado " & _
                                                            "con �xito anteriormente, es muy problable que halla problemas con la base " & _
                                                            "de datos de la aplicaci�n o con alg�n componente del hardware de su equipo."
                        bibloAvisoErrores.propContactarse = "Para solucionar el problema comun�quese con nosotros."
                        bibloAvisoErrores.MostrarMensaje
                End Select
            'destruyo instancia utilizada para inicializar la aplicaci�n
            Set funInicializarAplicacion = Nothing
            
        Case 515    'error
            'la fecha del sistema fue retrocedida.
            bibloAvisoErrores.propMsgError = "Error: " & codigoLicencia & Chr(10) & _
                                                "Problemas con la fecha del sistema."
            bibloAvisoErrores.propDescMsgError = "La fecha del sistema es menor a la fecha en la cual se " & _
                                                "ejecut� la aplicaci�n por �ltima vez. Como usted est� ejecutando, " & _
                                                "una versi�n de evaluaci�n no registrada, esta distorsi�n en la fecha, " & _
                                                "impide ejecutar la aplicaci�n."
            bibloAvisoErrores.propContactarse = "Para solucionar el problema comun�quese con nosotros."
            bibloAvisoErrores.MostrarMensaje
            
        Case 516    'error
            'no se puede identificar el tipo de licencia de la aplici�n
            bibloAvisoErrores.propMsgError = "Error: " & codigoLicencia & Chr(10) & _
                                                "No se puede identificar el tipo de licencia de la aplicaci�n."
            bibloAvisoErrores.propDescMsgError = "La informaci�n que se posee sobre la licencia de la aplicaci�n es " & _
                                                    "inchoerente. Esto puede deberse a problemas en la base de datos, " & _
                                                    "originados por una instalaci�n incorrecta o por problemas con alg�n componente " & _
                                                    "del  hardware de su equipo. " & _
                                                    "La aplicaci�n no posee informaci�n suficiente para poder ejecutarse."
            bibloAvisoErrores.propContactarse = "Para solucionar el problema comun�quese con nosotros."
            bibloAvisoErrores.MostrarMensaje
            
        Case 517    'error
            'la aplicaci�n se esta ejecutando en un disco duro distinto al que fue instalada.
            bibloAvisoErrores.propMsgError = "Error: " & codigoLicencia & Chr(10) & _
                                             "Versi�n registrada: incorrecta."
            bibloAvisoErrores.propDescMsgError = "La aplicaci�n no cumple con las condiciones necesarias para ejecutarse como versi�n registrada. " & _
                                                "La aplicaci�n se esta ejecutando en un disco duro distinto al que fue instalada y para el cual se " & _
                                                "adquiri� la licencia. Si esto no es as�, el error se debe a que ha ocurrido un fallo inesperado " & _
                                                "en las rutinas de seguridad."
            bibloAvisoErrores.propContactarse = "Para solucionar el problema comun�quese con nosotros."
            bibloAvisoErrores.MostrarMensaje
                                                
        Case 518    'error
            'no coinciden los tipos
            bibloAvisoErrores.propMsgError = "Error: " & codigoLicencia & Chr(10) & _
                                             "No coinciden los tipos."
            bibloAvisoErrores.propDescMsgError = "La informaci�n en el archivo Id de la aplicaci�n es erronea. " & _
                                                "Posiblemente el mismo fue modificado, alterando sus valores originales. " & _
                                                "La aplicaci�n no puede obtener informaci�n correcta para poder ejecutarse."
            bibloAvisoErrores.propContactarse = "Para solucionar el problema comun�quese con nosotros."
            bibloAvisoErrores.MostrarMensaje
            
        Case 519    'error
            'no existe el archivo
            bibloAvisoErrores.propMsgError = "Error: " & codigoLicencia & Chr(10) & _
                                             "No existe archivo Id de la aplicaci�n."
            bibloAvisoErrores.propDescMsgError = "El archivo Id de la aplicaci�n no se pudo localizar. " & _
                                                "El mismo fue cambiado de lugar dentro de su disco duro o eliminado. " & _
                                                "Si es la primera vez que ejecuta la aplicaci�n, entonces el problema se puede " & _
                                                "originar por una instalaci�n incorrecta."
            bibloAvisoErrores.propContactarse = "Para solucionar el problema comun�quese con nosotros."
            bibloAvisoErrores.MostrarMensaje
            
        Case 520    'error
            'error en ejecuci�n de algunas de las rutinas de la biblioteca
            bibloAvisoErrores.propMsgError = "Error: " & codigoLicencia & Chr(10) & _
                                             "Error no determinado."
            bibloAvisoErrores.propDescMsgError = "Se produjo un error inesperado al ejecutar las rutinas de seguridad. " & _
                                                "Imposible continuar con la ejecuci�n de la aplicaci�n."
            bibloAvisoErrores.propContactarse = "Para solucionar el problema comun�quese con nosotros."
            bibloAvisoErrores.MostrarMensaje
        
    End Select

    'Destruyo las instancias creadas.
    Set bibloControlVersionDemo = Nothing
    Set bibloAvisoErrores = Nothing
End Function

Private Function funActualizoFechaEjecucion() As Boolean
    'Actualiza la fecha de la �ltima ejecuci�n de una versi�n demo.
    '--------------------------------------------------------------------
    'Par�metros.
    '
    '   Salida: True    se pudo actualizar la �ltima fecha de ejecuci�n
    '           False   no se pudo actualizar
    '--------------------------------------------------------------------
    Set bibloControlVersionDemo = New ControlVersionDemo
        If bibloControlVersionDemo.funActualizarUltimoDiaEjecuci�nVD(tbSISTEMA_LICENCIA, idApli, Date) Then
            'la fecha se actualiz� correctamente
            funActualizoFechaEjecucion = True
        Else
            'se produjo un error al actualizar la fecha del sistema
            funActualizoFechaEjecucion = False
            Set bibloAvisoErrores = New AvisoErrores
            bibloAvisoErrores.propMsgError = "Error: " & 712 & Chr(10) & _
                                             "No se puede actualizar fecha de �ltima ejecuci�n."
            bibloAvisoErrores.propDescMsgError = "El procedo de actualizaci�n de la �ltima fecha de ejecuci�n " & _
                                                "de la aplicaci�n, no se pudo realizar correctamente." & _
                                                "La aplicaci�n no podr� seguir ejecut�ndose."
            bibloAvisoErrores.propContactarse = "Para solucionar el problema comun�quese con nosotros."
            bibloAvisoErrores.MostrarMensaje
            Set bibloAvisoErrores = Nothing
        End If
    Set bibloControlVersionDemo = Nothing
End Function

Public Sub subMuestroAvisoVersionDemo()
    'Muestra un aviso que indica que la aplicaci�n es una versi�n demo.
    'NOTA: Este procedimiento se declara Public ya que es utilizado por el formulario frmMain
    
    'creo instancias
    Set avisoDemo = New AvisoVersionDemo
    Set bibloControlVersionDemo = New ControlVersionDemo
    
    'establesco propiedades del forumlario de versi�n demo
    avisoDemo.AvisoVersionDemoPropTituloForm = "Versi�n de evaluaci�n(no registrada)."
    avisoDemo.AvisoVersionDemoPropNomAplicacion = "Perfiles"
    avisoDemo.AvisoVersionDemoPropSistemaAplicacion = "Para Window 9x/2000/XP/NT"
    avisoDemo.AvisoVersionDemoPropVersionAplicacion = "Version de evaluaci�n 1.0"
    avisoDemo.AvisoVersionDemoPropDiasDemos = bibloControlVersionDemo.funObtenerCantDiasAutorizadosVD(tbSISTEMA_LICENCIA, idApli)
    avisoDemo.AvisoVersionDemoPropPeriodoDeUso = "D�as utilizado: " & _
                                                    bibloControlVersionDemo.funObtenerCantDiasUtilizadosVD(tbSISTEMA_LICENCIA, Date, idApli) & _
                                                    " de su per�odo de " & avisoDemo.AvisoVersionDemoPropDiasDemos & " d�as."
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
    'Muestra un aviso que indica que la instalci�n no se puede ejecutar
    'porque el per�odo de evaluaci�n ha terminado.
    
    'creo instancias
    Set avisoFin = New AvisoFinPeriodoDemo
    Set bibloControlVersionDemo = New ControlVersionDemo
    
    'establesco todas las propiedades del objeto creado AvisoFinPeriodoDemo
    avisoFin.AvisoFinPeriodoDemoPropNomAplicacion = "Perfiles"
    avisoFin.AvisoFinPeriodoDemoPropSistemaAplicacion = "Para Window 9x/2000/XP/NT"
    avisoFin.AvisoFinPeriodoDemoPropVersionAplicacion = "Version de evaluaci�n 1.0"
    avisoFin.AvisoFinPeriodoDemoPropPeriodoTerminado = "D�as utilizado: " & _
                                                        bibloControlVersionDemo.funObtenerCantDiasUtilizadosVD(tbSISTEMA_LICENCIA, Date, idApli) & _
                                                        " de su per�odo de " & _
                                                        bibloControlVersionDemo.funObtenerCantDiasAutorizadosVD(tbSISTEMA_LICENCIA, idApli) & _
                                                        " d�as."
    avisoFin.AvisoFinPeriodoDemoPropTituloForm = "Fin del per�odo de evaluaci�n."
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
    'Esta aplicaci�n tiene la caracter�stica que luego que al misma se instala,
    'modfica una campo de la tabla par�metros, el cual tiene com objetivo, indicarle
    'a la aplicaci�n principal, que esta activado el control de usuarios.
    'Si est� activado el control de usuarios, la aplicaci�n principal requiere contrase�as y
    'autorizaci�n para ejecutar las diferentes opciones del sistema, seg�n lo establecido
    'por el administrador.
    'Cuando se deja de utilizar la versi�n demo, debido que finaliz� el per�odo de evaluaci�n,
    'es necesario desactivar este control, para que la aplicaci�n no se siga utiliznado
    'el sistema de control establecido.
    On Error Resume Next
    'modifico tabla par�metros
    tbSISTEMA_PARAMETROS.Edit
        'desactivo control
        tbSISTEMA_PARAMETROS("SisAdminTF") = 0
    tbSISTEMA_PARAMETROS.Update
End Sub
