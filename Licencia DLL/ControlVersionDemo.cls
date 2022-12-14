VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ControlVersionDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Utilizada para obtener el n?mero de serie del disco duro
Private Declare Function GetVolumeInformation Lib "Kernel32" _
    Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, _
                                    ByVal lpVolumeNameBuffer As String, _
                                    ByVal nVolumeNameSize As Long, _
                                    lpVolumeSerialNumber As Long, _
                                    lpMaximumComponentLength As Long, _
                                    lpFileSystemFlags As Long, _
                                    ByVal lpFileSystemNameBuffer As String, _
                                    ByVal nFileSystemNameSize As Long) As Long

Public Function funObtengoSerieDisco(unidad As String) As String
    'Obtengo el n?mero de serie del disco duro de la m?quina donde esjecuto
    'esta funci?n
    'Acci?n

    Dim lVSN As Long, n As Long, s1 As String, s2 As String
    Dim sTmp As String

    On Local Error Resume Next

    'Reservar espacio para las cadenas que se pasar?n al API
    s1 = String$(255, Chr$(0))
    s2 = String$(255, Chr$(0))

    n = GetVolumeInformation(unidad, s1, Len(s1), lVSN, 0, 0, s2, Len(s2))

    's1 ser? la etiqueta del volumen
    'lVSN tendr? el valor del Volume Serial Number (n?mero de serie del volumen)
    's2 el tipo de archivos: FAT, etc.

    'Convertirlo a hexadecimal para mostrarlo como en el Dir.
    sTmp = Hex$(lVSN)
    funObtengoSerieDisco = sTmp
End Function

Public Function funControloLicenciaAplicacion(archivo As String, _
                                                tabla As Recordset, _
                                                fechaActual As Date, _
                                                serieDisco As String) As Integer
    'Determina el tipo de licencia de una aplicaci?n.
    'Si la aplicaci?n es demo, muestra un mensaje de aviso y controla el per?odo de ejecuci?n.
    'Si la aplicaci?n esta registrada, val?da que sea una copia v?lida(n?mero de serie del disco)
    'Esta funci?n devuleve los siguientes c?digo que son tratados en la aplicaci?n cliente:
    
    '513: no existe informaci?n de la aplicaci?n
    'Se produce cuando no existe un registro con la clave igual al numeroId
    'en el archivo de licencias de la aplicaci?n.
    
    '514: se cumpli? el per?odo de evaluaci?n.
    
    '515: la fecha del sistema fue retrocedida
    'Se produce cuando la fecha de la ?ltima ejecuci?n del programa es mayor
    'a la fecha actual.
    
    '516: no se puede identificar el tipo de licencia de la aplici?n.
    'Ocurre cuando la informaci?n en la tabla de licencias es incoherente, debido a que se
    'modificaron los datos erroneamente en la tabla Lisencia o la instalaci?n no se realiz?
    'correctamente.
    
    '517: la aplicaci?n se esta ejecutando en un disco duro distinto al que fue instalada.
    
    '518: no coinciden los tipos
    'Se produce si el archivo contiene caracteres alfanum?ricos.
    
    '519: el archivo no existe
    'Se produce cuando el archivoId no existe.
    
    '520: se produjo un error de ejecuci?n
    
    '621: es una plicaci?n demo correcta
    '622: es una aplicaci?n registrda correcta
    
    On Error GoTo error
    
    Dim numeroId As Long
    
    'obtengo identificaci?n de la aplicaci?n desde archivo de identificacion
    numeroId = funObtengoIdAplicacion(archivo)
    If numeroId <> 518 And numeroId <> 519 Then
        'obtengo informaci?n de la aplicaci?n desde archivo de licencias
        tabla.Index = "pkLicencia"
        tabla.Seek "=", numeroId
        If Not tabla.NoMatch Then
            'verifico si la aplicaci?n es legal
            funControloLicenciaAplicacion = funAplicacionLegal(tabla, fechaActual, serieDisco)
        Else
            'no existe informaci?n de la aplicaci?n en archivo de licencias
            funControloLicenciaAplicacion = 513
        End If
    Else
        'no se pudo obtener el n?mero de identifcaci?n
        funControloLicenciaAplicacion = numeroId
    End If
Exit Function
error:
    'se produjo un error al ejecutar algunas de las funciones:
    '    funControlLicenciaAplicacion
    '        funAplicacionLegal
    funControloLicenciaAplicacion = 520
End Function

Public Function funObtengoIdAplicacion(archivo As String) As String
    'Obtiene el n?mero de Id de una aplicaci?n determinada.
    'Dicho n?mero de Id se almacena en un archivo de texto, que se debe
    'de encontrar en el directorio donde se ejecuta la aplicaci?n.
    'Si se produce un error se devuelve el n?mero de error producido.
    Dim numeroId As String
    On Error GoTo error
    
    'abro archivo para lectura
    Open archivo For Input As #1
    'si el archivo existe leo el n?mero de identificaci?n
    Line Input #1, numeroId
    'cierro el archivo
    Close #1
    funObtengoIdAplicacion = CLng(numeroId)
Exit Function
error:
    Select Case Err.Number
        Case 13 'no coinciden los tipos
            funObtengoIdAplicacion = 518
        Case 53 'el archivo no existe
            funObtengoIdAplicacion = 519
    End Select
    'inicializo el c?digo de error
    Err.Number = 0
End Function

Public Function funObtenerCantDiasUtilizadosVD(tabla As Recordset, fechaActual As Date, idAplicacion As Long) As Integer
    'Devuelve el total de d?as que van desde el d?a de la instalaci?n hasta el d?a de la
    'fecha del sistema.
    '-------------------------------------------------------------------------------------
    ' Par?metros.
    '   Entrada:
    '       [tabla]         tabla LICENCIA, donde se encuntra el d?a de instalci?n.
    '       [idAplicacion]  clave para acceder al registro de la tabla LICENCIA
    '       [fechaActual]   fecha actual del sistema
    '   Salida:
    '       Diferencia entre LICENCIA.AplicacionDiaInstalacion y fechActual
    '------------------------------------------------------------------------------------
    
    On Error Resume Next
    'posiciono en registro correspondiente
    tabla.Index = "pkLicencia"
    tabla.Seek "=", idAplicacion
    If Not tabla.NoMatch Then
        'obtengo valor y calculo diferencia
        funObtenerCantDiasUtilizadosVD = fechaActual - tabla("aplicacionDiaInstalacion")
    End If
End Function

Public Function funObtenerCantDiasAutorizadosVD(tabla As Recordset, idAplicacion As Long) As Integer
    'Devuelve la cantidad de d?as que se permite ejecutar la aplicaci?n, como versi?n demo.
    '-------------------------------------------------------------------------------------
    ' Par?metros.
    '   Entrada:
    '       [tabla]         tabla LICENCIA, donde se encuntra el d?a de instalci?n.
    '       [idAplicacion]  clave para acceder al registro de la tabla LICENCIA
    '   Salida:
    '       Valor del campo LICENCIA.AplicacionVDCantDias
    '------------------------------------------------------------------------------------
    On Error Resume Next
    'posiciono en registro correspondiente
    tabla.Index = "pkLicencia"
    tabla.Seek "=", idAplicacion
    If Not tabla.NoMatch Then
        'obtengo valor
        funObtenerCantDiasAutorizadosVD = tabla("aplicacionVDCantDias")
    End If
End Function

Public Function funActualizarUltimoDiaEjecuci?nVD(tabla As Recordset, idAplicacion As Long, fechaActual As Date) As Boolean
    'Actualiza el campo tabla.AplicacionVDUltimodiaEjecutada con la fecha actual del sistema.
    'El contenido de este campo es utilizado para validadar que la fecha de la m?quina no
    'se halla retrocedido, para enga?ar a la versi?n demo.
    '------------------------------------------------------------------
    'Par?metros.
    '   Entrada:
    '       [tabla] tabla LICENCIA, donde se encuntra el ?ltimo d?a de ejecuci?n
    '       [idAplicacion]  clave para acceder al registro de la tabla LICENCIA
    '   Salida:
    '       True:   se pudo actualizar el valor
    '       False:  no se pudo actualizar el valor
    '------------------------------------------------------------------
    On Error GoTo error
    'por defecto asumo que no se puede actualizar al fecha
    funActualizarUltimoDiaEjecuci?nVD = False
    
    'posiciono en registro correspondiente
    tabla.Index = "pkLicencia"
    tabla.Seek "=", idAplicacion
    If Not tabla.NoMatch Then
        'cambio valor del campo
        tabla.Edit
            tabla("aplicacionVDUltimoDiaEjecutada") = fechaActual
        tabla.Update
        'se pudo actualizar la fecha correctamente
        funActualizarUltimoDiaEjecuci?nVD = True
    End If
Exit Function
error:
    'se produjo un error
    funActualizarUltimoDiaEjecuci?nVD = False
End Function

Private Function funAplicacionLegal(tabla As Recordset, fechaActual As Date, serieDiscoDuroActual As String)
    'Determina si la aplicaci?n se puede ejecutar, y si es as?
    'que tipo de licencia tiene.
    Dim fechaFin As Date
    'determino si es una versi?n demo
    'Para que sea una versi?n demo: se tiene que dar estas dos condiciones
    If tabla("aplicacionVD") = True And _
        tabla("aplicacionNroLicencia") = "0" And _
        tabla("aplicacionSerieDisco") = "0" Then
        'valido que la fecha de la m?quina se la correcta
        If tabla("aplicacionVDUltimoDiaEjecutada") > fechaActual Then
            'la fecha del sistema fue retrocedida
            funAplicacionLegal = 515
        Else
            'valido el per?odo de ejecuci?n actual
            fechaFin = tabla("aplicacionDiaInstalacion") + _
                    tabla("aplicacionVDCantDias")
            If fechaFin < fechaActual Then
                'el per?odo de evaluaci?n lleg? a su fin
                funAplicacionLegal = 514
            Else
                'es una versi?n demo correcta
                funAplicacionLegal = 621
            End If
        End If
    Else
        'verifico si es una versi?n con licencia
        If tabla("aplicacionVD") = False And _
        tabla("aplicacionNroLicencia") <> "0" And _
        tabla("aplicacionSerieDisco") <> "0" Then
            'verifico si no es una copia, comparando los n?mero de serie de los discos duros
            If serieDiscoDuroActual <> tabla("aplicacionSerieDisco") Then
                'la aplicaci?n se esta ejecutando en un disco duro distinto al que fue instalada.
                funAplicacionLegal = 517
            Else
                'es una versi?n con licencia correcta
                funAplicacionLegal = 622
            End If
        Else
            'no se puede identificar el tipo de licencia de la versi?n
            funAplicacionLegal = 516
        End If
    End If
End Function
