Attribute VB_Name = "BaseDeDatos"
'Declaraci�n de constante de contrase�a
Public Const cContrase�aBD As String = ";PWD=manyacapo;"

Public Sub mSubAbroBaseDeDatos()
On Error GoTo errores
    'asigna espacio trabajo
    Set bdWK = DBEngine.Workspaces(0)
    
    'obtengo el directorio de ejecuci�n del exe.
    m_vardirBD = CaminoBaseDeDatos                'directorio para BD
    m_vardirRpt = App.Path & "\"                  'directorio para reportes
        
    'abre base de datos
    Set bdAplicacion = bdWK.OpenDatabase(m_vardirBD, False, False, cContrase�aBD)
    
    'abre tablas
    Set tbSISTEMA_USUARIOS = bdAplicacion.OpenRecordset("SISTEMA_USUARIOS", dbOpenTable)
    Set tbSISTEMA_PARAMETROS = bdAplicacion.OpenRecordset("SISTEMA_PARAMETROS", dbOpenTable)
    Set tbSISTEMA_PERFILES = bdAplicacion.OpenRecordset("SISTEMA_PERFILES", dbOpenTable)
    Set tbSISTEMA_OPERACIONES = bdAplicacion.OpenRecordset("SISTEMA_OPERACIONES", dbOpenTable)
    Set tbSISTEMA_LICENCIA = bdAplicacion.OpenRecordset("SISTEMA_LICENCIA", dbOpenTable)
    Set tbSISTEMA_LISTADOS = bdAplicacion.OpenRecordset("SISTEMA_LISTADOS", dbOpenTable)
    Set tbSISTEMA_MENSAJES = bdAplicacion.OpenRecordset("SISTEMA_MENSAJES", dbOpenTable)
errores:
    If Err.Number <> 0 Then
        MsgBox "Aviso del sistema n�mero: " & Err.Number & Chr(10) & _
        Err.Description & Chr(10) & _
        "Imposible continuar con la ejecuci�n" & Chr(10) & _
        "La base de datos seleccionada no es la correcta.", vbExclamation
        End
    End If
End Sub

Public Function funBuscoOperacionTF(CodOpr As Integer)
    'Busco una operaci�n
    funBuscoOperacionTF = True
    tbSISTEMA_OPERACIONES.Index = "pk_Opr"
    tbSISTEMA_OPERACIONES.Seek "=", CodOpr
    If tbSISTEMA_OPERACIONES.NoMatch Then
        funBuscoOperacionTF = False
    End If
End Function

Public Sub subInicializoControlData(controlData As Object)
    'Inicializa el control data pasado como par�metro
    'con la base de datos que utiliza la aplicaci�n
    
    controlData.Connect = cContrase�aBD 'establece la contrase�a de la base de datos
    controlData.DatabaseName = m_vardirBD
End Sub

