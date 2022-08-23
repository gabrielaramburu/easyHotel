Attribute VB_Name = "BaseDeDatos"
Option Explicit

Public Sub mSubAbroBaseDeDatos()
On Error GoTo errores
    'asigna espacio trabajo
    Set bdWK = DBEngine.Workspaces(0)
    
    'obtengo el directorio de ejecución del exe.
    m_vardirBD = CaminoBaseDeDatos                'directorio para BD
    m_vardirRpt = App.Path & "\"                  'directorio para reportes
        
    'abre base de datos
    Set bdAplicacion = bdWK.OpenDatabase(m_vardirBD, False, False, ";PWD=manyacapo;")
    
    'abre tablas
    Set tbSISTEMA_BITACORA = bdAplicacion.OpenRecordset("SISTEMA_BITACORA", dbOpenTable)
    Set tbSISTEMA_PARAMETROS = bdAplicacion.OpenRecordset("SISTEMA_PARAMETROS", dbOpenTable)
    Set tbSISTEMA_USUARIOS = bdAplicacion.OpenRecordset("SISTEMA_USUARIOS", dbOpenTable)
    Set tbSISTEMA_OPERACIONES = bdAplicacion.OpenRecordset("SISTEMA_OPERACIONES", dbOpenTable)
    Set tbSISTEMA_BITACORAlistados = bdAplicacion.OpenRecordset("SISTEMA_BITACORA_listados", dbOpenTable)
    Set tbSISTEMA_BITACORAparametros = bdAplicacion.OpenRecordset("SISTEMA_BITACORA_parametros", dbOpenTable)
    Set tbSISTEMA_PERFILES = bdAplicacion.OpenRecordset("SISTEMA_PERFILES", dbOpenTable)
    Set tbSISTEMA_LICENCIA = bdAplicacion.OpenRecordset("SISTEMA_LICENCIA", dbOpenTable)
errores:
    If Err.Number <> 0 Then
        MsgBox "Aviso del sistema número: " & Err.Number & Chr(10) & _
        Err.Description & Chr(10) & _
        "Imposible continuar con la ejecución" & Chr(10) & _
        "La base de datos seleccionada no es la correcta.", vbExclamation
        End
    End If
End Sub

Public Function mfunBuscoListado(lst As String)
    'busco un listado específico
    mfunBuscoListado = False
    tbSISTEMA_BITACORAlistados.Index = "pk_listado"
    tbSISTEMA_BITACORAlistados.Seek "=", lst
    If Not tbSISTEMA_BITACORAlistados.NoMatch Then 'existe
        mfunBuscoListado = True
    End If
End Function

Public Function mFunBuscoDescOpr(Opr As Integer)
    'busco la descripción de la operación
    mFunBuscoDescOpr = ""
    tbSISTEMA_OPERACIONES.Index = "pk_Opr"
    tbSISTEMA_OPERACIONES.Seek "=", Opr
    If Not tbSISTEMA_OPERACIONES.NoMatch Then   'existe
        mFunBuscoDescOpr = tbSISTEMA_OPERACIONES("DescOpr")
    End If
End Function

