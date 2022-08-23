Attribute VB_Name = "ModBaseDeDatos"
'Declaraci�n de variable de base de datos
Public bdLISTA As Database, bdWK As Workspace

'Declaraci�n de variable de tipo tabla
Public tbARCHIVOS As Recordset
Public tbSISTEMA_MENSAJES As Recordset
Public tbINFO As Recordset

Public Sub mSubAbroBaseDeDatos()
    'asigna espacio trabajo
    Set bdWK = DBEngine.Workspaces(0)
    
    'obtengo el directorio de ejecuci�n del exe.
    varDirBaseDeDatos = App.Path & "\listaMP3"
        
    'abre base de datos
    Set bdLISTA = bdWK.OpenDatabase(varDirBaseDeDatos)
    
    'abre tablas
    Set tbARCHIVOS = bdLISTA.OpenRecordset("Archivos", dbOpenTable)
    Set tbSISTEMA_MENSAJES = bdLISTA.OpenRecordset("SISTEMA_MENSAJES", dbOpenTable)
    Set tbINFO = bdLISTA.OpenRecordset("Informacion", dbOpenTable)
End Sub
