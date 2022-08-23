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

Public Function mFunExistenUsuariosDef(Optional mostrarMsg As Byte) As Boolean
    'Determino si hay usuarios definidos
    '-----------------------------------------------------------------------------------------
    'Parámetros.
    '       Entrada:
    '               [mostrarMsg] Si el valor de este parámetro es 1, no se muestra
    '                            el mensaje al usuario.
    '               Esta función es llamada para dos objetivos:
    '                   a) determinar si se puede ejecutar una opción determinada, la cual
    '                      necesita que existan usuarios definidos.
    '                   b) o para verificar si luego de eliminar un usuario existen
    '                      más usuarios o era el último. Para este punto no es necesario
    '                      mostrar un aviso al usuario, por lo que se utiliza el parámetro
    '                      para que el mismo no aparezca, en caso de que el resultado de la
    '                      función sea negativo.
    '
    '       Salida: True, existen 1 o más registros en la tabla SISTEMA_USUARIOS
    '               False, no existen registros en dicha tabla.
    '------------------------------------------------------------------------------------------
    Dim rstUsr As Recordset
    Dim qdfUsr As QueryDef
    Dim consulta As String
    Dim mensaje As String
    
    'defino consulta
    consulta = "select * from SISTEMA_USUARIOS"
    'ejecuto consulta
    Set qdfUsr = bdAplicacion.CreateQueryDef("")
    qdfUsr.SQL = consulta
    Set rstUsr = qdfUsr.OpenRecordset(dbOpenSnapshot)
       
    'cuento cantidad de registros
    If rstUsr.RecordCount > 0 Then
        'existen usuarios
        mFunExistenUsuariosDef = True
    Else
        'no existen usuarios
        mFunExistenUsuariosDef = False
        'verifico si tengo que mostrar mensaje
        If mostrarMsg <> 1 Then
            mensaje = "No existen usuarios definidos. " & Chr(10) & _
                    "Para ejecutar esta opción primero debe de definir al menos 1 usuario."
            MsgBox mensaje, vbExclamation, "No se puede ejecutar esta opción."
        End If
    End If
    Set qdfUsr = Nothing
    Set rstUsr = Nothing
End Function

