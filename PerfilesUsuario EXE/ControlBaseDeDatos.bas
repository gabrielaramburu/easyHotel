Attribute VB_Name = "ControlBaseDeDatos"
Option Explicit

'Al ejecutar un formulario, se asume que existen datos en la base de datos choerentes,
'que permitir�n la ejecuci�n correcta del c�digo de dichos formularios o m�dulos.

'Sin embargo, cuando se instala la aplicaci�n, las tablas de la base de datos estan vac�as,
'o sin inicializar. Esto puede originar que ciertos procesos no se puedan ejecutar, ya que
'es impresindible que los mismos cuenten con informaci�n b�sica.

'En este m�dulo se realiza el control de existencia de datos m�nimos, es decir, antes de
'abrir un determinado formulario, se valida que existan los datos m�nimos necesarios
'en la base de datos para que el mismo se pueda ejecutar.

Public Function mFunExistenUsuariosDef(Optional mostrarMsg As Byte) As Boolean
    'Determino si hay usuarios definidos
    '-----------------------------------------------------------------------------------------
    'Par�metros.
    '       Entrada:
    '               [mostrarMsg] Si el valor de este par�metro es 1, no se muestra
    '                            el mensaje al usuario.
    '               Esta funci�n es llamada para dos objetivos:
    '                   a) determinar si se puede ejecutar una opci�n determinada, la cual
    '                      necesita que existan usuarios definidos.
    '                   b) o para verificar si luego de eliminar un usuario existen
    '                      m�s usuarios o era el �ltimo. Para este punto no es necesario
    '                      mostrar un aviso al usuario, por lo que se utiliza el par�metro
    '                      para que el mismo no aparezca, en caso de que el resultado de la
    '                      funci�n sea negativo.
    '
    '       Salida: True, existen 1 o m�s registros en la tabla SISTEMA_USUARIOS
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
                    "Para ejecutar esta opci�n primero debe de definir al menos 1 usuario."
            MsgBox mensaje, vbExclamation, "No se puede ejecutar esta opci�n."
        End If
    End If
    Set qdfUsr = Nothing
    Set rstUsr = Nothing
End Function

