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

Public Function mFunExistenListados() As Boolean
    'Determina si existen listados definidos.
    'Esta funci�n es llamada desde el men� principal,
    'antes de ejecutar las operaciones que trabajan con listados.
    '--------------------------------------------------------------------------
    'Par�metros.
    '   Salida: True, existen registros en la tabla SISTEMA_BITACORAlistados
    '           False, no existen registros en dicha tabla
    '---------------------------------------------------------------------------
    Dim mensaje As String
    
    
    If tbSISTEMA_BITACORAlistados.RecordCount > 0 Then
        'existe 1 o m�s listados
        mFunExistenListados = True
    Else
        'no existen listados
        mensaje = "No existen listados definidos." & Chr(10) & _
                "Para ejecutar esta opci�n primero debe de definir al menos 1 listado."
                
        MsgBox mensaje, vbExclamation, "No se puede ejecutar esta opci�n."
        mFunExistenListados = False
    End If
End Function

