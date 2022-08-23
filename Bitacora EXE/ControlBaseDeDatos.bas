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

Public Function mFunExistenListados() As Boolean
    'Determina si existen listados definidos.
    'Esta función es llamada desde el menú principal,
    'antes de ejecutar las operaciones que trabajan con listados.
    '--------------------------------------------------------------------------
    'Parámetros.
    '   Salida: True, existen registros en la tabla SISTEMA_BITACORAlistados
    '           False, no existen registros en dicha tabla
    '---------------------------------------------------------------------------
    Dim mensaje As String
    
    
    If tbSISTEMA_BITACORAlistados.RecordCount > 0 Then
        'existe 1 o más listados
        mFunExistenListados = True
    Else
        'no existen listados
        mensaje = "No existen listados definidos." & Chr(10) & _
                "Para ejecutar esta opción primero debe de definir al menos 1 listado."
                
        MsgBox mensaje, vbExclamation, "No se puede ejecutar esta opción."
        mFunExistenListados = False
    End If
End Function

