VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InformacionApli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Este m?dulo se creo para implementar las funciones que brindan informaci?n del sistema,
'utilizada, por ejemplo en el formulario de ayuda y en el menu Main.

Public Function mFunObtenerLicenciaApli(idApli As Long, tabla As Recordset, tipoInf) As String
    'Devuelve informaci?n sobre la licencia de una aplicaci?n.
    '---------------------------------------------------------------------------------
    'Par?metros:
    '   Entrada   [idApli]  n?mero de id de la aplicaci?n de la cual deseo obtener datos.
    '                       Con este n?mero accedo directamente al registro de la aplicaci?n
    '                       en el archivo tbSISTEMA_LICENCIA
    '             [tabla]   tabla de control de licencia
    '
    '             [tipoInf] tipo de informaci?n que se desea obtener
    '                       1 = nro. de licencia
    '                       2 = nombre empresa
    '                       3 = due?o aplicaci?n
    '
    '   Salida:   si [tipoInf] = 1 : valor del campo tbSISTEMA_LICENCIA("AplicacionNroLicencia")
    '             si [tipoInf] = 2 : valor del campo tbSISTEMA_LICENCIA("AplicacionEmpresa")
    '             si [tipoInf] = 3 : valor del campo tbSISTEMA_LICENCIA("AplicacionDue?o")
    '----------------------------------------------------------------------------------
    On Error Resume Next
    'por defecto devuelvo empty
    mFunObtenerLicenciaApli = Empty
    'posiciono en registro correspondiente
    tabla.Index = "pkLicencia"
    tabla.Seek "=", idApli
    If Not tabla.NoMatch Then
        'existe registro
        Select Case tipoInf
            Case 1
                mFunObtenerLicenciaApli = tabla("AplicacionNroLicencia")
            Case 2
                mFunObtenerLicenciaApli = tabla("AplicacionEmpresa")
            Case 3
                mFunObtenerLicenciaApli = tabla("AplicacionDue?o")
        End Select
    End If
End Function

