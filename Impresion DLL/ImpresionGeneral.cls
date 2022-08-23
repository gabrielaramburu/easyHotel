VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ImpresionGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function mFunExisteImpresoraInstalada(impaVerificar As String) As Boolean
    'Determina si la impresora que se pasa como parámetro esta instalada en el sistema.
    '------------------------------------------------------------------------------------
    'Parámetros:
    '   Entrada [impAVerificar]  string que corresponde a la propiedad printers.DeviceName
    '
    '   Salida  true = si existe la impresora
    '           false= si no existe la impresora
    '-------------------------------------------------------------------------------------
    On Error Resume Next
    Dim impresoraInstalada As Printer
    
    'por defecto asumo que la impresora no existe
    mFunExisteImpresoraInstalada = False
    'recorro todas las impresoras del sistema
    For Each impresoraInstalada In Printers
        If impresoraInstalada.DeviceName = impaVerificar Then
            'la impresora esta instalada en el sistema
            mFunExisteImpresoraInstalada = True
            'salgo del bucle
            Exit For
        End If
    Next
End Function

Public Function mFunCantidadImpresorasInstaladas() As Integer
    'Determina cuantas impresoras hay instaladas en el sistema.
    '-----------------------------------------------------------------------------
    'Parámetros:
    '   Salida =    total de impresoras instaladas
    '               0 si no existen impresoras instaladas
    '------------------------------------------------------------------------------
    On Error Resume Next
    mFunCantidadImpresorasInstaladas = Printers.Count
End Function

