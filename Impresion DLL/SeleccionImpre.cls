VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SeleccionImpre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Declaro varible para utilización de formulario
Private WithEvents dlgSeleccionImpre As frmSeleccionImpre
Attribute dlgSeleccionImpre.VB_VarHelpID = -1

'Declaro variables para poder utilizar los objetos de ImpresionGeneral
Private biblioImpresionGeneral As ImpresionGeneral

'Variable utilizada para almacenar la impresora seleccionada en el formulario
Private mImpSel As String

Public Function mFunSeleccionoImpresora(impPorDefecto As String) As String
    'Permite seleccionar una impresora de la lista de impresoras del sistema.
    'Por defecto ubico el cursor del combo en la impresora por defecto.
    'Antes de ejecutar la función se verifica que existan impresoras instaladas.
    '-------------------------------------------------------------------------------------
    'Parámetros:
    '   Entrada: [impPorDefecto]    valor correspondiente a la impresora que se tiene que
    '                               mostrar por defecto en el cuendro de dialogo.
    '   Salida: Si se selecciona impresora
    '               Valor correspondienye a la impresora seleccionada.
    '               Esta impresora será la que se utilizará para emitir el listado.
    '           Si no se selecciona impresora (boton cancelar) se devuelve ""
    '           Si no hay impresoras instaladas se devuelve ""
    '--------------------------------------------------------------------------------------
    On Error Resume Next
    Dim X As Printer
    
    'verifico si hay impresoras instaladas en el sistema
    If biblioImpresionGeneral.mFunCantidadImpresorasInstaladas > 0 Then
        'recorro todas las impresoras del sistema
        For Each X In Printers
            dlgSeleccionImpre.cboImpreSistema.AddItem X.DeviceName
        Next
        'determino si la impresora que se quiere mostrar por defecto existe
        If Not biblioImpresionGeneral.mFunExisteImpresoraInstalada(impPorDefecto) Then
            'la impresora no existe
            'muestro la impresora por defecto del sistema
            dlgSeleccionImpre.cboImpreSistema.Text = Printer.DeviceName
            'muestro aviso en formulario
            dlgSeleccionImpre.lblAviso.Caption = _
            "La impresora por defecto del listado no esta instalada. " & _
            "Se muestra entonces la impresora predeterminada del sistema."
            dlgSeleccionImpre.lblAviso.Visible = True
        Else
            'la impresora si existe
            dlgSeleccionImpre.cboImpreSistema.Text = impPorDefecto
        End If
        'muestro formulario
        dlgSeleccionImpre.Show 1
        
        mFunSeleccionoImpresora = mImpSel
    Else
        'no se asigna impresora
        mFunSeleccionoImpresora = ""
    End If
End Function

Private Sub Class_Initialize()
    'Creo una instamcia privada del formulario
    Set dlgSeleccionImpre = New frmSeleccionImpre
    'Creo una instancia privada para acceder a ImpresionGeneral
    Set biblioImpresionGeneral = New ImpresionGeneral
End Sub

Private Sub Class_Terminate()
    'Destruyo la instancia del formulario al destruir la instancia de la clase SeleccionImpre
    Unload dlgSeleccionImpre
    Set dlgSeleccionImpre = Nothing
    'Destruyo instancia privadas creadas
    Set biblioImpresionGeneral = Nothing
End Sub

Private Sub dlgSeleccionImpre_notificoCliente(boton As Byte)
    'Determino operacion realizada en el formulario
    Select Case boton
        Case 1  'boton aceptar
            'devuelvo la impresora seleccionada.
            mImpSel = dlgSeleccionImpre.cboImpreSistema.Text
        Case 0  'boton cancelar
            mImpSel = ""
    End Select
End Sub
