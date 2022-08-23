VERSION 5.00
Object = "{EB8C3860-8603-11D0-A35D-7062E4000000}#1.1#0"; "VCBND.OCX"
Begin VB.Form frmPoblaciónFlotante 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Listado de población flotante"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   7905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Emisión de listado de población flotante "
      Height          =   3375
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   7695
      Begin VB.CommandButton botImprimir 
         Height          =   375
         Left            =   5040
         Picture         =   "frmPoblaciónFlotante.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         Tag             =   "Imprimir"
         Top             =   2880
         Width           =   1215
      End
      Begin VB.CommandButton botSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   375
         Left            =   6360
         TabIndex        =   3
         Top             =   2880
         Width           =   1215
      End
      Begin VcBndCtl.VcCalCombo fechaListadoPoblacionFlotante 
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   840
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _0              =   $"frmPoblaciónFlotante.frx":0942
         _1              =   $"frmPoblaciónFlotante.frx":0D4B
         _2              =   $"frmPoblaciónFlotante.frx":1154
         _3              =   "-@@@)f@@@E@@@A@@@@@@@@@'@@@E@@@A@@@@@@@@@)h@@@E@@@A@@@@@@@@@,467D"
         _count          =   4
         _ver            =   2
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "&Ingrese fecha a imprimir"
         Height          =   240
         Left            =   240
         TabIndex        =   0
         Top             =   480
         Width           =   2130
      End
   End
   Begin Hotel_Nana.gaHOTELbarra gaHOTELbarra1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   4
      Top             =   3420
      Width           =   7905
      _ExtentX        =   13944
      _ExtentY        =   582
   End
   Begin VB.Menu mnuFormulario 
      Caption         =   "&Formulario"
      Begin VB.Menu mnuFormularioImprimir 
         Caption         =   "Imprimir"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuFormularioSalir 
         Caption         =   "Salir"
         Shortcut        =   {F12}
      End
   End
End
Attribute VB_Name = "frmPoblaciónFlotante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim fechaListadoDef As Date
    
    'por defecto trabajo con la fecha del día anterior al
    'último cierre diario
    fechaListadoDef = m_FechaSis - 1
    fechaListadoPoblacionFlotante = fechaListadoDef
End Sub

Private Sub botImprimir_Click()
    Dim fechaImp As Date
    'valido que la fecha de impresión sea menor a la de la fecha del sistema
    If funValidoFecha Then
        fechaImp = Me.fechaListadoPoblacionFlotante.Value
        'verifico si existe listado(datos) de población flotante
        If funExisteListado(fechaImp) Then
            'verifico configuración listado
            If mfunAplicoConfImp(2, 17) = 1 Then
                subArmoReportePoblacionFlotante fechaImp
            End If
        Else
            mSubMensaje 4, 131
        End If
    End If
End Sub

Private Sub subArmoReportePoblacionFlotante(fechaImp As Date)
    '----------------------------------------------------------------------------------
    'Configuro el control Crystal genérico,obtiene datos y emite el listado
    'de población flotante.
    '-------------------------------------------------------------------------------
    'Parámetros.
    '   Entrada:    [fechaImp] fecha de la cual se quiere emitir el listado de
    '               población flotante.
    '-------------------------------------------------------------------------------
        
    'inicializo control data
     subInicializoControlData frmMAIN.Data1CrystalReport
    
    'NOTA: esta consulta debe de ser la misma que se utilizo para crear el reporte,
    'por ese motivo no debe de ser modificada en aspectos tales como campos a mostrar y
    'tablas utilizadas.
    
    'seleccino los registros del archivo de población flotante pertenecientes a la
    'fecha a imprimir
    frmMAIN.Data1CrystalReport.RecordSource = _
    "select * from poblacion_flotante where fechaListado = " & fechaSQL(fechaImp)
    
    'ejecuto consulta control data
    frmMAIN.Data1CrystalReport.Refresh
    
    'como de aquí en más voy a trabajar directamente con el control data
    'me aseguro de que se allan encontrado reservas
    If frmMAIN.Data1CrystalReport.Recordset.RecordCount > 0 Then
        'determino el nombre del reporte a utilizar
        frmMAIN.CrystalReport1.ReportFileName = vardir2 + "rptPflot.rpt"
        
        'inicializo fórmulas
        With frmMAIN.CrystalReport1
            .Formulas(0) = "cabVersion = '" & mFunImpVersionAplicacion & "'"        'nombre y versión aplicación
            .Formulas(1) = "cabRegistro = '" & mFunImpRegistro & "'"                'hotel propietario de la aplicación
            .Formulas(2) = "cabHoraImp = ' " & CStr(Time()) & "'"                   'hora actual
            .Formulas(4) = "parte1Fecha = '" & Format(fechaImp, "dddd dd mmmm yyyy") & "'"
        End With
        'genero listado
        frmMAIN.CrystalReport1.Action = 1
        'aviso de confirmación de impresión de reporte
        mSubMensaje 4, 140  'se imprimió listado de población flotante
        'inicializo fórmulas
        mSubInicializoFormulas 4
    Else
        'aviso de que no hay datos para imprimir
        mSubMensaje 3, 9
    End If
End Sub

Private Function funExisteListado(fechaListado As Date) As Boolean
    'Determina si existe el listado de población flotante.
    '--------------------------------------------------------------------------------------
    'Parámetros.
    '   Entrada: [fechaListado] fecha de la cual se desea imprimir la población flotante
    '                           (movimiento de personas)
    '
    '   Salida: [True]  existe listado generado
    '           [False] no existe listado, esto solo debe de ocurrir para las fechas
    '                   anteriores a la instalación del sistema.
    '--------------------------------------------------------------------------------------
    Dim tablaListadoPoblacion As Recordset
    Set tablaListadoPoblacion = tbPOBLACION_FLOTANTE
    
    'por defecto asumo que no existe listado generado
    funExisteListado = False
    tablaListadoPoblacion.Index = "pk_listado"
    tablaListadoPoblacion.Seek ">=", fechaListado, 0
    If Not tablaListadoPoblacion.NoMatch Then
        If tablaListadoPoblacion("fechaListado") = fechaListado Then
            'existe listado
            funExisteListado = True
        End If
    End If
    
    Set tablaListadoPoblacion = Nothing
End Function

Private Function funValidoFecha() As Boolean
    'Determina si la fecha de impresión del listado es correcta.
    'La misma debe de ser menor a la fecha del sistema.
    
    If Me.fechaListadoPoblacionFlotante.Value < m_FechaSis Then
        funValidoFecha = True
    Else
        'aviso de fecha erronea
        mSubMensaje 3, 6
        funValidoFecha = False
    End If
End Function

Private Sub botSalir_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Descargo formulario de memoria
    Set frmPoblaciónFlotante = Nothing
End Sub

Private Sub mnuFormularioImprimir_Click()
    'Se ejecuta también cuando presiono Ctrl+I
    botImprimir_Click
End Sub

Private Sub mnuFormularioSalir_Click()
    'Se ejecuta también cuando presiono F12
    botSalir_Click
End Sub

'*****************************************************
'*
'* Asistencia a usuarios
'*
'****************************************************

Private Sub fechaListadoPoblacionFlotante_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 214
End Sub

Private Sub botSalir_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 1, 1
End Sub

Private Sub botImprimir_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 213
End Sub

Private Sub fechaListadoPoblacionFlotante_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub botImprimir_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub botSalir_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

