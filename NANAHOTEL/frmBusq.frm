VERSION 5.00
Object = "{08825A62-8182-11D6-AE38-FDECBDCC172B}#14.0#0"; "SeleccionRegistrosBD.ocx"
Begin VB.Form frmBusq 
   Caption         =   "Form1"
   ClientHeight    =   3930
   ClientLeft      =   315
   ClientTop       =   1470
   ClientWidth     =   7935
   ClipControls    =   0   'False
   Icon            =   "frmBusq.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   7935
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   7680
      Top             =   480
   End
   Begin SeleccionRegistrosBD.SeleccionBD SeleccionBD1 
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   6800
      GrillaForeColor =   -2147483640
      GrillaBackColor =   -2147483643
      BeginProperty GrillaFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuFormulario 
      Caption         =   "&Formulario"
      Begin VB.Menu mnuFormularioSeleccionar 
         Caption         =   "Seleccionar"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuFormularioCancelar 
         Caption         =   "Cancelar"
      End
      Begin VB.Menu mnuDiv1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFormularioCambiarCriterio 
         Caption         =   "Cambiar criterio"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuFormularioIngresoCriterio 
         Caption         =   "Ir a ingreso de criterio"
         Shortcut        =   ^I
      End
   End
End
Attribute VB_Name = "frmBusq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Declaración de propiedades

Public propAncho As Integer                     'Determina el ancho del formulario
Public propLargo As Integer                     'Determina el largo del formulario
Public propTeclaSeleccion As Integer            'Determina la tecla con la cual se selecciona
                                                'una fila de la grilla
Public propNroCampoInicial As Integer           'Determina la columna por la cual se ordena la grilla
                                                'por defecto
Public propIndiceCampoRetorno As Integer        'Determina la columna de donde se obtiene el valor
                                                'que se devuelve después de seleccionar una fila de la grilla.
Public propTablasRelacionadas As String         'Determina las tablas auxiliares a las cuales se acceden
Public propCampos As String                     'Determina los campos que se muestran en la consulta
Public propTabla As String                      'Determina la tabla principal a la cual se accede
Public propSeleccionComplementaria              'Determina el criterio de selección complementaria (ver documentación control)
Public propTituloFormulario As String           'Determina el título del formulario
Public propRetorno As Variant                   'En esta propiedad se carga el valor devuelto por el control para
                                                'que la misma sea consultada desde el formulario que utiliza la ayuda

Private Sub Form_KeyPress(KeyAscii As Integer)
    'Cierro el formulario al digitar Esc
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    'Cambio el tamaño del formulario
    Me.Height = propAncho
    Me.Width = propLargo
    'Inicializo las propiedades del control de seleccion
    Me.SeleccionBD1.BaseDeDatos = vardir
    Me.SeleccionBD1.TeclaSeleccion = propTeclaSeleccion
    Me.SeleccionBD1.NroCampoInicial = propNroCampoInicial
    Me.SeleccionBD1.IndiceCampoRetorno = propIndiceCampoRetorno
    Me.SeleccionBD1.TablasRelacionadas = propTablasRelacionadas
    Me.SeleccionBD1.campos = propCampos
    Me.SeleccionBD1.tabla = propTabla
    Me.SeleccionBD1.SeleccionComplementaria = propSeleccionComplementaria
    
    'cambio el título del formulario
    Me.Caption = propTituloFormulario
    'desencadeno el evento timer
    Timer1.Enabled = True
End Sub

Private Sub Form_Resize()
    'Al modificar el tamaño del formulario también modifico el tamaño del
    'control de selección.
    Me.SeleccionBD1.Width = Me.Width - 400
    Me.SeleccionBD1.Height = Me.Height - 800
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Descargo formulario de memoria
    Set frmBusq = Nothing
End Sub

Private Sub mnuFormularioCambiarCriterio_Click()
    'Equivale a presionar el boton de cambiar criterios
    Me.SeleccionBD1.CambiarCriterios
End Sub

Private Sub mnuFormularioCancelar_Click()
    Unload Me
End Sub

Private Sub mnuFormularioIngresoCriterio_Click()
    'Le doy el focus al control que ingresa el criterio
    Me.SeleccionBD1.CambiarValorCriterio
End Sub

Private Sub mnuFormularioSeleccionar_Click()
    'Simulo que aprieto la tecla de selección de filas
    SendKeys (Chr(propTeclaSeleccion))
End Sub

Private Sub SeleccionBD1_Seleccionar(ValorClaveTablaPrincipal As Variant)
    'Se selecciono una fila
    frmBusq.propRetorno = ValorClaveTablaPrincipal
    'oculto el formulario
    Me.Visible = False
End Sub

Private Sub Timer1_Timer()
    'Primer cargo el formulario y después desencadeno este evento para
    'que se pueda ver la barra de progreso en funcionamiento
    'Si no incluyo este evento no se ve dicha barra.
    Timer1.Enabled = False
    'muestro datos en control
    Me.SeleccionBD1.Mostrar
End Sub
