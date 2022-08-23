VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.UserControl MantBD 
   ClientHeight    =   3840
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5310
   PropertyPages   =   "MantBD.ctx":0000
   ScaleHeight     =   3840
   ScaleWidth      =   5310
   ToolboxBitmap   =   "MantBD.ctx":003D
   Begin ComctlLib.Toolbar toolMenu 
      Height          =   630
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5250
      _ExtentX        =   9260
      _ExtentY        =   1111
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   7
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   ""
            Object.ToolTipText     =   "Guardar"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   ""
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   ""
            Object.ToolTipText     =   "Borrar"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Key             =   ""
            Object.ToolTipText     =   "Próximo libre"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Limpiar datos para nuvo ingreso"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Error en ocx."
      Height          =   1695
      Left            =   0
      TabIndex        =   8
      Top             =   720
      Visible         =   0   'False
      Width           =   4695
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "MantBD.ctx":034F
         Top             =   240
         Width           =   480
      End
      Begin VB.Label lblError 
         Caption         =   "Etiqueta de errores"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   720
         TabIndex        =   9
         Top             =   240
         Width           =   3855
      End
   End
   Begin VB.ListBox cboIngOpcF 
      Height          =   315
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   7
      Top             =   3120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ListBox cboIngOpcA 
      Height          =   315
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   6
      Top             =   3480
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3480
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Data Data2 
      Caption         =   "Data2 Este es el principal"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3480
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txtIngstrMemo 
      Height          =   645
      Left            =   2760
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   5
      Top             =   2760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtIngFecha 
      Height          =   285
      Left            =   2400
      TabIndex        =   4
      Top             =   2760
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtIngNum 
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      Top             =   2760
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtIngStr 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   2760
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid gTabla 
      Height          =   2055
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   3625
      _Version        =   393216
      Cols            =   4
      Enabled         =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   0
      GridLinesFixed  =   0
      ScrollBars      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   4680
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   8
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MantBD.ctx":0791
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MantBD.ctx":13E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MantBD.ctx":2035
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MantBD.ctx":2C87
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MantBD.ctx":38D9
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MantBD.ctx":452B
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MantBD.ctx":4845
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MantBD.ctx":4B5F
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "MantBD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Declaración de constantes
Private Const contsAnchoMemo As Byte = 3            'valor por el cual se multiplica el
                                                    'ancho del controle de tipo memo
Private Const cLargoMinimoCeldaDatos As Long = 1000
Private Const cAnchoMinimoCelda As Long = 240       'ancho mínimo permitido para las filas
                                                    'de la grilla (240 valor por defecto del control MSFGrid)
Private Const cAnchoLista As Long = 1               'valor por el cual se multiplica el
                                                    'ancho del controle de tipo lista
Private Const cMsgErr As String = "Error grabe en Mantenimiento.ocx"

'Utilizada para trabajar con la grilla de ingreso, en esta varible estará almacenado
'el objeto con el cual se está trabajando actualmente
Private ControlSeleccionadoEnGrilla As Object

'Utilizadas para determinar si estoy modificando un campo clave para ocx
'con tipo de clave = 1 (determinada por usuario)
Private campoClaveNum As Boolean
Private campoClaveStr As Boolean
Private campoClaveCboA As Boolean
Private campoClaveCboF As Boolean

'Como el evento de la grilla SelChanged se ejecuta al realizar siertas operaciones en la grilla
'como por ejemplo seleccionar rango de celdas, utilizo esta variable para controlar su ejecución
'durante la ejecución del método IniciarMantenimiento.
Private habilitoSelChange As Boolean

'Utilizada para determinar que fila de la grilla contiene el campo clave
'se inicializa en el evento resize
Private filaCampoClave As Integer

'Determino si al ingresar o modificra un dato se produjo un error de ingreso provocado
'por el usuario al no respetar las propiedades del campo
'Si es así no inicializo la grilla como ocurriría si la modificación o alta estuvieran bien
Private errorEnIngresoDeDatos As Boolean

'Declaración de variables de propiedades
Private propiedadCampo As String
Private propiedadCaminoBaseDeDatos As String
Private propiedadContraseñaBaseDeDatos As String
Private propiedadTabla As String
Private propiedadIndiceCampoClave  As Integer
Private propiedadTipoClave As Byte                  'determina el tipo de clave principal de la tabla
                                                    '0 la clave la determina el usuario
                                                    '1 la clave es de tipo correlativo
Private propiedadSugerirProxLibre As Byte
Private propiedadTablaContador As String
Private propiedadIndiceCampoCont As Integer
Private propiedadIndiceCampoClaveContador As Integer
Private propiedadIntegridad As String

'propiedades de apariencia
Private propAnchoCeldas As Long
Private propLargoCeldas As Long
Private propFuenteNombreCampo As Byte
Private propFuenteDatosIngresados As Byte
Private propFuenteDatosAIngresar As Byte
Private propColorCaracteresIngreso As OLE_COLOR
Private propColorFondoDatos As OLE_COLOR
Private propColorFondoGrilla As OLE_COLOR
Private propColorFondoCampoIngreso As OLE_COLOR
Private propMostrarLineas As Byte

'Declaración de constantes de propiedades
Private Const constCaminoBaseDeDatos As String = ""
Private Const constContraseñaBaseDeDatos As String = ""
Private Const constTabla As String = ""
Private Const constIndiceCampoClave As Integer = 0
Private Const constTipoClave As Byte = 0
Private Const constSugerirProxLibre As Byte = 0 'por defecto no sugiero
Private Const constTablaContador As String = ""
Private Const constIndiceCampoCont As Integer = 0
Private Const constIndiceCampoClaveContador As Integer = 0
Private Const constIntegridad As String = ""

'constantes de apariencia
Private Const cAnchoCeldas As Long = cAnchoMinimoCelda
Private Const cLargoCeldas As Long = cLargoMinimoCeldaDatos
Private Const cFuenteNombreCampo As Byte = 8
Private Const cFuenteDatosAIngresar As Byte = 8
Private Const cFuenteDatosIngresados As Byte = 8
Private Const cMostrarLineas As Byte = 1 'por defecto las muestro
Private Const cColorCaracteresIngreso As Long = &H80000012  'negro
Private Const cColorFondoDatos  As Long = &H80000005    'blanco
Private Const cColorFondoGrilla As Long = &H80000005    'blanco
Private Const cColorFondoCampoIngreso As Long = &H80000005   'blanco

'Declaración de eventos
Public Event ErrorEnIngreso(tipo As Byte, desc As String)
Public Event CambioOperacion(operacionActual As Byte)
Public Event SeGraboTabla(claveGrabada As Variant)
Public Event SeModificoTabla(claveModifica As Variant)
Public Event SeEliminoTabla(claveEliminada As Variant)
Public Event NoHayDatosSuficientes(archivo As String)   'se produce cuando se quiere trabajar
                                                        'con un campo de tipo comboArchivo, y el archivo
                                                        'no tiene datos.

'Estas variales son generales ya que la implemnetación del control así lo requieren
'En ellas se almacenan los valores de alguna de las pripedades que deben de tener los
'controles que se utilizan para el ingreso de datos en la grilla, como así también
'información que permita validadr elingreso de datos a los mismos.

'propiedades para ingresar números
Private ValorDecimal As Boolean
Private ValorNegativo As Boolean

Private Sub UserControl_InitProperties()
    On Error Resume Next
    'se ejecuta cuando se coloca un nuevo control en el contenedor
        
    'Cargo propiedades del control con los valores predefinidos
    propiedadCaminoBaseDeDatos = constCaminoBaseDeDatos
    propiedadContraseñaBaseDeDatos = constContraseñaBaseDeDatos
    propiedadTabla = constTabla
    propiedadIndiceCampoClave = constIndiceCampoClave
    propiedadTipoClave = constTipoClave
    propiedadSugerirProxLibre = constSugerirProxLibre
    propiedadTablaContador = constTablaContador
    propiedadIndiceCampoCont = constIndiceCampoCont     'determina el índice de la tabla donde
                                                        'se almacena el campo contador
    
    propiedadIndiceCampoClaveContador = constIndiceCampoClaveContador   'determina el ínice del campo clave de la tabla
                                                                        'con la cual trabaja el ocx.
    propAnchoCeldas = cAnchoCeldas
    propLargoCeldas = cLargoCeldas
    propFuenteNombreCampo = cFuenteNombreCampo
    propFuenteDatosAIngresar = cFuenteDatosAIngresar
    propFuenteDatosIngresados = cFuenteDatosIngresados
    propMostrarLineas = cMostrarLineas
    propColorFondoGrilla = cColorFondoGrilla
    propColorFondoDatos = cColorFondoDatos
    propColorCaracteresIngreso = cColorCaracteresIngreso
    propColorFondoCampoIngreso = cColorFondoCampoIngreso
    propiedadIntegridad = constIntegridad
End Sub

Private Sub UserControl_Initialize()
    On Error Resume Next
    'Se ejecuta cada vez que se crea una nueva instancia del control
    
End Sub

Private Sub UserControl_Terminate()
    On Error Resume Next
    'Al finalizar el control eleimino la asignación
    Set ControlSeleccionadoEnGrilla = Nothing
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    'se produce al crearse una nueva instancia en diseño o ejecución
    'del control
    'propiedad campos
    campo = PropBag.ReadProperty("Campo", "")
    'propiedad camino de base de datos
    CaminoBaseDeDatos = PropBag.ReadProperty("CaminoBaseDeDatos", constCaminoBaseDeDatos)
    ContraseñaBaseDeDatos = PropBag.ReadProperty("ContraseñaBaseDeDatos", constContraseñaBaseDeDatos)
    'propiedad tabla principal
    tabla = PropBag.ReadProperty("Tabla", constTabla)
    'propiedad indice campo clave tabla
    IndiceCampoClave = PropBag.ReadProperty("IndiceCampoClave", constIndiceCampoClave)
    'propiedad tipo de clave 0=determinada por usuario,1 = correlativo
    TipoClave = PropBag.ReadProperty("TipoClave", constTipoClave)
    'propiedad sugerir proximo
    SugerirProxLibre = PropBag.ReadProperty("SugerirProxLibre", constSugerirProxLibre)
    'propiedad tabla origen de del contador
    TablaContador = PropBag.ReadProperty("TablaContador", constTablaContador)
    'propiedad ínide del campo contador en la tabla contador
    CampoCont = PropBag.ReadProperty("CampoCont", constIndiceCampoCont)
    'propiedad del indice del campo clave de tipo contador en la tabala principal
    IndiceCampoClaveContador = PropBag.ReadProperty("IndiceCampoClaveContador", constIndiceCampoClaveContador)
    'propiedad control de integridad
    integridad = PropBag.ReadProperty("Integridad", constIntegridad)
    
    'propiedades de apariencia
    ColorFondoDatos = PropBag.ReadProperty("ColorFondoDatos", cColorFondoDatos)
    ColorFondoGrilla = PropBag.ReadProperty("ColorFondoGrilla", cColorFondoGrilla)
    ColorCaracteresIngreso = PropBag.ReadProperty("ColorCaracteresIngreso", cColorCaracteresIngreso)
    ColorFondoCampoIngreso = PropBag.ReadProperty("ColorFondoCampoIngreso", cColorFondoCampoIngreso)
    AnchoCeldas = PropBag.ReadProperty("AnchoCeldas", cAnchoCeldas)
    LargoCeldas = PropBag.ReadProperty("LargoCeldas", cLargoCeldas)
    FuenteNombreCampo = PropBag.ReadProperty("FuenteNombreCampo", cFuenteNombreCampo)
    FuenteDatosIngresados = PropBag.ReadProperty("FuenteDatosIngresados", cFuenteDatosIngresados)
    FuenteDatosAIngresar = PropBag.ReadProperty("FuenteDatosAIngresar", cFuenteDatosAIngresar)
    MostrarLineas = PropBag.ReadProperty("MostrarLineas", cMostrarLineas)
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    On Error Resume Next
    'se graba al destruirse una instancia en tiemo de diseño
    
    'propiedad campos (solo en perzonalizado)
    PropBag.WriteProperty "Campo", propiedadCampo       'Creo que no es posible tener
                                                        'valores predeterminados de esta propiedad
    
    'propiedad camino de base de datos
    PropBag.WriteProperty "CaminoBaseDeDatos", propiedadCaminoBaseDeDatos, constCaminoBaseDeDatos
    PropBag.WriteProperty "ContraseñaBaseDeDatos", propiedadContraseñaBaseDeDatos, constContraseñaBaseDeDatos
    'propiedad tabla principal
    PropBag.WriteProperty "Tabla", propiedadTabla, constTabla
    'propiedad indice campo clave tabla
    PropBag.WriteProperty "IndiceCampoClave", propiedadIndiceCampoClave, constIndiceCampoClave
    'propiedad tipo de clave 0=determinada por usuario,1 = correlativo
    PropBag.WriteProperty "TipoClave", propiedadTipoClave, constTipoClave
    'propiedad sugerir proximo
    PropBag.WriteProperty "SugerirProxLibre", propiedadSugerirProxLibre, constSugerirProxLibre
    'propiedad tabla origen de del contador
    PropBag.WriteProperty "TablaContador", propiedadTablaContador, constTablaContador
    'propiedad ínide del campo contador en la tabla contador
    PropBag.WriteProperty "CampoCont", propiedadIndiceCampoCont, constIndiceCampoCont
    'propiedad del indice del campo clave de tipo contador en la tabala principal
    PropBag.WriteProperty "IndiceCampoClaveContador", propiedadIndiceCampoClaveContador, constIndiceCampoClaveContador
    'propiedad integridad
    PropBag.WriteProperty "Integridad", propiedadIntegridad, constIntegridad
    
    'propiedades de apariencia
    'NOTA: a las priopiedades de tipo OLE_COLOR no le puede asignar valor por defecto
    PropBag.WriteProperty "ColorFondoDatos", propColorFondoDatos, cColorFondoDatos
    PropBag.WriteProperty "ColorFondoGrilla", propColorFondoGrilla, cColorFondoGrilla
    PropBag.WriteProperty "ColorCaracteresIngreso", propColorCaracteresIngreso, cColorCaracteresIngreso
    PropBag.WriteProperty "ColorFondoCampoIngreso", propColorFondoCampoIngreso, cColorFondoCampoIngreso
    PropBag.WriteProperty "AnchoCeldas", propAnchoCeldas, cAnchoCeldas
    PropBag.WriteProperty "LargoCeldas", propLargoCeldas, cAnchoCeldas
    PropBag.WriteProperty "FuenteNombreCampo", propFuenteNombreCampo, cFuenteNombreCampo
    PropBag.WriteProperty "FuenteDatosIngresados", propFuenteDatosIngresados, cFuenteDatosIngresados
    PropBag.WriteProperty "FuenteDatosAIngresar", propFuenteDatosAIngresar, cFuenteDatosAIngresar
    PropBag.WriteProperty "MostrarLineas", propMostrarLineas, cMostrarLineas
End Sub

Private Sub UserControl_Resize()
    'Modifico el tamaño del control al respecificado por el usuario
    'modifico el tamaño de la grilla, el - 100 es para dejar un
    'pequeño margen.
    'NOTA: este evento ocurre después de haber ReadProperties
    On Error GoTo error
    'establesco tamaño de grilla
    gTabla.Width = UserControl.Width - 100
    gTabla.Height = UserControl.Height - UserControl.gTabla.Top
    
    'modifico el tamaño de la toolbar
    toolMenu.Width = UserControl.gTabla.Width
    
    'Oculto la columna de la grilla que se utiliza para almacenar las
    'porpiedades del campo correspondiente
    UserControl.gTabla.ColWidth(1) = 0
    
    'Oculto la columna de la grilla que se utiliza para almacenar datos en tiempo de ejecución
    'correspondientes a los datos ingresados por el usuario, por ejemplo
    'el item data del la opción de los combos seleccionados
    UserControl.gTabla.ColWidth(2) = 0
    
    'oculto la primera fila
    UserControl.gTabla.RowHeight(1) = 0
    
    'Establesco tamaño de la primer columna dependiendo de la propiedad
    UserControl.gTabla.ColWidth(0) = propLargoCeldas
    
    'Cambio el tamaño de los fuentes de la grilla
    gTabla.Font.Size = propFuenteDatosIngresados
        
    'Cambio de color el fondo de la grilla
    gTabla.BackColor = propColorFondoGrilla
    
    'Muestro o oculto líneas de la grilla dependiendo del valor de la propiedad
    gTabla.GridLines = propMostrarLineas
    gTabla.GridLinesFixed = propMostrarLineas
        
    Exit Sub
error:
    subControloErrores 515, "UserControl.Resize"
End Sub

Private Sub subCambioAnchoFilas(ancho As Long)
    'Recorro todas las filas de la grilla y establesco su ancho
    Dim i As Integer
    On Error GoTo error
    i = 2 'modifico a partir de la segunda fila
    Do While i < gTabla.Rows
        gTabla.Row = i
        gTabla.RowHeight(i) = ancho
        'modifico el tamaño del fuente de la primera columna
        gTabla.Col = 0
        gTabla.CellFontSize = propFuenteNombreCampo
        i = i + 1
    Loop
    Exit Sub
error:
    subControloErrores 515, "subCambioAnchoFilas"
End Sub
'***************************************************************
'*
'* Creo propiedades
'*
'***************************************************************

'Determina los campos de la tabla a mostrar en el ocx
Public Property Get campo() As String
Attribute campo.VB_MemberFlags = "400"
    On Error Resume Next
    campo = propiedadCampo
End Property

Public Property Let campo(ByVal nuevoCampo As String)
    'A) Obtengo la propiedad campo del cuadro de propiedades si la misma se cambia
    
    propiedadCampo = nuevoCampo 'obtengo el nuevo valor de la propiedad
    'la propiedad campo cambió por lo tanto se ejecuta el procedimiento WriteProperties
    PropertyChanged "Campo"
End Property

'Determina los copntroles de integridad referencial que se deben realizar
Public Property Get integridad() As String
Attribute integridad.VB_MemberFlags = "400"
    integridad = propiedadIntegridad
End Property

Public Property Let integridad(ByVal nuevaIntegridad As String)
    propiedadIntegridad = nuevaIntegridad
    PropertyChanged "Integridad"
End Property

'Determina la base de datos a trabajar con el ocx
Public Property Get CaminoBaseDeDatos() As String
    'Propiedad camino y nombre de la base de datos
    CaminoBaseDeDatos = propiedadCaminoBaseDeDatos
End Property

Public Property Let CaminoBaseDeDatos(ByVal nuevoCamino As String)
    propiedadCaminoBaseDeDatos = nuevoCamino
    PropertyChanged "CaminoBaseDeDatos"
    'cargo propiedad leída a  los controles data
    Data1.DatabaseName = propiedadCaminoBaseDeDatos 'utilizado para cargar los combos desde archivos
    Data2.DatabaseName = propiedadCaminoBaseDeDatos 'utilizado para manejar datos de la tabla princiapal
    Data3.DatabaseName = propiedadCaminoBaseDeDatos 'utilizado para:
                                                    'a) buscar próximo registro libre (sugerir próximo)
                                                    'b) busqueda y actualización de correlativos
End Property

Public Property Get ContraseñaBaseDeDatos() As String
    ContraseñaBaseDeDatos = propiedadContraseñaBaseDeDatos
End Property

Public Property Let ContraseñaBaseDeDatos(ByVal nuevoValor As String)
    propiedadContraseñaBaseDeDatos = nuevoValor
    Data1.Connect = propiedadContraseñaBaseDeDatos
    Data2.Connect = propiedadContraseñaBaseDeDatos
    Data3.Connect = propiedadContraseñaBaseDeDatos
    PropertyChanged "ContraseñaBaseDeDatos"
End Property

'Determina la tabla a trabajar con el ocx
Public Property Get tabla() As String
Attribute tabla.VB_MemberFlags = "400"
    tabla = propiedadTabla
End Property

Public Property Let tabla(ByVal nuevaTabla As String)
    propiedadTabla = nuevaTabla
    PropertyChanged "Tabla"
    'asigno propiedad a data2
    Data2.RecordSource = "Select * from " & propiedadTabla
End Property

'Determina el índice del campo clave de la tabla
Public Property Get IndiceCampoClave() As String
Attribute IndiceCampoClave.VB_MemberFlags = "400"
    IndiceCampoClave = propiedadIndiceCampoClave
End Property

Public Property Let IndiceCampoClave(ByVal nuevoIndice As String)
    propiedadIndiceCampoClave = nuevoIndice
    PropertyChanged "IndiceCampoClave"
End Property

Public Property Get IndiceCampoClaveContador() As Integer
Attribute IndiceCampoClaveContador.VB_MemberFlags = "400"
    IndiceCampoClaveContador = propiedadIndiceCampoClaveContador
End Property

Public Property Let IndiceCampoClaveContador(ByVal nuevoIndice As Integer)
    propiedadIndiceCampoClaveContador = nuevoIndice
    PropertyChanged "IndiceCampoClaveContador"
End Property

'Determina el típo de la clave de la tabla
'0= determinado por el usuario
'1= número correlativo
Public Property Get TipoClave() As Byte
Attribute TipoClave.VB_MemberFlags = "400"
    TipoClave = propiedadTipoClave
End Property

Public Property Let TipoClave(ByVal nuevoTipoClave As Byte)
    propiedadTipoClave = nuevoTipoClave
    PropertyChanged "TipoClave"
End Property

'Determina si sugiero o no próximo número libre si el tipo de clave
'es determinado por el usuario
Public Property Get SugerirProxLibre() As Byte
Attribute SugerirProxLibre.VB_MemberFlags = "400"
    SugerirProxLibre = propiedadSugerirProxLibre
End Property

Public Property Let SugerirProxLibre(nuevoLibre As Byte)
    If nuevoLibre = 0 Or nuevoLibre = 1 Then
        'valido que esta propiedad solo pueda tener dos valores
        'permitido: 0 y 1
        propiedadSugerirProxLibre = nuevoLibre
        PropertyChanged "SugerirProxLibre"
    End If
End Property

'Determina la tabla de donde obtengo el próximo número libre para el
'tipo de clave por número correlativo
Public Property Get TablaContador() As String
Attribute TablaContador.VB_MemberFlags = "400"
    TablaContador = propiedadTablaContador
End Property

Public Property Let TablaContador(ByVal nuevaTablaContador As String)
    propiedadTablaContador = nuevaTablaContador
    PropertyChanged "TablaContador"
End Property

'Determino el campo de la tabla donde se almacena el número de contador
Public Property Get CampoCont() As Integer
Attribute CampoCont.VB_MemberFlags = "400"
    CampoCont = propiedadIndiceCampoCont
End Property

Public Property Let CampoCont(ByVal nuevoIndiceCampoCont As Integer)
    propiedadIndiceCampoCont = nuevoIndiceCampoCont
    PropertyChanged "CampoCont"
End Property

'Determino propiedades de apariencia
Public Property Get ColorFondoDatos() As OLE_COLOR
    ColorFondoDatos = propColorFondoDatos
End Property

Public Property Let ColorFondoDatos(ByVal nuevoColorFondo As OLE_COLOR)
    propColorFondoDatos = nuevoColorFondo
    PropertyChanged "ColorFondoDatos"
End Property

Public Property Get ColorFondoGrilla() As OLE_COLOR
    ColorFondoGrilla = propColorFondoGrilla
End Property

Public Property Let ColorFondoGrilla(ByVal nuevoColorFondoGrilla As OLE_COLOR)
    propColorFondoGrilla = nuevoColorFondoGrilla
    PropertyChanged "ColorFondoGrilla"
End Property

Public Property Get ColorCaracteresIngreso() As OLE_COLOR
    ColorCaracteresIngreso = propColorCaracteresIngreso
End Property

Public Property Let ColorCaracteresIngreso(ByVal nuevoCaracter As OLE_COLOR)
    propColorCaracteresIngreso = nuevoCaracter
    PropertyChanged "ColorCaracteresIngreso"
End Property

Public Property Get ColorFondoCampoIngreso() As OLE_COLOR
    ColorFondoCampoIngreso = propColorFondoCampoIngreso
End Property

Public Property Let ColorFondoCampoIngreso(ByVal nuevoColor As OLE_COLOR)
    propColorFondoCampoIngreso = nuevoColor
    PropertyChanged "ColorFondoCampoIngreso"
End Property
Public Property Get AnchoCeldas() As Long
    AnchoCeldas = propAnchoCeldas
End Property

Public Property Let AnchoCeldas(ByVal nuevoAncho As Long)
    propAnchoCeldas = nuevoAncho
    If propAnchoCeldas < cAnchoMinimoCelda Then
        'no permito que las celdas sean menores al ancho mínimo
        propAnchoCeldas = cAnchoCeldas
    End If
    PropertyChanged "AnchoCeldas"
End Property

Public Property Get LargoCeldas() As Long
    LargoCeldas = propLargoCeldas
End Property

Public Property Let LargoCeldas(ByVal nuevoLargo As Long)
    On Error GoTo error
    propLargoCeldas = nuevoLargo
    If propLargoCeldas > (gTabla.Width - cLargoMinimoCeldaDatos) Then
        'si el largo de la primer celda sobrepasa de tal manera al laro de
        'la celda de datos, haciendo que esta última no cumpla con su medida
        'mínima, entonces modifico el tamaño de la primer celda
        propLargoCeldas = gTabla.Width - cLargoMinimoCeldaDatos
    End If
    PropertyChanged "LargoCeldas"
    Exit Property
error:
    subControloErrores 515, "Public property Let LargoCeldas"
End Property

Public Property Get FuenteNombreCampo() As Byte
    FuenteNombreCampo = propFuenteNombreCampo
End Property

Public Property Let FuenteNombreCampo(ByVal nuevoFuente As Byte)
    propFuenteNombreCampo = nuevoFuente
    PropertyChanged "FuenteNombreCampo"
End Property

Public Property Get FuenteDatosIngresados() As Byte
    FuenteDatosIngresados = propFuenteDatosIngresados
End Property

Public Property Let FuenteDatosIngresados(ByVal nuevoFuente As Byte)
    propFuenteDatosIngresados = nuevoFuente
    PropertyChanged "FuenteDatosIngresados"
End Property

Public Property Get FuenteDatosAIngresar() As Byte
    FuenteDatosAIngresar = propFuenteDatosAIngresar
End Property

Public Property Let FuenteDatosAIngresar(ByVal nuevoFuente As Byte)
    propFuenteDatosAIngresar = nuevoFuente
    PropertyChanged "FuenteDatosAIngresar"
End Property

Public Property Get MostrarLineas() As Byte
    MostrarLineas = propMostrarLineas
End Property

Public Property Let MostrarLineas(ByVal nuevaLinea As Byte)
    propMostrarLineas = nuevaLinea
    PropertyChanged "MostrarLineas"
End Property

'****************************************************************
'*
'*  Creo los campo en la grilla del control y asigno a la columna
'*  oculta los valores de las propiedades de cada campo creado
'*
'****************************************************************

Private Sub subCreoCampos(campos As String)
    'Recorro el string de propiedades y almaceno en la grilla
    'Recorro el string y muestro información de campos en grilla
    On Error GoTo error
    Dim largo As String
    Dim cadaCampo As String
    Dim caracter As String
    
    Dim i As Integer
    
    largo = Len(campos)
    i = 1
    Do While i <= largo
        caracter = Mid(campos, i, 1)    'obtengo todos y cada uno de los caracteres
        If caracter = "@" Then  '@ indica que finaliza el campo
            'creo una fila en la grilla de campos
            subMuestroCampoEnGrilla cadaCampo
            cadaCampo = ""  'inicializo para cargar información de un nuevo campo
        Else
            cadaCampo = cadaCampo & caracter
        End If
        i = i + 1
    Loop
    Exit Sub
error:
    subControloErrores 515, "subCreoCampos"
End Sub

Private Sub subMuestroCampoEnGrilla(campo As String)
    'Creo una nueva fila en la grilla, guardando las propiedades en la misma
    'correspondiente al campo creado.
    On Error GoTo error
    Dim tipoCampo As String
    Dim PropiedadesCampo As String
    Dim PropiedadesCampoAux As String
    Dim descriCampo As String
    
    PropiedadesCampo = ""
    PropiedadesCampoAux = ""
    
    tipoCampo = mfunObtengoValorDesdeStr(campo, 2, ";")     'tipo de dato del campo
    descriCampo = mfunObtengoValorDesdeStr(campo, 1, ";")   'descripción del campo
    Select Case tipoCampo
        Case 0  'string
            PropiedadesCampoAux = mfunObtengoValorDesdeStr(campo, 5, ";") & ";" & _
                                    mfunObtengoValorDesdeStr(campo, 6, ";") & ";" & _
                                    mfunObtengoValorDesdeStr(campo, 7, ";")
                                    'mínimo
                                    'máximo
                                    'campo memo

        Case 1  'numérico
            PropiedadesCampoAux = mfunObtengoValorDesdeStr(campo, 8, ";") & ";" & _
                                    mfunObtengoValorDesdeStr(campo, 9, ";") & ";" & _
                                    mfunObtengoValorDesdeStr(campo, 10, ";") & ";" & _
                                    mfunObtengoValorDesdeStr(campo, 11, ";")
                                    'mínimo
                                    'máximo
                                    'permitir decimales
                                    'permitir negativos
        Case 2  'fecha
            PropiedadesCampoAux = mfunObtengoValorDesdeStr(campo, 12, ";")
                                    'tipo de validación en ingreso de fecha
        Case 3  'combo fijio
            PropiedadesCampoAux = mfunObtengoValorDesdeStr(campo, 13, ";")
                                    'opciones del combo de opciones
        Case 4  'combo arch
            PropiedadesCampoAux = mfunObtengoValorDesdeStr(campo, 14, ";") & ";" & _
                                    mfunObtengoValorDesdeStr(campo, 15, ";") & ";" & _
                                    mfunObtengoValorDesdeStr(campo, 16, ";")
                                    'nombre de la tabla
                                    'indice del campo de la tabla
                                    'indice del campo que se asigna al itemData del combo
                                    
    End Select
    
    'agrego propiedades comunes a todos los campos
    PropiedadesCampo = tipoCampo & ";" & _
                    mfunObtengoValorDesdeStr(campo, 3, ";") & ";" & _
                    mfunObtengoValorDesdeStr(campo, 4, ";") & ";" & _
                    PropiedadesCampoAux
                    'tipo del campo
                    'indice del campo en la tabla
                    'permitir ingresar nulo
                    'RESTO DE LAS PROPIEDADES
    
    'creo nueva linea
    UserControl.gTabla.AddItem descriCampo
    'asigno propiedades a la columna oculata
    UserControl.gTabla.Col = 1
    UserControl.gTabla.Row = UserControl.gTabla.Rows - 1
    UserControl.gTabla.Text = PropiedadesCampo
Exit Sub
error:
    subControloErrores 515, "subMuestroCampoEnGrilla"
End Sub

'*******************************************************************
'*
'* Ingreso datos en la grilla.
'*
'*******************************************************************

Private Sub gTabla_Scroll()
    'Cuando utilizo la barra de scroll, también muevo el control de ingreso de datos.
    gTabla_SelChange
End Sub

Private Sub gTabla_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then 'enter
        'cada vez que digito enter le doy el focus al control de ingreso de datos
        'correspondiente a la fila donde estoy parado, permitiendo de esta manera
        'el ingreso o la modificación de un dato
        
        ControlSeleccionadoEnGrilla.SetFocus    'hace referecncia la objeto seleccionado actualmente
        
        'es posible que el control todavía no este cargado por lo que si se
        'presiona enter se producirá un error que será interceptado, pero que no
        'afecta al comportamiento del ocx.
    End If
End Sub

Private Sub gTabla_SelChange()
    'Utilizo este evento de la grilla para mostrar un control que
    'permitirá ingresar datos a la misma.
    On Error GoTo error
    Dim tipoControlAMostrar As String
    Dim PropControlAMostrar As String
    Dim CampoMemo As Boolean
    Dim descTipoDeDato As String
    'propiedades para ingresar caracteres
    Dim StrMax As Integer
    
    If habilitoSelChange = True Then
        'obtengo propiedades de la columna de propiedades, correspondiente a la nueva fila
        'en la que se desea ingresar datos
        PropControlAMostrar = UserControl.gTabla.TextMatrix(UserControl.gTabla.Row, 1)
        
        'obtengo tipo del control a mostrar
        tipoControlAMostrar = mfunObtengoValorDesdeStr(PropControlAMostrar, 1, ";")
        Select Case tipoControlAMostrar
            Case 0  'mustro un textbox
                descTipoDeDato = "Texto"
                'Para el ingreso de los campos memo utilizo otro control por eso
                'es necesario determinar si el string a ingresasr es o no del tipo memo.
                CampoMemo = CBool(mfunObtengoValorDesdeStr(PropControlAMostrar, 6, ";"))
                If CampoMemo Then
                    subMuestroControlEnGrilla UserControl.txtIngstrMemo
                    'asigno control a varible para poder trabajar con él
                    Set ControlSeleccionadoEnGrilla = txtIngstrMemo
                Else
                    subMuestroControlEnGrilla UserControl.txtIngStr
                    'asigno control a varible para poder trabajar con él
                    Set ControlSeleccionadoEnGrilla = txtIngStr
                End If
                'establesco porpiedades
                StrMax = mfunObtengoValorDesdeStr(PropControlAMostrar, 5, ";")
                'asigo maximo determinado para el string a la propiedad maxlengh
                UserControl.txtIngStr.MaxLength = StrMax
                UserControl.txtIngstrMemo.MaxLength = StrMax
                'muestro valor de la grilla en control
                ControlSeleccionadoEnGrilla.Text = gTabla.TextMatrix(gTabla.Row, 3)
                
            Case 1  'muestro un txtbox solo númerico
                descTipoDeDato = "Numérico"
                subMuestroControlEnGrilla UserControl.txtIngNum
                'establesco propiedades
                ValorDecimal = CBool(mfunObtengoValorDesdeStr(PropControlAMostrar, 6, ";"))
                ValorNegativo = CBool(mfunObtengoValorDesdeStr(PropControlAMostrar, 7, ";"))
                'asigno control a varible para poder trabajar con él
                Set ControlSeleccionadoEnGrilla = txtIngNum
                'muestro valor de la grilla en control
                ControlSeleccionadoEnGrilla.Text = gTabla.TextMatrix(gTabla.Row, 3)
                
            Case 2  'muestro un textbox solo fecha
                descTipoDeDato = "Fecha dd/mm/aa"
                subMuestroControlEnGrilla UserControl.txtIngFecha
                'asigno control a varible para poder trabajar con él
                Set ControlSeleccionadoEnGrilla = txtIngFecha
                'muestro valor de la grilla en control
                ControlSeleccionadoEnGrilla.Text = gTabla.TextMatrix(gTabla.Row, 3)
                
            Case 3  'muestro un combobox con opciones predeterminadas
                descTipoDeDato = "Opciones"
                'establesco elemento del combo
                subCargoComboDesdeLista mfunObtengoValorDesdeStr(PropControlAMostrar, 4, ";")
                subMuestroControlEnGrilla UserControl.cboIngOpcF
                'asigno control a varible para poder trabajar con él
                Set ControlSeleccionadoEnGrilla = cboIngOpcF
                'muestro valor de la grilla en control
                mSubPosicionoCombo UserControl.cboIngOpcF, Val(gTabla.TextMatrix(gTabla.Row, 2))
                
            Case 4  'muestro un combobox con información desde archivo
                descTipoDeDato = "Opciones"
                'establesco elemetos del combo
                subCargoComboDesdeArch PropControlAMostrar
                subMuestroControlEnGrilla UserControl.cboIngOpcA
                'asigno control a varible para poder trabajar con él
                Set ControlSeleccionadoEnGrilla = cboIngOpcA
                'muestro valor de la grilla en control
                mSubPosicionoCombo UserControl.cboIngOpcA, Val(gTabla.TextMatrix(gTabla.Row, 2))
                
        End Select
        'Muestro información en grilla del campo que estoy ingresado
        gTabla.TextMatrix(0, 3) = "Ingrese " & gTabla.TextMatrix(gTabla.Row, 0)
        gTabla.TextMatrix(0, 0) = descTipoDeDato
    End If
Exit Sub
error:
    subControloErrores 515, "gTabla_SelChange"
End Sub

Private Sub subCargoComboDesdeArch(propComboArch As String)
    'Para caragr este combo es necesario obtener la información desde
    'un campo de un archivo determinado , ambos datos se especifican
    'en la porpiedad del campo
    Dim tablaOrigen As String
    Dim campoTablaOrigen As Integer
    Dim consulta As String
    Dim clave As Integer
    
    On Error GoTo error
    'inicializo combo
    UserControl.cboIngOpcA.Clear
    
    tablaOrigen = mfunObtengoValorDesdeStr(propComboArch, 4, ";")
    campoTablaOrigen = mfunObtengoValorDesdeStr(propComboArch, 5, ";")
    clave = mfunObtengoValorDesdeStr(propComboArch, 6, ";")
    
    'Por medio de esta consulta sql obtengo todos los registros
    'de la tabla indicada, como así tambien todos sus campos
    consulta = _
        "Select * from " & tablaOrigen
        
    Data1.RecordSource = consulta
    Data1.Refresh
    
    'verifico si el archivo tiene registros
    If Data1.Recordset.RecordCount > 0 Then
        Data1.Recordset.MoveFirst
        'recorro el recordset creado y cargo el valor del campo correspondiente
        'en el combo
        Do While Not Data1.Recordset.EOF
            'agrego elemento al combo
            UserControl.cboIngOpcA.AddItem Data1.Recordset(campoTablaOrigen)
            UserControl.cboIngOpcA.ItemData(UserControl.cboIngOpcA.NewIndex) = Data1.Recordset(clave)
            Data1.Recordset.MoveNext
        Loop
    Else
        'desencadeno evento a la aplicación cliente
        RaiseEvent NoHayDatosSuficientes(tablaOrigen)
    End If
Exit Sub
error:
    subControloErrores 512, "subCarboComboDesdeArch"
End Sub

Private Sub subCargoComboDesdeLista(lista As String)
    'Se encarga de cargar el combo con los elementos de la propiedad correspondiente
    'el formato de la lista recivida es:
    '1nombre_elemento#2nombre_elemento#Nnombre_elemento
    On Error GoTo error
    Dim largo As Integer
    Dim i As Integer
    Dim caracter As String
    Dim elemento As String
    Dim contEle As Integer
    
    contEle = 1
    i = 1
    largo = Len(lista)
    'inicializo combo para elminar opciones anteriores
    cboIngOpcF.Clear
    'recorro la cadena y obtengo cada elemento
    Do While i <= largo
        caracter = Mid(lista, i, 1)
        If caracter = "#" Then  'indica un nuevo elmento
            'agrego un nuevo elemento a la lista
            cboIngOpcF.AddItem elemento
            
            'Cuando se almacene en el archivo el valor de este combo
            'en realidad lo que se almacenará sera el valor almacenado
            'en itemdata, para así poder manejar la información en la
            'aplicación de forma más choerente que si se almacenara el
            'valor de la propiedad text del combo
            cboIngOpcF.ItemData(cboIngOpcF.NewIndex) = contEle
            elemento = ""
            contEle = contEle + 1
        Else
            elemento = elemento & caracter
        End If
        i = i + 1
    Loop
Exit Sub
error:
    subControloErrores 515, "subCargoComboDesdeLista"
End Sub

Private Sub subMuestroControlEnGrilla(Control As Object)
    'Muestro un control en la grilla, adaptando su tamaño y su posición
    'con respecto a la columna visible nro. 3
    On Error GoTo error
    Dim controlAMostrar As Object
    Dim difPuntoC As Integer
    Dim restoAMostrar As Integer
    Set controlAMostrar = Control
    
    'oculto el control del campo anterior, como no se cual es oculto todos
    txtIngStr.Visible = False
    txtIngNum.Visible = False
    cboIngOpcF.Visible = False
    cboIngOpcA.Visible = False
    txtIngFecha.Visible = False
    txtIngstrMemo.Visible = False
    
    'posiciono control
    controlAMostrar.Top = gTabla.RowPos(gTabla.Row) + gTabla.Top + 50
    controlAMostrar.Left = (gTabla.ColPos(3) + gTabla.Left) + 50    '-50 es para corregir pequeño def. vis.
    'determino tamaño control
    'cambio ancho del control
    controlAMostrar.Height = gTabla.RowHeight(gTabla.Row)
    If TypeOf Control Is TextBox Then
        If Control.MultiLine = True Then    'es un campo memo
            'aumento el ancho del campo tantas veces como indique la constante
            controlAMostrar.Height = _
                    (gTabla.RowHeight(gTabla.Row)) * contsAnchoMemo
            'ademas existe la posibilidad de que el control memo no entre dento del control
            'por aparecer(por ejemplo) en la última fila de la grilla. Por este motivo
            'es necesario poder reposicionar el control para que puede trabajar comodamente con él.
            difPuntoC = gTabla.Height - (gTabla.RowPos(gTabla.Row) + gTabla.Top)
            If difPuntoC < controlAMostrar.Height Then
                restoAMostrar = controlAMostrar.Height - difPuntoC
                controlAMostrar.Top = (gTabla.RowPos(gTabla.Row) + gTabla.Top) - restoAMostrar _
                + gTabla.RowHeight(gTabla.Row)
            End If
        End If
    Else
        If TypeOf Control Is ListBox Then
            'aumento el ancho del campo tantas veces como indique la constante
            controlAMostrar.Height = gTabla.RowHeight(gTabla.Row) * cAnchoLista
        End If
    End If
    'cambio largo de control
    controlAMostrar.Width = gTabla.ColWidth(3) - 10 '- 10 es para corregir pequeño defecto visual
    'establesco color del fondo del control
    controlAMostrar.BackColor = propColorFondoCampoIngreso
    'establesco tamaño del fuente
    controlAMostrar.FontSize = propFuenteDatosAIngresar
    'establesco el color del fuente igual al de la grilla
    controlAMostrar.ForeColor = gTabla.ForeColor
    'muestro control
    controlAMostrar.Visible = True
    Set controlAMostrar = Nothing
    'Estas varibles determinan si estoy trabajando con el campo clave de la tabla
    'su contenido es relevante cuando el tipo de clave de la tabla = 0 (determinada por usuario)
    'Para simplificar el código asigno false a las variables correspondientes a todos los tipos
    'Las mismas se inicializan a True cundo le doy el focus al control correspondiente al campo clave
    campoClaveCboA = False
    campoClaveCboF = False
    campoClaveNum = False
    campoClaveStr = False
Exit Sub
error:
    subControloErrores 515, "subMuestroControlEnGrilla"
End Sub

'Descripción de la utilidad de los eventos
'GotFocus
    'a) Cuando se digita enter se le da el focus al control activo
    'para que se permita el ingreso de datos en el control, modificando las propiedades
    'del mismo para que el usuario perciva que puede ingresar datos
    'b) Cuando el usuario hace click sobre el control también sucede lo mismo

'Keypress
    'a) si digito la tecla enter se graban los datos a la grilla
    'b)En en los campo de tipo numérico se utiliza para controlar que se ingresen solo
    'caracteres numérico
    
'Change
    'Es utilizado si el campo que estoy modificando es el campo clave de la tabla
    'Si es así busco información de ese registro y muestro en la grilla
    
'LostFocus
    'Inicializo las propiedades del control para indicar que perdió el focus, es decir
    'que no se pueden ingresar más datos en el.
    
'**************************
'* Tipo de campo numérico
'**************************
Private Sub txtIngNum_GotFocus()
    'Cuando el control obtiene el foco: modifico las pripiedades
    'para indicar al usuario que se puede ingresasr datos en el
    On Error GoTo error
    subHabilitoControlParaIngresoDeDatos txtIngNum
    'verifico si estoy modificando el campo clave
    If gTabla.Row = filaCampoClave Then
        'estoy modificando un campo clave
        campoClaveNum = True    'busco dato ingresado en la tabla
    End If
Exit Sub
error:
    subControloErrores 515, "txtIngNum_GotFocus"
End Sub

Private Sub txtIngNum_KeyPress(KeyAscii As Integer)
    On Error GoTo error
    'El usuario confirma los datos
    If KeyAscii = 13 Then   'enter
        'asigno dato ingresado en el control a la grilla
        gTabla.TextMatrix(gTabla.Row, 3) = txtIngNum.Text
        gTabla.SetFocus
    End If
    If KeyAscii = 27 Then   'esc
        'Inicializo control con los datos anteriores
        txtIngNum = gTabla.TextMatrix(gTabla.Row, 3)
        gTabla.SetFocus
    End If
    'Verifico que el tipo de dato sea el correcto
    'Voy a permitir el ingreso de números negativos y o decimales
    'si las propiedades del campo así lo indican
    mSubValidoNum KeyAscii, ValorDecimal, ValorNegativo
Exit Sub
error:
    subControloErrores 515, "txtIngNum_KeyPress"
End Sub

Private Sub txtIngNum_Change()
    On Error GoTo error
    'si estoy modificando un campo clave
    If campoClaveNum Then
        'busco registro por dicha clave
        If funBuscoRegistro(Val(txtIngNum.Text)) Then
            'si existe muestro los datos
            subMuestroDatosEnGrilla
            subInicializoBotones 1  'modificar
        Else
            'inicializo grilla para nuevo ingreso
            subInicializoGrilla
            subInicializoBotones 0  'guardar
        End If
    End If
Exit Sub
error:
    subControloErrores 515, "txtIngNum_Change"
End Sub

Private Sub txtIngNum_LostFocus()
    On Error GoTo error
    subDeshabilitoControlParaIngresoDeDatos txtIngNum
Exit Sub
error:
    subControloErrores 515, "txtIngNum_LostFocus"
End Sub
'**************************
'* Tipo de campo string
'**************************
Private Sub txtIngStr_GotFocus()
    'Cunado el control obtiene el foco modifico las propiedades
    'para indicar al usuario que se puede ingresasr datos en el
    On Error GoTo error
    subHabilitoControlParaIngresoDeDatos txtIngStr
    'verifico si estoy modificando el campo clave
    If gTabla.Row = filaCampoClave Then
        'estoy modificando un campo clave
        campoClaveStr = True    'busco dato ingresado en la tabla
    End If
Exit Sub
error:
    subControloErrores 515, "txtIngStr_GotFocus"
End Sub

Private Sub txtIngStr_KeyPress(KeyAscii As Integer)
    On Error GoTo error
    'El usuario confirma los datos
    If KeyAscii = 13 Then   'enter
        'asigno dato ingresado en el control a la grilla
        gTabla.TextMatrix(gTabla.Row, 3) = txtIngStr.Text
        gTabla.SetFocus
    End If
    If KeyAscii = 27 Then   'esc
        'Inicializo control con los datos anteriores
        txtIngStr.Text = gTabla.TextMatrix(gTabla.Row, 3)
        gTabla.SetFocus
    End If
Exit Sub
error:
    subControloErrores 515, "txtIngStr_KeyPress"
End Sub

Private Sub txtIngStr_Change()
    On Error GoTo error
    'si estoy modificando un campo clave
    If campoClaveStr Then
        'busco registro por dicha clave
        If funBuscoRegistro(txtIngStr.Text) Then
            'si existe muestro los datos
            subMuestroDatosEnGrilla
            subInicializoBotones 1  'modificar
        Else
            'inicializo grilla para nuevo ingreso
            subInicializoGrilla
            subInicializoBotones 0  'guardar
        End If
    End If
Exit Sub
error:
    subControloErrores 515, "txtIngStr_Change"
End Sub

Private Sub txtIngStr_LostFocus()
    On Error GoTo error
    subDeshabilitoControlParaIngresoDeDatos txtIngStr
Exit Sub
error:
    subControloErrores 515, "txtIngStr_LostFocus"
End Sub

'**************************
'* Tipo de campo memo
'**************************
Private Sub txtIngstrMemo_GotFocus()
    'Cunado el control obtiene el foco modifico las pripiedades
    'para indicar al usuario que se puede ingresasr datos en el
    On Error GoTo error
    subHabilitoControlParaIngresoDeDatos txtIngstrMemo
Exit Sub
error:
    subControloErrores 515, "txtIngStrMemo_GotFocus"
End Sub

Private Sub txtIngstrMemo_KeyPress(KeyAscii As Integer)
    On Error GoTo error
    If KeyAscii = 27 Then   'esc
        'Inicializo control con los datos anteriores
        txtIngstrMemo.Text = gTabla.TextMatrix(gTabla.Row, 3)
        gTabla.SetFocus
    End If
    
    If KeyAscii = 13 Then
        'asigno dato ingresado en el control a la grilla
        gTabla.TextMatrix(gTabla.Row, 3) = txtIngstrMemo.Text
        gTabla.SetFocus
    End If
Exit Sub
error:
    subControloErrores 515, "txtIngStrMemo_KeyPress"
End Sub

Private Sub txtIngstrMemo_LostFocus()
    On Error GoTo error
    subDeshabilitoControlParaIngresoDeDatos txtIngstrMemo
Exit Sub
error:
    subControloErrores 515, "txtIngStrMemo_LostFocus"
End Sub

'**************************
'* Tipo de campo fecha
'**************************
Private Sub txtIngFecha_GotFocus()
    On Error GoTo error
    'Cuando el control obtiene el foco modifico las pripiedades
    'para indicar al usuario que se puede ingresasr datos en el
    subHabilitoControlParaIngresoDeDatos txtIngFecha
Exit Sub
error:
    subControloErrores 515, "txtIngFecha_GotFocus"
End Sub

Private Sub txtIngFecha_KeyPress(KeyAscii As Integer)
    On Error GoTo error
    'El usuario confirma los datos
    If KeyAscii = 13 Then   'enter
        'asigno dato ingresado en el control a la grilla
        gTabla.TextMatrix(gTabla.Row, 3) = txtIngFecha.Text
        gTabla.SetFocus
    End If
    If KeyAscii = 27 Then   'esc
        'Inicializo el controlcon los datos anteriores
        txtIngFecha.Text = gTabla.TextMatrix(gTabla.Row, 3)
        gTabla.SetFocus
    End If
Exit Sub
error:
    subControloErrores 515, "txtIngFecha_KeyPress"
End Sub

Private Sub txtIngFecha_LostFocus()
    On Error GoTo error
    subDeshabilitoControlParaIngresoDeDatos txtIngFecha
Exit Sub
error:
    subControloErrores 515, "txtIngFecha_LostFocus"
End Sub

'*********************************
'* Tipo de campo combo opc. arch.
'*********************************
Private Sub cboIngOpcA_GotFocus()
    On Error GoTo error
    'Cuando el control obtiene el foco modifico las pripiedades
    'para indicar al usuario que se puede ingresasr datos en el
    subHabilitoControlParaIngresoDeDatos cboIngOpcA
    'verifico si estoy modificando el campo clave
    If gTabla.Row = filaCampoClave Then
        'estoy modificando un campo clave
        campoClaveCboA = True    'busco dato ingresado en la tabla
    End If
Exit Sub
error:
    subControloErrores 515, "cboIngOpcA_GotFocus"
End Sub

Private Sub cboIngOpcA_KeyPress(KeyAscii As Integer)
    On Error GoTo error
    'El usuario confirma los datos
    If KeyAscii = 13 Then   'enter
        'si no tengo seleccionado ningún elemento no ingreso nada
        If cboIngOpcA.ListIndex <> -1 Then
            'asigno la propiedad texto del combo a la columna visible
            gTabla.TextMatrix(gTabla.Row, 3) = cboIngOpcA.Text
            'asigno la priopiedad itemdata que es el valor que voy a guardar en la tablas
            'a la columna 2 que esta oculta
            gTabla.TextMatrix(gTabla.Row, 2) = cboIngOpcA.ItemData(cboIngOpcA.ListIndex)
            gTabla.SetFocus
        End If
    End If
    If KeyAscii = 27 Then   'esc
        'inicializo el control con los datos anteriores
        If gTabla.TextMatrix(gTabla.Row, 3) <> Empty Then
            cboIngOpcA.Text = gTabla.TextMatrix(gTabla.Row, 3)
        Else
            cboIngOpcA.ListIndex = -1
        End If
        gTabla.SetFocus
    End If
Exit Sub
error:
    subControloErrores 515, "cboIngOpcA_KeyPress"
End Sub

Private Sub cboIngOpcA_Scroll()
On Error GoTo error
    If campoClaveCboA Then
        'busco registro por dicha clave
        If funBuscoRegistro(cboIngOpcA.ItemData(cboIngOpcA.ListIndex)) Then
            'si existe muestro los datos
            subMuestroDatosEnGrilla
            subInicializoBotones 1  'modificar
        Else
            'inicializo grilla para nuevo ingreso
            subInicializoGrilla
            subInicializoBotones 0  'guardar
        End If
    End If
Exit Sub
error:
    subControloErrores 515, "cboIngOpcA_Scroll"
End Sub

Private Sub cboIngOpcA_LostFocus()
    On Error GoTo error
    subDeshabilitoControlParaIngresoDeDatos cboIngOpcA
Exit Sub
error:
    subControloErrores 515, "cboIngOpcA_LostFocus"
End Sub

'*********************************
'* Tipo de campo combo opc. fijas
'*********************************
Private Sub cboIngOpcF_GotFocus()
    On Error GoTo error
    'Cuando el control obtiene el foco modifico las pripiedades
    'para indicar al usuario que se puede ingresasr datos en el
    subHabilitoControlParaIngresoDeDatos cboIngOpcF
    'verifico si estoy modificando el campo clave
    If gTabla.Row = filaCampoClave Then
        'estoy modificando un campo clave
        campoClaveCboF = True    'busco dato ingresado en la tabla
    Else
        'no es un campo clave
        campoClaveCboF = False   'no hago nada
    End If
Exit Sub
error:
    subControloErrores 515, "cboIngOpcF_GotFocus"
End Sub

Private Sub cboIngOpcF_KeyPress(KeyAscii As Integer)
    On Error GoTo error
    'El usuario confirmó los datos
    If KeyAscii = 13 Then   'enter
        'si no tengo seleccionado ningún elemento no ingreso nada
        If cboIngOpcF.ListIndex <> -1 Then
            'asigno la propiedad texto del combo a la columna visible
            gTabla.TextMatrix(gTabla.Row, 3) = cboIngOpcF.Text
            'asigno la priopiedad itemdata que es el valor que voy a guardar en la tablas
            'a la columna 2 que esta oculta
            gTabla.TextMatrix(gTabla.Row, 2) = cboIngOpcF.ItemData(cboIngOpcF.ListIndex)
            gTabla.SetFocus
        End If
    End If
    If KeyAscii = 27 Then   'esc
        'inicializo el control con los datos anteriores
        If gTabla.TextMatrix(gTabla.Row, 3) <> Empty Then
            cboIngOpcF.Text = gTabla.TextMatrix(gTabla.Row, 3)
        Else
            cboIngOpcF.ListIndex = -1
        End If
        gTabla.SetFocus
    End If
Exit Sub
error:
    subControloErrores 515, "CboIngOpcF_KeyPress"
End Sub

Private Sub cboIngOpcF_LostFocus()
    On Error GoTo error
    subDeshabilitoControlParaIngresoDeDatos cboIngOpcF
Exit Sub
error:
    subControloErrores 515, "cboIngOpcF_LostFocus"
End Sub

Private Sub subHabilitoControlParaIngresoDeDatos(controlActivo As Object)
    'Cuando el control con el que se está trabajando tiene el focus se inicializan
    'ciertas prpopiedades para idicar al usuario que puede modificar su contenido
    On Error GoTo error
    If TypeOf controlActivo Is TextBox Then
        'selecciono los datos ya ingresados para mejorar interface
        'pero solo a los controles de tipo textbox
        controlActivo.SelStart = 0
        controlActivo.SelLength = Len(controlActivo)
    End If
    'modifico color del fondo del control
    controlActivo.BackColor = propColorFondoDatos
    'modifico el color de la fuente del color
    controlActivo.ForeColor = propColorCaracteresIngreso
    'modifico el tamaño del fuente
    controlActivo.FontSize = propFuenteDatosAIngresar
Exit Sub
error:
    subControloErrores 515, "subHabilitoControlParaIngresoDeDatos"
End Sub

Private Sub subDeshabilitoControlParaIngresoDeDatos(controlActivo As Object)
    'Cuando el control que está ingresando los datos pierde el focus
    'inicializo nuevamente las propiedades indicando que no se puede ingresasr datos
    'en en control
    On Error GoTo error
    'modifico color del fondo del control
    controlActivo.BackColor = propColorFondoCampoIngreso
    'modifico el color de la fuente del color
    controlActivo.ForeColor = propColorCaracteresIngreso
    'modifico el tamaño del fuente
    controlActivo.FontSize = propFuenteDatosIngresados
    'establesco el color del fuente igual al de la grilla
    controlActivo.ForeColor = gTabla.ForeColor
Exit Sub
error:
    subControloErrores 515, "subDeshabilitoControlParaIngresoDeDatos"
End Sub

'********************************************************************
'*
'*  Procedimientos para que el control maneje los datos de la tabla
'*  correspondiente
'*
'*
'*********************************************************************

Private Function funBuscoRegistro(clave As Variant) As Boolean
    'Busco un registro determinado en la tabla
    On Error GoTo error
    Dim consulta As String
    Dim nombreCampoClave As String
    
    funBuscoRegistro = False
    'verifico el tipo de la clave
    If Not IsNumeric(clave) Then
        'si la clave es de tipo texto tengo que anexarle comillas para que la consulta
        'sql se realize correctamente
        clave = "'" & clave & "'"
    End If
    'Con este refresh cargo el control data con todos los registro de la tabla
    'ya que es es necesario que la tabla este cargada para poder obtener
    'el nombre de cada campo.
    
    'La propiedad recordsource de este control se inicializa en el property let
    'de la propiedad tabla o en el metodo MostrarRegistro
    Data2.Refresh
    
    'Obtengo el nombre del campo clave de la tabla
    nombreCampoClave = Data2.Recordset(propiedadIndiceCampoClave).Name
    'busco registro
    consulta = "Select * from " & propiedadTabla & _
                " Where " & nombreCampoClave & " = " & clave
    Data2.RecordSource = consulta
    'Con este refresh tengo que obtener el registro que coincida con la clave
    Data2.Refresh
    If Data2.Recordset.RecordCount = 1 Then
        'encontré registro
        funBuscoRegistro = True
    End If
Exit Function
error:
    subControloErrores 513, "funBuscoRegistro"
End Function

Private Sub subMuestroDatosEnGrilla()
    'Muestra los campos de un registro correspondiente en el control
    'tipoCbo es optional ya que se utiliza solamente para mostrar los campos de tipo combo
    On Error GoTo error
    Dim i As Integer
    Dim Pcontrol As String
    Dim tipoControl As String
    Dim listaCombo As String
    Dim indiceCampoActual As Integer
    
    'recorro la grilla y verifico el indice de cada campo
    i = 2 'comienzo a partir de la segunda fila
    Do While i < gTabla.Rows
        'obtengo todas las propiedades del control
        Pcontrol = UserControl.gTabla.TextMatrix(i, 1)
        'obtengo tipo del control a mostrar
        tipoControl = mfunObtengoValorDesdeStr(Pcontrol, 1, ";")
        'obtengo índice del campo actualmente seleccionado
        indiceCampoActual = mfunObtengoValorDesdeStr(Pcontrol, 2, ";")
        'muestro el valor del campo en la grilla
        Select Case tipoControl
            Case 3 'estoy trabajando con combo de opciones fijas
                'Para mostrar la descripción a la cual pertenece el código almacenado en la tabla:
                'tengo que cargar el combo de opciones nuevamente pero sin mostrarlo por pantalla
                listaCombo = mfunObtengoValorDesdeStr(Pcontrol, 4, ";")
                subCargoComboDesdeLista (listaCombo)
                'asigno la priopiedad itemdata que es el valor que voy a guardar en la tablas
                'a la columna 2 que esta oculta
                gTabla.TextMatrix(i, 2) = Data2.Recordset(indiceCampoActual)
                'busco en el combo el valor del itemdata
                mSubPosicionoCombo UserControl.cboIngOpcF, Data2.Recordset(indiceCampoActual)
                'ahora que tengo el string en el combo lo asigno a la grilla
                gTabla.TextMatrix(i, 3) = UserControl.cboIngOpcF.Text
                
            Case 4  'estoy trabajando con un combo con opciones desde archivos
                'Para mostrar la descripción a la cual pertenece el código almacenado en la tabla:
                'tengo que cargar el combo de opciones nuevamente pero sin mostrarlo por pantalla
                subCargoComboDesdeArch (Pcontrol)
                
                'asigno la priopiedad itemdata que es el valor que voy a guardar en la tablas
                'a la columna 2 que esta oculta
                gTabla.TextMatrix(i, 2) = Data2.Recordset(indiceCampoActual)
                'busco en el combo el valor del itemdata
                mSubPosicionoCombo UserControl.cboIngOpcA, Data2.Recordset(indiceCampoActual)
                'ahora que tengo el string en el combo lo asigno a la grilla
                gTabla.TextMatrix(i, 3) = UserControl.cboIngOpcA.Text
            Case Else
                'sino simplemente cargo el dato del archivo a la grilla
                If Not IsNull(Data2.Recordset(indiceCampoActual)) Then
                    gTabla.TextMatrix(i, 3) = Data2.Recordset(indiceCampoActual)
                End If
                
        End Select
        i = i + 1
    Loop
Exit Sub
error:
    subControloErrores 515, "subMuestroDatosEnGrilla"
End Sub

Private Sub subInicializoGrilla()
    'Inicializo al grilla para nuevo ingreso de datos
    'recorro la grilla y verifico el indice de cada campo
    On Error GoTo error
    Dim i As Integer
    
    i = 2 'comienzo a partir de la segunda fila
    Do While i < gTabla.Rows
        gTabla.TextMatrix(i, 2) = ""    'inicializo columna de itemData
        gTabla.TextMatrix(i, 3) = ""    'inicializo columna de datos
        i = i + 1
    Loop
Exit Sub
error:
    subControloErrores 515, "subInicializoGrilla"
End Sub

Private Sub subInicializoBotones(tipo As Byte)
    'Muestra los botones de la barra de botones, dependiendo de la operación a realizar
    On Error GoTo error
    If tipo = 0 Then    'inicializo barra para poder grabar
        UserControl.toolMenu.Buttons(1).Enabled = True  'bot. guardar
        UserControl.toolMenu.Buttons(2).Enabled = False  'bot. modificar
        UserControl.toolMenu.Buttons(3).Enabled = False  'bot. borrar
        'este evento indica que en el control se estan ingresando datos.
        '(boton guardar activado)
        RaiseEvent CambioOperacion(1)
    End If
    If tipo = 1 Then    'inicializo barra para poder modificar
        UserControl.toolMenu.Buttons(1).Enabled = False  'bot. modificar
        UserControl.toolMenu.Buttons(2).Enabled = True  'bot. modificar
        UserControl.toolMenu.Buttons(3).Enabled = True  'bot. borrar
        'este evento indica que en el control se estan modificando datos.
        '(boton modificar activado)
        RaiseEvent CambioOperacion(2)
    End If
    If tipo = 2 Then    'desabilito todos los controles
        UserControl.toolMenu.Buttons(1).Enabled = False  'bot. modificar
        UserControl.toolMenu.Buttons(2).Enabled = False  'bot. modificar
        UserControl.toolMenu.Buttons(3).Enabled = False  'bot. borrar
        'este evento indica que en el control no esta en uso.
        '(botones guardar,modificar, y borrar desactivados)
        RaiseEvent CambioOperacion(0)
    End If
Exit Sub
error:
    subControloErrores 515, "subInicializoBotones"
End Sub

'************************************************************************************
'*
'*          Trabajo con botones de la barra
'*
'*************************************************************************************

Private Sub toolMenu_ButtonClick(ByVal Button As ComctlLib.Button)
    'Determino que boton fue presionado
    On Error GoTo error
    subDigitoBoton Button.Index
Exit Sub
error:
    subControloErrores 515, "toolMenu_ButtonClick"
End Sub

Private Sub subDigitoBoton(boton As Byte)
    'Determino que boton fue digitado y realizo las operaciones correspondientes
    'Este procedimiento puede ser llamado cuando se hace un click sobre el boton
    'de la toolbar o cuando se ejecuta algunos de los métodos proporcionados
    'por el ocx.(grabardatos,limpiardatos,ect)
    On Error GoTo error
    Select Case boton
        Case 6  'sugiero próximo libre
            subInicializoGrilla     'limpio grilla
            subInicializoBotones 2  'desabilito controles
            gTabla.Row = filaCampoClave     'obtengo campo clave
            gTabla_SelChange                'ejecutando este evento muestro el control de ingreso de datos
            gTabla_KeyPress (13)            'ejecutando este evento habilito el control para ingresar datos
            
            'como se de antemano que solo puedo trabajar con el boton de sugerir próximo cuando el campo clave
            'es de tipo numérico inicializo la variable que permite la búsqueda del registro a true.
            campoClaveNum = True
            'simulo que se ingreso el número, provocando el evento changed
            ControlSeleccionadoEnGrilla.Text = funObtengoProximoLibre
                        
        Case 7  'limpio para nuevo ingreso
            subInicializoGrilla     'limpio grilla
            gTabla.Row = 2          'muestro el primer control de ingreso
            gTabla_SelChange
            subInicializoBotones 2  'desabilito controles
            'determino tipo de clave de la tabla
            If propiedadTipoClave = 1 Then
                'Si la clave es de tipo correlativo
                'habilito boton de grabar
                subInicializoBotones 0
            End If
            
        Case 1 'guardar
            subGraboDatos
            If Not errorEnIngresoDeDatos Then
                subInicializoGrilla     'limpio grilla
                gTabla.Row = 2          'muestro el primer control de ingreso
                gTabla_SelChange
                subInicializoBotones 2  'desabilito controles
                'determino tipo de clave de la tabla
                If propiedadTipoClave = 1 Then
                    'Si la clave es de tipo correlativo
                    'habilito boton de grabar
                    subInicializoBotones 0
                End If
            End If
        Case 2  'modificar
            subModificoDatos
            If Not errorEnIngresoDeDatos Then
                subInicializoGrilla     'limpio grilla
                gTabla.Row = 2          'muestro el primer control de ingreso
                gTabla_SelChange
                subInicializoBotones 2  'desabilito controles
                'determino tipo de clave de la tabla
                If propiedadTipoClave = 1 Then
                    'Si la clave es de tipo correlativo
                    'habilito boton de grabar
                    subInicializoBotones 0
                End If
            End If
        Case 3  'borrar
            'cargo propiedades del formulario
            frmBorrarRegistro.propFormControlIntegridad = propiedadIntegridad
            frmBorrarRegistro.propCaminoBase = propiedadCaminoBaseDeDatos
            Set frmBorrarRegistro.propRegistroEliminar = Data2
            'obtengo el valor a buscar
            If propiedadTipoClave = 0 Then
                frmBorrarRegistro.propValorAEliminar = _
                Data2.Recordset(propiedadIndiceCampoClave).Value
            Else
                frmBorrarRegistro.propValorAEliminar = _
                Data2.Recordset(propiedadIndiceCampoClaveContador).Value
            End If
            frmBorrarRegistro.Show 1
            If frmBorrarRegistro.propResultadoEliminacion = True Then
                RaiseEvent SeEliminoTabla(frmBorrarRegistro.propValorAEliminar)
            End If
            Unload frmBorrarRegistro
            subInicializoGrilla     'limpio grilla
            gTabla.Row = 2          'muestro el primer control de ingreso
            gTabla_SelChange
            subInicializoBotones 2  'desabilito controles
    End Select
Exit Sub
error:
    subControloErrores 515, "subDigitoBoton"
End Sub

Private Function funObtengoFilaClave() As Integer
    'Verifico que tipo de clave tiene asignada la propiedad del control
    'Si el tipo de clave es 1: verifico si estoy trabajando con el control correspondiente
    'recorro la grilla y obtengo el la fila en la que se encuantra el campo clave
    On Error GoTo error
    Dim Pcontrol As String
    Dim indiceCampoActual As String
    Dim i As Integer
    
    i = 2 'primera fila válida de la grilla
    funObtengoFilaClave = 0 'si no encuentro el campo clave o el tipo de clave
                            'es correlativo: debulevo 0
    
    'determino el tipo de clave utilizada por el control
    If propiedadTipoClave = 0 Then  '0= el usuario ingresa la clave
                                    '1= la clave se determina por número correlativo
                                    'obtengo todas las propiedades del control
        Do While i < gTabla.Rows
            'recorro grilla
            Pcontrol = UserControl.gTabla.TextMatrix(i, 1)
            'obtengo índice del campo actualmente seleccionado
            indiceCampoActual = mfunObtengoValorDesdeStr(Pcontrol, 2, ";")
            'verifico si estoy trabajando con el campo clave
            If propiedadIndiceCampoClave = indiceCampoActual Then
                'encontre la fila correspondiente al campo clave
                funObtengoFilaClave = i
            End If
            i = i + 1
        Loop
    End If
Exit Function
error:
    subControloErrores 515, "funObtengoFilaClave"
End Function

Private Function funObtengoProximoLibre() As Long
    'Recorre el archivo con el cual trabaja el ocx, ordenado por la clave principal
    'y obtiene el primer número libre comenzando desde 1 inclusive.
    'Este procedimiento solo es válido cuando la clave del archivo es de tipo numérico
    Dim consulta As String
    Dim contReg As Long
    On Error GoTo error
    funObtengoProximoLibre = 1  'si la tabla esta vacía comienzo desde 1
    'Obtengo nombre de los campos de la tabla a trabajar
    consulta = "Select * from " & propiedadTabla
    Data3.RecordSource = consulta
    Data3.Refresh
    'Obtengo todos los números claves oredenados en forma ascendente
    consulta = "select " & Data3.Recordset(propiedadIndiceCampoClave).Name & _
                " from " & propiedadTabla & _
                " order by 1"
    Data3.RecordSource = consulta
    Data3.Refresh
    'recorro el recordset y obtengo el primer número vacío
    contReg = 1
    If Data3.Recordset.RecordCount > 0 Then
        Data3.Recordset.MoveFirst
        Do While Not Data3.Recordset.EOF
            If Data3.Recordset(0).Value <> contReg Then
                Exit Do
            End If
            Data3.Recordset.MoveNext
            contReg = contReg + 1
        Loop
    End If
    funObtengoProximoLibre = contReg
Exit Function
error:
    subControloErrores 514, "funObtengoProximoLibre"
End Function

Private Sub subGraboDatos()
    'Grabo los datos que estan en la grilla, a la tabla correspondiente
    On Error GoTo error
    Dim consulta As String
    Dim correlativo As Boolean
    Dim proxCorr As Long
    'determino si estoy trabajando con un registro correlativo
    If propiedadTipoClave = 1 Then
        correlativo = True
    Else
        correlativo = False
    End If
    
    If correlativo Then
        'Obtengo correlativo
        proxCorr = funObtengoProximoCorr
        'inicializo data2
        consulta = "Select * from " & propiedadTabla
        Data2.RecordSource = consulta
        Data2.Refresh
        'creo un nuevo registro en el recordSet (tabla)
        Data2.Recordset.AddNew
            'cargo datos en el registro creado
            subCargoCorrelativo False, proxCorr
        'genero próximo número para el siguiente registro
        subGeneroProximoContador proxCorr
    Else
        'La clave esta determinada por el usuario
        'creo un nuevo registro en el recordSet (tabla)
        Data2.Recordset.AddNew
            'cargo datos en el registro creado
            subCargoDatos True 'true determina que cargo el campo clave
    End If
Exit Sub
error:
    subControloErrores 515, "subGraboDatos"
End Sub

Private Sub subModificoDatos()
    'Avtualizo la tabla con los datos que estan en la tabla
    Dim correlativo As Boolean
    On Error GoTo error
    'determino si estoy trabajando con un registro correlativo
    If propiedadTipoClave = 1 Then
        correlativo = True
    Else
        correlativo = False
    End If
    If correlativo Then
        'Si estoy modificando un registro correspondiente a una tabla de clave tipo correlativo
        'significa que se ejecutó el procedimiento MostrarRegistro, el cuál carga la grilla del control
        'con los datos correspondientes a un registro determinado, habilitando el botón de modificar.

        'Preparo registro actual para modificar
        Data2.Recordset.Edit
            'cargo nuevos datos en el registro
            subCargoCorrelativo True    'true significa que no tengo que modificar el valor clave del registro
    Else
        'La clave es determinada por el usuario
        
        'El registro actual ya esta determinado ya que si estoy en este punto
        'significa que se digitó el boton de modificar, el cuál solo se habilita tras
        'modificar el valor del campo clave, lo que produce que se busque ese nuevo valor
        'de la clave en la tabla correspondiente.
        
        'Preparo registro actual para modificar
        Data2.Recordset.Edit
            'cargo nuevos datos en el registro
            subCargoDatos False 'false determina que no cargo el campo clave
    
    End If
Exit Sub
error:
    subControloErrores 515, "subModificoDatos"
End Sub

Private Sub subCargoCorrelativo(Modificacion As Boolean, Optional proxCorr As Long)
    'Cargo los datos que estan en la grilla pero tomando en cuanta que la clave es un número
    'correlativo que se obtiene de una función
    On Error GoTo error
    Dim i As Integer
    Dim Pcontrol As String
    Dim indiceCampoActual As Integer
    Dim tipoCampo As Byte
    Dim claveArchivo As Variant 'usada para enviar como parámetros de los eventos
                                'SeGraboTabla y SeModificoTabla
    
    
    i = 2 'comienzo desde la primer fila de la cual se ingresan datos
    errorEnIngresoDeDatos = False
    Do While i < gTabla.Rows
        'obtengo todas las propiedades del control
        Pcontrol = UserControl.gTabla.TextMatrix(i, 1)
        'obtengo índice del campo actualmente seleccionado
        indiceCampoActual = mfunObtengoValorDesdeStr(Pcontrol, 2, ";")
        'obtengo tipo del campo
        tipoCampo = mfunObtengoValorDesdeStr(Pcontrol, 1, ";")
        'valido si los datos ingresados son correctos
        If funValidoDatos(Pcontrol, tipoCampo, i) Then
            If tipoCampo = 3 Or tipoCampo = 4 Then  'de tipo combos
                'asigno itemdata a la tabla
                Data2.Recordset(indiceCampoActual).Value = Val(gTabla.TextMatrix(i, 2))
            Else
                'asigno el valor de la columna visible a la tabla
                Data2.Recordset(indiceCampoActual).Value = gTabla.TextMatrix(i, 3)
            End If
        Else
            errorEnIngresoDeDatos = True
            Exit Do
        End If
        i = i + 1
    Loop
    'si estoy modificando un registro no cargo la clave
    If Not Modificacion Then
        If Not errorEnIngresoDeDatos Then
            'si se cargaron al registro todos los campos correctamente
            'cargo campo clave
            Data2.Recordset(propiedadIndiceCampoClave).Value = proxCorr
        End If
    End If
    'verifico si los datos se cargaron correctamente en el registro
    If Not errorEnIngresoDeDatos Then
        'tengo que guardar el valor de la clave sino luego de ejecutar el update
        'pierdo el mismo
        claveArchivo = Data2.Recordset(propiedadIndiceCampoClave).Value
    
        'si se cargaron correctamente grabo el registro
        Data2.Recordset.Update
        
        'Desencadeno evento que indica que se grabó o actualizó un registro.
        'Este en particular se desata cuando se graba un registro en una tabla
        'cuya clave es de tipo correlativo.
        'Le paso como parámetro la clave de la tabla
        If Not Modificacion Then
            RaiseEvent SeGraboTabla(claveArchivo)
        Else
            RaiseEvent SeModificoTabla(claveArchivo)
        End If
        'NOTA: debo de llamar al evento depues de actualizar la tabla
    Else
        'si hubo un error no grabo
        Data2.Recordset.CancelUpdate
    End If
Exit Sub
error:
    subControloErrores 516, "subCargoCorrelativo"
End Sub

Private Function funObtengoProximoCorr() As Long
    'Obtengo el próximo número libre de la tabla de correlativos
    On Error GoTo error
    Dim consulta As String
    'asumo que la tabla donde se encuentra el próximo número es la tabla de parámetros del
    'sistema y por lo tanto cuenta con un solo registro
    'cargo tabla en control data
    Data3.RecordSource = propiedadTablaContador
    Data3.Refresh
    'me posiciono en el primer registro
    Data3.Recordset.MoveFirst
    'obtengo el valor del campo contador
    funObtengoProximoCorr = Val(Data3.Recordset(propiedadIndiceCampoCont).Value)
    If funObtengoProximoCorr <= 0 Then
        'asumo por defecto el valor 1 para comenzar con el primer registro ingresado
        funObtengoProximoCorr = 1
    End If
Exit Function
error:
    subControloErrores 515, "funObtengoProximoCorr"
End Function

Private Sub subGeneroProximoContador(ultimoCorr As Long)
    'Aumento en 1 el valor del campo contador para que este disponible al ingresasr
    'el próximo nuevo registro
    On Error GoTo error
    'cargo tabla en control data
    Data3.RecordSource = propiedadTablaContador
    'me posiciono en el primer registro
    Data3.Recordset.MoveFirst
    If Data3.Recordset(propiedadIndiceCampoCont).Value = ultimoCorr Then
        'no se produjo ningúna alta en el interín
        Data3.Recordset.Edit
            Data3.Recordset(propiedadIndiceCampoCont).Value = ultimoCorr + 1
        Data3.Recordset.Update
    Else
        'en el interín se produjo otra alta
        ultimoCorr = Data3.Recordset(propiedadIndiceCampoCont).Value
        Data3.Recordset.Edit
            Data3.Recordset(propiedadIndiceCampoCont).Value = ultimoCorr + 1
        Data3.Recordset.Update
    End If
Exit Sub
error:
    subControloErrores 515, "subGeneroProximoContador"
End Sub

Private Sub subCargoDatos(cargoClave As Boolean)
    'Cargo los datos que estan en la grilla a la tabla
    'Dependiendo del valor de cargoCalve incluyo o no el campo clave
    On Error GoTo error
    Dim i As Integer
    Dim indiceCampoActual As Integer
    Dim Pcontrol As String
    Dim tipoCampo As Byte
    Dim tomoEnCuenta As Boolean
    Dim claveArchivo As Variant 'usada para enviar como parámetros de los eventos
                                'SeGraboTabla y SeModificoTabla
    
    i = 2 'comienzo desde la primer fila de la cual se ingresan datos
    errorEnIngresoDeDatos = False
    Do While i < gTabla.Rows
        'obtengo todas las propiedades del control
        Pcontrol = UserControl.gTabla.TextMatrix(i, 1)
        'obtengo índice del campo actualmente seleccionado
        indiceCampoActual = mfunObtengoValorDesdeStr(Pcontrol, 2, ";")
        'obtengo tipo del campo
        tipoCampo = mfunObtengoValorDesdeStr(Pcontrol, 1, ";")
        'valido si los datos ingresados son correctos
        If funValidoDatos(Pcontrol, tipoCampo, i) Then
            'determino si es campo clave
            If indiceCampoActual = propiedadIndiceCampoClave Then
                'cuando modifico un registro no es necesario modificar el campo clave
                tomoEnCuenta = cargoClave
            Else
                'si no es campo clave lo proceso siempre
                tomoEnCuenta = True
            End If
        
            If tomoEnCuenta Then
                If tipoCampo = 3 Or tipoCampo = 4 Then  'de tipo combos
                    'asigno itemdata a la tabla
                    Data2.Recordset(indiceCampoActual).Value = Val(gTabla.TextMatrix(i, 2))
                Else
                    'asigno el valor de la columna visible a la tabla
                    Data2.Recordset(indiceCampoActual).Value = gTabla.TextMatrix(i, 3)
                End If
            End If
            i = i + 1
        Else
            errorEnIngresoDeDatos = True
            Exit Do
        End If
    Loop
    'verifico si los datos se cargaron correctamente en el registro
    If Not errorEnIngresoDeDatos Then
        'tengo que guardar el valor de la clave sino luego de ejecutar el update
        'pierdo el mismo
        claveArchivo = Data2.Recordset(propiedadIndiceCampoClave).Value
        
        'si se cargaron correctamente grabo el registro
        Data2.Recordset.Update
        
        'Desencadeno evento que indica que se grabó o actualizó un registro.
        'Este en particular se desata cuando se graba un registro de una tabla
        'cuya clave la determina el usuario.
        'Le paso como parámetro la clave de la tabla
        If cargoClave Then
            RaiseEvent SeGraboTabla(claveArchivo)
        Else
            RaiseEvent SeModificoTabla(claveArchivo)
        End If
        'NOTA: debo de llamar al evento depues de actualizar la tabla
    Else
        'si hubo un error no grabo
        Data2.Recordset.CancelUpdate
    End If
Exit Sub
error:
    subControloErrores 516, "subCargoDatos"
End Sub

Private Function funValidoDatos(Pcontrol As String, tipoCampo As Byte, indice As Integer) As Boolean
    'Determino si el dato del campo a grabar(modificar) es correcto es decir cumple con las
    'pripiedades predeterminadas para el mismo
    On Error GoTo error
    Dim codErr As Byte
    Dim descErr As String
    Dim strMin As Integer
    Dim datoStrAValidar As String
    Dim datoNumAValidar As Double
    Dim datoFecha As Date
    Dim permitoNulo As Boolean
    Dim valorMin As Double
    Dim valorMax As Double
    Dim tipoValidacionFecha As Byte
    codErr = 0  'por defecto asumo que no hay errores
    funValidoDatos = True
    'verifico si permite ingresar datos en blanco
    permitoNulo = CBool(mfunObtengoValorDesdeStr(Pcontrol, 3, ";"))
    Select Case tipoCampo
        Case 0  'string
            'Para los datos de tipo string comparo tamaño de campo
            
            If Not permitoNulo Then 'si permito no controlo nada
                'Controlo mínimo caracteres (el máximo se controla por la propiedad maxlengh)
                'obtengo dato a validar
                datoStrAValidar = gTabla.TextMatrix(indice, 3)
                'obtengo propiedad mínimo de caracteres
                strMin = Val(mfunObtengoValorDesdeStr(Pcontrol, 4, ";"))
                'comparo largo
                If strMin > Len(datoStrAValidar) Then
                    'el largo no es el correcto
                    codErr = 1
                End If
            End If
            
        Case 1  'numérico
            'Para los datos de tipo númerico comparo valores de datos
            
            'valido ingreso de nulos
            If (Trim(gTabla.TextMatrix(indice, 3)) = Empty) And Not permitoNulo Then
                'no se permite ingresar nulos
                codErr = 4
            Else
                If IsNumeric(gTabla.TextMatrix(indice, 3)) Then
                    'obtengo dato a validar
                    datoNumAValidar = gTabla.TextMatrix(indice, 3)
                    'obtengo propiedad: valor mínimo
                    valorMin = mfunObtengoValorDesdeStr(Pcontrol, 4, ";")
                    'obtengo propiedad: valor máximo
                    valorMax = mfunObtengoValorDesdeStr(Pcontrol, 5, ";")
                    'controlo valor mínimo
                    If datoNumAValidar < valorMin Then
                        'el mínimo no es el correcto
                        codErr = 2
                    Else
                        If datoNumAValidar > valorMax Then
                            'el máximo no es el correcto
                            codErr = 3
                        End If
                    End If
                Else
                    codErr = 8
                End If
            End If
            
        Case 2  'fecha
            'valido ingreso de nulos
            If (Trim(gTabla.TextMatrix(indice, 3)) = Empty) And Not permitoNulo Then
                'no se permite ingresar nulos
                codErr = 4
            Else
                If IsDate(gTabla.TextMatrix(indice, 3)) Then    'si es una fecha
                    datoFecha = gTabla.TextMatrix(indice, 3)
                    'obtengo tipo de validación
                    tipoValidacionFecha = mfunObtengoValorDesdeStr(Pcontrol, 4, ";")
                    If tipoValidacionFecha = 2 Then 'menor igual a la fecha de hoy
                        If datoFecha <= Date Then
                            'esta ok
                        Else
                            codErr = 6
                        End If
                    Else
                        If tipoValidacionFecha = 3 Then 'mayor igual a la fecha de hoy
                            If datoFecha >= Date Then
                                'esta ok
                            Else
                               codErr = 7
                            End If
                        End If
                    End If
                Else
                    codErr = 5
                End If
            End If
        Case 3 To 4  'combo fijo y combo arch
            'valido ingreso de nulos
            If Trim(gTabla.TextMatrix(indice, 3)) = Empty Then
                'no se permite ingresar nulos
                codErr = 4
            End If
    End Select
    'verifico errores
    
    If codErr > 0 Then  'se produjo un error en el ingreso de datos
        Select Case codErr
            Case 1
                descErr = "La cantidad de caracteres ingresado es menor que el mínimo permitido."
            Case 2
                descErr = "El valor ingresado es menor al mínimo permitido."
            Case 3
                descErr = "El valor ingresado es mayor al máximo permitido."
            Case 4
                descErr = "No se permiten valores nulos."
            Case 5
                descErr = "El formato de la fecha no es correcto."
            Case 6
                descErr = "La fecha ingresada debe de ser menor igual a la fecha de hoy."
            Case 7
                descErr = "La fecha ingresada debe de ser mayor igual a la fecha de hoy."
            Case 8
                descErr = "El formato del número ingresado no es el correcto."
            
        End Select
        'desencadeno evento del ocx
        RaiseEvent ErrorEnIngreso(codErr, descErr)
        'posiciono control en el lugar donde se produjo el error
        gTabla.Row = indice
        gTabla_SelChange
        funValidoDatos = False
    End If
Exit Function
error:
    subControloErrores 515, "funValidoDatos"
End Function

'***********************************************************
'*
'* Métodos proporcionados al cliente
'*
'**********************************************************

Public Sub MostrarBoton(boton As Byte, mostrar As Boolean)
    'En alguno casos no será posibles agregar nuevos registros a un tabla determinada,
    'por lo que es necesario ocultar el boton de guardar.
    'Por otro lado también puede estar limitado el poder borrar registros.
    'Para éstos casos se implementa este método que oculta el boton especificado en al argumento
    'boton.
    On Error GoTo error
    If boton = 1 Or boton = 2 Or boton = 3 Or boton = 6 Or boton = 7 Then
        UserControl.toolMenu.Buttons(boton).Visible = mostrar
    End If
Exit Sub
error:
    subControloErrores 515, "MostrarBoton"
End Sub

Public Sub MostrarRegistro(clave As Variant)
    'Muestro un registro correspondiente en el control
    'si existe muestro los datos
    On Error GoTo error
    'si no paso nada como parámetro entonces no ejecuto nada
    If Trim(clave) <> Empty Then
        Data2.RecordSource = "Select * from " & propiedadTabla
        If funBuscoRegistro(clave) Then
            subMuestroDatosEnGrilla
            subInicializoBotones 1  'modificar
            'forzando este evento logro que el control de ingreso de datos se carge
            'con el valor cargado en la tabla
            gTabla_SelChange
        End If
    End If
Exit Sub
error:
    subControloErrores 515, "MuestroRegistro"
End Sub

Public Sub IniciarMantenimiento()
    'Prepara el control para su funcionamiento
    'creando las filas en la grilla, correspondientes a los campos de la tabla,
    'almacenando las propiedades de los mismos en columnas ocultas
    'Estos datos son impresindibles para que en tiempo de ejecución el control funcione.
    On Error GoTo error
    Dim anchoAux As Long
    
    habilitoSelChange = False   'no permito ejecutar SelChanged por ahora
    '1) VALIDO COHERENCIA DE PROPIEDADES ASIGNADAS EN TIEMPO DE EJECUCIÓN
    If funValidoCoherenciaPropiedades Then
        '2) INICIALIZO GRILLA
        'Elimino todas las filas de la grilla , el False indica que no elimino las columnas
        mSubLimpioGrilla UserControl.gTabla, False
        'oculto todos los controles de ingreso de datos
        UserControl.txtIngFecha.Visible = False
        UserControl.txtIngNum.Visible = False
        UserControl.txtIngStr.Visible = False
        UserControl.txtIngstrMemo.Visible = False
        UserControl.cboIngOpcA.Visible = False
        UserControl.cboIngOpcF.Visible = False
        '3) DIBUJO CAMPOS DE LA GRILLA
        'dibujo campos en la grilla
        'con la información establecida para la propiedad campo creo los campos del control
        subCreoCampos propiedadCampo
        '4) ESTABLESCO NUEVO TAMAÑO DE LA GRILLA DEPENDIENDO SI SE MUESTRAN O NO LAS BARRAS DE DESPLAZAMIENTO
        'Redibujo control si muestro barras de desplazamineto
        'verifico si estoy mostrando más linas que las que el control puede
        'mostrar sin tener que dibujar la barra de desplazamiento
        'todas las filas menos la primera y la oculta tienen el mismo ancho
        
        'El tamaño de la columna de ingreso de datos se adapta al tamaño de la grilla
        'Esta línea de código es impresindible que se cargue en este procedimiento y no en el
        'evento Resize, ya que el ancho de la columna de datos puede variar al ejecutar el método
        'IniciarMantenimiento.
        UserControl.gTabla.ColWidth(3) = (UserControl.gTabla.Width - UserControl.gTabla.ColWidth(0)) - 100
    
        anchoAux = ((gTabla.Rows - 2) * propAnchoCeldas) _
                    + gTabla.RowHeight(0)   'sumo el ancho de la primera columna
        If anchoAux > gTabla.Height Then
            'se muestra grilla de desplazamiento por lo que tengo
            'que reestablecer el largo de la fila de datos, restándole 255
            'que es el ancho de la barra de desplazamiento
            UserControl.gTabla.ColWidth(3) = UserControl.gTabla.ColWidth(3) - 255
        End If
        
        '5) INICIALIZO BOTONES
        'por defecto muestro todos los botones de operaciones
        UserControl.toolMenu.Buttons(1).Visible = True 'guardar
        UserControl.toolMenu.Buttons(2).Visible = True 'mod.
        UserControl.toolMenu.Buttons(3).Visible = True 'eli.
        UserControl.toolMenu.Buttons(7).Visible = True 'limpiar
        
        'por defecto no muestro botón de sugerir próximo
        UserControl.toolMenu.Buttons(6).Visible = False
        
        'Obtengo fila del campo clave
        filaCampoClave = funObtengoFilaClave
        If (propiedadTipoClave = 0) And filaCampoClave > 0 Then 'si la clave es de tipo determinado por el usuario
                                                                'y se especificó clave
            'verifico si la clave es de tipo númerico
            If Val(mfunObtengoValorDesdeStr(gTabla.TextMatrix(filaCampoClave, 1), 1, ";")) = 1 Then
                'verifico si se establecio en propiedades sugerir próximo
                If CBool(propiedadSugerirProxLibre) Then
                    'habilito botón de sugerir próximo
                    UserControl.toolMenu.Buttons(6).Visible = True
                End If
            End If
        End If
        
        '6) INICIALIZO APARIENCIA DE GRILLA
        'Cambio ancho de las filas y tamaño de la fuente de la primera columna
        subCambioAnchoFilas propAnchoCeldas
        'centro los datos que se muestran en las celdas hacia la izquierda
        mSubRangoCeldas gTabla, 3, 3, 2, gTabla.Rows - 1
        gTabla.CellAlignment = flexAlignLeftCenter
    
        '7) INICIALIZO BARRA DE TAREAS
        'determino si el control trabaja con una tabla de correlativos
        If propiedadTipoClave = 1 Then  'tipo clave correlativo
            'habilito el boton de guardar
            subInicializoBotones 0
        Else
            'inicializo botones
            subInicializoBotones 2
        End If
        
        '8) PERMITO TRABAJAR CON EL CONTROL
        'habilito barra
        UserControl.toolMenu.Enabled = True
        habilitoSelChange = True   'permito ejecutar SelChanged
        UserControl.Refresh
        'gTabla.SetFocus
        'Esta línea produce error, cuando ejecuto la aplicación cliente fuera de Visual
        'es decir desde un exe.
    End If
Exit Sub
error:
subControloErrores 515, "IniciarMantenimiento"
End Sub

Public Sub GragarRegistro()
    'Es lo mismo que hacer click sobre el boton de grabar
    'Puede ser utilizado para trabajar con tecla de función en la aplicación cliente
    
    'verifico si es posible desencadenar el evento click del boton
    If UserControl.toolMenu.Buttons(1).Enabled = True _
        And UserControl.toolMenu.Buttons(1).Visible = True _
        And UserControl.Enabled = True Then
        subDigitoBoton 1
    End If
End Sub

Public Sub SugerirProximo()
    'Es lo mismo que hacer click sobre el boton de sugerir próximo
    'Puede ser utilizado para trabajar con tecla de función en la aplicación cliente
    
    'verifico si es posible desencadenar el evento click del boton
    If UserControl.toolMenu.Buttons(6).Enabled = True _
        And UserControl.toolMenu.Buttons(6).Visible = True _
        And UserControl.Enabled = True Then
            subDigitoBoton 6
    End If
End Sub

Public Sub BorrarRegistro()
    'Es lo mismo que hacer click sobre el boton de borrar
    'Puede ser utilizado para trabajar con tecla de función en la aplicación cliente
    
    'verifico si es posible desencadenar el evento click del boton
    If UserControl.toolMenu.Buttons(3).Enabled = True _
        And UserControl.toolMenu.Buttons(3).Visible = True _
        And UserControl.Enabled = True Then
        subDigitoBoton 3
    End If
End Sub
    
Public Sub ModificarRegistro()
    'Es lo mismo que hacer click sobre el boton de modificar
    'Puede ser utilizado para trabajar con tecla de función en la aplicación cliente
    
    'verifico si es posible desencadenar el evento click del boton
    If UserControl.toolMenu.Buttons(2).Enabled = True _
    And UserControl.toolMenu.Buttons(2).Visible = True _
    And UserControl.Enabled = True Then
        subDigitoBoton 2
    End If
End Sub

Public Sub LimpioRegistro()
    'Es lo mismo que hacer click sobre el boton de limpiar
    'Puede ser utilizado para trabajar con tecla de función en la aplicación cliente
    
    'verifico si es posible desencadenar el evento click del boton
    If UserControl.toolMenu.Buttons(7).Enabled = True _
        And UserControl.toolMenu.Buttons(7).Visible = True _
        And UserControl.Enabled = True Then
            subDigitoBoton 7
    End If
End Sub

Public Sub MuestroSeñalDeFocus(muestro As Boolean)
    'Cuando le doy el focus al control muestro un ícono que le indica al usuario esto.
    On Error GoTo error
    gTabla.Col = 3
    gTabla.Row = 0
    gTabla.CellPictureAlignment = 7
    If muestro Then
        'muestro ícono indicando que el control tiene el fócus
        Set gTabla.CellPicture = ImageList1.ListImages(7).Picture
    Else
        'oculto ícono
        Set gTabla.CellPicture = ImageList1.ListImages(8).Picture
    End If
Exit Sub
error:
    subControloErrores 515, "MuestroSeñalDeFocus"
End Sub
'***********************************************************
'*
'* Fin métodos proporcionados
'*
'**********************************************************

Private Function funValidoCoherenciaPropiedades() As Boolean
    'Valido que las propiedades del control seán choerentes entre sí.
    'Las propiedades de un control se pueden establecer de diferentes formas:
    'Desde la página de propiedades personalizada
        'Al establecer las diferentes propiedades del control desde la página de
        'propiedades, las mismas se validadan desde el punto de vista de la choerencia
        'que tengan dichas propiedades entre sí, no permitiendo guardar valores que no sean
        'validos.
        'También se valida que el tipo de dato ingresado como propiedad coincida con el tipo
        'de dato definido para la propiedad.
    'Desde la página de propiedades estandar de visual
        'Desde aquí solo tengo acceso a las propiedades que no requieren de chequeo de choerencia
        'Visual Basic se encarga de validar que los tipos de datos ingresados para cada propiedad
        'coincida con el tipo de dato con el cual esta definida la propiedad.
        'Es necesario implementar en el procedimiento Property Get, procedimiento que
        'controlen el rango de los valores establecidos, por ejemplo una propiedad de tipo
        'numerico que solo acepte los valores 0 y 1. (ver *)
    'Desde el código del programa que utiliza el ocx
        'La tercer opción para determinar las propiedades de un control, es la que puede originar
        'error al asignar valores a las propiedades, que hagan cancelar la aplicación.
        'Estos errores pueden ser originados al asignar tipos de datos distintos a los establecidos
        'para las propiedades, o pueden ser originados al establecer datos en las propiedades,
        'que aunque si sean del tipo correcto, no son choerentes con otras propiedades.
        'Este último tipo de error producirá que el ocx, no funcione correctamente.
        'Para evitar este tipo de errores es necesario implementar este procedimiento que
        'que se encarga de validar las propiedades establecidas para el control.
        
    '(*)Para evitar validar dos veces los valores de las propiedades( aquí y en los procedimientos de
    'property let) solo se realizan los controles de rango de valores desde este procedimiento
    
    'Si ocurre un error sea del origen que sea, se envía un error a la aplicación cliente.

    'DESCRIPCIÓN DE LOS CONTROLES A REALIZAR A LAS PRPIEDADES
    
    'NOTA: programar todos los posibles errores que puedan tener las propiedades requiere un
    'tiempo que para esta primera versión el ocx, no es redituable implemnetar.
    
    funValidoCoherenciaPropiedades = True
End Function

'*****************************************************
'*
'*  Control de errores
'*
'******************************************************

Private Sub subControloErrores(numErr As Integer, desde As String)
    'Al recivir un error indico en el ocx que se produjo el mismo
    Dim msgErr As String
    Dim descAux As String
    Dim errDesc
    Select Case numErr
        Case 512
            'error en carga de combo de archivos
            descAux = " Error en carga de combo de tipo archivo"
        Case 513
            'error en función buscar registro
            descAux = " Error al buscar registro en tabla"
        Case 514
            'error en función que busca próximo registro
            descAux = " Error al buscar próximo registro"
        Case 515
            'error de programa
            descAux = " Error desconocido"
        Case 516
            descAux = " Posiblemente está tratando de cargar un campo de la tabla, con un valor de tipo diferente " & _
                    "al tipo del campo."
        Case Else
            
    End Select
    errDesc = Err.Number & " " & Err.Description & Chr(10) & _
                numErr & descAux & Chr(10) & _
                desde & Chr(10) & _
                "Consulte con su proveedor de sowftware"
    'muestro error en grilla
    'UserControl.Frame1.Visible = True
    UserControl.lblError.Caption = errDesc
    'Al presentarse un error aparece un cuadro de díalogo sobre el control indicando
    'que se ha producido un error en el mismo. Como no es bueno modificar la
    'interface de la aplicación con los mensajes de error de los componentes
    'se trata de que éste mensaje no afecte demasiado la interfaz, pero sí que brinde información
    'al usuario o al programador de la aplicación que usa este componenete.
End Sub



