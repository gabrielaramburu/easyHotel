VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.UserControl SeleccionBD 
   ClientHeight    =   3540
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7485
   PropertyPages   =   "SeleccionBD.ctx":0000
   ScaleHeight     =   3540
   ScaleWidth      =   7485
   ToolboxBitmap   =   "SeleccionBD.ctx":0017
   Begin ComctlLib.ProgressBar pbProgreso 
      Height          =   135
      Left            =   0
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   238
      _Version        =   327682
      Appearance      =   1
   End
   Begin MSFlexGridLib.MSFlexGrid gInfCampos 
      Height          =   855
      Left            =   1440
      TabIndex        =   1
      Top             =   2280
      Visible         =   0   'False
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   1508
      _Version        =   393216
      Rows            =   3
      FixedCols       =   0
      Enabled         =   -1  'True
   End
   Begin VB.Frame Frame1 
      Caption         =   "Error en ocx."
      Height          =   1695
      Left            =   480
      TabIndex        =   3
      Top             =   1080
      Visible         =   0   'False
      Width           =   6135
      Begin VB.TextBox lblError 
         BackColor       =   &H80000000&
         Height          =   1215
         Left            =   720
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Text            =   "SeleccionBD.ctx":0329
         Top             =   240
         Width           =   5055
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "SeleccionBD.ctx":032F
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.TextBox txtCriterio 
      Enabled         =   0   'False
      Height          =   315
      Left            =   0
      MaxLength       =   255
      TabIndex        =   5
      Top             =   320
      Width           =   6975
   End
   Begin VB.CommandButton botCambioCriterio 
      Enabled         =   0   'False
      Height          =   315
      Left            =   7050
      Picture         =   "SeleccionBD.ctx":0771
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Cambia el criterio de seleción"
      Top             =   320
      Width           =   375
   End
   Begin VB.ListBox lwCriterios 
      Height          =   2700
      Left            =   6000
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   ""
      Top             =   2760
      Visible         =   0   'False
      Width           =   2295
   End
   Begin MSDBGrid.DBGrid gSeleccion 
      Bindings        =   "SeleccionBD.ctx":0A47
      Height          =   2775
      Left            =   0
      OleObjectBlob   =   "SeleccionBD.ctx":0A57
      TabIndex        =   0
      Top             =   720
      Width           =   7455
   End
   Begin VB.Label lblCriterio 
      AutoSize        =   -1  'True
      Caption         =   "Ordenado por:"
      Height          =   195
      Left            =   0
      TabIndex        =   6
      Top             =   80
      Width           =   1020
   End
End
Attribute VB_Name = "SeleccionBD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Declaración de constantes
Private Const cAnchoMinimoControl As Integer = 2000
Private Const cLargoMinimoControl As Integer = 2000

'Declaración de constantes de propiedades
Private Const cTabla As String = ""
Private Const cBaseDeDatos As String = ""
Private Const cContraseñaBaseDeDatos As String = ""
Private Const cCampos As String = ""
Private Const cTablasRelacionadas As String = ""
Private Const cNroCampoInicial As Integer = 0
Private Const cIndiceCampoRetorno As Integer = 0
Private Const cTeclaSeleccion As Integer = 13 'por defecto es la tecla Enter
Private Const cSeleccionComplementaria As String = ""


'Declaración de variables de propiedad
Private propTabla As String
Private propBaseDeDatos As String
Private propContraseñaBaseDeDatos As String
Private propCampos As String
Private propNroCampoInicial As Integer
Private propTablasRelacionadas As String
Private propIndiceCampoRetorno As Integer
Private propTeclaSeleccion As Integer
Private propSeleccionComplementaria As String   'no aparece en las propiedades personalizadas
                                                'La finalidad es poder realizar consultas más complejas
                                                'incluyendo en esta propiedad, nuevos criterios de selección.

'Property Variables:
Dim m_GrillaFont As Font

'Declaración de variables generales

Private gCampoSel As Integer            'contiene el número de columna seleccionada en la grilla
                                        'de selección
                                        
Private gConsulta As String             'almacena la consulta principal (sin criterios de selección)
                                        'Se utiliza una variable global porque se carga al ejecutar el método Mostrar
                                        'y se utiliza en el evento changed de txtCriterio
                                        
Private gContieneJoin As Boolean        'para poder implementar la búsqueda en el evento txtCriterio_change
                                        'tengo que sabre si la consulta es simple o realiza joins.
                                        'Dependiendo del caso se incluirá en la misma la palabra AND o WHERE.
'Declaración de eventos
Public Event Seleccionar(ValorClaveTablaPrincipal As Variant)
                                        'Se ejecuta al hacer doble click sobre la gilla o al digitar
                                        'la tecla predeterminada de selección. El objetivo de este evento es
                                        'informar a la aplicación cliente que se seleccionó una fila de la grilla,
                                        'devolviendo el valor correspondiente al campo clave de la tabla principal,
                                        'del registro seleccionado (fila seleccionada).

Private Sub botCambioCriterio_Click()
    On Error GoTo error
    lwCriterios.Visible = True
    lwCriterios.SetFocus
    lwCriterios.ListIndex = 0   'ilumino el primer criterio de la lista
Exit Sub
error:
 subControloErrores 515, "botCambioCriterio"
End Sub

Private Sub UserControl_InitProperties()
    'Se ejecuta cuando se coloca un nuevo control en el contenedor
    'Cargo propiedades del control con los valores predefinidos
    On Error Resume Next
    propTabla = cTabla
    propBaseDeDatos = cBaseDeDatos
    propContraseñaBaseDeDatos = cContraseñaBaseDeDatos
    propCampos = cCampos
    propNroCampoInicial = cNroCampoInicial
    Set m_GrillaFont = Ambient.Font
    propIndiceCampoRetorno = cIndiceCampoRetorno
    propTablasRelacionadas = cTablasRelacionadas
    propTeclaSeleccion = cTeclaSeleccion
    propSeleccionComplementaria = cSeleccionComplementaria
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    'Se produce al crearse una nueva instancia en diseño o ejecución
    'del control
    
    tabla = PropBag.ReadProperty("Tabla", cTabla)
    BaseDeDatos = PropBag.ReadProperty("BaseDeDatos", cBaseDeDatos)
    ContraseñaBaseDeDatos = PropBag.ReadProperty("ContraseñaBaseDeDatos", cContraseñaBaseDeDatos)
    campos = PropBag.ReadProperty("Campos", cCampos)
    NroCampoInicial = PropBag.ReadProperty("NroCampoInicial", cNroCampoInicial)
    SeleccionComplementaria = PropBag.ReadProperty("SeleccionComplementaria", cSeleccionComplementaria)
    'propiedades de apariencia
    GrillaBackColor = PropBag.ReadProperty("GrillaBackColor", Ambient.BackColor)
    GrillaForeColor = PropBag.ReadProperty("GrillaForeColor", Ambient.ForeColor)
    Set m_GrillaFont = PropBag.ReadProperty("GrillaFont", Ambient.Font)
    TablasRelacionadas = PropBag.ReadProperty("TablasRelacionadas", cTablasRelacionadas)
    IndiceCampoRetorno = PropBag.ReadProperty("IndiceCampoRetorno", cIndiceCampoRetorno)
    TeclaSeleccion = PropBag.ReadProperty("TeclaSeleccion", cTeclaSeleccion)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    On Error Resume Next
    'Se graba al destruirse una instancia en tiempo de diseño
    
    PropBag.WriteProperty "Tabla", propTabla, cTabla
    PropBag.WriteProperty "BaseDeDatos", propBaseDeDatos, cBaseDeDatos
    PropBag.WriteProperty "ContraseñaBaseDeDatos", propContraseñaBaseDeDatos, cContraseñaBaseDeDatos
    PropBag.WriteProperty "Campos", propCampos, cCampos
    PropBag.WriteProperty "NroCampoInicial", propNroCampoInicial, cNroCampoInicial
    PropBag.WriteProperty "SeleccionComplementaria", propSeleccionComplementaria, cSeleccionComplementaria
    'propiedades de apariencia
    PropBag.WriteProperty "GrillaForeColor", gSeleccion.ForeColor, Ambient.ForeColor
    PropBag.WriteProperty "GrillaBackColor", gSeleccion.BackColor, Ambient.BackColor
    Call PropBag.WriteProperty("GrillaFont", m_GrillaFont, Ambient.Font)
    PropBag.WriteProperty "IndiceCampoRetorno", propIndiceCampoRetorno, cIndiceCampoRetorno
    PropBag.WriteProperty "TablasRelacionadas", propTablasRelacionadas, cTablasRelacionadas
    PropBag.WriteProperty "TeclaSeleccion", propTeclaSeleccion, cTeclaSeleccion
End Sub

'Declaración de procedimientos de propiedades
'Propiedad tabla
Public Property Get tabla() As String
    tabla = propTabla
End Property

Public Property Let tabla(ByVal nuevoValor As String)
    propTabla = nuevoValor
    PropertyChanged "Tabla"
End Property

'Propiedad BaseDeDatos
Public Property Get BaseDeDatos() As String
    BaseDeDatos = propBaseDeDatos
End Property

Public Property Let BaseDeDatos(ByVal nuevoValor As String)
    'Inicializo control data
    propBaseDeDatos = nuevoValor
    Data1.DatabaseName = propBaseDeDatos
    PropertyChanged "BaseDeDatos"
End Property

'Propiedad ContraseñaBaseDeDatos
Public Property Get ContraseñaBaseDeDatos() As String
    ContraseñaBaseDeDatos = propContraseñaBaseDeDatos
End Property

Public Property Let ContraseñaBaseDeDatos(ByVal nuevoValor As String)
    propContraseñaBaseDeDatos = nuevoValor
    Data1.Connect = propContraseñaBaseDeDatos
    PropertyChanged "ContraseñaBaseDeDatos"
End Property

'Propiedad Campos
Public Property Get campos() As String
Attribute campos.VB_MemberFlags = "400"
    campos = propCampos
End Property

Public Property Let campos(ByVal nuevoValor As String)
    propCampos = nuevoValor
    PropertyChanged "Campos"
End Property

'Propiedad NroCampoInicial
Public Property Get NroCampoInicial() As Integer
    NroCampoInicial = propNroCampoInicial
End Property

Public Property Let NroCampoInicial(ByVal nuevoValor As Integer)
    propNroCampoInicial = nuevoValor
    PropertyChanged "NroCampoInicial"
End Property

'Propiedad indiceCampoRetorno
Public Property Get IndiceCampoRetorno() As Integer
    IndiceCampoRetorno = propIndiceCampoRetorno
End Property

Public Property Let IndiceCampoRetorno(ByVal nuevoIndice As Integer)
    propIndiceCampoRetorno = nuevoIndice
    PropertyChanged "IndiceCampoRetorno"
End Property

'Propiedad TablasRelacionadas
Public Property Get TablasRelacionadas() As String
Attribute TablasRelacionadas.VB_MemberFlags = "400"
    TablasRelacionadas = propTablasRelacionadas
End Property

Public Property Let TablasRelacionadas(ByVal nuevoValor As String)
    propTablasRelacionadas = nuevoValor
    PropertyChanged "TablasRelacionadas"
End Property

'Propiedad SeleccionComplementaria
Public Property Get SeleccionComplementaria() As String
    SeleccionComplementaria = propSeleccionComplementaria
End Property

Public Property Let SeleccionComplementaria(ByVal nuevoValor As String)
    propSeleccionComplementaria = nuevoValor
End Property

'Declaración de propiedades de apariencia por delegación
'Propiedad BackColor
Public Property Get GrillaBackColor() As OLE_COLOR
    GrillaBackColor = gSeleccion.BackColor
End Property

Public Property Let GrillaBackColor(ByVal nuevoColor As OLE_COLOR)
    gSeleccion.BackColor = nuevoColor
    PropertyChanged "GrillaBackColor"
End Property

'Propiedad ForeColor
Public Property Get GrillaForeColor() As OLE_COLOR
    GrillaForeColor = gSeleccion.ForeColor
End Property

Public Property Let GrillaForeColor(ByVal nuevoColor As OLE_COLOR)
    gSeleccion.ForeColor = nuevoColor
    PropertyChanged "GrillaForeColor"
End Property

'Propiedad font
Public Property Get GrillaFont() As Font
    Set GrillaFont = m_GrillaFont
End Property

Public Property Set GrillaFont(ByVal New_GrillaFont As Font)
    Set m_GrillaFont = New_GrillaFont
    gSeleccion.Font = m_GrillaFont
    PropertyChanged "GrillaFont"
End Property

'Propiedad TeclaSeleccion
Public Property Get TeclaSeleccion() As Integer
    TeclaSeleccion = propTeclaSeleccion
End Property

Public Property Let TeclaSeleccion(ByVal nuevoValor As Integer)
    propTeclaSeleccion = nuevoValor
    PropertyChanged "TeclaSeleccion"
End Property

Private Sub UserControl_Resize()
    'Adapto el tamaño de los controles que contiene el ocx
    'con relación al tamaño del mismo.
    On Error GoTo error
    'controlo tamaño mínimo del control
    If UserControl.Height < cAnchoMinimoControl Then
        UserControl.Height = cAnchoMinimoControl
    Else
        If UserControl.Width < cLargoMinimoControl Then
            UserControl.Width = cLargoMinimoControl
        End If
    End If
    'adapto largo grilla
    gSeleccion.Width = UserControl.Width
    'adapto largo textBox criterio
    txtCriterio.Width = UserControl.Width - 450 '450 es el espacio para el boton de cambio criterio
    'adapto largo barrra progreso
    pbProgreso.Width = UserControl.txtCriterio.Width
    'adapto ancho de grilla
    'el ancho de la grilla esta determinado por el ancho del control - 750 tips
    'que se dejan libre para la presentación del textbox y etiqueta del criterio de selección.
    gSeleccion.Height = UserControl.Height - 750
    gSeleccion.Top = 750
    botCambioCriterio.Left = gSeleccion.Width - botCambioCriterio.Width
    'cambio tamaño del control List utilizado para cambiar criterio de ordenación
    lwCriterios.Height = gSeleccion.Height
    lwCriterios.Left = gSeleccion.Width - lwCriterios.Width
    lwCriterios.Top = gSeleccion.Top
Exit Sub
error:
 subControloErrores 515, "UserControl Resize"
End Sub

Private Sub txtCriterio_Change()
    'Cada vez que el usuario modifica el contenido de este control tengo que realizar
    'una búsqueda utilizando como criterio el valor de este control, el cual lo comparo
    'con el campo de la consulta que esté actualmente seleccionado.
    'Este campo se selecciona mediante la propiedad NroCampoInicial, o es modificado por
    'el usuario al hacer click sobre alguna de las columnas de la grilla.
    On Error GoTo error
    Dim nuevaCond As String
    Dim colActual As Integer
    Dim clausulaVariable As String
    'La parte de la consulta SQL correspondiente al Where, es la más compleja.
    'Esta parte se puede dividir en otras tres partes:
    '   1: parte del join, la cual encontramos en consultas que acceden a más de una tabla
    '   2: parte de selección complementaria, la cual encontramos en consultas que deben
    '      de aplicar uno o varios criterios de selección
    '   3: parte de selección de registros, según criterio establecido por usuario.
    'Las partes 1 y 2 pueden no existir.
    'Lo importante a saber, es que dependiendo de las partes 1 y 2, dependerá
    'como se forme la parte 3, que siempre existirá.
    'Las siguientes líneas de código determinan como quedará conformada la tercer parte
    'en función de las partes 1 y 2.
    
    
    'verifico si hay parte 1
    If gContieneJoin Then
        clausulaVariable = " and "
    Else
        If Trim(propSeleccionComplementaria) = Empty Then
            clausulaVariable = " where "
        Else
            clausulaVariable = " and "
        End If
    End If

    'agrego a la consulta original una nueva condición de búsqueda y ordenación
    nuevaCond = gInfCampos.TextMatrix(1, gCampoSel) & "." & gInfCampos.TextMatrix(0, gCampoSel) & _
                " LIKE '" & txtCriterio.Text & "*'" & " order by " & gInfCampos.TextMatrix(1, gCampoSel) & "." & gInfCampos.TextMatrix(0, gCampoSel)
                
    'inicializo nuevamente el data para realizar consulta
    Data1.RecordSource = gConsulta & clausulaVariable & nuevaCond
    
    colActual = gSeleccion.LeftCol
    Data1.Refresh                   'realizo consulta
    subCambioAnchoColumnas          'reconfiguro tamaño de columnas
    subPosColumnaTrabajo colActual  'posiciono nuevamente en columna actual
Exit Sub
error:
 subControloErrores 515, "txtCriterio_Changed"
End Sub

Private Sub gSeleccion_KeyDown(KeyCode As Integer, Shift As Integer)
    'txtCriterio.SetFocus
    txtCriterio_KeyPress (KeyCode)
End Sub

Private Sub gSeleccion_HeadClick(ByVal ColIndex As Integer)
    'Cuando el usuario hace un click sobre el cabezal de la grilla en una columna determinada
    'cargo variable global con el índice de la columna seleccionada
    Dim c As Column
    Dim colActual As Integer
    On Error GoTo error
    'determino el tipo de campo a ordenar
    'si el campo es de tipo memo no ordeno ya que no se puede ordenar una consulta
    'por un campo de tipo memo.
    If Data1.Recordset.Fields(ColIndex).Type <> dbMemo Then
        gCampoSel = ColIndex
        'cambio etiqueta de criterio
        Set c = gSeleccion.Columns(ColIndex)
        lblCriterio.Caption = "Ordenado por : " & c.Caption
        Set c = Nothing
        'ordeno la grilla por el nuevo campo seleccionado
        'NOTA: siempre que trabajo con un nombre de un campo tengo que indicar a
        'que tabla pertenece, para que no se produzca un error en caso de existir dos campos
        'con el mismo nombre en las tablas que estoy trabajando.
        Data1.RecordSource = gConsulta & " order by " & gInfCampos.TextMatrix(1, gCampoSel) & "." & gInfCampos.TextMatrix(0, gCampoSel)
        colActual = gSeleccion.LeftCol
        Data1.Refresh
        subCambioAnchoColumnas          'reconfiguro tamaño de columnas
        subPosColumnaTrabajo colActual  'posiciono en columna actual de trabajo
        'le doy el focus al control de ingreso de criterio para mejorar interface de usuario
        txtCriterio.SetFocus
    End If
    
Exit Sub
error:
 subControloErrores 517, "gSeleccion_HeadClick"
End Sub

Private Sub subPosColumnaTrabajo(col As Integer)
    'Al ejecutar nuevamente una consulta, es decir al ejecutar el método Refresh
    'del control data, la grilla se inicializa, por lo que la primer columna vidible
    'de la grilla pasa a ser la columna 0. Esto puede ser molesto para el caso
    'de que el usuario trabaje con alguna columna que no este visible por falta de espacio en la grilla.
    'Este procedimiento es llamado después de cada refresh y tiene como objetivo mantener
    'simpre como primer columna visible la que el usuario determine.
    gSeleccion.LeftCol = col
End Sub

Private Sub subCambioAnchoColumnas()
    'Cambio el ancho de las columnas de la grilla, al ancho determinado por el usuario
    'en las propiedades del control
    'Este procedimiento es llamado después de ejecutar un data1.refresh
    On Error GoTo error
    Dim c As Column
    Dim i As Integer
    Dim totCol As Integer
    i = 0   'comienzo en la primer columna
    totCol = gSeleccion.Columns.Count
    'recorro todas las columnas de la grilla de selección
    Do While i < totCol
        Set c = gSeleccion.Columns(i)
        c.Width = Val(gInfCampos.TextMatrix(2, i))  'establesco el ancho de la columna
        i = i + 1
    Loop
    Set c = Nothing
Exit Sub
error:
 subControloErrores 515, "subCambioAnchoColumnas"
End Sub

Private Sub txtCriterio_KeyPress(KeyAscii As Integer)
    'Valido que no se ingresen caracteres prohibidos
    'Como el valor de este control se utiliza para formar una consulta
    'SQL es necesario controlar que no contenga caracteres  que hagan cancelar la misma.
    On Error GoTo error
    Select Case KeyAscii
        Case 39 'caracter de '
            KeyAscii = 0
        Case 124 'caracter de |
            KeyAscii = 0
        Case propTeclaSeleccion
            'La tecla de seleccion depende del valor de la variable
            'digitar esta tecla,es lo mismo que hacer dobleclik sobre la grilla
            gSeleccion_DblClick
    End Select
Exit Sub
error:
 subControloErrores 515, "txtCriterio_KeyPress"
End Sub

Private Sub txtCriterio_KeyDown(KeyCode As Integer, Shift As Integer)
    'Al presionar las teclas de subir y bajar en este control cambio la fila seleccionada
    'tal cual si estuviera posicionado (con el focus) sobre la grilla.
    On Error GoTo error
    Select Case KeyCode
        Case 38 'tecla subir
            'valido que no se balla de rango
            If gSeleccion.Row > 0 Then
                gSeleccion.Row = gSeleccion.Row - 1
            End If
        Case 40 'tecla bajar
            'valido que no se balla de rango
            If gSeleccion.Row < gSeleccion.VisibleRows - 1 Then
                gSeleccion.Row = gSeleccion.Row + 1
            End If
    End Select
Exit Sub
error:
 subControloErrores 515, "txtCriterio_KeyDown"
End Sub

Private Sub gSeleccion_DblClick()
    'Al hacer dobleclick sobre un registro de la grilla se
    'ejecuta el evento seleccionar del control. También al presionar la tecla predeterminada.
    'Es necesario: a) saber que campo de la tabla voy a devolver (asumo que este campo que se indica
    'en la propiedad IndiceCampoRetorno, es el campo clave de la tabla principal)
    'b) después de tener el campo, debo de asegurarme que se devuleva el valor correspondiente
    'al registro que corresponde con la fila seleccionada en la grilla, esto no es dificil
    'ya que visual establece la relación de la grilla con el control data automaticamente.
    On Error GoTo error
    'verifico si existen registros seleccionados
    If Data1.Recordset.RecordCount > 0 Then
        RaiseEvent Seleccionar(Data1.Recordset.Fields(propIndiceCampoRetorno).Value)
    End If
Exit Sub
error:
 subControloErrores 516, "gSeleccion_DblClick"
End Sub

'***************************************************************
'*
'*          Métodos proporcionados a la aplicación: CambiarValorCriterio
'*
'***************************************************************

Public Sub CambiarValorCriterio()
    'Por medio de este evento le doy al programador, la posibilidad de poder
    'darle el focus a este control por medio de una tecla de función
    On Error GoTo error
    If txtCriterio.Enabled = True And txtCriterio.Visible = True Then
        txtCriterio.SetFocus
    End If
Exit Sub
error:
    subControloErrores 516, "CambiarValorCriterio"
End Sub

'***************************************************************
'*
'*          Métodos proporcionados a la aplicación: CambiarCriterios
'*
'***************************************************************
Public Sub CambiarCriterios()
    'Por medio de este evento le doy al programador, la posibilidad de asignarle una tecla
    'de función al boton de cambio de criterio.
    If botCambioCriterio.Enabled = True And botCambioCriterio.Visible = True Then
        botCambioCriterio_Click
    End If
End Sub

'***************************************************************
'*
'*          Métodos proporcionados a la aplicación: Mostrar
'*
'***************************************************************

Public Sub ActualizarDatos()
    'Es utilizado para ejecutra nuevamente la consulta de selección
    On Error Resume Next
    Data1.Refresh
    subCambioAnchoColumnas          'reconfiguro tamaño de columnas
End Sub

Public Sub MostrarRapido(consulta As String, Optional posicion As Byte)
    'Muestra información en la grilla, con la característica que no hay que crear
    'la consulta SQL, sino que la misma se pasa drectamente como parámetros
    'aumentando el rendimiento del control.
    '-----------------------------------------------------------------------------------
    'Parámetros.
    '   Entrada:    [consulta]  consulta SQl la cual obtendrá los datos a mostrar
    '               [posicion]   determina donde se muestra los controles de criterio
    '                           1= los controles de criterio se posicionan abajo
    '                           de la grilla
    '-----------------------------------------------------------------------------------
    On Error GoTo error
    
    Dim colActual As Integer
    If posicion = 1 Then
        subPosicionoControles
    End If
    'obtengo parte select
    gConsulta = consulta
    Data1.RecordSource = gConsulta
    colActual = gSeleccion.LeftCol
    Data1.Refresh
    
    'subCambioAnchoColumnas          'reconfiguro tamaño de columnas
    subPosColumnaTrabajo colActual  'posiciono nuevamente en la columna actual
    subCargoListaDeCriterios        'cargo la lista de criterios
    'habilito los controles componentes para trabajar con ellos
    gSeleccion.Enabled = True
    txtCriterio.Enabled = True
    botCambioCriterio.Enabled = True
    'para mejorar interface le doy el focus al textBox de criterio
    txtCriterio.SetFocus
    'Para ordenar la grilla por el campo de defecto indicado en la propiedad NroCampoInicial,
    'desencadeno el evento HeadClick de la grilla, pasándole como parámetro el valor de
    'esta propiedad.
    gSeleccion_HeadClick propNroCampoInicial
Exit Sub
error:
    subControloErrores 518, "MostrarRapido"
End Sub

Public Sub Mostrar(Optional posicion As Byte)
    'Muestra información en la grilla, correspondiente al valor de las
    'propiedades.
    'posiciono controles componentes de acuerdo al valor de la propiedad posicion
    
    'NOTA: Al mostrar consultas complejas(string de propiedades muy largo),
    'la demora puede ser de considerable(un par de segundos).
    'Esto se debe a que hay que recorrer el string caracter por caracter.
    'Lo que quiero resaltar es que el control data no demora en absoluto su ejecución,
    'por más compleja que sea la consulta.
    On Error GoTo error
    
    Dim colActual As Integer
    If posicion = 1 Then
        subPosicionoControles
    End If
    'muestro barra de progreso
    subInicializoBarraProgreso 0
    
    'obtengo parte select
    gConsulta = "Select " & funObtengoParteSelect(propCampos)
    gConsulta = gConsulta & funObtengoParteFromWhere(propTablasRelacionadas, propTabla)
    gConsulta = gConsulta & funObtengoParteSeleccionComplementaria(propSeleccionComplementaria)
    Data1.RecordSource = gConsulta
    colActual = gSeleccion.LeftCol
    Data1.Refresh
    'oculto barra de progreso
    subInicializoBarraProgreso 2
    
    subCambioAnchoColumnas          'reconfiguro tamaño de columnas
    subPosColumnaTrabajo colActual  'posiciono nuevamente en la columna actual
    subCargoListaDeCriterios        'cargo la lista de criterios
    'habilito los controles componentes para trabajar con ellos
    gSeleccion.Enabled = True
    txtCriterio.Enabled = True
    botCambioCriterio.Enabled = True
    'para mejorar interface le doy el focus al textBox de criterio
    txtCriterio.SetFocus
    'Para ordenar la grilla por el campo de defecto indicado en la propiedad NroCampoInicial,
    'desencadeno el evento HeadClick de la grilla, pasándole como parámetro el valor de
    'esta propiedad.
    gSeleccion_HeadClick propNroCampoInicial
Exit Sub
error:
    subControloErrores 518, "Mostrar"
End Sub

Public Sub ClickColumna(columna As Byte)
    'Con este método permito al programador poder implementar teclas
    'que permitan simular un click sobre la columna determinada de la grilla
    'De esta forma podemos cambiar la columna por la cual se ordena la consulta, de tres maneras:
    '   1) con el mouse, haciendo click sobre la columna determinada
    '   2) seleccionando la columna desde la lista de columnas
    '   3) mediante teclas de acceso, por ejemplo Ctrol+F1
    On Error GoTo error
    Dim totalCol As Byte
    
    totalCol = gSeleccion.Columns.Count - 1
    'valido que el número de columna séa valido
    If columna >= 0 And columna <= totalCol Then
        gSeleccion_HeadClick (columna)
    End If
Exit Sub
error:
    subControloErrores 520, "ClickColumna"
End Sub

'***************************************************************
'*
'*  Fin Métodos proporcionados
'*
'***************************************************************

Private Sub subInicializoBarraProgreso(muestro As Byte)
    'Como existe una demora mínima entre el momento de ejecutar el evento
    'mostrar y la aparición de los datos en la grilla, es necesario mostrar
    'una barra de progreso que mejore la interfaz con el usuario.
    
    Select Case muestro
        Case 0         'inicializo para comenzar a mostrar
            UserControl.pbProgreso.Min = 0
    
            UserControl.pbProgreso.Max = Len(propCampos) + _
                                    Len(propTablasRelacionadas)
            UserControl.pbProgreso.Value = 0
            UserControl.pbProgreso.Visible = True
            
        Case 1          'estoy creando la consulta
            UserControl.pbProgreso.Value = UserControl.pbProgreso.Value + 1
            
        Case 2          'finalizó el armado de la consulta
            UserControl.pbProgreso.Visible = False
    End Select
End Sub

Private Sub Image1_Click()
    'Si realizo dobleClick sobre el ícono de error
    'muestro la consulta sql que se generó. De esta manera puedo saber que
    'se ejecutó mal.
    MsgBox gConsulta
End Sub

Private Sub subPosicionoControles()
    'Ubico los controles de criterio por debajo de la grilla de seleccion
    gSeleccion.Top = 0
    lwCriterios.Top = 0
    
    lblCriterio.Top = gSeleccion.Height + 100
    pbProgreso.Top = gSeleccion.Height + 100
    txtCriterio.Top = gSeleccion.Height + lblCriterio.Height + 150
    botCambioCriterio.Top = gSeleccion.Height + lblCriterio.Height + 150
End Sub

Private Sub subCargoListaDeCriterios()
    'Para mejorar la interface con el usuario, se implementa una boton de cambio de criterio.
    'Este boton se puede utilizar como alternativa a hacer click con el mouse sobre la columna
    'por la cual se quiere ordenar la grilla, con la ventaja de utilizar el teclado en vez del mouse.
    'En este procedimiento cargo en el control ListView todos los criterios existentes
    On Error GoTo error
    Dim c As Column
    Dim i As Integer
    Dim totCol As Integer
    i = 0
    'obtengo total de columnas
    totCol = gSeleccion.Columns.Count
    'inicializo lista
    lwCriterios.Clear
    'recorro todas las columnas de la grilla de selección
    Do While i < totCol
        'obtengo título de cada columna
        Set c = gSeleccion.Columns(i)
        'creo un nuevo elmento en el ListView
        lwCriterios.AddItem c.Caption
        lwCriterios.ItemData(lwCriterios.NewIndex) = i  'con este valor ejecuto el eveto
                                                        'HedClick de la grilla
        i = i + 1
    Loop
Exit Sub
error:
     subControloErrores 515, "subCargoListaDeCriterios"
End Sub

Private Sub lwCriterios_DblClick()
    'Selecciono criterio de la lista de criterios
    On Error GoTo error
    subCambioCriterio lwCriterios.ItemData(lwCriterios.ListIndex)
Exit Sub
error:
    subControloErrores 515, "lwCriterios_DblClick"
End Sub

Private Sub lwCriterios_KeyPress(KeyAscii As Integer)
    'Selecciono criterio de la lista de criterios
    On Error GoTo error
    If KeyAscii = propTeclaSeleccion Then
        subCambioCriterio lwCriterios.ItemData(lwCriterios.ListIndex)
    End If
Exit Sub
error:
    subControloErrores 515, "lwCriterios_KeyPress"
End Sub

Private Sub subCambioCriterio(columna As Integer)
    'Ejecuto el evento HeadClick de la grilla, para cambiar el criterio de ordenación
    On Error GoTo error
    gSeleccion_HeadClick (columna)
    'después de seleccionar el nuevo criterio le doy el focus al control de ingreso de criterio
    txtCriterio.SetFocus
Exit Sub
error:
    subControloErrores 515, "subCambioCriterio"
End Sub

Private Sub lwCriterios_LostFocus()
    'Si se pierde el focus de la lista oculta la misma
    lwCriterios.Visible = False
End Sub

Private Function funObtengoParteFromWhere(tablasR As String, tablaP As String) As String
    'Genero un string con formato correspondiente a la parte From y Where de la consulta SQL
    'recorro todos los campos de la propiedad campos
    On Error GoTo error
    Dim largo As String
    Dim cadaCampo As String
    Dim caracter As String
    Dim i As Integer
    
    Dim nomTablaR As String
    Dim indCampoTablaR As Integer
    Dim nomCampoTablaR As String
    Dim indCampoTablaP As Integer
    Dim nomCampoTablaP As String
    Dim parteFrom As String
    Dim parteWhere As String
    
    
    largo = Len(tablasR)
    i = 1
    Do While i <= largo
        caracter = Mid(tablasR, i, 1)    'obtengo todos y cada uno de los caracteres
        If caracter = "@" Then  '@ indica que finaliza el campo
            'parte from
            nomTablaR = mfunObtengoValorDesdeStr(cadaCampo, 2, ";")
            parteFrom = parteFrom & nomTablaR & ","
            'parte where
            indCampoTablaR = mfunObtengoValorDesdeStr(cadaCampo, 3, ";")
            indCampoTablaP = mfunObtengoValorDesdeStr(cadaCampo, 1, ";")
            'obtengo nombre campo tabla principal
            Data1.RecordSource = tablaP
            Data1.Refresh
            nomCampoTablaP = Data1.Recordset.Fields(indCampoTablaP).Name
            'obtengo nombre de campo de tabla relacionada
            Data1.RecordSource = nomTablaR
            Data1.Refresh
            nomCampoTablaR = Data1.Recordset.Fields(indCampoTablaR).Name
            'armo parte where
            parteWhere = parteWhere & tablaP & "." & nomCampoTablaP & "=" & nomTablaR & "." & nomCampoTablaR & " and "
            'incicializo para nuevo campo
            cadaCampo = ""
        Else
            cadaCampo = cadaCampo & caracter
        End If
        i = i + 1
        'muestro barra de progreso
        subInicializoBarraProgreso 1
    Loop
    'a la parte from la sumo la tabla principal
    parteFrom = parteFrom & tablaP
    'verifico si obtuve parte where. Para las consultas que no realizan Join esta variables
    'no se inicializa.
    If Len(parteWhere) > 0 Then
        'realizo Join ya que la consulta obtiene información de más de una tabla
        'resto el último and de la parte where
        parteWhere = Mid(parteWhere, 1, Len(parteWhere) - 4)
        funObtengoParteFromWhere = " from " & parteFrom & " Where " & parteWhere
        gContieneJoin = True
    Else
        'la consulta es simple, es decir solo toma información de una sola tabla
        funObtengoParteFromWhere = " from " & parteFrom
        gContieneJoin = False
    End If
Exit Function
error:
 subControloErrores 515, "funObtengoParteFromWhere"
End Function

Private Function funObtengoParteSelect(campo As String) As String
    'Genero un string con formato correspondiente a sentencia SQL, correspondiente a la
    'parte Select.
    On Error GoTo error
    Dim largo As String
    Dim cadaCampo As String
    Dim caracter As String
    Dim parteSelect As String
    Dim i As Integer
    Dim nomTabla As String
    Dim indCampo As Integer
    Dim descCampo As String
    Dim nomCampo As String
    Dim anchoCampo As Integer
    Dim cabezalCampos As String
    Dim cabezalTablas As String
    Dim cabezalAncho As String
    
    'recorro todos los campos de la propiedad campos
    largo = Len(campo)
    i = 1
    Do While i <= largo
        caracter = Mid(campo, i, 1)    'obtengo todos y cada uno de los caracteres
        If caracter = "@" Then  '@ indica que finaliza el campo
            'obtengo cada valor por separado
            nomTabla = mfunObtengoValorDesdeStr(cadaCampo, 2, ";")
            indCampo = mfunObtengoValorDesdeStr(cadaCampo, 1, ";")
            descCampo = mfunObtengoValorDesdeStr(cadaCampo, 3, ";")
            anchoCampo = Val(mfunObtengoValorDesdeStr(cadaCampo, 4, ";"))
            'inicializo data para obtener nombres de campos
            Data1.RecordSource = nomTabla
            Data1.Refresh
            'obtengo nombre de campo
            nomCampo = Data1.Recordset.Fields(indCampo).Name
            parteSelect = parteSelect & nomTabla & "." & nomCampo & " as " & descCampo & ","
            cadaCampo = ""  'inicializo para cargar información de un nuevo campo
            'obtengo información de campos en grilla secundaria
            cabezalCampos = cabezalCampos & nomCampo & "|"
            cabezalTablas = cabezalTablas & nomTabla & Chr(9)
            cabezalAncho = cabezalAncho & anchoCampo & Chr(9)
        Else
            cadaCampo = cadaCampo & caracter
        End If
        i = i + 1
        'muestro barra de progreso
        subInicializoBarraProgreso 1
    Loop
    'resto última coma
    parteSelect = Mid(parteSelect, 1, Len(parteSelect) - 1)
    'grabo información de campos y tablas en grilla auxiliar
    subGraboInfCampos cabezalTablas, cabezalCampos, cabezalAncho
    funObtengoParteSelect = parteSelect
Exit Function
error:
 subControloErrores 515, "funObtengoParteSelect: " & parteSelect
End Function

Private Function funObtengoParteSeleccionComplementaria(selComple As String) As String
    'Se denomina selección complementaria a la propiedad que contiene una
    'parte de una instrucción SQL la cual realiza una selección de registros determinada.
    'Ejemplo: si se quieren mostrar solo los pasajeros del hotel que esten en este momento
    'hospedados(tabla de pasajeros), hay que incluir una condición en la clausura
    'Where de la instrucción que filtre estos registros.
    'Esta función se encarga de anexar a la variable gConsulta la instrucción pasada
    'como parámetro.
    On Error GoTo error
    Dim clausuraVariable As String
    
    'verifico si la propiedad esta vacía
    If Trim(selComple) <> Empty Then
        'determino si hay parte 1 (ver exolicación de partes en txtCriterio_change)
        If gContieneJoin Then
            clausuraVariable = " and "
        Else
            clausuraVariable = " where "
        End If
    Else
        clausuraVariable = ""
    End If
    funObtengoParteSeleccionComplementaria = clausuraVariable & selComple
Exit Function
error:
 subControloErrores 519, "funObtengoParteSeleccionComplementaria"
End Function

Private Sub subGraboInfCampos(infTablas As String, infCampos As String, infAncho As String)
    'Como en la grilla principal no puede almacenar información referente
    'a los campos que se muestran en la misma, debo de crear una grilla auxiliar no visible,
    'en donde almacenar dicha información.
    'La primer columna de la fila principal corresponde con la primer columna de la fila
    'auxiliar.Utilizo un grilla porque no me interesa crear listas en memoria
    'En la primera fila almaceno el nombre del campo
    'En en la segunda el nombre de la tabla
    'En la tercera fila almaceno el ancho de cada columna
    On Error GoTo error
    UserControl.gInfCampos.FormatString = infCampos
    'cargo en la segunda fila de la grilla los valores correspondientes
    'a los nombres de la tabla
    mSubRangoCeldas gInfCampos, 0, gInfCampos.Cols - 1, 1, 1
    gInfCampos.Clip = infTablas
    'cargo en la tercer fila de la grilla los valores correspondientes
    'a los anchos de las columnas
    mSubRangoCeldas gInfCampos, 0, gInfCampos.Cols - 1, 2, 2
    gInfCampos.Clip = infAncho
Exit Sub
error:
 subControloErrores 515, "subGraboInfCampos"
End Sub

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
        Case 515
            'error de programa
            descAux = " Error desconocido, causado posiblemente por asignación de valores erróneos a propiedades."
        Case 516
            descAux = " El índice del campo de retorno que se indicó como propiedad " & _
                    "no existe en la tabla prncipal."
        Case 517
            descAux = " El número de campo por el cual se quiere ordenar la grilla por defecto, " & _
                    "no es un número de columna válido para ésta."
        Case 518
            descAux = " No se puede formar instrucción Select a causa de propiedades incorrectas." & _
                    "Verifique nombres de tablas, índices de campos, camino y nombre de la BD."
        Case 519
            descAux = " No se puede formar instrucción Select a causa de problemas con la propiedad " & _
                        " SeleccionComplementaria."
        Case 520
            descAux = " Error al cambiar de columna por la cual se ordena la grilla."
    End Select
    errDesc = Err.Number & " " & Err.Description & _
                numErr & descAux & _
                desde & _
                "Consulte con su proveedor de sowftware"
    'muestro error en grilla
    UserControl.Frame1.Visible = True
    UserControl.lblError.Text = errDesc
    'Al presentarse un error aparece un cuadro de díalogo sobre el control indicando
    'que se ha producido un error en el mismo. Como no es bueno modificar la
    'interface de la aplicación con los mensajes de error de los componentes
    'se trata de que éste mensaje no afecte demasiado la interfaz, pero sí que brinde información
    'al usuario o al programador de la aplicación que usa este componenete.
End Sub





