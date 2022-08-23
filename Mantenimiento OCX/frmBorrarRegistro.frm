VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmBorrarRegistro 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Eliminación de regristros"
   ClientHeight    =   2580
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5805
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3240
      Top             =   120
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2280
      Visible         =   0   'False
      Width           =   2340
   End
   Begin MSFlexGridLib.MSFlexGrid gPropiedad 
      Height          =   1095
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Visible         =   0   'False
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   1931
      _Version        =   393216
      Rows            =   0
      Cols            =   4
      FixedRows       =   0
      FixedCols       =   0
   End
   Begin ComctlLib.ListView lstTareas 
      Height          =   1695
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   2990
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Verificación realizada"
         Object.Width           =   7832
      EndProperty
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   2280
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Frame Frame1 
      Caption         =   "Error en ocx."
      Height          =   2055
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   5055
      Begin VB.Label lblError 
         Caption         =   "Etiqueta de errores"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   720
         TabIndex        =   5
         Top             =   240
         Width           =   3855
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   120
         Picture         =   "frmBorrarRegistro.frx":0000
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.Label lblEstado 
      AutoSize        =   -1  'True
      Caption         =   "Verificando si es posible eliminar el regristro ..."
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   3210
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   5040
      Picture         =   "frmBorrarRegistro.frx":0442
      Top             =   80
      Visible         =   0   'False
      Width           =   480
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   3720
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmBorrarRegistro.frx":0884
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmBorrarRegistro.frx":0BD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmBorrarRegistro.frx":0F28
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmBorrarRegistro.frx":1242
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmBorrarRegistro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Declaro propiedad del formulario

'Esta propiedad contiene información de los controles a realizar
Public propFormControlIntegridad As String
'Esta propiedad contiene el valor que se desea eliminar
Public propValorAEliminar As Variant
'Esta propiedad determina la base de datos con la que trabaja en control data
Public propCaminoBase As String
'Esta propiedad determina el registro a eliminar ya que no esta disponible en este formulario
Public propRegistroEliminar As Data

Public propResultadoEliminacion As Boolean      'Esta propiedad es leída por el control
                                                'una vez cerrado el formulario para determinar si
                                                'se pudo borrar el registro

Private Sub Form_KeyPress(KeyAscii As Integer)
    'Con la tecla ESC cierro el formulario
    If KeyAscii = 27 Then
        'oculto el formulario porque después de cerrado necesito acceder a información
        'para poder desencadenar el evento que indica que se borro un registro.
        Me.Hide
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo error
    'por defecto asumo que todo salió bien, es decir que no se produjo ningún error
    'de funcionamiento y que el registro a borrar no contiene referencias en otras tablas.
    propResultadoEliminacion = True
    'inicializo grilla de propiedades
    subInicializoGrilla
    'cargo la grilla de propiedades con los controles a realizar
    subCargoGrillaPropiedades
    'establesco propiedades de la barra de progreso
    Me.ProgressBar1.Min = 0
    Me.ProgressBar1.Max = gPropiedad.Rows + 1 'obtengo cantidad de controles a ralizar
                                              'El elemento que se suma pertenece
                                              'al tiempo que lleva borrar el registro
    Me.ProgressBar1.Value = 0
    'establezco propiedad del control data
    Data1.DatabaseName = propCaminoBase
    'desencadeno evento
    Timer1.Enabled = True
Exit Sub
error:
    subControloErroresForm 515, "form_Load frmBorrarRegistro"
End Sub

Private Sub subInicializoGrilla()
    'creo el cabezal de la grilla de propiedades
    gPropiedad.FormatString = "Tabla | " & _
                                "Indice campo clave |" & _
                                "Desc. tabla | " & _
                                "Menaje err "
End Sub

Private Sub subCargoGrillaPropiedades()
    'Recorro el string y muestro información de campos en grilla
    On Error GoTo error
    Dim largo As String
    Dim cadaCampo As String
    Dim caracter As String
    
    Dim i As Integer
    
    largo = Len(propFormControlIntegridad)
    i = 1
    Do While i <= largo
        'proceso todos y cada uno de los caracteres de la pripiedad de integridad
        caracter = Mid(propFormControlIntegridad, i, 1)
        If caracter = ";" Then  '; indica nuevo campo
            'cada campo se asigna a una columna diferente en la grilla de propiedad
            caracter = Chr(9)   'chr(9) indica nueva columna
        End If
        If caracter = "@" Then  '@ indica que finaliza el campo
            'creo una fila en la grilla de propiedades
            gPropiedad.AddItem ""
            mSubRangoCeldas gPropiedad, 0, _
                            gPropiedad.Cols - 1, _
                            gPropiedad.Rows - 1, _
                            gPropiedad.Rows - 1
            gPropiedad.Clip = cadaCampo
            cadaCampo = ""  'inicializo para cargar información de un nuevo control
        Else
            cadaCampo = cadaCampo & caracter
        End If
        i = i + 1
    Loop
Exit Sub
error:
    subControloErroresForm 515, "subCargoGrillaPropiedades frmBorrarRegistro"

End Sub

Private Function funVerificoRegistros() As Boolean
    'Verifico si es posible eliminar el registro
    On Error GoTo error
    Dim i As Integer
    Dim camposProcesados As String
    Dim descErr As String
    Dim descVer As String
    Dim nomTabla As String
    Dim indCampo As Integer
    
    
    funVerificoRegistros = True
    i = 0   'comienzo por la primer fila que contenga datos
    'recorro grilla
    Do While i < gPropiedad.Rows
        nomTabla = gPropiedad.TextMatrix(i, 0)  'nombre de la tabla
        indCampo = gPropiedad.TextMatrix(i, 1)  'indice del campo a buscar
        'busco registro en la tabla
        If Not funBuscoRegistro(nomTabla, indCampo) Then
            'no permito borrar
            funVerificoRegistros = False
            'obtengo mensaje de error
            descErr = gPropiedad.TextMatrix(i, 3)  'mensaje de error
            'muestro línea de error en la grilla
            subMuestroLineaLista descErr, 1
            Exit Do
        Else
            'obtengo mensaje de verificación
            descVer = "Verificación en tabla de " & gPropiedad.TextMatrix(i, 2) & " ok."  'mensaje de verificación
            'muestro la tarea en la lista
            subMuestroLineaLista descVer, 2
        End If
      
        i = i + 1
        'paso a un nuevo control
        Me.ProgressBar1.Value = i
    Loop
Exit Function
error:
    subControloErroresForm 515, "funVerificoRegistros frmBorrarRegistro"
End Function

Private Sub subMuestroLineaLista(desc As String, tipoIcono As Byte)
    'Creo una línea en la grilla indicando la tarea que se acaba de realizar
    On Error GoTo error
    lstTareas.ListItems.Add , , desc, , tipoIcono
    lstTareas.View = lvwReport
Exit Sub
error:
    subControloErroresForm 515, "subMuestroLineaLista frmborrarRegistro"
End Sub

Private Function funBuscoRegistro(tabla As String, indCampo As Integer) As Boolean
    'Sea A ---> B una relacion N a 1:
    'Si quiero eliminar un registro de la tabla B, debo verificar que no exista en A
    'Si existe no puedo eliminar porque debo de mantener la integridad referencial.
    'El parámetro tabla, hace referencias a todas las posibles tablas A y indCampo,
    'es el campo de la tabla A donde se hacer referencia a B.
    On Error GoTo error
    Dim consulta As String
    funBuscoRegistro = True
    'determino el tipo de la clave
    If Not IsNumeric(propValorAEliminar) Then
        'si string tengo que anexar comillas para que se ejecute correctamente
        propValorAEliminar = "'" & propValorAEliminar & "'"
    End If
    'obtengo el nombre del campo a buscar
    consulta = tabla
    Data1.RecordSource = consulta
    Data1.Refresh
    'realizo consulta para buscar dato
    consulta = "select " & Data1.Recordset(indCampo).Name & _
                " from " & tabla & _
                " where " & Data1.Recordset(indCampo).Name & _
                " = " & propValorAEliminar
    Data1.RecordSource = consulta
    Data1.Refresh
    'verifico si encontré el valor en la tabla
    If Data1.Recordset.RecordCount > 0 Then
        'si encontré no permito borrar
        funBuscoRegistro = False
    End If
Exit Function
error:
    subControloErroresForm 515, "funBuscoRegistro frmBorrarRegistro"
End Function

Private Sub Timer1_Timer()
    'Luego de mostrar el formulario necesito un evento que permita
    'comenzar con la verificación y ese evento lo creo con este temporizador
    On Error GoTo error
    Timer1.Enabled = False
    'inicio proceso de verificación
    If funVerificoRegistros Then
        If propResultadoEliminacion = True Then 'verifico si no se produjo ningún error
            'elimino registro
            propRegistroEliminar.Recordset.Delete
            'muestro resultado de la eliminación
            subMuestroLineaLista "Eliminación de registro ok", 4
            lblEstado.Caption = "El registro se eliminó correctamente"
        End If
    Else
        'el registro contiene referencias a otras tablas
        propResultadoEliminacion = False
        'muestro resultado de la eliminación
        subMuestroLineaLista "No se eliminó el registro", 3
        lblEstado.Caption = "El registro no se eliminó"
        'muestro ícono que indica que no se borró el registro
        Me.Image1.Visible = True
    End If
    'barra de progreso
    Me.ProgressBar1.Visible = False
Exit Sub
error:
    subControloErroresForm 515, "Timer1_Timer frmBorrarRegistro"
End Sub

Private Sub subControloErroresForm(numErr As Integer, desde As String)
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
        Case Else
            
    End Select
    propResultadoEliminacion = False
    errDesc = Err.Number & " " & Err.Description & Chr(10) & _
                numErr & descAux & Chr(10) & _
                desde & Chr(10) & _
                "Consulte con su proveedor de sowftware"
    'muestro error en grilla
    Frame1.Visible = True
    lblError.Caption = errDesc
    'oculto demás controles
    lstTareas.Visible = False

End Sub
