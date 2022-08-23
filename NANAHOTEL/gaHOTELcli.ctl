VERSION 5.00
Begin VB.UserControl gaHOTELcli 
   ClientHeight    =   1215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2475
   ScaleHeight     =   1215
   ScaleWidth      =   2475
   ToolboxBitmap   =   "gaHOTELcli.ctx":0000
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   840
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2415
      Begin VB.Label lblNomCli 
         Caption         =   "lblNomCli"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblCodCli 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "lblCodCli"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
   End
End
Attribute VB_Name = "gaHOTELcli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const def_borde = 200
Const def_ancho = 735   'El ancho del control
Const m_def_MsgError = "Error grave en: gaHOTELcli.ocx"


'Declaración de variables de propiedades
Dim m_BaseDatos As String
Dim m_CodCli As Long

'Valor de propiedades por defecto
Const m_def_CodCli = 0
Const m_def_NomCli = "Cli. -desc-"
Const m_def_BackColor = 0
Const m_def_BaseDatos = ""



Private Sub UserControl_AmbientChanged(PropertyName As String)
    Dim objctl As Object
    If PropertyName = "BackColor" Then
        'Determino las propiedades (básicas) del control
        'igual a las del contenedor
        For Each objctl In Controls
            objctl.BackColor = Ambient.BackColor
        Next
        UserControl.BackColor = Ambient.BackColor
        'cuando vuelvo a crear una instancia del control devo
        'de establecer la propiedad iguale a la nueva del contenedor.
        PropertyChanged "BackColor"
    End If
End Sub

Private Sub UserControl_Initialize()
    'Cuando creo una nueva instancia se despliegan en las etiquetas los
    'valores por defecto.
    lblCodCli.Caption = m_def_CodCli    '0
    lblNomCli.Caption = m_def_NomCli    '"Cli. -desc-"
End Sub

Private Sub UserControl_InitProperties()
    'Determino las propiedades (básicas) del control
    'igual a las del contenedor
    Dim objctl As Object
    
    'Para cada control componente
    For Each objctl In Controls
        objctl.BackColor = Ambient.BackColor
    Next
    'Para el UserControl
    UserControl.BackColor = Ambient.BackColor
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    Dim objctl As Object

    UserControl.BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    
    'Determino las propiedades (básicas) del control
    'igual a las del contenedor
    For Each objctl In Controls
        objctl.BackColor = UserControl.BackColor
    Next
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    'El ancho del control siempre es el mismo
    Height = def_ancho
    'Si se cambia el largo del control, cambio el largo de la etiqueta
    'lblnomcli para que se adapte al nuevo borde derecho
    'y tambien el largo del frame
    
    Frame1.Width = Width - def_borde + 150
    lblNomCli.Width = Frame1.Width - lblNomCli.Left - def_borde
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, m_def_BackColor)
End Sub

Public Property Get CodigoCliente() As Long
    CodigoCliente = lblCodCli.Caption
End Property

Public Property Let CodigoCliente(ByVal New_CodigoCliente As Long)
    On Error GoTo error
    'Esta propiedad esta disponible solo en tiempo de ejecución
    If Ambient.UserMode = True Then
        m_CodCli = New_CodigoCliente
        realizo_consulta m_CodCli
    End If
error:
    control_errores
End Property

Public Property Get CaminoBaseDeDatos() As String
    CaminoBaseDeDatos = m_def_BaseDatos
End Property

Public Property Let CaminoBaseDeDatos(ByVal New_CaminoBaseDeDatos As String)
    If Ambient.UserMode = True Then
        m_BaseDatos = New_CaminoBaseDeDatos
        Data1.DatabaseName = m_BaseDatos
    End If
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    Dim objctl As Object
    
    'Cuando cambia el valor de la propiedad ForeColor,
    'la cambio a todos los controles componentes
    For Each objctl In Controls
        objctl.BackColor = New_BackColor
    Next
    'Hago lo propio con UserControl
    UserControl.BackColor = New_BackColor

    PropertyChanged "BackColor"
End Property

Private Sub realizo_consulta(CodCli As Long)
    On Error GoTo error
    Dim consulta As String
    consulta = "select nombre_completo_titular " & _
                "from clientes " & _
                "where nrocorr = " & CodCli
    Data1.RecordSource = consulta
    Data1.Refresh
    If Data1.Recordset.RecordCount > 0 Then
        lblCodCli.Caption = CodCli
        lblNomCli.Caption = Data1.Recordset("nombre_completo_titular")
    Else
        'Si el número de cliente no existe, despliego valores por defecto
        lblCodCli.Caption = m_def_CodCli
        lblNomCli.Caption = m_def_NomCli
    End If
error:
    control_errores
End Sub

Private Sub control_errores()
    Dim msg_error As String
    Select Case err.Number
        Case 0
            Exit Sub
        Case 91
            msg_error = "No se especificó el camino de la base de datos."
        Case 3024
            msg_error = "No se encontró la base de datos."
    End Select
    MsgBox "Error: " & err.Number & " " & err.Description & Chr(10) & msg_error, vbCritical, m_def_MsgError
End Sub




