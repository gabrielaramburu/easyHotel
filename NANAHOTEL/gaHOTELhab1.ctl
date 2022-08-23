VERSION 5.00
Begin VB.UserControl gaHOTELtipo 
   ClientHeight    =   900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3480
   ScaleHeight     =   900
   ScaleWidth      =   3480
   ToolboxBitmap   =   "gaHOTELhab1.ctx":0000
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   600
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Suite"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   -37
      Width           =   615
   End
   Begin VB.Label lblNroHab 
      Caption         =   "lblNroHab"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   -37
      Width           =   855
   End
   Begin VB.Label lblTipoHab 
      Caption         =   "lblTipoHab"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   -37
      Width           =   1455
   End
End
Attribute VB_Name = "gaHOTELtipo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const def_borde = 50
Const def_ancho = 300   'El ancho del control
Const m_def_MsgError = "Error grave en: gaHOTELtipo.ocx"

'Declaración de Eventos
Event Click()

'Declaración de variables de propiedades
Dim m_BaseDatos As String
Dim m_NroHab As Long

'Valor de propiedades por defecto
Const m_def_NroHab = 0
Const m_def_TipoHab = "Hab. -desc-"
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
        PropertyChanged "BackColor"
    End If
End Sub

Private Sub UserControl_Initialize()
    lblNroHab.Caption = m_def_NroHab
    lblTipoHab.Caption = m_def_TipoHab
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
    'El ancho del control siempre es el mismo
    Height = def_ancho
    'Si se cambia el largo del control, cambio el largo de la etiqueta lblTipoHab
    'para que se adapte al nuevo borde derecho
    lblTipoHab.Width = Width - lblTipoHab.Left - def_borde
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, m_def_BackColor)
End Sub

Public Property Get NumeroHabitacion() As Long
Attribute NumeroHabitacion.VB_Description = "Determina el número de habitación que se desea mostrar  en control. Se asigna en tiempo de ejecución."
Attribute NumeroHabitacion.VB_ProcData.VB_Invoke_Property = ";Comportamiento"
Attribute NumeroHabitacion.VB_MemberFlags = "200"
    NumeroHabitacion = lblNroHab.Caption
End Property

Public Property Let NumeroHabitacion(ByVal New_NumeroHabitacion As Long)
    On Error GoTo error
    If Ambient.UserMode = True Then
        m_NroHab = New_NumeroHabitacion
        realizo_consulta m_NroHab
    End If
error:
    control_errores
End Property

Public Property Get CaminoBaseDeDatos() As String
Attribute CaminoBaseDeDatos.VB_Description = "Determina la ubicación de la base de datos. Se asigna en tiempo de ejecución."
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

Private Sub realizo_consulta(nrohab As Long)
    On Error GoTo error
    Dim consulta As String
    consulta = "select descripcion " & _
                "from habitaciones,tipo_habitaciones " & _
                "where habitaciones.nrohab = " & nrohab & " and " & _
                "habitaciones.tipohab = tipo_habitaciones.tipohab"
    Data1.RecordSource = consulta
    Data1.Refresh
    If Data1.Recordset.RecordCount > 0 Then
        lblNroHab.Caption = nrohab
        lblTipoHab.Caption = Data1.Recordset("descripcion")
    Else
        lblNroHab.Caption = m_def_NroHab
        lblTipoHab.Caption = m_def_TipoHab
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


