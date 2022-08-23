VERSION 5.00
Begin VB.UserControl gaHOTELtitular 
   ClientHeight    =   1335
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6435
   ScaleHeight     =   1335
   ScaleWidth      =   6435
   ToolboxBitmap   =   "gaHOTELhab2.ctx":0000
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6375
      Begin Hotel_Nana.gaHOTELtipo gaHOTELtipo1 
         Height          =   300
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   529
         BackColor       =   -2147483633
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   4560
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   120
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label lblTipoTit2 
         Caption         =   "lblTipoTit2"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label lblTipoTit1 
         Caption         =   "lblTipoTit1"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblNomTit2 
         Caption         =   "lblNomTit2"
         Height          =   255
         Left            =   1320
         TabIndex        =   2
         Top             =   960
         Width           =   4575
      End
      Begin VB.Label lblNomTit1 
         Caption         =   "lblNomTit1"
         Height          =   255
         Left            =   1320
         TabIndex        =   1
         Top             =   720
         Width           =   4575
      End
   End
End
Attribute VB_Name = "gaHOTELtitular"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Const def_ancho_control = 1335
Const def_borde_control = 200

'Defino constantes
Const m_def_tipo1 = "T. Unico"
Const m_def_tipo2 = "T. Aloja."
Const m_def_tipo3 = "T. Extras"
Const m_def_NoExisteCli = "Cliente -desc-"

Const m_def_MsgError = "Error grave en: gaHOTELtitular.ocx "

'Defino constantes de propidades
Const m_def_NroHab = 0
Const m_def_BaseDatos = ""
Const m_def_BackColor = 0

'Defino variables de propiedades
Dim m_NroHab As Long
Dim m_BaseDatos As String

'Variables
Dim m_titularAloja As Long
Dim m_titularExtra As Long




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
    lblNomTit1.Caption = m_def_NoExisteCli
    lblNomTit2.Caption = m_def_NoExisteCli
    lblTipoTit1.Caption = "1er Titular"
    lblTipoTit2.Caption = "2do Titular"
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

Private Sub UserControl_Resize()
    'El ancho del control siempre es el mismo
    Height = def_ancho_control
    'Si se cambia el largo del control,
    'cambio el largo de las etiquetas y el frame,
    'para que se adapten al nuevo borde derecho.
    Frame1.Width = Width - def_borde_control + 150
    lblNomTit1.Width = Frame1.Width - lblNomTit1.Left - def_borde_control
    lblNomTit2.Width = Frame1.Width - lblNomTit2.Left - def_borde_control
    gaHOTELtipo1.Width = Frame1.Width - gaHOTELtipo1.Left - def_borde_control
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "BackColor", UserControl.BackColor, m_def_BackColor
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Dim objctl As Object
        
    UserControl.BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    
    'Para cada control componente
    For Each objctl In Controls
        objctl.BackColor = UserControl.BackColor
    Next
    'Para el UserControl
    UserControl.BackColor = UserControl.BackColor
End Sub

Public Property Get NumeroHabitacion() As Long
Attribute NumeroHabitacion.VB_Description = "Número de habitación a mostrar. Se asigna en tiempo de ejecución."
    NumeroHabitacion = m_def_NroHab
End Property

Public Property Let NumeroHabitacion(ByVal New_NumeroHabitacion As Long)
    On Error GoTo error
    If Ambient.UserMode = True Then
        m_NroHab = New_NumeroHabitacion
        realizo_consulta
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
    
    UserControl.BackColor = New_BackColor
    
    'Para cada control componente
    For Each objctl In Controls
        objctl.BackColor = New_BackColor
    Next
    'Para el UserControl
    UserControl.BackColor = New_BackColor

    PropertyChanged "BackColor"
End Property

Private Sub realizo_consulta()
    On Error GoTo error
    Dim consulta As String
    consulta = "Select titular_unica,titular_aloja,titular_extra " & _
               "from habitaciones " & _
               "where nrohab=" & m_NroHab
    Data1.RecordSource = consulta
    Data1.Refresh
    If Data1.Recordset.RecordCount > 0 Then 'Si encontre habitación
        'muestro tipo de habitación, aunque la habitación no este ocupada
        gaHOTELtipo1.CaminoBaseDeDatos = m_BaseDatos
        gaHOTELtipo1.NumeroHabitacion = m_NroHab
        If habitacionOcupada Then
            'muestro titulares
            If Data1.Recordset("titular_unica") <> 0 Then
                cuenta_unica
            Else
                m_titularAloja = Data1.Recordset("titular_aloja")
                m_titularExtra = Data1.Recordset("titular_extra")
                cuentas_separadas
            End If
        End If
    End If
    
error:
    control_errores
End Sub

Private Sub cuenta_unica()
    Dim consulta As String
    consulta = obtengo_nombre_cliente(Data1.Recordset("titular_unica"))
    Data1.RecordSource = consulta
    Data1.Refresh
    'Muestro nombre y tipo titular
    If Data1.Recordset.RecordCount > 0 Then
        lblNomTit1.Caption = Data1.Recordset("nombre_completo_titular")
    Else
        lblNomTit1.Caption = m_def_NoExisteCli
    End If
    lblTipoTit1.Caption = m_def_tipo1
    
    'No muestro segundo titular (porque no existe).
    lblNomTit2.Visible = False
    lblTipoTit2.Visible = False
End Sub

Private Sub cuentas_separadas()
    Dim consulta As String
    'Muestro nombre y tipo titular Alojamiento
    consulta = obtengo_nombre_cliente(m_titularAloja)
    Data1.RecordSource = consulta
    Data1.Refresh
    If Data1.Recordset.RecordCount > 0 Then
        lblNomTit1.Caption = Data1.Recordset("nombre_completo_titular")
    Else
        lblNomTit1.Caption = m_def_NoExisteCli
    End If
    lblTipoTit1.Caption = m_def_tipo2
    
    'Muestro nombre y tipo titular Extras
    consulta = obtengo_nombre_cliente(m_titularExtra)
    Data1.RecordSource = consulta
    Data1.Refresh
    If Data1.Recordset.RecordCount > 0 Then
        lblNomTit2.Caption = Data1.Recordset("nombre_completo_titular")
    Else
        lblNomTit2.Caption = m_def_NoExisteCli
    End If
    lblTipoTit2.Caption = m_def_tipo3
End Sub

Private Function obtengo_nombre_cliente(cli As Long)
    obtengo_nombre_cliente = _
    "Select nombre_completo_titular " & _
    "from clientes " & _
    "where nrocorr = " & cli
End Function

Private Function habitacionOcupada()
    'Determino si la habitación esta ocupada, mediante el método de titulares,
    'es decir si no tiene titulares asignados es porque no hay nadie alojado en ella.
    habitacionOcupada = True
    If Data1.Recordset("titular_extra") = 0 _
    And Data1.Recordset("titular_aloja") = 0 _
    And Data1.Recordset("titular_unica") = 0 Then
        habitacionOcupada = False
    End If
End Function

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

