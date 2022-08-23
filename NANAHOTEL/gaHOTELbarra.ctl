VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.UserControl gaHOTELbarra 
   Alignable       =   -1  'True
   ClientHeight    =   330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4995
   ScaleHeight     =   330
   ScaleWidth      =   4995
   ToolboxBitmap   =   "gaHOTELbarra.ctx":0000
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   135
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   238
      _Version        =   327682
      Appearance      =   1
   End
   Begin ComctlLib.StatusBar stbBarra 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   529
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   1764
            MinWidth        =   1764
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   1764
            MinWidth        =   1764
            Picture         =   "gaHOTELbarra.ctx":0312
            Text            =   "Usurio"
            TextSave        =   "Usurio"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   1764
            MinWidth        =   1764
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "gaHOTELbarra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Const m_def_ancho = 330         'El ancho del control
Const m_def_largo_panel3 = 1000 'Largo minimo del panel fecha
Const m_def_largo_panel2 = 2000 'Largo minimo del panel usuario
Const m_def_MsgError = "Eror en: gaHOTELbarra.ocx"

'Valor por defecto de las propiedades
Const m_def_leyenda = ""
Const m_def_alineacion = 2

'Este es un evento es un evento del control constituyente ProgresBarr
'que es exspuesto para añadir utilidad al control ocx.
'concretamente cada vez que se hace dlbclick sobre usuarios
'se muestra pantalla de cambio de usuario.
Public Event DblClickSobreUsuario()

Private Sub stbBarra_PanelDblClick(ByVal Panel As ComctlLib.Panel)
    If Panel.Index = 2 Then
        RaiseEvent DblClickSobreUsuario
    End If
End Sub

Private Sub UserControl_InitProperties()
    'Tamaño por defecto
    'Primera instancia del formulario
    UserControl.Height = m_def_ancho
    
    'Creo la barra en el borde inferior de formulario
    UserControl.Extender.Align = 2
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error GoTo error
    'Cuando creo una nueva instancia en tiempo de ejecución
    'cargo los valores de usuario y de fecha
    If Ambient.UserMode = True Then
        stbBarra.Panels(2).Text = m_UsuarioSisNom
        stbBarra.Panels(3).Text = m_FechaSis
    End If

error:
    control_errores
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    'El ancho del control es fijo
    UserControl.Height = m_def_ancho
                        
    'Tamaño de paneles
    stbBarra.Panels(2).Width = m_def_largo_panel2
    stbBarra.Panels(3).Width = m_def_largo_panel3
    'El panel de leyenda se adapta al resto que quede libre
    stbBarra.Panels(1).Width = UserControl.Width - _
                            m_def_largo_panel3 - _
                            m_def_largo_panel2
End Sub

Private Sub control_errores()
    Dim msg_error As String
    Select Case err.Number
        Case 0
            Exit Sub
        Case Else
            msg_error = "Error: " & err.Number & " " & err.Description & Chr(10) _
            & m_def_MsgError
    End Select
End Sub

Public Sub Leyenda(texto As String)
    'Muestra leyenda
    stbBarra.Panels(1).Text = texto
    
    'El tipo de la barra de tareas se trensforma en normal
    UserControl.stbBarra.Style = sbrNormal
    
    'Oculta barra de progreso
    UserControl.ProgressBar1.Visible = False
End Sub

Public Sub Progreso(min As Long, max As Long, valor As Long)
    'Muestra barra de progreso
    
    'Configuro largo de barra
    UserControl.ProgressBar1.Width = 4000
    
    'El tipo de la barra de tareas se transforma en simple
    UserControl.stbBarra.Style = sbrSimple
    
    'Muestro barra de progreso
    UserControl.ProgressBar1.Visible = True
    
    'Cargo valores a barra de proceso
    UserControl.ProgressBar1.min = min
    UserControl.ProgressBar1.max = max
    UserControl.ProgressBar1.Value = valor
End Sub

Public Sub ProgresoFin()
    UserControl.ProgressBar1.Visible = False
    'El tipo de la barra de tareas se transforma en normal
    UserControl.stbBarra.Style = sbrNormal
End Sub

Public Sub InicializoUsuario()
    stbBarra.Panels(2).Text = m_UsuarioSisNom
End Sub

Public Sub InicializoFecha()
    stbBarra.Panels(3).Text = m_FechaSis
End Sub

