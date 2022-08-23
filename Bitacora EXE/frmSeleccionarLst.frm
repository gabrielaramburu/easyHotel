VERSION 5.00
Begin VB.Form frmSeleccionarLst 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seleccionar listado"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6450
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   6450
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Seleccione un listado "
      Height          =   5415
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6255
      Begin VB.CommandButton botNingunoPredeterminado 
         Caption         =   "Ninguno"
         Default         =   -1  'True
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   4920
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtInfLst 
         BackColor       =   &H80000000&
         Height          =   1335
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   3120
         Width           =   5775
      End
      Begin VB.CommandButton botImprimir 
         Caption         =   "Imprimir"
         Height          =   375
         Left            =   3120
         TabIndex        =   7
         Top             =   4920
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton botPredeterminar 
         Caption         =   "Predeterminar"
         Height          =   375
         Left            =   2040
         TabIndex        =   6
         Top             =   4920
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton botEliminar 
         Caption         =   "Eliminar"
         Height          =   375
         Left            =   960
         TabIndex        =   5
         Top             =   4920
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ListBox lstListados 
         Columns         =   1
         Height          =   1980
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   5775
      End
      Begin VB.CommandButton botEjecutar 
         Caption         =   "Ejecutar"
         Height          =   375
         Left            =   3600
         TabIndex        =   4
         Top             =   4920
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton botCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   4920
         TabIndex        =   3
         Top             =   4920
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Información del listado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&Listados disponibles"
         Height          =   240
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Width           =   1860
      End
   End
End
Attribute VB_Name = "frmSeleccionarLst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub botCancelar_Click()
    Unload Me
End Sub

Private Sub botEjecutar_Click()
    'Ejecuto el listado por pantalla
    Me.Hide
    mSubEjecutoListado Me.lstListados.List(lstListados.ListIndex)
    Unload Me
End Sub

Private Sub botEliminar_Click()
    'Elimino el listado seleccionado
    mSubEliminoListado Me.lstListados.List(lstListados.ListIndex)
    Unload Me
End Sub

Private Sub botImprimir_Click()
    'Imprimo listado
    Me.Hide
    mSubImprimoListado Me.lstListados.List(lstListados.ListIndex)
    Unload Me
End Sub

Private Sub botNingunoPredeterminado_Click()
    'No establece ningún listado como predeterminado
    mSubPredeterminarListado ""
    Unload Me
End Sub

Private Sub botPredeterminar_Click()
    'Establece el listado seleccionado como el predeterminado
    mSubPredeterminarListado Me.lstListados.List(lstListados.ListIndex)
    Unload Me
End Sub

Private Sub Form_Load()
    'Cargo los listados existentes
    If tbSISTEMA_BITACORA.RecordCount > 0 Then
        tbSISTEMA_BITACORAlistados.Index = "pk_listado"
        tbSISTEMA_BITACORAlistados.MoveFirst
        Do While Not tbSISTEMA_BITACORAlistados.EOF
            'creo nuevo elemento
            Me.lstListados.AddItem tbSISTEMA_BITACORAlistados("NomLst")
            tbSISTEMA_BITACORAlistados.MoveNext
        Loop
        'Según la acción a realizar depende el boton que se muestra
        subMuestroBoton
        If lstListados.ListCount > 0 Then
          lstListados.ListIndex = 0
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Descargo formulario de memoria
    Set frmSeleccionarLst = Nothing
End Sub

Private Sub lstListados_Click()
    'Cada vez que cambio de listado muestro la selección
    'Muestro la descripción del listado
    'en la venta de descripción
    If mfunBuscoListado(lstListados.List(lstListados.ListIndex)) Then    'si existe listado
        txtInfLst.Text = tbSISTEMA_BITACORAlistados("InfLst")
    Else
        txtInfLst.Text = ""
    End If
End Sub

Private Sub subMuestroBoton()
    Dim bot As CommandButton
    Select Case tipo_accion_selec
        Case 1  'ejecutar
            Set bot = Me.botEjecutar
        Case 2  'imprimir
            Set bot = Me.botImprimir
        Case 3  'eliminar
            Set bot = Me.botEliminar
        Case 4  'predeterminar
            Set bot = Me.botPredeterminar
            Me.botNingunoPredeterminado.Visible = True
    End Select
    bot.Visible = True
    bot.Left = 3400
    bot.Default = True
End Sub

