VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmConsultaTitular 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta de titular"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   9510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Hotel_Nana.gaHOTELbarra gaHOTELbarra1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   6780
      Width           =   9510
      _ExtentX        =   16775
      _ExtentY        =   582
   End
   Begin VB.Frame Frame1 
      Caption         =   "Consulta de titular"
      Height          =   6615
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   9255
      Begin VB.CommandButton botImprimir 
         Height          =   375
         Left            =   6600
         Picture         =   "frmConsultaTitular.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Tag             =   "Imprimir"
         Top             =   6120
         Width           =   1215
      End
      Begin MSFlexGridLib.MSFlexGrid msfgTit 
         Height          =   5175
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   9128
         _Version        =   393216
         FocusRect       =   0
         HighLight       =   0
         AllowUserResizing=   1
      End
      Begin VB.TextBox txtCriterio 
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   6120
         Width           =   5175
      End
      Begin VB.CommandButton botSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   375
         Left            =   7920
         TabIndex        =   5
         Top             =   6120
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "&Grilla de titulares"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lblCriterio 
         AutoSize        =   -1  'True
         Caption         =   "&Ordenado por: "
         Height          =   195
         Left            =   240
         TabIndex        =   0
         Top             =   5880
         Width           =   1065
      End
   End
   Begin VB.Menu mnuFormulario 
      Caption         =   "&Formulario"
      Begin VB.Menu mnuFormularioAceptar 
         Caption         =   "Salir"
         Shortcut        =   {F12}
      End
   End
   Begin VB.Menu mnuOrden 
      Caption         =   "&Ordenado por ..."
      Begin VB.Menu mnuOrdenNombre 
         Caption         =   "Por nombre de titular"
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu mnuOrdenHab 
         Caption         =   "Por número de habitación"
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu mnuOrdentipoTit 
         Caption         =   "Por  tipo de titular"
         Shortcut        =   ^{F3}
      End
      Begin VB.Menu mnuOrdenTipoAloja 
         Caption         =   "Por tipo de alojamiento"
         Shortcut        =   ^{F4}
      End
   End
End
Attribute VB_Name = "frmConsultaTitular"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    subCabezalGrilla
    subCargoGrilla
    'por defecto ordeno por nombre de titular
    mnuOrdenNombre_Click
End Sub

Private Sub subCargoGrilla()
    'Recorro todas las habitaciones
    tbHABITACIONES.MoveFirst
    Do While Not tbHABITACIONES.EOF
        'si la habitación esta ocupada
        If busco_habita_checkin(tbHABITACIONES("nrohab")) Then
            If tbHABITACIONES("titular_unica") <> 0 Then
                subCreoLinea _
                "Unico", tbHABITACIONES("titular_unica"), tbHABITACIONES("nrohab")
            Else
                subCreoLinea _
                "Alojamiento", tbHABITACIONES("titular_aloja"), tbHABITACIONES("nrohab")
                subCreoLinea _
                "Extras", tbHABITACIONES("titular_extra"), tbHABITACIONES("nrohab")
            End If
        End If
        tbHABITACIONES.MoveNext
    Loop
End Sub

Private Sub subCreoLinea(TipoTit As String, cli As Long, hab As Long)
    'Genera una nueva linea en la grilla,
    'la cual representa a un titular
    Dim linea As String

    linea = _
    Chr(9) & _
    obtengo_nombre_pasajero(cli) & _
    Chr(9) & _
    hab & _
    Chr(9) & _
    TipoTit & _
    Chr(9) & _
    funObtengoAlojado(cli, hab)

    msfgTit.AddItem linea
End Sub

Private Sub subCabezalGrilla()
    msfgTit.FormatString = _
    "   |Nombre del titular                                                                     |" & _
    "Habitación         |" & _
    "Tipo de titular    |" & _
    "Alojamiento        "
End Sub
    
Private Function funObtengoAlojado(cli As Long, hab As Long)
    'Determino la ubicación de un cliente
    If busco_pasajero(cli) = hab Then
        funObtengoAlojado = "Dentro"
    Else
        If busco_pasajero(cli) = 0 Then
            funObtengoAlojado = "No Alojado"
        Else
            funObtengoAlojado = "Fuera " & tbCHECKIN("nrohab")
        End If
    End If
End Function

Private Sub Form_Unload(Cancel As Integer)
    'Descargo formulario de memoria
    Set frmConsultaTitular = Nothing
End Sub

Private Sub mnuOrdenHab_Click()
    'Ordeno por número de habitación
    mSubMuestroIcono Me.msfgTit, 2
    subDesmarcoTodas
    mnuOrdenHab.Checked = True
    subOrdenoGrilla 2
End Sub

Private Sub mnuOrdenNombre_Click()
    'Ordeno por nombre de titular
    mSubMuestroIcono Me.msfgTit, 1
    subDesmarcoTodas
    mnuOrdenNombre.Checked = True
    subOrdenoGrilla 1
End Sub

Private Sub mnuOrdenTipoAloja_Click()
    'Ordeno por tipo de alojamiento
    mSubMuestroIcono Me.msfgTit, 4
    subDesmarcoTodas
    mnuOrdenTipoAloja.Checked = True
    subOrdenoGrilla 4
End Sub

Private Sub mnuOrdentipoTit_Click()
    'Ordeno por tipo de titular
    mSubMuestroIcono Me.msfgTit, 3
    subDesmarcoTodas
    mnuOrdentipoTit.Checked = True
    subOrdenoGrilla 3
End Sub

Private Sub txtCriterio_Change()
    'Selecciono filas que cumpla con el criterio
    
    'La columna por la cual se compara el texto es la determinada
    'por el criterio de ordenación
    
    Dim i As Integer
    Dim resultado As Boolean
    i = 1
    'recorro todas las filas de la grilla
    Do While i < msfgTit.Rows
        msfgTit.Row = i
        'resultado = cadena Like patrón
        resultado = msfgTit.Text Like txtCriterio.Text & "*"
        If resultado Then
            'muestro fila
            msfgTit.RowHeight(msfgTit.Row) = 240
        Else
            '0culto fila
            msfgTit.RowHeight(msfgTit.Row) = 0
        End If
        i = i + 1
    Loop
End Sub

Private Sub subOrdenoGrilla(criterio As Byte)
    'Ordeno grilla
    Select Case criterio
        Case 1
            Me.lblCriterio.Caption = Me.lblCriterio.Caption & "nombre de titular"
        Case 2
            Me.lblCriterio.Caption = Me.lblCriterio.Caption & "número de habitación"
        Case 3
            Me.lblCriterio.Caption = Me.lblCriterio.Caption & "tipo de tiular"
        Case 4
            Me.lblCriterio.Caption = Me.lblCriterio.Caption & "tipo de alojamiento"
    End Select
    msfgTit.col = criterio          'determino la columna por la cual se ordena
    msfgTit.Row = msfgTit.RowSel  'ordeno todas las filas
    msfgTit.Sort = 5               'Cadena ascendente. No distingue mayúsculas y minúsculas
    
End Sub

Private Sub subDesmarcoTodas()
    'Desmarco todas las opciones del menu, de esta manera solo
    'se puede ver marcada la opción por la cual esta ordenada la grilla en ese momento.
    Me.mnuOrdenHab.Checked = False
    Me.mnuOrdenNombre.Checked = False
    Me.mnuOrdenTipoAloja.Checked = False
    Me.mnuOrdentipoTit.Checked = False
End Sub

Private Sub botSalir_Click()
    Unload Me
End Sub

Private Sub msfgTit_GotFocus()
    'Para mejorar la interface del usuario siempre le doy el focus al control
    'que ingresa el criterio
    Me.txtCriterio.SetFocus
End Sub

Private Sub mnuFormularioAceptar_Click()
    'Equivale a presionar el boton aceptar o la tecla F12
    botSalir_Click
End Sub

'************************************************************
'*
'* Asistencia al usuario
'*
'************************************************************

Private Sub botSalir_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 1, 1
End Sub

Private Sub botImprimir_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 23
End Sub

Private Sub txtCriterio_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 24
End Sub

Private Sub botSalir_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub botImprimir_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub txtCriterio_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub
