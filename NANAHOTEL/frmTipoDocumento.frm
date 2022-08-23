VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmTipoDocumento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Selección de documentos"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   4920
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   4920
   StartUpPosition =   2  'CenterScreen
   Begin Hotel_Nana.gaHOTELbarra gaHOTELbarra1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5280
      Width           =   4920
      _ExtentX        =   8678
      _ExtentY        =   582
   End
   Begin VB.Frame Frame1 
      Caption         =   "&Tipos de documentos"
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.CommandButton BotAyuda 
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   4
         Top             =   3120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtNroDoc 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   240
         MaxLength       =   8
         TabIndex        =   3
         Top             =   3120
         Width           =   2655
      End
      Begin VB.ListBox LstTipoDoc 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2160
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   2655
      End
      Begin VB.CommandButton botConfirmar 
         Height          =   375
         Left            =   2040
         Picture         =   "frmTipoDocumento.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Tag             =   "Aceptar"
         Top             =   4680
         Width           =   1215
      End
      Begin VB.CommandButton botCancelar 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   3360
         Picture         =   "frmTipoDocumento.frx":08B6
         Style           =   1  'Graphical
         TabIndex        =   6
         Tag             =   "Cancelar"
         Top             =   4680
         Width           =   1215
      End
      Begin VB.Label lblIngNumeroDoc 
         AutoSize        =   -1  'True
         Caption         =   "&Número de documento"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   2880
         Width           =   1620
      End
      Begin ComctlLib.ImageList ImageList1 
         Left            =   3840
         Top             =   1800
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   12
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmTipoDocumento.frx":1178
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmTipoDocumento.frx":1492
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmTipoDocumento.frx":17AC
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmTipoDocumento.frx":1AC6
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmTipoDocumento.frx":1DE0
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmTipoDocumento.frx":20FA
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmTipoDocumento.frx":2414
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmTipoDocumento.frx":272E
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmTipoDocumento.frx":2A48
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmTipoDocumento.frx":2D62
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmTipoDocumento.frx":307C
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmTipoDocumento.frx":3396
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Image Image1 
         Height          =   1215
         Left            =   3120
         Stretch         =   -1  'True
         Top             =   360
         Visible         =   0   'False
         Width           =   1335
      End
   End
   Begin VB.Menu mnuFormulario 
      Caption         =   "&Formulario"
      Begin VB.Menu mnuFormularioAceptar 
         Caption         =   "Aceptar"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuFormularioCancelar 
         Caption         =   "Cancelar"
      End
   End
   Begin VB.Menu mnubuscar 
      Caption         =   "&Buscar..."
      Visible         =   0   'False
      Begin VB.Menu mnuBuscarDocumentos 
         Caption         =   "Documentos..."
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmTipoDocumento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub BotAyuda_Click()
    Select Case tipo_accion_tipodocumento
        Case 1  'consulto facturas
            txtNroDoc.Text = _
                mFunBusqueda(4, LstTipoDoc.ItemData(LstTipoDoc.ListIndex))
        Case 3  'consulto devoluciones
            txtNroDoc.Text = _
                mFunBusqueda(4, LstTipoDoc.ItemData(LstTipoDoc.ListIndex))
        Case 7  'consulto recivos manuales
            txtNroDoc.Text = _
                mFunBusqueda(5, 2)  '2=recivos manuales
    End Select
End Sub

Private Sub Form_Load()
    'Apariencia formulario
    mSubConfiguroFuentesControlesSistema Me
    
    Select Case tipo_accion_tipodocumento
        Case 1  'Consulto facturas
            cargo_facturas_y_boletas
            botAyuda.Visible = True
            Me.mnuBuscar.Visible = True
        Case 2
            cargo_facturas_y_boletas
        Case 3  'Consulto devoluciones
            cargo_devoluciones
            botAyuda.Visible = True
            Me.mnuBuscar.Visible = True
        Case 4
            'consulto o elimino recivo automatico
            Me.Caption = "Recivos"
            cargo_recivo_automatico
        Case 5
            Me.Caption = "Nuevo recivo"
            'nuevo recivo automático
            cargo_recivos
            'no hay que ingresar nro. de documento
            txtNroDoc.Visible = False
            lblIngNumeroDoc.Visible = False
        Case 6
            'nuevo recivo manual
            Me.Caption = "Nuevo recivos manual"
            cargo_recivos
        Case 7  'consulto Recivos
            cargo_recivo_manual
            Me.Caption = "Recivos manuales"
            botAyuda.Visible = True
            Me.mnuBuscar.Visible = True
    End Select
    'me posiciono en el primer documento de la lista
    LstTipoDoc.ListIndex = 0
End Sub

Private Sub botConfirmar_Click()
    Select Case tipo_accion_tipodocumento
        Case 1
            If valido_fac_bol Then
                Me.Hide
                frmFacturacion.Show 1
            End If
        Case 2
            'nueva devolución
            If valido_nrodocu Then
                Me.Hide
                frmDevolucion.Show 1
            End If
        Case 3
            'consulto o anulo devolución
            If valido_devolu Then
                Me.Hide
                frmDevolucion.Show 1
            End If
        Case 4
            'consulto o elimino recivo automatico o manual
            If valido_recivo_automatico Then
                Me.Hide
                frmRecivo.Show 1
            End If
        Case 5
            'nuevo recivo automático
            Me.Hide
            frmRecivo.Show 1
        Case 6
            'nuevo recivo manual
            If valido_nuevo_recivo_manual Then
                Me.Hide
                frmRecivo.Show 1
            End If
        Case 7
            If valido_recivo_manual Then
                Me.Hide
                frmRecivo.Show 1
            End If
    End Select
End Sub

Private Function valido_devolu()
    Dim tipo_docu As Byte
    'Valido que el número de devolución a consultar o anular exista.
    valido_devolu = False
    tipo_docu = Me.LstTipoDoc.ItemData(Me.LstTipoDoc.ListIndex)
    If busco_documentoTF(tipo_docu, Val(Me.txtNroDoc.Text)) Then
        valido_devolu = True
    Else
        'no existe la devolución
        mSubMensaje 4, 40
        txtNroDoc.Text = Empty
    End If
End Function

Private Function valido_nrodocu()
    Dim tipo_docu As Byte
    'Valido si el número de documento que ingreso existe y además.
    valido_nrodocu = False
    tipo_docu = Me.LstTipoDoc.ItemData(Me.LstTipoDoc.ListIndex)
    If busco_documentoTF(tipo_docu, Val(Me.txtNroDoc.Text)) Then
        valido_nrodocu = True
    Else
        'el documento no existe
        mSubMensaje 4, 41
        txtNroDoc.Text = Empty
    End If
End Function

Private Function valido_nuevo_recivo_manual()
    'Valida que el número de resivo ingresado no exista
    valido_nuevo_recivo_manual = False
    If Val(txtNroDoc.Text) <> 0 Then
        'si no existe
        If Not busco_recivoTF(2, Val(txtNroDoc.Text)) Then  'no existe recivo
            valido_nuevo_recivo_manual = True
        Else
            'ya existe ese recivo
            mSubMensaje 4, 42
            txtNroDoc.Text = Empty
        End If
    End If
End Function

Private Function valido_recivo_automatico()
    'Valida que el número de recivo ingresado exista
    valido_recivo_automatico = False
    If busco_recivoTF(1, Val(txtNroDoc.Text)) Then
        valido_recivo_automatico = True
    Else
        'no existe ese recivo
        mSubMensaje 4, 43
        txtNroDoc.Text = Empty
    End If
End Function

Private Function valido_recivo_manual()
    'Valida que el número de recivo ingresado exista
    valido_recivo_manual = False
    If busco_recivoTF(2, Val(txtNroDoc.Text)) Then
        valido_recivo_manual = True
    Else
        'no existe ese recivo
        mSubMensaje 4, 43
        txtNroDoc.Text = Empty
    End If
End Function

Private Function valido_fac_bol()
    valido_fac_bol = False
    'si existe tipo de documento continúo
    If busco_documentoTF(Me.LstTipoDoc.ItemData(Me.LstTipoDoc.ListIndex), Val(Me.txtNroDoc.Text)) Then
        valido_fac_bol = True
    Else
        'no existe el documento seleccionado
        mSubMensaje 4, 41
        txtNroDoc.Text = Empty
    End If
End Function

Private Sub cargo_facturas_y_boletas()
    LstTipoDoc.AddItem "Contado M/N"
    LstTipoDoc.ItemData(LstTipoDoc.NewIndex) = 1
    LstTipoDoc.AddItem "Contado U$S"
    LstTipoDoc.ItemData(LstTipoDoc.NewIndex) = 2
    LstTipoDoc.AddItem "Factura M/N"
    LstTipoDoc.ItemData(LstTipoDoc.NewIndex) = 3
    LstTipoDoc.AddItem "Factura U$S"
    LstTipoDoc.ItemData(LstTipoDoc.NewIndex) = 4
End Sub

Private Sub cargo_devoluciones()
    LstTipoDoc.AddItem "Dev. Contado M/N"
    LstTipoDoc.ItemData(LstTipoDoc.NewIndex) = 5
    LstTipoDoc.AddItem "Dev. Contado U$S"
    LstTipoDoc.ItemData(LstTipoDoc.NewIndex) = 6
    LstTipoDoc.AddItem "Dev. Crédito M/N"
    LstTipoDoc.ItemData(LstTipoDoc.NewIndex) = 7
    LstTipoDoc.AddItem "Dev. Crédito U$S"
    LstTipoDoc.ItemData(LstTipoDoc.NewIndex) = 8
End Sub

Private Sub cargo_recivos()
    LstTipoDoc.AddItem "Recivo M/N"
    LstTipoDoc.ItemData(LstTipoDoc.NewIndex) = 9
    LstTipoDoc.AddItem "Recivo U$S"
    LstTipoDoc.ItemData(LstTipoDoc.NewIndex) = 10
End Sub

Private Sub cargo_recivo_automatico()
    LstTipoDoc.AddItem "Recivo "
    LstTipoDoc.ItemData(LstTipoDoc.NewIndex) = 11
End Sub

Private Sub cargo_recivo_manual()
    LstTipoDoc.AddItem "Recivo manual"
    LstTipoDoc.ItemData(LstTipoDoc.NewIndex) = 12
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Descargo formulario de memoria
    Set frmTipoDocumento = Nothing
End Sub

Private Sub LstTipoDoc_Click()
    txtNroDoc.Text = ""
    'Cargo la imagen correspondiente al documento seleccionado
    Me.Image1.Picture = _
    ImageList1.ListImages.Item(LstTipoDoc.ItemData(LstTipoDoc.ListIndex)).Picture
End Sub

Private Sub LstTipoDoc_DblClick()
    botConfirmar_Click
End Sub

Private Sub botCancelar_Click()
    Unload Me
End Sub

Private Sub txtNroDoc_KeyPress(KeyAscii As Integer)
    'Valido números
    ValidoNum KeyAscii, False, False
End Sub

Private Sub mnuBuscarDocumentos_Click()
    'Equivale a presionar el boton de ayuda o la tecla F1
    If botAyuda.Visible = True Then
        BotAyuda_Click
    End If
End Sub

Private Sub mnuFormularioAceptar_Click()
    'Equivale a presionar el boton de aceptar o la tecla F12
    botConfirmar_Click
End Sub

Private Sub mnuFormularioCancelar_Click()
    'Equivale a presionar el boton de cancelar
    botCancelar_Click
End Sub

