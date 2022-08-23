VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form frmInicial 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Inicio de aplicación"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5160
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
      DialogTitle     =   "Seleccione base de datos"
   End
   Begin VB.Frame Frame1 
      Caption         =   "Configuración "
      Height          =   3975
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   5895
      Begin VB.ComboBox cboPantalla 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   3240
         Width           =   1935
      End
      Begin VB.TextBox txtBaseDeDatos 
         BackColor       =   &H8000000B&
         Height          =   375
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   2400
         Width           =   4095
      End
      Begin VB.CommandButton botExaminar 
         Caption         =   "&Examinar"
         Height          =   375
         Left            =   4440
         TabIndex        =   0
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   1500
         Left            =   120
         Picture         =   "frmInicial.frx":0000
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1500
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   1815
         Left            =   1800
         TabIndex        =   8
         Top             =   240
         Width           =   3855
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "&Tamaño pantalla"
         Height          =   240
         Left            =   240
         TabIndex        =   1
         Top             =   3000
         Width           =   1530
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Base de datos"
         Height          =   240
         Left            =   240
         TabIndex        =   6
         Top             =   2160
         Width           =   1320
      End
   End
   Begin VB.CommandButton botAceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton botCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4800
      TabIndex        =   4
      Top             =   4200
      Width           =   1215
   End
End
Attribute VB_Name = "frmInicial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Event NotificoCliente(boton As Byte)

Private Sub botAceptar_Click()
    'valido que se halla ingresado base de datos
    If Len(txtBaseDeDatos.Text) > 0 Then
        Me.Visible = False
        RaiseEvent NotificoCliente(1)   'boton aceptar
    Else
        MsgBox "No se ingreso base de datos.", vbExclamation
        botExaminar.SetFocus
    End If
End Sub

Private Sub botCancelar_Click()
    Me.Visible = False
    RaiseEvent NotificoCliente(2)   'boton cancelar
End Sub

Private Sub botExaminar_Click()
    Me.CommonDialog1.Filter = "Acces (*.mdb)|*.mdb"
    Me.CommonDialog1.InitDir = "c:\"
    
    'cdlOFNHideReadOnly no muestro click de solo lectura
    'cdlOFNFileMustExist  el archivo debe de existir
    Me.CommonDialog1.Flags = cdlOFNHideReadOnly Or cdlOFNFileMustExist
    Me.CommonDialog1.ShowOpen
    Me.txtBaseDeDatos.Text = Me.CommonDialog1.filename
End Sub

Private Sub Form_Load()
    'Tipos de pantalla
    cboPantalla.AddItem "800x600"
    cboPantalla.AddItem "640x480"
    'Por defecto marco 800x600
    cboPantalla.Text = "800x600"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'si se hace click en el boton cerrar, oculta el
    'cuadro de diálogo en lugar de descargarlo
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
        Me.Visible = False
    End If
End Sub

