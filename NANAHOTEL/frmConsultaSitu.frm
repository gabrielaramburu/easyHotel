VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{0B963941-6B05-11D5-AE38-892C4BE92F2B}#4.0#0"; "gaHOTELhabitaciones.ocx"
Begin VB.Form frmConsultaSitu 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta de cambios de situación"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\NANAHOTEL\hotel.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select * from situacion_his,tipo_estado_hab"
      Top             =   120
      Visible         =   0   'False
      Width           =   5055
   End
   Begin Hotel_Nana.gaHOTELbarra gaHOTELbarra1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   2
      Top             =   6495
      Width           =   9510
      _ExtentX        =   16775
      _ExtentY        =   582
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cambios de situación"
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9285
      Begin VB.CommandButton botSalir 
         Caption         =   "Salir"
         Height          =   375
         Left            =   7920
         TabIndex        =   3
         Top             =   5760
         Width           =   1215
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmConsultaSitu.frx":0000
         Height          =   4455
         Left            =   240
         OleObjectBlob   =   "frmConsultaSitu.frx":0010
         TabIndex        =   1
         Top             =   960
         Width           =   8775
      End
      Begin gaHOTELhabitaciones.gaHOTELtipo gaHOTELtipo1 
         Height          =   300
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   529
         BackColor       =   -2147483633
      End
   End
End
Attribute VB_Name = "frmConsultaSitu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private hab_cuenta As Long

Private Sub botSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    'Apariencia formulario
    ConfiguroFuentesControlesSistema Me
    
    'obtengo habitacion
    hab_cuenta = Val(frmIngHabitacion2.txtNroHab.Text)
    cabezal_formulario
    genero_consulta
End Sub

Private Sub cabezal_formulario()
    Me.gaHOTELtipo1.CaminoBaseDeDatos = vardir
    Me.gaHOTELtipo1.NumeroHabitacion = hab_cuenta
End Sub

Private Sub genero_consulta()
    Dim consulta As String
    consulta = "select * from situacion_his,tipo_estado_hab where nrohab_situ = " & Str(hab_cuenta) & _
    " and situacion_situ = cod " & _
    " and tipo_cod = 2 " & _
    " order by fechacambio_situ  DESC"
    Data1.RecordSource = consulta
    Data1.Refresh
End Sub
