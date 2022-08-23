VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Form1 
   Caption         =   "Formulario de pruba"
   ClientHeight    =   6330
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10860
   LinkTopic       =   "Form1"
   ScaleHeight     =   6330
   ScaleWidth      =   10860
   StartUpPosition =   3  'Windows Default
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Form1.frx":0000
      Height          =   5295
      Left            =   120
      OleObjectBlob   =   "Form1.frx":0010
      TabIndex        =   1
      Top             =   840
      Width           =   10455
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\NANAHOTEL\hotel.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SISTEMA_BITACORA"
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ejecutar"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public bdHOTEL As Database, bdWK As Workspace
Public tbBITACORA As Recordset

Private Controloperaciones As GraboOperacion

Private Sub Command1_Click()
    Set Controloperaciones = New GraboOperacion
    Controloperaciones.GraboOperacionEnBaseDeDatos Date, "", 99, Time, Time, "todo bien cholulo?", tbBITACORA
    Data1.Refresh
End Sub

Private Sub Form_Load()



    
    'asigna espacio trabajo
    Set bdWK = DBEngine.Workspaces(0)
    
    'obtengo el directorio de ejecución del exe.
        
    'abre base de datos
    Set bdHOTEL = bdWK.OpenDatabase("C:\NANAHOTEL\HOTEL.MDB")
    
    'abre tablas
    Set tbBITACORA = bdHOTEL.OpenRecordset("SISTEMA_BITACORA", dbOpenTable)
    
End Sub
