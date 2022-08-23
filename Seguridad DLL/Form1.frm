VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4785
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4785
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command9 
      Caption         =   "Mensaje"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Width           =   2295
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Stand By"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   2295
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Nuevo usuario Admin"
      Height          =   495
      Left            =   2400
      TabIndex        =   6
      Top             =   4080
      Width           =   2175
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Elimino usuario"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   4080
      Width           =   2175
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Modifico contraseña"
      Height          =   495
      Left            =   2400
      TabIndex        =   4
      Top             =   3480
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Nuevo usuario"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   735
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Administrador"
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   1680
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "No muestro usuario"
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   960
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Muestro usuarios"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public bdHOTEL As Database, bdWK As Workspace
Public tbUSUARIOS As Recordset
Public tbParametros As Recordset
Public tbperfiles As Recordset
    
Private WithEvents PidoClave As UsuarioMuestro
Attribute PidoClave.VB_VarHelpID = -1
Private WithEvents pidoclave2 As NoMuestroUsuario
Attribute pidoclave2.VB_VarHelpID = -1
Private WithEvents pido3 As NoMuestroUsuario
Attribute pido3.VB_VarHelpID = -1
Private WithEvents nuevousuario As Contraseñas
Attribute nuevousuario.VB_VarHelpID = -1
Private WithEvents modificocontra As Contraseñas
Attribute modificocontra.VB_VarHelpID = -1
Private WithEvents EliminOusuario As Contraseñas
Attribute EliminOusuario.VB_VarHelpID = -1
Private WithEvents nuevousuarioadmin As Contraseñas
Attribute nuevousuarioadmin.VB_VarHelpID = -1
Private WithEvents standbyusr As UsuarioMuestro
Attribute standbyusr.VB_VarHelpID = -1
Private WithEvents MuestroMensaje As Mensaje
Attribute MuestroMensaje.VB_VarHelpID = -1

Private Sub Command4_Click()
    Set nuevousuario = New Contraseñas

        nuevousuario.nuevousuario tbUSUARIOS
        
    Set nuevousuario = Nothing

End Sub

Private Sub Command7_Click()
    Set nuevousuarioadmin = New Contraseñas
    nuevousuarioadmin.nuevousuarioadmin tbParametros
    Set nuevousuarioadmin = Nothing
End Sub

Private Sub Command6_Click()

    Set EliminOusuario = New Contraseñas
    EliminOusuario.EliminOusuario tbUSUARIOS, tbperfiles, "gabriel"
    Set EliminOusuario = Nothing
    
End Sub

Private Sub Command5_Click()
    Set modificocontra = New Contraseñas
    modificocontra.ModificoUsuario tbUSUARIOS

    Set modificocontra = Nothing
End Sub

Private Sub Command1_Click()
    'creo un nuevo objeto thing
    
    Set PidoClave = New UsuarioMuestro
    PidoClave.MuestroUsuario tbUSUARIOS
    Set PidoClave = Nothing
End Sub

Private Sub Command2_Click()
    'creo un nuevo objeto thing
    
    Set pidoclave2 = New NoMuestroUsuario
    pidoclave2.MuestroSinUsuario tbUSUARIOS
    Set pidoclave2 = Nothing
End Sub

Private Sub Command3_Click()
    'creo un nuevo objeto thing
    
    Set pido3 = New NoMuestroUsuario
    pido3.MuestroAdmin "hola"
    Set pido3 = Nothing
End Sub

Private Sub Command8_Click()
    'stand by
    
    Set standbyusr = New UsuarioMuestro
    standbyusr.MuestroUsuarioStandBy tbUSUARIOS
    Set standbyusr = Nothing
End Sub

Private Sub Command9_Click()
    'mensaje
    
    Set MuestroMensaje = New Mensaje
    MuestroMensaje.MensajeAccesoDenegado "Gabriel"
    Set MuestroMensaje = Nothing
End Sub

Private Sub Form_Load()
    Dim VARDIR As String
    Dim VARDIR2 As String
    'asigna espacio trabajo
    Set bdWK = DBEngine.Workspaces(0)
    
    'obtengo el directorio de ejecución del exe.
    VARDIR = App.Path & "\hotel.mdb"         'directorio para BD
    VARDIR2 = App.Path & "\reportes\"        'directorio para reportes
        
    'abre base de datos
    Set bdHOTEL = bdWK.OpenDatabase("C:\NANAHOTEL\HOTEL.MDB")
    
    'abre tablas
    Set tbUSUARIOS = bdHOTEL.OpenRecordset("SISTEMA_USUARIOS", dbOpenTable)
    Set tbParametros = bdHOTEL.OpenRecordset("sistema_parametros", dbOpenTable)
    Set tbperfiles = bdHOTEL.OpenRecordset("sistema_perfiles", dbOpenTable)

End Sub

Private Sub standbyusr_NotificoCliente(usuario As String, boton As Byte)
    Debug.Print "Hola"
End Sub
