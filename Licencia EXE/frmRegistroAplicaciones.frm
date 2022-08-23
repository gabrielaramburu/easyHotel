VERSION 5.00
Begin VB.Form frmRegistroAplicaciones 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro de aplicaciones."
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6075
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "&Datos de la aplicación "
      Height          =   6495
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   5775
      Begin VB.TextBox txtContrseña 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   240
         MaxLength       =   50
         PasswordChar    =   "*"
         TabIndex        =   11
         Top             =   4200
         Width           =   1455
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   405
         Left            =   1800
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   6000
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtAplicación 
         Height          =   360
         Left            =   240
         MaxLength       =   255
         TabIndex        =   7
         Top             =   2760
         Width           =   5055
      End
      Begin VB.TextBox txtBasDatos 
         Height          =   360
         Left            =   240
         MaxLength       =   255
         TabIndex        =   9
         Top             =   3480
         Width           =   5055
      End
      Begin VB.TextBox txtUnidad 
         Height          =   360
         Left            =   240
         TabIndex        =   5
         Top             =   2040
         Width           =   2295
      End
      Begin VB.CommandButton botIniciar 
         Caption         =   "&Iniciar"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   6000
         Width           =   1215
      End
      Begin VB.CommandButton botCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   2880
         TabIndex        =   13
         Top             =   6000
         Width           =   1215
      End
      Begin VB.CommandButton botRegistrar 
         Caption         =   "&Registrar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4200
         TabIndex        =   14
         Top             =   6000
         Width           =   1215
      End
      Begin VB.TextBox txtNomPropietario 
         Height          =   360
         Left            =   240
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1320
         Width           =   5055
      End
      Begin VB.TextBox txtNomEmpresa 
         Height          =   360
         Left            =   240
         MaxLength       =   50
         TabIndex        =   1
         Top             =   600
         Width           =   5055
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "&Contraseña"
         Height          =   240
         Left            =   240
         TabIndex        =   10
         Top             =   3960
         Width           =   1035
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "&Aplicación"
         Height          =   240
         Left            =   240
         TabIndex        =   6
         Top             =   2520
         Width           =   945
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "&Base de datos"
         Height          =   240
         Left            =   240
         TabIndex        =   8
         Top             =   3240
         Width           =   1320
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "&Unidad "
         Height          =   240
         Left            =   240
         TabIndex        =   4
         Top             =   1800
         Width           =   705
      End
      Begin VB.Label lblInformacionObtenida 
         BorderStyle     =   1  'Fixed Single
         Height          =   1140
         Left            =   240
         TabIndex        =   16
         Top             =   4680
         Width           =   5145
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "N&ombre propietario"
         Height          =   240
         Left            =   240
         TabIndex        =   2
         Top             =   1080
         Width           =   1755
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&Nombre empresa"
         Height          =   240
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Width           =   1590
      End
   End
End
Attribute VB_Name = "frmRegistroAplicaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Declaración de variables
Public claveIngresada As String
Public inicioPresionado As Boolean
Public permitirRegistro As Boolean
Public serieDisco As String
Public aplicacionId As Long

'Utilizada para obtener el número de serie del disco duro
Private Declare Function GetVolumeInformation Lib "Kernel32" _
    Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, _
                                    ByVal lpVolumeNameBuffer As String, _
                                    ByVal nVolumeNameSize As Long, _
                                    lpVolumeSerialNumber As Long, _
                                    lpMaximumComponentLength As Long, _
                                    lpFileSystemFlags As Long, _
                                    ByVal lpFileSystemNameBuffer As String, _
                                    ByVal nFileSystemNameSize As Long) As Long

Private Sub Form_Load()
    'indica si se presionó el boton de inicio
    inicioPresionado = False
    'indica si se pudieron obtener todos los datos para el registro
    permitirRegistro = True
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'Determino los caracteres digitados después de digitar el botón
    'de inicio
    On Error Resume Next
    If KeyAscii = vbKeyEscape Then
        claveIngresada = ""
        Exit Sub
    End If
    If inicioPresionado Then
        claveIngresada = claveIngresada & Chr(KeyAscii)
    End If
    'si la contraseña es correcta y no se produjo ningún error
    If claveIngresada = "manyacapo" And permitirRegistro Then
        botRegistrar.Enabled = True
        'no permito modificar datos del formulario
        Me.txtAplicación.Enabled = False
        Me.txtBasDatos.Enabled = False
        Me.txtContrseña.Enabled = False
        Me.txtNomEmpresa.Enabled = False
        Me.txtNomPropietario.Enabled = False
        Me.txtUnidad.Enabled = False
        Me.botIniciar.Enabled = False
    End If
End Sub

Private Sub botIniciar_Click()
    'Obtengo datos para el registro
    On Error Resume Next
    
    Dim resultadoBD As String
    
    'obtengo serie del disco duro
    serieDisco = funObtengoSerieDisco(Me.txtUnidad)
    'obtengo archivo de identificación de la aplicación a registrar
    aplicacionId = funObtengoIdAplicacion(Me.txtAplicación)
    'accedo al registro correspondiente de la tabla de licencias
    resultadoBD = funAccedoARegistro(Me.txtBasDatos, aplicacionId)
    
    'muestro datos obtenidos en etiqueta
    Me.lblInformacionObtenida = "Serie disco: " & serieDisco & Chr(10)
    Me.lblInformacionObtenida = Me.lblInformacionObtenida & "AplicacionID: " & aplicacionId & Chr(10)
    Me.lblInformacionObtenida = Me.lblInformacionObtenida & "Base de Datos: " & resultadoBD
    
    'indico que se presionó el boton de inicio
    inicioPresionado = True
End Sub

Private Sub botRegistrar_Click()
    'Grabo datos en la base de datos de la aplicación
    On Error GoTo error
    'grabo datos
    Data1.Recordset.Edit
        Data1.Recordset("aplicacionNroLicencia") = aplicacionId & "-" & Hex$(serieDisco)
        Data1.Recordset("aplicacionSerieDisco") = Hex$(serieDisco)
        Data1.Recordset("aplicacionFechaLicencia") = Date
        Data1.Recordset("aplicacionEmpresa") = Me.txtNomEmpresa
        Data1.Recordset("aplicacionDueño") = Me.txtNomPropietario
        Data1.Recordset("aplicacionVD") = False
    Data1.Recordset.Update
error:
    'controlo el resultado de la operación
    Select Case Err.Number
        Case 0
            'se registró la base de datos
            MsgBox "La aplicación se registró correctamente.", vbInformation, "Registración correcta"
            botCancelar_Click
        Case Else
            'no se pudo registrar la base de datos
            MsgBox "No se puedo realizar la registración de la aplicación" & _
                    Chr(10) & _
                    Err.Number & _
                    Chr(10) _
                    & Err.Description, vbExclamation, "Error en registro de aplicación"
    End Select
End Sub

Private Function funAccedoARegistro(base As String, idApli As Long) As String
    'Abro la base de datos y accedo a la tabla de licencias.
    On Error GoTo error
    Me.Data1.Connect = ";PWD=" & Me.txtContrseña.Text & ";"
    Me.Data1.DatabaseName = base
    Me.Data1.RecordSource = "select * from SISTEMA_LICENCIA" & _
                            " where AplicacionId = " & idApli
    Me.Data1.Refresh
    'verifico si obtuve valor
    If Me.Data1.Recordset.RecordCount > 0 Then
        'se encontró el registro
        funAccedoARegistro = "Se accedió al registro."
        'permito registrar la aplicación
        permitirRegistro = True
    Else
        'no se encontró el registro
        funAccedoARegistro = "No se encontró registro con clave igual a Id."
        'no permito registrar la aplicación
        permitirRegistro = False
    End If
Exit Function
error:
    'no permito registrar la aplicación
    permitirRegistro = False
    funAccedoARegistro = "Error " & Err.Number
End Function

Private Function funObtengoSerieDisco(unidad As String) As String
    'Obtengo el número de serie del disco duro de la máquina donde esjecuto
    'esta función
    'Acción

    Dim lVSN As Long, n As Long, s1 As String, s2 As String
    Dim sTmp As String

    On Error GoTo error

    'Reservar espacio para las cadenas que se pasarán al API
    s1 = String$(255, Chr$(0))
    s2 = String$(255, Chr$(0))

    n = GetVolumeInformation(unidad, s1, Len(s1), lVSN, 0, 0, s2, Len(s2))

    's1 será la etiqueta del volumen
    'lVSN tendrá el valor del Volume Serial Number (número de serie del volumen)
    's2 el tipo de archivos: FAT, etc.

    'Convertirlo a hexadecimal para mostrarlo como en el Dir.
    'sTmp = Hex$(lVSN)
    funObtengoSerieDisco = lVSN
    'permito registrar la aplicación
    permitirRegistro = True
Exit Function
error:
    'no permito registrar la aplicación
    permitirRegistro = False
End Function

Private Function funObtengoIdAplicacion(archivo As String) As Long
    'Obtiene el número de Id de una aplicación determinada.
    'Dicho número de Id se almacena en un archivo de texto, que se debe
    'de encontrar en el directorio donde se ejecuta la aplicación.
    'Si se produce un error se devuelve el número de error producido.
    Dim numeroId As String
    On Error GoTo error
    
    'abro archivo para lectura
    Open archivo For Input As #1
    'si el archivo existe leo el número de identificación
    Line Input #1, numeroId
    'cierro el archivo
    Close #1
    funObtengoIdAplicacion = CLng(numeroId)
    'permito registrar la aplicación
    permitirRegistro = True
Exit Function
error:
    Select Case Err.Number
        Case 13 'no coinciden los tipos
            funObtengoIdAplicacion = 518
        Case 53 'el archivo no existe
            funObtengoIdAplicacion = 519
        Case Else
            funObtengoIdAplicacion = Err.Number
    End Select
    'no permito registrar la aplicación
    permitirRegistro = False
    'inicializo el código de error
    Err.Number = 0
End Function

Private Sub botCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Set frmRegistroAplicaciones = Nothing
End Sub
