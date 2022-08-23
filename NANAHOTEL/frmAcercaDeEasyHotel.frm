VERSION 5.00
Begin VB.Form frmAcercaDeEasyHotel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Acerca de EasyHotel"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6855
   Icon            =   "frmAcercaDeEasyHotel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton botInfSistema 
      Caption         =   "&Información del sistema..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   7
      Top             =   3840
      Width           =   2055
   End
   Begin VB.CommandButton botAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   6
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   2985
      Left            =   120
      Picture         =   "frmAcercaDeEasyHotel.frx":000C
      Top             =   120
      Width           =   1260
   End
   Begin VB.Label lblUsoAutorizado 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblUsoAutorizado"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1680
      TabIndex        =   10
      Top             =   1920
      Width           =   4215
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "lblVersion"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   5085
      TabIndex        =   9
      Top             =   840
      Width           =   720
   End
   Begin VB.Label lblRevision 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "lblRevision"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   5040
      TabIndex        =   8
      Top             =   1080
      Width           =   765
   End
   Begin VB.Label lblAdvertencia 
      Caption         =   "lblAdvertencia"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   5
      Top             =   3240
      Width           =   4335
   End
   Begin VB.Image Image1 
      Height          =   105
      Left            =   120
      Picture         =   "frmAcercaDeEasyHotel.frx":B71E
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   6570
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Se autoriza este producto a:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1680
      TabIndex        =   4
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Marcos Bernini"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1680
      TabIndex        =   3
      Top             =   1080
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Gabriel Aramburu"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1680
      TabIndex        =   2
      Top             =   840
      Width           =   1275
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Copyright (c) 2000 - 2002"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1680
      TabIndex        =   1
      Top             =   600
      Width           =   1860
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "EasyHotel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   1155
   End
End
Attribute VB_Name = "frmAcercaDeEasyHotel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Todas estas chingadas son para mostrar la información del sistema
' Opciones de seguridad de claves del Registro...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Tipos principales de claves del Registro...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Cadena Unicode terminada en Null
Const REG_DWORD = 4                      ' Número de 32 bits

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Private Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' Prueba a obtener del Registro la información del sistema sobre el nombre y la ruta del programa...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ' Prueba a obtener del Registro la información del sistema sobre la ruta del programa...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' Comprueba la existencia de una versión conocida de un archivo de 32 bits
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
        ' Error - Imposible encontrar el archivo...
        Else
            GoTo SysInfoErr
        End If
    ' Error - Imposible encontrar la entrada de Registro...
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox "System Information Is Unavailable At This Time", vbOKOnly
End Sub

Private Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' Contador de bucle
    Dim rc As Long                                          ' Código de retorno
    Dim hKey As Long                                        ' Controlador a una clave de Registro abierta
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Tipo de dato de una clave de Registro
    Dim tmpVal As String                                    ' Almacén temporal de una valor de clave de Registro
    Dim KeyValSize As Long                                  ' Tamaño de la variable de la clave de Registro
    '------------------------------------------------------------
    ' Abre la clave de Registro en la raíz {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Abre la clave de Registro
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Trata el error...
    
    tmpVal = String$(1024, 0)                               ' Asigna espacio para la variable
    KeyValSize = 1024                                       ' Marca el tamaño de la variable
    
    '------------------------------------------------------------
    ' Recupera valores de claves de Registro...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Obtiene o crea un valor de clave
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Trata el error
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 agrega una cadena terminada en Null...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Se encontró Null, se extrae de la cadena
    Else                                                    ' WinNT no tiene una cadena terminada en Null...
        tmpVal = Left(tmpVal, KeyValSize)                   ' No se encontró Null, sólo se extrae la cadena
    End If
    '------------------------------------------------------------
    ' Determina el tipo de valor de la clave para conversión...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Busca tipos de datos...
    Case REG_SZ                                             ' Tipo de dato de la cadena de la clave de Registro
        KeyVal = tmpVal                                     ' Copia el valor de la cadena
    Case REG_DWORD                                          ' El tipo de dato de la cadena de la clave es Double Word
        For i = Len(tmpVal) To 1 Step -1                    ' Convierte cada byte
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Genera el valor carácter a carácter
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Convierte Double Word a String
    End Select
    
    GetKeyValue = True                                      ' Vuelve con éxito
    rc = RegCloseKey(hKey)                                  ' Cierra la clave de Registro
    Exit Function                                           ' Salir
    
GetKeyError:      ' Restaurar después de que ocurra un error...
    KeyVal = ""                                             ' Establece el valor de retorno para una cadena vacía
    GetKeyValue = False                                     ' Devuelve un error
    rc = RegCloseKey(hKey)                                  ' Cierra la clave de Registro
End Function
'Finaliza chingadas de información del sistema
'-----------------------------------------------------------------------------------------------------------------------------
Private Sub botInfSistema_Click()
    StartSysInfo
End Sub

Private Sub botAceptar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim infLicenciaApli As InformacionApli
    
    Dim nomHotel As String
    Dim nomPropietario As String
    Dim nroLicencia As String
    
    'inicializo etiqueta de advertencia
    Me.lblAdvertencia.Caption = _
    "Advertencia: este programa está protegido por las leyes de " & _
    "derechos de autor y otros tratados internacionales. " & _
    "La reproducción o ditribución no autorizada de este programa, o de " & _
    "cualquier parte del mismo, está penada por la ley con sanciones civiles y penales, " & _
    "y será objeto de todas las acciones judiciales que correspondan. "
    'muestro versión
    Me.lblVersion.Caption = "Versión 1.0"
    'muestro revisión
    Me.lblRevision.Caption = "Revisión " & App.Major & "." & _
                                        App.Minor & "." & App.Revision
    
    'determino si la aplicación es una versión demo o registrada
    If gEsUnaVersionDemo Then
        'es una versión demo
        nomHotel = "Versión no registrada."
        nomPropietario = "Se autoriza el uso durante el período de evaluación."
        nroLicencia = ""
    Else
        'es una versión registrada
        'creo nueva instancia y otengo datos de la licencia de la aplicación
        Set infLicenciaApli = New InformacionApli
            nomHotel = infLicenciaApli.mFunObtenerLicenciaApli(idApli, tbSISTEMA_LICENCIA, 2)
            nomPropietario = infLicenciaApli.mFunObtenerLicenciaApli(idApli, tbSISTEMA_LICENCIA, 3)
            nroLicencia = "Número de licencia: " & _
            infLicenciaApli.mFunObtenerLicenciaApli(idApli, tbSISTEMA_LICENCIA, 1)
        'destruyo instancia
        Set infLicenciaApli = Nothing
    End If
    
    'muestro autorización de uso
    Me.lblUsoAutorizado.Caption = _
    "  " & nomHotel & Chr(10) & _
    "   " & nomPropietario & Chr(10) & _
    "   " & nroLicencia
End Sub

