VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Prueba."
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   5895
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton botObtengoLicencia 
      Caption         =   "Obtengo Licencia"
      Height          =   375
      Left            =   2520
      TabIndex        =   9
      Top             =   1200
      Width           =   1815
   End
   Begin VB.CommandButton botObtenerDue�o 
      Caption         =   "Obtengo Due�o"
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton botEmpresa 
      Caption         =   "Obtengo empresa"
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton botActualidoDia 
      Caption         =   "Actualizo ultimo d�a"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2640
      Width           =   2175
   End
   Begin VB.CommandButton botCantDiasPermitidos 
      Caption         =   "Cant. d�as permitidos"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   2175
   End
   Begin VB.CommandButton botCantDiasPeriodo 
      Caption         =   "Cantidad d�as periodo"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   2175
   End
   Begin VB.CommandButton botObtenerId 
      Caption         =   "&Controlar lisencia"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   2175
   End
   Begin VB.CommandButton botAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   4440
      TabIndex        =   2
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton botFinEvaluaci�n 
      Caption         =   "&Fin Evaluaci�n"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   2175
   End
   Begin VB.CommandButton botAvisoVersionDemo 
      Caption         =   "Aviso &versi�n demo"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public bdHOTEL As Database, bdWK As Workspace
Public tbLICENCIA As Recordset

Private avisoFin As AvisoFinPeriodoDemo
Private avisoDemo As AvisoVersionDemo
Private bibloControlVersionDemo As ControlVersionDemo
Private bibloAvisoErrores As AvisoErrores
Private funInicializarAplicacion As InicializarAplicacion
Private informacionAplicacion As InformacionApli

Private Sub botActualidoDia_Click()
    Dim idapli As Long
    Set bibloControlVersionDemo = New ControlVersionDemo
        'obtengo c�digo de aplicaci�n
        idapli = bibloControlVersionDemo.funObtengoIdAplicacion(App.Path & "\Aplicacion.id.txt")
        If bibloControlVersionDemo.funActualizarUltimoDiaEjecuci�nVD(tbLICENCIA, idapli, Date) Then
            MsgBox "Se actualiz� la fecha del �ltimo d�a de ejecuci�n"
        Else
            MsgBox "Error al actualizar fecha del �ltimo d�a de ejecuci�n", vbExclamation
        End If
    Set bibloControlVersionDemo = Nothing
End Sub

Private Sub botCantDiasPeriodo_Click()
    Dim idapli As Long
    Set bibloControlVersionDemo = New ControlVersionDemo
        'obtengo c�digo de aplicaci�n
        idapli = bibloControlVersionDemo.funObtengoIdAplicacion(App.Path & "\Aplicacion.id.txt")
        MsgBox "Cantidad de d�as de uso: " & bibloControlVersionDemo.funObtenerCantDiasUtilizadosVD(tbLICENCIA, Date, idapli)
    Set bibloControlVersionDemo = Nothing
End Sub

Private Sub botCantDiasPermitidos_Click()
    Dim idapli As Long
    Set bibloControlVersionDemo = New ControlVersionDemo
        'obtengo c�digo de aplicaci�n
        idapli = bibloControlVersionDemo.funObtengoIdAplicacion(App.Path & "\Aplicacion.id.txt")
        MsgBox "Cantidad de d�as permitidos: " & bibloControlVersionDemo.FunObtenerCantDiasAutorizadosVD(tbLICENCIA, idapli)
    Set bibloControlVersionDemo = Nothing
End Sub


Private Sub botObtenerDue�o_Click()
    Dim idapli As Long
    Set bibloControlVersionDemo = New ControlVersionDemo
    Set informacionAplicacion = New InformacionApli
    idapli = bibloControlVersionDemo.funObtengoIdAplicacion(App.Path & "\Aplicacion.id.txt")
    MsgBox informacionAplicacion.mFunObtenerLicenciaApli(idapli, tbLICENCIA, 3)
End Sub

Private Sub botObtengoLicencia_Click()
    Dim idapli As Long
    Set bibloControlVersionDemo = New ControlVersionDemo
    Set informacionAplicacion = New InformacionApli
    idapli = bibloControlVersionDemo.funObtengoIdAplicacion(App.Path & "\Aplicacion.id.txt")
    MsgBox informacionAplicacion.mFunObtenerLicenciaApli(idapli, tbLICENCIA, 1)
End Sub

Private Sub Form_Load()
    'asigna espacio trabajo
    Set bdWK = DBEngine.Workspaces(0)
    
    'abre base de datos
    Set bdHOTEL = bdWK.OpenDatabase("C:\NANAHOTEL\HOTEL.MDB", False, False, ";PWD=manyacapo;")
    
    'abre tablas
    Set tbLICENCIA = bdHOTEL.OpenRecordset("SISTEMA_LICENCIA", dbOpenTable)
End Sub

Private Sub botAceptar_Click()
    Unload Me
End Sub

Private Sub botAvisoVersionDemo_Click()
    Set avisoDemo = New AvisoVersionDemo
    'Establesco propiedades del forumlario de versi�n demo
    
    avisoDemo.AvisoVersionDemoPropTituloForm = "Versi�n de evaluaci�n(no registrada)."
    avisoDemo.AvisoVersionDemoPropNomAplicacion = "Hotel2000"
    avisoDemo.AvisoVersionDemoPropSistemaAplicacion = "Para Window 9x/2000/XT/NT"
    avisoDemo.AvisoVersionDemoPropVersionAplicacion = "Version de evaluaci�n 1.0"
    avisoDemo.AvisoVersionDemoPropDiasDemos = "60"
    avisoDemo.AvisoVersionDemoPropPeriodoDeUso = "D�as utilizado: " & "1" & " de su per�odo de " & avisoDemo.AvisoVersionDemoPropDiasDemos & " d�as."
    avisoDemo.AvisoVersionDemoPropDerechos = "Copyright(c) 2000-2002 by Chupacabras Sowftware" & Chr(10) & _
                                        "All Rights Reserved." & Chr(10) & _
                                        "Maldonado - Uruguay" & Chr(10) & _
                                        "www.chupacabrasventanita.com"
    avisoDemo.MostrarAvisoVersionDemo
    Set avisoDemo = Nothing
End Sub

Private Sub botFinEvaluaci�n_Click()
    Set avisoFin = New AvisoFinPeriodoDemo
    'Establesco todas las propiedades del objeto creado AvisoFinPeriodoDemo
    avisoFin.AvisoFinPeriodoDemoPropNomAplicacion = "Hotel2000"
    avisoFin.AvisoFinPeriodoDemoPropSistemaAplicacion = "Para Window 9x/2000/XT/NT"
    avisoFin.AvisoFinPeriodoDemoPropVersionAplicacion = "Version de evaluaci�n 1.0"
    avisoFin.AvisoFinPeriodoDemoPropPeriodoTerminado = "D�as utilizado: " & "30" & " de su per�odo de 30 d�as."
    avisoFin.AvisoFinPeriodoDemoPropTituloForm = "Fin del per�odo de evaluaci�n."
    avisoFin.AvisoFinPeriodoDemoPropExtension = True
    
    avisoFin.AvisoFinPeriodoDemoPropDerechos = "Copyright(c) 2000-2002 by Chupacabras Sowftware" & Chr(10) & _
                                        "All Rights Reserved." & Chr(10) & _
                                        "Maldonado - Uruguay" & Chr(10) & _
                                        "www.chupacabrasventanita.com"
    avisoFin.MostrarAvisoFinPeriodoDemo
    Set avisoFin = Nothing
End Sub

Private Sub botEmpresa_Click()
    Dim idapli As Long
    Set bibloControlVersionDemo = New ControlVersionDemo
    Set informacionAplicacion = New InformacionApli
    idapli = bibloControlVersionDemo.funObtengoIdAplicacion(App.Path & "\Aplicacion.id.txt")
    MsgBox informacionAplicacion.mFunObtenerLicenciaApli(idapli, tbLICENCIA, 2)
End Sub

Private Sub botObtenerId_Click()
    Dim codigoLicencia As Integer
    Dim serieDisco As String
    Dim idapli As Long
    Dim resIniApli As Integer
    
    
    'Creo nuevas instancias de las clases que contienen los componentes de c�digo
    'que se utilizan para validar la aplicaci�n.
    Set bibloControlVersionDemo = New ControlVersionDemo
    Set bibloAvisoErrores = New AvisoErrores
    
    'obtengo serie del disco duro
    serieDisco = bibloControlVersionDemo.funObtengoSerieDisco("C:\")
    'verifico tipo de licencia
    codigoLicencia = bibloControlVersionDemo.funControloLicenciaAplicacion(App.Path & "\Aplicacion.id.txt", _
                                                                            tbLICENCIA, _
                                                                            Date, _
                                                                            serieDisco)
    Select Case codigoLicencia
        Case 621
            botAvisoVersionDemo_Click
        Case 622
            'Es una versi�n registrada
        Case 514
            botFinEvaluaci�n_Click
        Case 513
            'No se encontr� un registro con la clave primaria igual al n�mero de identificaci�n
            'de la aplicaci�n que estoy ejecutando.
            
            'obtengo c�digo de aplicaci�n
            idapli = bibloControlVersionDemo.funObtengoIdAplicacion(App.Path & "\Aplicacion.id.txt")
            'ejecuto inicializaci�n de la aplicaci�n
            Set funInicializarAplicacion = New InicializarAplicacion
                resIniApli = funInicializarAplicacion.funInicializarAplicacion(tbLICENCIA, _
                                                                            idapli, _
                                                                            Date)
                Select Case resIniApli
                    Case 0
                        'despu�s de inicializar la aplicaci�n muestro aviso de versi�n demo.
                         botAvisoVersionDemo_Click
                    Case Else
                        'problemas al inicalizar la aplicaci�n
                        bibloAvisoErrores.propMsgError = "Error: " & resIniApli & Chr(10) & _
                                                            "Problemas al inicializar la aplicaci�n."
                        bibloAvisoErrores.propDescMsgError = "El proceso de inicializaci�n de la aplicaci�n, " & _
                                                            "no se pudo realizar correctamente. Si es la primera vez que " & _
                                                            "ejecuta esta aplicaci�n, es posible que el problema se origine " & _
                                                            "por una instalaci�n incorrecta. Si por el contrario ustes ya la ha ejecutado " & _
                                                            "con exito anteriormente, es muy problable que halla problemas con la base " & _
                                                            "de datos de la aplicaci�n o con alg�n componente del hardware de su equipo."
                        bibloAvisoErrores.propContactarse = "Para solucionar el problema comun�quese con nosotros."
                        bibloAvisoErrores.MostrarMensaje
                End Select
            Set funInicializarAplicacion = Nothing
        Case 515
            'la fecha del sistema fue retrocedida.
            bibloAvisoErrores.propMsgError = "Error: " & codigoLicencia & Chr(10) & _
                                                "Problemas con la fecha del sistema."
            bibloAvisoErrores.propDescMsgError = "La fecha del sistema es menor a la fecha en la cual se " & _
                                                "ejecut� la aplicaci�n por �ltima vez. Como usted est� ejecutando, " & _
                                                "una versi�n de evaluaci�n no registrada, esta distorsi�n en la fecha, " & _
                                                "impide ejecutar la aplicaci�n."
            bibloAvisoErrores.propContactarse = "Para solucionar el problema comun�quese con nosotros."
            bibloAvisoErrores.MostrarMensaje

        Case 516
            'no se puede identificar el tipo de licencia de la aplici�n
            bibloAvisoErrores.propMsgError = "Error: " & codigoLicencia & Chr(10) & _
                                                "No se puede identificar el tipo de licencia de la aplicaci�n."
            bibloAvisoErrores.propDescMsgError = "La informaci�n que se posee sobre la licencia de la aplicaci�n es " & _
                                                    "inchoerente. Esto puede deberse a problemas en la base de datos, " & _
                                                    "originados por una instalaci�n incorrecta o por problemas con alg�n componente " & _
                                                    "del  hardware de su equipo. " & _
                                                    "La aplicaci�n no posee informaci�n suficiente para poder ejecutarse."
            bibloAvisoErrores.propContactarse = "Para solucionar el problema comun�quese con nosotros."
            bibloAvisoErrores.MostrarMensaje
        Case 517
            'la aplicaci�n se esta ejecutando en un disco duro distinto al que fue instalada.
            bibloAvisoErrores.propMsgError = "Error: " & codigoLicencia & Chr(10) & _
                                             "Versi�n registrada: incorrecta."
            bibloAvisoErrores.propDescMsgError = "La aplicaci�n no cumple con las condiciones necesarias para ejecutarse como versi�n registrada. " & _
                                                "La aplicaci�n se esta ejecutando en un disco duro distinto al que fue instalada y para el cual se " & _
                                                "adquiri� la licencia. Si esto no es as�, el error se debe a que ha ocurrido un fallo inesperado " & _
                                                "en las rutinas de seguridad."
            bibloAvisoErrores.propContactarse = "Para solucionar el problema comun�quese con nosotros."
            bibloAvisoErrores.MostrarMensaje
                                                
        Case 518
            'no coinciden los tipos
            bibloAvisoErrores.propMsgError = "Error: " & codigoLicencia & Chr(10) & _
                                             "No coinciden los tipos."
            bibloAvisoErrores.propDescMsgError = "La informaci�n en el archivo Id de la aplicaci�n es erronea. " & _
                                                "Posiblemente el mismo fue modificado, alterando sus valores originales. " & _
                                                "La aplicaci�n no puede obtener informaci�n correcta para poder ejecutarse."
            bibloAvisoErrores.propContactarse = "Para solucionar el problema comun�quese con nosotros."
            bibloAvisoErrores.MostrarMensaje
            
        Case 519
            'no existe el archivo
            bibloAvisoErrores.propMsgError = "Error: " & codigoLicencia & Chr(10) & _
                                             "No existe archivo Id de la aplicaci�n."
            bibloAvisoErrores.propDescMsgError = "El archivo Id de la aplicaci�n no se pudo localizar. " & _
                                                "El mismo fue cambiado de lugar dentro de su disco duro o eliminado. " & _
                                                "Si es la primera vez que ejecuta la aplicaci�n, entonces el problema se puede " & _
                                                "originar por una instalaci�n incorrecta."
            bibloAvisoErrores.propContactarse = "Para solucionar el problema comun�quese con nosotros."
            bibloAvisoErrores.MostrarMensaje

        Case 520
            'error en ejecuci�n de algunas de las rutinas de la biblioteca
            bibloAvisoErrores.propMsgError = "Error: " & codigoLicencia & Chr(10) & _
                                             "Error no determinado."
            bibloAvisoErrores.propDescMsgError = "Se produjo un error inesperado al ejecutar las rutinas de seguridad. " & _
                                                "Imposible continuar con la ejecuci�n de la aplicaci�n."
            bibloAvisoErrores.propContactarse = "Para solucionar el problema comun�quese con nosotros."
            bibloAvisoErrores.MostrarMensaje


    End Select
    'Destruyo las instancias creadas.
    Set bibloControlVersionDemo = Nothing
    Set bibloAvisoErrores = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Form1 = Nothing
End Sub
