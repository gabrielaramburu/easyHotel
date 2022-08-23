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
   Begin VB.CommandButton botObtenerDueño 
      Caption         =   "Obtengo Dueño"
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
      Caption         =   "Actualizo ultimo día"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2640
      Width           =   2175
   End
   Begin VB.CommandButton botCantDiasPermitidos 
      Caption         =   "Cant. días permitidos"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   2175
   End
   Begin VB.CommandButton botCantDiasPeriodo 
      Caption         =   "Cantidad días periodo"
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
   Begin VB.CommandButton botFinEvaluación 
      Caption         =   "&Fin Evaluación"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   2175
   End
   Begin VB.CommandButton botAvisoVersionDemo 
      Caption         =   "Aviso &versión demo"
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
        'obtengo código de aplicación
        idapli = bibloControlVersionDemo.funObtengoIdAplicacion(App.Path & "\Aplicacion.id.txt")
        If bibloControlVersionDemo.funActualizarUltimoDiaEjecuciónVD(tbLICENCIA, idapli, Date) Then
            MsgBox "Se actualizó la fecha del último día de ejecución"
        Else
            MsgBox "Error al actualizar fecha del último día de ejecución", vbExclamation
        End If
    Set bibloControlVersionDemo = Nothing
End Sub

Private Sub botCantDiasPeriodo_Click()
    Dim idapli As Long
    Set bibloControlVersionDemo = New ControlVersionDemo
        'obtengo código de aplicación
        idapli = bibloControlVersionDemo.funObtengoIdAplicacion(App.Path & "\Aplicacion.id.txt")
        MsgBox "Cantidad de días de uso: " & bibloControlVersionDemo.funObtenerCantDiasUtilizadosVD(tbLICENCIA, Date, idapli)
    Set bibloControlVersionDemo = Nothing
End Sub

Private Sub botCantDiasPermitidos_Click()
    Dim idapli As Long
    Set bibloControlVersionDemo = New ControlVersionDemo
        'obtengo código de aplicación
        idapli = bibloControlVersionDemo.funObtengoIdAplicacion(App.Path & "\Aplicacion.id.txt")
        MsgBox "Cantidad de días permitidos: " & bibloControlVersionDemo.FunObtenerCantDiasAutorizadosVD(tbLICENCIA, idapli)
    Set bibloControlVersionDemo = Nothing
End Sub


Private Sub botObtenerDueño_Click()
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
    'Establesco propiedades del forumlario de versión demo
    
    avisoDemo.AvisoVersionDemoPropTituloForm = "Versión de evaluación(no registrada)."
    avisoDemo.AvisoVersionDemoPropNomAplicacion = "Hotel2000"
    avisoDemo.AvisoVersionDemoPropSistemaAplicacion = "Para Window 9x/2000/XT/NT"
    avisoDemo.AvisoVersionDemoPropVersionAplicacion = "Version de evaluación 1.0"
    avisoDemo.AvisoVersionDemoPropDiasDemos = "60"
    avisoDemo.AvisoVersionDemoPropPeriodoDeUso = "Días utilizado: " & "1" & " de su período de " & avisoDemo.AvisoVersionDemoPropDiasDemos & " días."
    avisoDemo.AvisoVersionDemoPropDerechos = "Copyright(c) 2000-2002 by Chupacabras Sowftware" & Chr(10) & _
                                        "All Rights Reserved." & Chr(10) & _
                                        "Maldonado - Uruguay" & Chr(10) & _
                                        "www.chupacabrasventanita.com"
    avisoDemo.MostrarAvisoVersionDemo
    Set avisoDemo = Nothing
End Sub

Private Sub botFinEvaluación_Click()
    Set avisoFin = New AvisoFinPeriodoDemo
    'Establesco todas las propiedades del objeto creado AvisoFinPeriodoDemo
    avisoFin.AvisoFinPeriodoDemoPropNomAplicacion = "Hotel2000"
    avisoFin.AvisoFinPeriodoDemoPropSistemaAplicacion = "Para Window 9x/2000/XT/NT"
    avisoFin.AvisoFinPeriodoDemoPropVersionAplicacion = "Version de evaluación 1.0"
    avisoFin.AvisoFinPeriodoDemoPropPeriodoTerminado = "Días utilizado: " & "30" & " de su período de 30 días."
    avisoFin.AvisoFinPeriodoDemoPropTituloForm = "Fin del período de evaluación."
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
    
    
    'Creo nuevas instancias de las clases que contienen los componentes de código
    'que se utilizan para validar la aplicación.
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
            'Es una versión registrada
        Case 514
            botFinEvaluación_Click
        Case 513
            'No se encontró un registro con la clave primaria igual al número de identificación
            'de la aplicación que estoy ejecutando.
            
            'obtengo código de aplicación
            idapli = bibloControlVersionDemo.funObtengoIdAplicacion(App.Path & "\Aplicacion.id.txt")
            'ejecuto inicialización de la aplicación
            Set funInicializarAplicacion = New InicializarAplicacion
                resIniApli = funInicializarAplicacion.funInicializarAplicacion(tbLICENCIA, _
                                                                            idapli, _
                                                                            Date)
                Select Case resIniApli
                    Case 0
                        'después de inicializar la aplicación muestro aviso de versión demo.
                         botAvisoVersionDemo_Click
                    Case Else
                        'problemas al inicalizar la aplicación
                        bibloAvisoErrores.propMsgError = "Error: " & resIniApli & Chr(10) & _
                                                            "Problemas al inicializar la aplicación."
                        bibloAvisoErrores.propDescMsgError = "El proceso de inicialización de la aplicación, " & _
                                                            "no se pudo realizar correctamente. Si es la primera vez que " & _
                                                            "ejecuta esta aplicación, es posible que el problema se origine " & _
                                                            "por una instalación incorrecta. Si por el contrario ustes ya la ha ejecutado " & _
                                                            "con exito anteriormente, es muy problable que halla problemas con la base " & _
                                                            "de datos de la aplicación o con algún componente del hardware de su equipo."
                        bibloAvisoErrores.propContactarse = "Para solucionar el problema comuníquese con nosotros."
                        bibloAvisoErrores.MostrarMensaje
                End Select
            Set funInicializarAplicacion = Nothing
        Case 515
            'la fecha del sistema fue retrocedida.
            bibloAvisoErrores.propMsgError = "Error: " & codigoLicencia & Chr(10) & _
                                                "Problemas con la fecha del sistema."
            bibloAvisoErrores.propDescMsgError = "La fecha del sistema es menor a la fecha en la cual se " & _
                                                "ejecutó la aplicación por última vez. Como usted está ejecutando, " & _
                                                "una versión de evaluación no registrada, esta distorsión en la fecha, " & _
                                                "impide ejecutar la aplicación."
            bibloAvisoErrores.propContactarse = "Para solucionar el problema comuníquese con nosotros."
            bibloAvisoErrores.MostrarMensaje

        Case 516
            'no se puede identificar el tipo de licencia de la aplición
            bibloAvisoErrores.propMsgError = "Error: " & codigoLicencia & Chr(10) & _
                                                "No se puede identificar el tipo de licencia de la aplicación."
            bibloAvisoErrores.propDescMsgError = "La información que se posee sobre la licencia de la aplicación es " & _
                                                    "inchoerente. Esto puede deberse a problemas en la base de datos, " & _
                                                    "originados por una instalación incorrecta o por problemas con algún componente " & _
                                                    "del  hardware de su equipo. " & _
                                                    "La aplicación no posee información suficiente para poder ejecutarse."
            bibloAvisoErrores.propContactarse = "Para solucionar el problema comuníquese con nosotros."
            bibloAvisoErrores.MostrarMensaje
        Case 517
            'la aplicación se esta ejecutando en un disco duro distinto al que fue instalada.
            bibloAvisoErrores.propMsgError = "Error: " & codigoLicencia & Chr(10) & _
                                             "Versión registrada: incorrecta."
            bibloAvisoErrores.propDescMsgError = "La aplicación no cumple con las condiciones necesarias para ejecutarse como versión registrada. " & _
                                                "La aplicación se esta ejecutando en un disco duro distinto al que fue instalada y para el cual se " & _
                                                "adquirió la licencia. Si esto no es así, el error se debe a que ha ocurrido un fallo inesperado " & _
                                                "en las rutinas de seguridad."
            bibloAvisoErrores.propContactarse = "Para solucionar el problema comuníquese con nosotros."
            bibloAvisoErrores.MostrarMensaje
                                                
        Case 518
            'no coinciden los tipos
            bibloAvisoErrores.propMsgError = "Error: " & codigoLicencia & Chr(10) & _
                                             "No coinciden los tipos."
            bibloAvisoErrores.propDescMsgError = "La información en el archivo Id de la aplicación es erronea. " & _
                                                "Posiblemente el mismo fue modificado, alterando sus valores originales. " & _
                                                "La aplicación no puede obtener información correcta para poder ejecutarse."
            bibloAvisoErrores.propContactarse = "Para solucionar el problema comuníquese con nosotros."
            bibloAvisoErrores.MostrarMensaje
            
        Case 519
            'no existe el archivo
            bibloAvisoErrores.propMsgError = "Error: " & codigoLicencia & Chr(10) & _
                                             "No existe archivo Id de la aplicación."
            bibloAvisoErrores.propDescMsgError = "El archivo Id de la aplicación no se pudo localizar. " & _
                                                "El mismo fue cambiado de lugar dentro de su disco duro o eliminado. " & _
                                                "Si es la primera vez que ejecuta la aplicación, entonces el problema se puede " & _
                                                "originar por una instalación incorrecta."
            bibloAvisoErrores.propContactarse = "Para solucionar el problema comuníquese con nosotros."
            bibloAvisoErrores.MostrarMensaje

        Case 520
            'error en ejecución de algunas de las rutinas de la biblioteca
            bibloAvisoErrores.propMsgError = "Error: " & codigoLicencia & Chr(10) & _
                                             "Error no determinado."
            bibloAvisoErrores.propDescMsgError = "Se produjo un error inesperado al ejecutar las rutinas de seguridad. " & _
                                                "Imposible continuar con la ejecución de la aplicación."
            bibloAvisoErrores.propContactarse = "Para solucionar el problema comuníquese con nosotros."
            bibloAvisoErrores.MostrarMensaje


    End Select
    'Destruyo las instancias creadas.
    Set bibloControlVersionDemo = Nothing
    Set bibloAvisoErrores = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Form1 = Nothing
End Sub
