VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConfiguroInicial"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private WithEvents dlgInicial As frmInicial
Attribute dlgInicial.VB_VarHelpID = -1
Private mCamino As String   'lugar donde genero archivo configuración
                            'posiblemente se pueda utilizar el metodo app
                            'pero no lo utilizo por no estar convenzido de
                            'que funcione correctamente en una dll.
Private mArchivo As String  'nombre archivo confg.
Event NotificoClientes(boton As Byte)

Public Sub MostrarPantallaConfigurar(NombreArch As String, _
                                    camino As String, _
                                    nombreApp As String)
    'Muestra pantalla de configuración la cual permite
    'seleccionar base de datos y resolución de la pantalla

    'Cargo variables
    mCamino = camino
    mArchivo = NombreArch
    
    'Etiqueta
    dlgInicial.Label1.Caption = _
    "La aplicación " & nombreApp & " se está ejecutando por primera vez. " & _
    "Para continuar debe de ingresar el nombre de " & _
    "la base de datos a utilizar, como así también la " & _
    "configuración de la pantalla que usted desee."
    
    'titulo formulario
    dlgInicial.Caption = dlgInicial.Caption & " " & nombreApp

    dlgInicial.Show 1
End Sub

Private Sub Class_Initialize()
    Set dlgInicial = New frmInicial
End Sub

Private Sub Class_Terminate()
    'Cuando se termina el objeto descargo el formulario
    'y  establesco su referencia a nothing para
    'evitar ocupar espacios de memoria huerfanos
    
    Unload dlgInicial
    Set dlgInicial = Nothing
End Sub

Private Sub dlgInicial_NotificoCliente(boton As Byte)
    Dim NroArch As Integer
    Dim BaseDeDatos As String
    Dim Resolucion As String
    
    If boton = 1 Then 'aceptar
        BaseDeDatos = dlgInicial.txtBaseDeDatos.Text
        Resolucion = dlgInicial.cboPantalla.Text
        
        'creo archivo de configuración
        NroArch = FreeFile
        Open mCamino & "\" & mArchivo For Append As NroArch
        'grabo camino de base de datos
        Print #NroArch, BaseDeDatos
        'grabo resolución pantalla
        Print #NroArch, Resolucion
        Close NroArch
    End If
    RaiseEvent NotificoClientes(boton)
End Sub
