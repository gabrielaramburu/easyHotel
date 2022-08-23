VERSION 5.00
Begin VB.Form frmAvisoErrorDatosMinimos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Imposible ejecutar opción."
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton botAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   5520
      TabIndex        =   0
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Image Image2 
      Height          =   210
      Left            =   240
      Picture         =   "frmAvisoErrorDatosMinimos.frx":0000
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   6420
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Detalle"
      Height          =   240
      Left            =   240
      TabIndex        =   3
      Top             =   1870
      Width           =   645
   End
   Begin VB.Label lblDetalle 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   1215
      Left            =   240
      TabIndex        =   2
      Top             =   2160
      Width           =   6495
   End
   Begin VB.Label lblError 
      Caption         =   "Label1"
      Height          =   1335
      Left            =   1080
      TabIndex        =   1
      Top             =   240
      Width           =   5655
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "frmAvisoErrorDatosMinimos.frx":4BBA
      Top             =   360
      Width           =   480
   End
End
Attribute VB_Name = "frmAvisoErrorDatosMinimos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Declaración de propiedades del formulario
Public propTipoDetalle As String    'determina la información a mostrar en la
                                    'ventana de detalle.

Private Sub botAceptar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.lblError = "Imposible de ejecutar esta opción." & Chr(10) & _
                "No existe datos mínimos necesarios para ejecutar esta opción del sistema. " & _
                "Si la aplicación ha sido recientemente instalada, usted " & _
                "aun no a configurado el perfil de la misma, o no ha realizado " & _
                "transacciones en las tablas de la base de datos."
    'muestro detalle
    Select Case propTipoDetalle
        Case "frmCargaReserva"
            Me.lblDetalle.Caption = "No existen habitaciones definidas."
            
        Case "frmBloquearHab"
            Me.lblDetalle.Caption = "No existen motivos de bloqueo definidos."
            
        Case "frmCambioSitu"
            'valido que exista situaciones de habitaciones
            Me.lblDetalle.Caption = "No existen situaciones de habitaciones definidas"
            
        Case "frmConsultaCompleta"
            'valido que existan habitaciones
            Me.lblDetalle.Caption = "No existen habitaciones definidas."
            
        Case "frmConsultaTitular"
            'valido que existan habitaciones
            Me.lblDetalle.Caption = "No existen habitaciones definidas."
        
        Case "frmCuadroHab"
            'valido que existan habitaciones
            Me.lblDetalle.Caption = "No existen habitaciones definidas."
        
        Case "frmIngExtras"
            'valido que existan artículos
            Me.lblDetalle.Caption = "No existen artículos en al base de datos."
          
        Case "frmListadoIngresos"
            'valido que existan tipos de habitaciones
            Me.lblDetalle.Caption = "No existen definidos tipos de habitaciones."
        
        Case "frmListadoEgresos"
            'valido que existan tipos de habitaciones
            Me.lblDetalle.Caption = "No existen definidos tipos de habitaciones."
            
        Case "frmVerDisponibilidad"
            'valido que existan tipos de habitaciones
            Me.lblDetalle.Caption = "No existen definidos tipos de habitaciones."

        Case "frmCierreDiario"
            'valido que existan habitaciones
            Me.lblDetalle.Caption = "No existen habitaciones definidas."
        
        Case "frmEstadoCuentas"
            'debe de existir por lo menos una cotización
            Me.lblDetalle.Caption = "No existe ningúna cotización establecida para la moneda dolar."
            
        Case "frmFacturacion"
            'debe de existir por lo menos una cotización
            Me.lblDetalle.Caption = "No existe ningúna cotización establecida para la moneda dolar."
            
        Case "frmConsultaCuentas"
            'debe de existir por lo menos una cotización
            Me.lblDetalle.Caption = "No existe ningúna cotización establecida para la moneda dolar."
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmAvisoErrorDatosMinimos = Nothing
End Sub
