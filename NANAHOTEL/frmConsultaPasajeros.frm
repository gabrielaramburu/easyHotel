VERSION 5.00
Begin VB.Form frmConsultaPasajeros 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ubicación de pasajeros"
   ClientHeight    =   6825
   ClientLeft      =   2550
   ClientTop       =   795
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton botSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   8160
      TabIndex        =   27
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Caption         =   "Datos del alojamiento"
      Height          =   1935
      Left            =   120
      TabIndex        =   20
      Top             =   4080
      Width           =   9255
      Begin VB.Label fhasta 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "fhasta"
         Height          =   375
         Left            =   3240
         TabIndex        =   26
         Top             =   840
         Width           =   3735
      End
      Begin VB.Label fdesde 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "fdesde"
         Height          =   375
         Left            =   3240
         TabIndex        =   25
         Top             =   360
         Width           =   3735
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Días alojados hasta el momento"
         Height          =   240
         Left            =   120
         TabIndex        =   24
         Top             =   1440
         Width           =   2895
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Fecha prevista de egreso"
         Height          =   240
         Left            =   120
         TabIndex        =   23
         Top             =   840
         Width           =   2310
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Ingresó al hotel"
         Height          =   240
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   1365
      End
      Begin VB.Label LTotDiasA 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LTotDiasA"
         Height          =   375
         Left            =   3240
         TabIndex        =   21
         Top             =   1440
         Width           =   975
      End
   End
   Begin Hotel_Nana.gaHOTELbarra gaHOTELbarra1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   6495
      Width           =   9510
      _ExtentX        =   16775
      _ExtentY        =   582
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos personales"
      Height          =   2415
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   9255
      Begin VB.Label lec 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lec"
         Height          =   375
         Left            =   5880
         TabIndex        =   18
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label Lfn 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Lfn"
         Height          =   375
         Left            =   1560
         TabIndex        =   17
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Estado civil"
         Height          =   240
         Left            =   4560
         TabIndex        =   16
         Top             =   1800
         Width           =   1035
      End
      Begin VB.Label fechanaccheck 
         AutoSize        =   -1  'True
         Caption         =   "Fecha nac."
         Height          =   240
         Left            =   120
         TabIndex        =   15
         Top             =   1800
         Width           =   990
      End
      Begin VB.Label Lpais 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Lpais"
         Height          =   375
         Left            =   5880
         TabIndex        =   14
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "País"
         Height          =   240
         Left            =   5160
         TabIndex        =   13
         Top             =   840
         Width           =   405
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Fax"
         Height          =   240
         Left            =   5160
         TabIndex        =   12
         Top             =   1320
         Width           =   330
      End
      Begin VB.Label Lfax 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Lfax"
         Height          =   375
         Left            =   5880
         TabIndex        =   11
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Ltel 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ltel"
         Height          =   375
         Left            =   1560
         TabIndex        =   10
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label Label7 
         Caption         =   "Teléfono"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Lloc 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Lloc"
         Height          =   375
         Left            =   1560
         TabIndex        =   8
         Top             =   840
         Width           =   3255
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Localidad"
         Height          =   240
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   900
      End
      Begin VB.Label Ldir 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ldir"
         Height          =   375
         Left            =   1560
         TabIndex        =   6
         Top             =   360
         Width           =   6255
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Dirección"
         Height          =   240
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pasajero"
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9255
      Begin VB.CommandButton botOtro 
         Caption         =   "&Otro "
         Height          =   375
         Left            =   7800
         TabIndex        =   3
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label lblHabitacion 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblHabitacion"
         Height          =   375
         Left            =   1920
         TabIndex        =   29
         Top             =   780
         Width           =   4335
      End
      Begin VB.Label lblPasajero 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblPasajero"
         Height          =   375
         Left            =   1920
         TabIndex        =   28
         Top             =   300
         Width           =   7095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Alojado en:"
         Height          =   240
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre Completo"
         Height          =   240
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1650
      End
   End
   Begin VB.Menu mnuFormulario 
      Caption         =   "&Formulario"
      Begin VB.Menu mnuFormularioAceptar 
         Caption         =   "Salir"
         Shortcut        =   {F12}
      End
      Begin VB.Menu div1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFormularioOtro 
         Caption         =   "Otro"
         Shortcut        =   {F9}
      End
   End
End
Attribute VB_Name = "frmConsultaPasajeros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private habpas As Long

Private Sub Form_Load()
    'Apariencia formulario
    mSubConfiguroFuentesControlesSistema Me
    
    inicializo
    muestro
End Sub

Private Sub botOtro_Click()
    Dim cli_aux As String
    
    inicializo
    'Muestro solo los clientes alojados
    cli_aux = mFunBusqueda(2)
    If Val(cli_aux) <> 0 Then
        cliente_a_ubicar = cli_aux
        muestro
    End If
End Sub

Private Sub muestro()
    If busco_clienteTF(cliente_a_ubicar) Then
        lblPasajero.Caption = tbCLIENTES("nombre_completo_titular")
        habpas = busco_pasajero(cliente_a_ubicar)
        If busco_habitaTF(habpas) Then
            If busco_tipo_habTF(tbHABITACIONES("tipohab")) Then
                lblHabitacion.Caption = "Suite " & habpas & " " & tbTIPO_HABITACIONES("descripcion")
                muestro_datos_extras
            End If
        End If
    End If
End Sub

Private Sub muestro_datos_extras()
    LTotDiasA.Caption = m_FechaSis - tbCHECKIN("fcheckdes")
    'muestro fechas
    fdesde.Caption = Format(tbCHECKIN("fcheckdes"), "dddd, d mmm yyyy")
    fhasta.Caption = Format(tbCHECKIN("fcheckhas"), "dddd, d mmm yyyy")
    cargo_cli
End Sub

Private Sub cargo_cli()
    Ldir.Caption = tbCLIENTES("direccion_titular")
    Lloc.Caption = tbCLIENTES("localidad_titular")
    Ltel.Caption = tbCLIENTES("tel_titular")
    If Not IsNull(tbCLIENTES("fecha_nac_titular")) Then Lfn.Caption = tbCLIENTES("fecha_nac_titular")
    Lpais.Caption = mFunBuscoDescPais(CInt(tbCLIENTES("pais_reside_titular")))
    Lfax.Caption = tbCLIENTES("fax_titular")
    lec.Caption = mFunBuscoDescEstadoCivil(CInt(tbCLIENTES("estado_civil_titular")))
End Sub

Private Sub inicializo()
    mSubEtiquetasInicializo Me
End Sub

Private Sub botSalir_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Descargo formulario de memoria
    Set frmConsultaPasajeros = Nothing
End Sub

Private Sub mnuFormularioAceptar_Click()
    'Esta opción equivale a presionar el botón aceptar o la tecla F12
    botSalir_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub mnuFormularioOtro_Click()
    'Equivale a presionar la tecla F9 o el botón de otro cliente
    botOtro_Click
End Sub

'*******************************************************************
'*
'*  Asistencia a usuario
'*
'*******************************************************************

Private Sub botOtro_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 22
End Sub

Private Sub botSalir_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 1, 1
End Sub

Private Sub botOtro_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub botSalir_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub



