VERSION 5.00
Begin VB.Form frmModificacionReserva 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Este formulario se utiliza para: modificacion, anulacion y consulta reserva y check-in"
   ClientHeight    =   5610
   ClientLeft      =   2130
   ClientTop       =   3330
   ClientWidth     =   4920
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   4920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Hotel_Nana.gaHOTELbarra gaHOTELbarra1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5280
      Width           =   4920
      _ExtentX        =   8678
      _ExtentY        =   582
   End
   Begin VB.Frame Frame1 
      Caption         =   "&Número de reserva"
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.TextBox txtNroReservaAnio 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   240
         MaxLength       =   4
         TabIndex        =   2
         Top             =   480
         Width           =   855
      End
      Begin VB.CommandButton Salir 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   3360
         Picture         =   "frmModificacionReserva.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Tag             =   "Cancelar"
         Top             =   4680
         Width           =   1215
      End
      Begin VB.CommandButton botCheckin 
         Height          =   375
         Left            =   2040
         Picture         =   "frmModificacionReserva.frx":08C2
         Style           =   1  'Graphical
         TabIndex        =   4
         Tag             =   "Aceptar"
         Top             =   4680
         Width           =   1215
      End
      Begin VB.CommandButton BusqTitular 
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2640
         TabIndex        =   3
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox txtNroReserva 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1320
         MaxLength       =   5
         TabIndex        =   1
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Menu mnuFormulario 
      Caption         =   "&Formulario"
      Begin VB.Menu mnuFormularioAceptar 
         Caption         =   "Aceptar"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuFormularioCancelar 
         Caption         =   "Cancelar"
      End
   End
   Begin VB.Menu mnuBuscar 
      Caption         =   "&Buscar..."
      Begin VB.Menu mnuBuscarReservas 
         Caption         =   "Reservas..."
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmModificacionReserva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ayuda As Boolean

Private Sub Form_Load()
    'Apariencia formulario
    mSubConfiguroFuentesControlesSistema Me
    
    configuracion_apariencia
    ayuda = False
    'es necesario para determinar si estoy consultado una reserva anulada
    consulta_reserva_anulada = False
    inicializo_numero_reserva
    'asigno titulo al formulario
    Select Case tipo_accion_reserva
        Case "MODIFICAR"
            frmModificacionReserva.Caption = "Modificación de Reserva"
            tipo_accion_busqueda_reserva = 2
        Case "ANULAR"
            frmModificacionReserva.Caption = "Anulación de Reserva"
            tipo_accion_busqueda_reserva = 2
        Case "CONSULTAR"
            frmModificacionReserva.Caption = "Consultar Reserva"
            tipo_accion_busqueda_reserva = 1
        Case "Check-in"
            tipo_accion_busqueda_reserva = 3
            tipo_accion_checkin = 1
            frmModificacionReserva.Caption = "Check-In"
    End Select
End Sub

Private Sub botCheckin_Click()
    'busco reserva
    nro_reserva = (Val(txtNroReservaAnio.Text) * 100000) + Val(txtNroReserva)
    If reserva_activa Then
        'Por las dudas descargo todos los formularios involucrados
        'tengo que revisar teoricamente si esto esta bien. 90% confirmado que si
        'Tengo un lío grande con esto. Voy a tener que estudiarlo bien.
        Set frmCargaReserva = Nothing
        Set frmCheck_in = Nothing
        Set frmReservaSeleHab = Nothing
        Set frmTitularesHabitacion = Nothing
        
        frmCargaReserva.Show 1
    End If
    inicializo_numero_reserva
End Sub

Private Function reserva_activa()
    reserva_activa = True
    Select Case tipo_accion_busqueda_reserva
        Case 1  'consultar
            If Not busco_reservaTF(nro_reserva) Then    'no existe en reservas
                If Not busco_reserva_anuladaTF(nro_reserva) Then 'existe en anuladas
                    reserva_activa = False
                    'reserva inexistente
                    mSubMensaje 4, 18
                Else
                    consulta_reserva_anulada = True 'cosulto una reserva anulada
                End If
            End If
            
        Case 2  'Modificar Anular
            If busco_reservaTF(nro_reserva) Then    'si existe
                'controlo que la reserva este activa
                If tbRESERVAS("fechaing") >= m_FechaSis Then
                    'Si ingresa hoy
                    If tbRESERVAS("fechaing") = m_FechaSis Then
                        'Valido que ya no hallan ingresado todas
                        'o alguna de la habitaciones.
                        If busco_reservaCheckinTF(nro_reserva) Then
                            'La reserva ingresa hoy, pero ya ingresó al hotel
                            'alguna de las habitaciones
                            mSubMensaje 4, 19
                            reserva_activa = False
                        Else
                        End If
                    End If
                Else
                    reserva_activa = False
                    'la reserva esta inactiva.
                    mSubMensaje 4, 20
                End If
            Else
                reserva_activa = False
                If busco_reserva_anuladaTF(nro_reserva) Then 'existe en anuladas
                    'reserva anulada
                    mSubMensaje 4, 21
                Else
                    'reserva inexistente
                    mSubMensaje 4, 22
                End If
            End If
            
        Case 3  'checkin
            If busco_reservaTF(nro_reserva) Then    'si existe
                'controlo que la reserva ingrese hoy
                If tbRESERVAS("fechaing") <> m_FechaSis Then
                    reserva_activa = False
                    'la reserva esta inactiva
                    mSubMensaje 4, 20
                End If
            Else
                reserva_activa = False
                If busco_reserva_anuladaTF(nro_reserva) Then 'existe en anuladas
                    'reserva anulada
                    mSubMensaje 4, 21
                Else
                    'reserva inexistente
                    mSubMensaje 4, 22
                End If
            End If
    End Select
End Function

Private Sub BusqTitular_Click()
    Dim nrores_aux As String
    ayuda = True
    nrores_aux = mFunBuscarReserva(tipo_accion_busqueda_reserva)
    If nrores_aux <> "" Then
        txtNroReservaAnio.Text = Mid(nrores_aux, 1, 4)
        txtNroReserva.Text = Val(Mid(nrores_aux, 5, 8))
        Me.botCheckin = True
    Else
        inicializo_numero_reserva
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Descargo formulario de memoria
    Set frmModificacionReserva = Nothing
End Sub

Private Sub txtNroReserva_KeyPress(KeyAscii As Integer)
    ValidoNum KeyAscii, False, False
End Sub

Private Sub txtNroReservaAnio_KeyPress(KeyAscii As Integer)
    CapturoEnter KeyAscii
    ValidoNum KeyAscii, False, False
End Sub

Private Sub inicializo_numero_reserva()
    txtNroReservaAnio.Text = Year(m_FechaSis)
    txtNroReserva.Text = ""
End Sub

Private Sub configuracion_apariencia()
    'Determina la apariencia del los elemento configurables del formulario
End Sub

Private Sub salir_Click()
    Unload Me
End Sub

Private Sub mnuFormularioAceptar_Click()
    'Equivale a presionar el boton aceptar o la tecl F12
    botCheckin_Click
End Sub

Private Sub mnuFormularioCancelar_Click()
    'Equivale a presionar el boton Cancelar o la tecla Esc
    salir_Click
End Sub

Private Sub mnuBuscarReservas_Click()
    'Equivale a presionar el boton de ayuda o la tecla F1
    BusqTitular_Click
End Sub

