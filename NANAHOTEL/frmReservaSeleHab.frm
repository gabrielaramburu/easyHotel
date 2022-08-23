VERSION 5.00
Begin VB.Form frmReservaSeleHab 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Selección de Habitación"
   ClientHeight    =   4890
   ClientLeft      =   4350
   ClientTop       =   2970
   ClientWidth     =   3975
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   3975
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3600
      Top             =   120
   End
   Begin VB.Frame Frame1 
      Caption         =   "&Habitaciones disponibles  "
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3735
      Begin VB.CommandButton botConfirmar 
         Height          =   375
         Left            =   2280
         Picture         =   "frmReservaSeleHab.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         Tag             =   "Aceptar"
         Top             =   3840
         Width           =   1215
      End
      Begin VB.CommandButton botNoAsignar 
         Caption         =   "&No Asignar"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   3840
         Width           =   1215
      End
      Begin VB.CommandButton botCerrar 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   2280
         Picture         =   "frmReservaSeleHab.frx":08B6
         Style           =   1  'Graphical
         TabIndex        =   3
         Tag             =   "Cancelar"
         Top             =   4320
         Width           =   1215
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   3060
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   3255
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
      Begin VB.Menu div1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFormularioNoAsignar 
         Caption         =   "No asignar"
      End
   End
End
Attribute VB_Name = "frmReservaSeleHab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    'Apariencia formulario
    mSubConfiguroFuentesControlesSistema Me
    
    configuracion_apariencia
    'Obtengo el título del formulario, mostrando el tipo de habitaciónes
    Select Case tipo_accion_SeleccionHab
        Case 1  'reserva
            
        Case 2  'checkin
            Frame1.Caption = "&Habitaciones disponibles y limpias."
            'cuando es un checkin siempre se tiene que asignar una habitación
            'por lo tanto no permito el boton de no asignar
            Me.botNoAsignar.Enabled = False
            Me.mnuFormularioNoAsignar.Enabled = False
            
        Case 3  'walkin libre
            Frame1.Caption = "&Habitaciones disponibles y limpias."
            'cuando es un walkinL siempre se tiene que asignar una habitación
            'por lo tanto no permito el boton de no asignar
            Me.botNoAsignar.Enabled = False
            Me.mnuFormularioNoAsignar.Enabled = False
            
        Case 4  'walkin ocupada
            Frame1.Caption = "&Habitaciones ocupadas " & frmCheck_in.cboTipo_habitacion.Text
            'cuando es un walkinO siempre se tiene que asignar una habitación
            'por lo tanto no permito el boton de no asignar
            Me.botNoAsignar.Enabled = False
            Me.mnuFormularioNoAsignar.Enabled = False
    End Select
    Timer1.Enabled = False

    If tipo_accion_SeleccionHab = 4 Then    'walkinO
        cargo_habitaciones_ocupadas
    Else
        If tipo_accion_SeleccionHab = 3 Or _
        tipo_accion_SeleccionHab = 2 Then 'walkinL o Checkin
            'para el walkin libre y el checkin las habitaciones además de disponibles
            'tiene que estar limpias.
            cargo_habitaciones True
        Else
            '1 reserva
            cargo_habitaciones False
        End If
    End If
End Sub

Private Sub botCerrar_Click()
    cancelo_seleccion_habitaciones = True
    Unload Me
End Sub

Private Sub botConfirmar_Click()
    asigno_hab
End Sub

Private Sub asigno_hab()
    cancelo_seleccion_habitaciones = False
    Select Case tipo_accion_SeleccionHab
        Case 1
            frmCargaReserva.txtHab.Text = corto_palabras(List1.List(List1.ListIndex))
        Case 2
            frmCheck_in.txtHabCheck.Text = corto_palabras(List1.List(List1.ListIndex))
        Case 3
            frmCheck_in.txtHabWalk.Text = corto_palabras(List1.List(List1.ListIndex))
        Case 4
            frmCheck_in.txtHabWalk.Text = corto_palabras(List1.List(List1.ListIndex))
    End Select
    Unload Me
End Sub

Private Sub botNoAsignar_Click()
    cancelo_seleccion_habitaciones = False
    Select Case tipo_accion_SeleccionHab
        Case 1
            frmCargaReserva.txtHab.Text = 0
            Unload Me
    End Select
End Sub

Private Sub cargo_habitaciones(habLimpia As Boolean)
    '--------------------------------------------------------------------------------
    'Parámetros.
    '   Entrada [habLimpias]
    '               Este parámetro determina si hay que validar que las
    '               habitaciones además de disponibles, tengan situación limpia.
    '               True : cuando estoy trabajando con un checkin o con un walkin libre.
    '               False: cuando estoy trabajando con una reserva.
    '--------------------------------------------------------------------------------
    'Cargo habitaciones libre
    Dim tipo As Integer
    Dim fd As Date, fh As Date
    
    List1.Clear
    tipo = obtengo_tipo_hab

    fd = obtengo_fecha_desde
    fh = obtengo_fecha_hasta
    tbHABITACIONES.MoveFirst
    tbHABITACIONES.Index = "i_tipohab"
    tbHABITACIONES.Seek ">=", tipo
    If Not tbHABITACIONES.NoMatch Then
        Do While Not tbHABITACIONES.EOF
            If tbHABITACIONES("tipohab") = tipo Then
                'verifico no esté reservada
                If Not habitacion_reservada(tbHABITACIONES("nrohab"), fd, fh) Then
                    'verifico no este ocupada
                    If Not habitacion_ocupada(tbHABITACIONES("nrohab"), fd) Then
                        'verifico no este bloqueada
                        If Not habitacion_bloqueada(tbHABITACIONES("nrohab"), fd, fh) Then
                            'la habitación está disponible
                            'verifico si tengo que validar situación
                            If habLimpia Then
                                'la habitación además de disponible tiene qu estar limpia
                                If tbHABITACIONES("situacionhab") = 1 Then  '1 = limpia
                                    'muestro habitación
                                    List1.AddItem tbHABITACIONES("nrohab") & " " & obtengo_tipo_hab_desc
                                End If
                            Else
                                'no tengo que validar situación
                                List1.AddItem tbHABITACIONES("nrohab") & " " & obtengo_tipo_hab_desc
                            End If
                        End If
                    End If
                End If
            Else
                Exit Do
            End If
            tbHABITACIONES.MoveNext
        Loop
        Timer1.Enabled = True
    End If
End Sub

Private Sub cargo_habitaciones_ocupadas()
    'Utilizado para realizar el Walkin Ocupadas.
    Dim tipo As Integer
    
    List1.Clear
    tipo = obtengo_tipo_hab
    
    tbHABITACIONES.MoveFirst
    tbHABITACIONES.Index = "i_tipohab"
    tbHABITACIONES.Seek ">=", tipo
    If Not tbHABITACIONES.NoMatch Then
        'recorro todas las habitaciones del tipo determinado
        Do While Not tbHABITACIONES.EOF
            If tbHABITACIONES("tipohab") = tipo Then
                'determino si la habitación esta ocupada actualmente
                If busco_habita_checkin(tbHABITACIONES("nrohab")) Then
                    'determino si el período de ocupación es válido
                    If mFunDeterminoOcupacionValida(tbHABITACIONES("nrohab")) Then
                        List1.AddItem tbHABITACIONES("nrohab") & " " & obtengo_tipo_hab_desc
                    End If
                End If
            Else
                Exit Do
            End If
            tbHABITACIONES.MoveNext
        Loop
        Timer1.Enabled = True
    End If
End Sub

'*-----------------
'1= desde reserva
'2= desde checkin
'3=desde walkinL
'4=desde walkinO
'*------------------

Private Function obtengo_tipo_hab()
    Select Case tipo_accion_SeleccionHab
        Case 1
            obtengo_tipo_hab = frmCargaReserva.cboTipo_habitacion.ItemData(frmCargaReserva.cboTipo_habitacion.ListIndex)
        Case 2
            obtengo_tipo_hab = frmCheck_in.gHabitaciones.TextMatrix(frmCheck_in.gHabitaciones.Row, 2)
        Case 3
            obtengo_tipo_hab = frmCheck_in.cboTipo_habitacion.ItemData(frmCheck_in.cboTipo_habitacion.ListIndex)
        Case 4
            obtengo_tipo_hab = frmCheck_in.cboTipo_habitacion.ItemData(frmCheck_in.cboTipo_habitacion.ListIndex)
    End Select
End Function

Private Function obtengo_fecha_desde()
    Select Case tipo_accion_SeleccionHab
        Case 1
            obtengo_fecha_desde = frmCargaReserva.fechaing.Text
        Case 2
            obtengo_fecha_desde = frmCheck_in.fechaingreso.Text
        Case 3
            obtengo_fecha_desde = m_FechaSis
    End Select
End Function

Private Function obtengo_fecha_hasta()
    Select Case tipo_accion_SeleccionHab
        Case 1
            obtengo_fecha_hasta = frmCargaReserva.fechaegr.Text
        Case 2
            obtengo_fecha_hasta = frmCheck_in.fechaegreso.Text
        Case 3
            obtengo_fecha_hasta = frmCheck_in.fEgreso.Text
    End Select
End Function
    
Private Function obtengo_tipo_hab_desc()
    'Devuelve un string indicando el tipo de habitación
    obtengo_tipo_hab_desc = "Suite " & mFun_BuscoDescriTipoHab(tbHABITACIONES("tipohab"))
End Function

Private Sub Form_Unload(Cancel As Integer)
    'Descargo formulario de memoria
    Set frmReservaSeleHab = Nothing
End Sub

Private Sub List1_DblClick()
    asigno_hab
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    If List1.ListCount < 1 Then
        If tipo_accion_SeleccionHab = 4 Then
            'no hay habitaciones ocupadas hoy.
            mSubMensaje 4, 36
        Else
            'no hay habitaciones disponibles en el período seleccionado
            mSubMensaje 4, 37
        End If
        cancelo_seleccion_habitaciones = True
        Me.botConfirmar.Enabled = False
    Else
        List1.ListIndex = 0
    End If
End Sub

Private Sub configuracion_apariencia()
    'Determina la apariencia del los elemento configurables del formulario
    If tipo_accion_SeleccionHab = 4 Then    'ocupadas
        Me.List1.ForeColor = mSisColor_12SeleccionHabOcupada
    Else
        Me.List1.ForeColor = mSisColor_11SeleccionHabLibre
    End If
End Sub

Private Sub mnuFormularioAceptar_Click()
    'Equivale a preionar la tecla F12 o el boton de aceptar
    botConfirmar_Click
End Sub

Private Sub mnuFormularioCancelar_Click()
    'Equivale a presionar la tecla Esc o el boton de cancelar
    botCerrar_Click
End Sub

Private Sub mnuFormularioNoAsignar_Click()
    'Equivale a presionar el boton de borrar
    botNoAsignar_Click
End Sub

