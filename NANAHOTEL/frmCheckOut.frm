VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCheck_Out 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Check-Out"
   ClientHeight    =   6825
   ClientLeft      =   2160
   ClientTop       =   1620
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Hotel_Nana.gaHOTELtitular gaHOTELtitular1 
      Height          =   1335
      Left            =   120
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   2355
      BackColor       =   -2147483633
   End
   Begin VB.Frame Frame2 
      Caption         =   "&Pasajeros alojados en la habitación"
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   9255
      Begin VB.CommandButton botCancelar 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   7920
         Picture         =   "frmCheckOut.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Tag             =   "Cancelar"
         Top             =   4440
         Width           =   1215
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   405
         Left            =   3840
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   2  'Snapshot
         RecordSource    =   ""
         Top             =   3960
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.CommandButton botConfirmar 
         Enabled         =   0   'False
         Height          =   375
         Left            =   6600
         Picture         =   "frmCheckOut.frx":08C2
         Style           =   1  'Graphical
         TabIndex        =   3
         Tag             =   "Aceptar"
         Top             =   4440
         Width           =   1215
      End
      Begin VB.CommandButton botTodos 
         Caption         =   "&Todos "
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   4440
         Width           =   1215
      End
      Begin MSFlexGridLib.MSFlexGrid dbgrid1 
         Bindings        =   "frmCheckOut.frx":1178
         Height          =   2655
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   4683
         _Version        =   393216
         Cols            =   3
         AllowBigSelection=   0   'False
         FocusRect       =   2
         HighLight       =   0
         ScrollBars      =   2
         SelectionMode   =   1
         FormatString    =   $"frmCheckOut.frx":1188
      End
   End
   Begin Hotel_Nana.gaHOTELbarra gaHOTELbarra1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   6495
      Width           =   9510
      _ExtentX        =   16775
      _ExtentY        =   582
   End
   Begin VB.Menu mnuformulario 
      Caption         =   "&Formulario"
      Begin VB.Menu mnuFormularioAceptar 
         Caption         =   "Aceptar"
         Enabled         =   0   'False
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuFormularioCancelar 
         Caption         =   "Cancelar"
      End
   End
End
Attribute VB_Name = "frmCheck_Out"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public hab_cuenta As Long
Private titular_unica As Long
Private titular_extra As Long
Private titular_aloja As Long

Private Sub botConfirmar_Click()
        Dim i As Integer
    Dim tipo_err As Byte
    Dim titular_ok As Boolean
    Dim marcado_incompleto As Boolean
    Dim inicializar_habitacion As Boolean
            
    'obtengo titulares de la habitación
    obtengo_titular
    
    DBGrid1.col = 1
    titular_ok = True
    
    'si true ->estan todos marcados
    'si false->hay alguno desmarcado
    marcado_incompleto = verifico_marcas2
    
    'Si la habitación queda vacía es necesatio inicializarla
    'para que quede lista para un nuevo ingreso
    inicializar_habitacion = False
    i = 1
    'Recorro la grilla y proceso cada pasajero seleccionado.
    Do While i < DBGrid1.Rows
        DBGrid1.col = 1
        DBGrid1.Row = i
        'si esta marcado para check-out
        If DBGrid1.CellBackColor = mSisColor_15FilaSeleccionada Then
            DBGrid1.col = 2
            '1) verifico si es titular de alguna habitacion
            'puede ser de la que actualmente se desea realizar el checkout o de
            'cualquier otra
            If titular_habitaciones(Val(DBGrid1.Text)) = True Then
                'si es titular de la habitación del checkout (alojado dentro)
                If Val(DBGrid1.Text) = titular_unica Or _
                Val(DBGrid1.Text) = titular_extra Or _
                Val(DBGrid1.Text) = titular_aloja Then
                    'si todos los pasajeros se van permito realizar el checkout
                    If marcado_incompleto = True Then
                        'la habitación queda libre y tengo que inicializarla
                        inicializar_habitacion = True
                    Else
                        'No puedo dejar a la habitación sin titular
                        titular_ok = False
                        tipo_err = 1
                        Exit Do 'ya no sigo procesando los demas pasajeros
                    End If
                Else
                'si es titular de otra habitación no lo puedo dejar ir
                'ya que sino la otra habitación quedaría sin titular.
                    titular_ok = False
                    tipo_err = 3
                    Exit Do
                End If
            End If
            
            'verifico que no tenga gastos pendientes de facturación
            'no es necesario que el pasajero sea titular para verificar sus gastos.
            'ver documentación.
        
            If verifico_gastos Then
                titular_ok = False
                tipo_err = 2
                Exit Do
            End If
        End If
        i = i + 1
    Loop
    
    If titular_ok Then
        'pido confirmación al usuario para realizar el checkout
        If mFunMensaje(4, 80) Then
            'recorro nuevamente para eliminar, ya que esta todo ok
            i = 1
            Do While i < DBGrid1.Rows
                DBGrid1.col = 1
                DBGrid1.Row = i
                'si esta marcado para check-out
                If DBGrid1.CellBackColor = mSisColor_15FilaSeleccionada Then    'si esta marcado
                    DBGrid1.col = 2
                    '1) borro  tbCHECKIN y gravo  tbCHECKOUT
                    borro_checkin Val(DBGrid1.Text)
                End If
                i = i + 1
            Loop
                        
            '3) inicializo habitación para que quede disponible nuevamente.
            'solo en el caso de que la habitación quede libre
            If inicializar_habitacion = True Then
                inicializo_habitacion hab_cuenta
                cambio_situacion hab_cuenta, 2  'sucia
            End If
            'grabo bitacora
            GraboBitacora "Hab. " & hab_cuenta
            'aviso de finalización de la operación
            mSubMensaje 4, 84
            Unload Me
            frmIngHabitacion.Show 1
        End If
    Else
        errores tipo_err
    End If
End Sub

Private Sub Form_Load()
    'Apariencia formulario
    mSubConfiguroFuentesControlesSistema Me
    
    'obtengo habitacion
    hab_cuenta = Val(frmIngHabitacion.txtNroHab.Text)
    cabezal_formulario
    DBGrid1.Font.Bold = True
    
    'inicializo control data
    subInicializoControlData Me.Data1
    
    'muestro pasajeros de la habitación
    SQLpasajeros_habitacion hab_cuenta, Data1
    'propiedades de la grilla
    DBGrid1.FormatString = " |Nombre pasajero                                                                                                      |                       "
End Sub

Private Sub BotConfirmar2_Click()
End Sub

Private Sub borro_checkin(pasajero As Long)
    tbCHECKIN.Index = "i_checkin_cli"
    tbCHECKIN.Seek "=", pasajero
    If Not tbCHECKIN.NoMatch Then
        '4) grabo los datos en el checkout, para poder sacar datos estadísticos.
        grabo_checkout Val(DBGrid1.Text)
        tbCHECKIN.Delete
    End If
End Sub

Private Sub grabo_checkout(pas As Long)
    tbCHECKOUT.AddNew
        tbCHECKOUT("nrohab") = tbCHECKIN("nrohab")
        tbCHECKOUT("nroreserva") = tbCHECKIN("nroreserva")
        tbCHECKOUT("nrocorrcli") = tbCHECKIN("nrocorrcli")
        tbCHECKOUT("fdes") = tbCHECKIN("fcheckdes")
        tbCHECKOUT("fhas") = tbCHECKIN("fcheckhas")
        tbCHECKOUT("horainghab") = tbCHECKIN("horainghab")
        tbCHECKOUT("finghab") = tbCHECKIN("finghab")
        tbCHECKOUT("horaegrhab") = Time
        tbCHECKOUT("fegrhab") = m_FechaSis
    tbCHECKOUT.Update
End Sub

Private Function titular_habitaciones(tit As Long)
    titular_habitaciones = False
    tbCHECKIN.MoveFirst
    'recorro checkin para buscar todas las habitaciones ocupadas
    Do While Not tbCHECKIN.EOF
        'accedo a cada habitación alojada en el hotel, en forma directa
        tbHABITACIONES.Index = "inrohab"
        tbHABITACIONES.Seek "=", tbCHECKIN("nrohab")
        'busco si es titular en alguna de esas habitaciones
        If Not tbHABITACIONES.NoMatch Then
            If tbHABITACIONES("titular_unica") = tit Or _
            tbHABITACIONES("titular_extra") = tit Or _
            tbHABITACIONES("titular_aloja") = tit Then
                'si es titular
                titular_habitaciones = True
                Exit Do
            End If
        End If
        tbCHECKIN.MoveNext
    Loop
End Function

Private Sub obtengo_titular()
    'obtengo tiular/es de la habitación
    If busco_habitaTF(hab_cuenta) Then
        titular_unica = tbHABITACIONES("titular_unica")
        titular_extra = tbHABITACIONES("titular_extra")
        titular_aloja = tbHABITACIONES("titular_aloja")
    End If
End Sub

Public Function verifico_gastos()
    'verfifico en archivo tbCUENTAS si tiene gastos
    'Me posiciono en el primer gasto no facturado de ese titular
    'Si me puedo posicionar es porque tengo gastos pendientes (no importa de que fecha)
    verifico_gastos = False
    tbCUENTAS.Index = "i_titular"
    tbCUENTAS.Seek ">=", 0, Val(DBGrid1.Text), 0  'pasajero
    If Not tbCUENTAS.NoMatch Then
        If tbCUENTAS("facturado") = 0 And _
        tbCUENTAS("titular_cuenta") = Val(DBGrid1.Text) Then
            verifico_gastos = True
            Exit Function
        End If
    End If
    'verifico si tiene gastos de alojamiento
    'idem que para gastos extras
    tbCUENTAS_ALOJA.Index = "i_titular"
    tbCUENTAS_ALOJA.Seek ">=", 0, Val(DBGrid1.Text), 0 'pasajero
    If Not tbCUENTAS_ALOJA.NoMatch Then
        If tbCUENTAS_ALOJA("facturado") = 0 And _
        tbCUENTAS_ALOJA("titular_aloja") = Val(DBGrid1.Text) Then
            verifico_gastos = True
        End If
    End If
End Function

Private Sub botTodos_Click()
    marcar_todos
    botConfirmar.Enabled = True
    Me.mnuFormularioAceptar.Enabled = True
End Sub

Private Sub dbgrid1_KeyPress(KeyAscii As Integer)
    'Permito seleccionar un pasajero con la tecla enter
    dbgrid1_DblClick
End Sub

Private Sub dbgrid1_DblClick()
    Dim colorback As String
    Dim colorfore As String
    If DBGrid1.CellBackColor = mSisColor_15FilaSeleccionada Then    'si esta marcado
        colorback = &H80000005                  'desmarco
        colorfore = &H80000008
    Else                                        'marco
        colorback = mSisColor_15FilaSeleccionada
        colorfore = mSisColor_19FilaSeleccionadaTexto
    End If
    
    DBGrid1.col = 1
    DBGrid1.CellBackColor = colorback
    DBGrid1.CellForeColor = colorfore
    
    If DBGrid1.CellBackColor = mSisColor_15FilaSeleccionada Then    'si se realizo una marca
        botConfirmar.Enabled = True
        Me.mnuFormularioAceptar.Enabled = True
    Else        ' si se desmarco ferifico que no sea la última
        If verifico_marcas = False Then
            botConfirmar.Enabled = False
            Me.mnuFormularioAceptar.Enabled = False
        End If
    End If
End Sub

Public Function verifico_marcas()
    Dim i As Integer
    verifico_marcas = False
    i = 1
    DBGrid1.Row = 1
    Do While i < DBGrid1.Rows
        DBGrid1.Row = i
        If DBGrid1.CellBackColor = mSisColor_15FilaSeleccionada Then
            verifico_marcas = True
            Exit Do
        End If
       i = i + 1
    Loop
End Function

Private Sub marcar_todos()
    Dim i As Integer
    i = 1
    Do While i < DBGrid1.Rows
       DBGrid1.Row = i
       DBGrid1.CellBackColor = mSisColor_15FilaSeleccionada
       i = i + 1
    Loop
End Sub

Private Function verifico_marcas2()
    'Recorro la grilla y si estan todos marcados(en azul) debuelvo true.
    Dim i As Integer
    verifico_marcas2 = True
    i = 1
    DBGrid1.Row = 1
    Do While i < DBGrid1.Rows
       DBGrid1.Row = i
       If DBGrid1.CellBackColor <> mSisColor_15FilaSeleccionada Then
            verifico_marcas2 = False
            Exit Do
       End If
       i = i + 1
    Loop
End Function

Private Sub cabezal_formulario()
    Me.gaHOTELtitular1.CaminoBaseDeDatos = vardir
    Me.gaHOTELtitular1.NumeroHabitacion = hab_cuenta
End Sub

Private Function obtengo_nombre()
    DBGrid1.col = 1 'columna nombre
    obtengo_nombre = DBGrid1.Text
    DBGrid1.col = 2 'columna nro
End Function

Private Sub errores(err As Byte)
    Select Case err
        Case 1
            'No puede dejar a la habitación sin titular
            mSubMensaje 4, 81
        Case 2
            'El muy atorrante tiene gastos pendientes
            mSubMensaje 4, 82, obtengo_nombre
        Case 3
            'un Pasajero es titular de otra habitación
            mSubMensaje 4, 83
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Descargo formulario de memoria
    Set frmCheck_Out = Nothing
End Sub

Private Sub mnuFormularioAceptar_Click()
    'Equivale a presionar el boton ed aceptar o la tecla F12
    botConfirmar_Click
End Sub

Private Sub mnuFormularioCancelar_Click()
    'Equivale a presionar el boton d cancelar o la tecla Esc
    botCancelar_Click
End Sub

Private Sub botCancelar_Click()
    Unload Me
    frmIngHabitacion.Show 1
End Sub

'******************************************************************************
'*
'*  Asistencia al usuario
'*
'******************************************************************************

Private Sub dbgrid1_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 17
End Sub

Private Sub botTodos_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 13
End Sub

Private Sub botCancelar_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 1, 3
End Sub

Private Sub botConfirmar_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 1, 2
End Sub

Private Sub botConfirmar_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub botCancelar_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub botTodos_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub dbgrid1_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub
        
