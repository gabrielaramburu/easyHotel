VERSION 5.00
Object = "{EB8C3860-8603-11D0-A35D-7062E4000000}#1.1#0"; "VCBND.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmBloquearHab 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bloqueo de habitaciones"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   9510
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton botConfirmar 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   7920
      TabIndex        =   16
      Tag             =   "Aceptar"
      Top             =   6120
      Width           =   1200
   End
   Begin VB.Frame Frame2 
      Caption         =   "Bloqueos activos de la habitación "
      Height          =   3135
      Left            =   120
      TabIndex        =   14
      Top             =   0
      Width           =   9255
      Begin Hotel_Nana.gaHOTELtipo gaHOTELtipo1 
         Height          =   300
         Left            =   120
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   360
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   529
         BackColor       =   -2147483633
      End
      Begin VB.CommandButton botEliminar 
         Caption         =   "D&esbloquear"
         Height          =   375
         Left            =   7560
         TabIndex        =   11
         Top             =   240
         Width           =   1440
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Bindings        =   "frmBloquearHab.frx":0000
         Height          =   2055
         Left            =   240
         TabIndex        =   10
         Top             =   960
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   3625
         _Version        =   393216
         Cols            =   6
         FocusRect       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   $"frmBloquearHab.frx":0010
      End
      Begin VB.Label Label1 
         Caption         =   "Bloqueos &activos"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Width           =   1575
      End
   End
   Begin Hotel_Nana.gaHOTELbarra gaHOTELbarra1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   6495
      Width           =   9510
      _ExtentX        =   16775
      _ExtentY        =   582
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   405
      Left            =   2040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   "select * from bloqueo_hab, tipo_estado_hab"
      Top             =   6120
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Caption         =   "Nuevo período de bloqueo"
      Height          =   2775
      Left            =   120
      TabIndex        =   12
      Top             =   3240
      Width           =   9255
      Begin VB.CommandButton botBloquear 
         Caption         =   "&Bloquear"
         Height          =   375
         Left            =   7800
         TabIndex        =   8
         Top             =   240
         Width           =   1200
      End
      Begin VB.ComboBox cboMotivoBloq 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   4935
      End
      Begin VB.TextBox txtObsBloq 
         Height          =   855
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   1680
         Width           =   8775
      End
      Begin VcBndCtl.VcCalCombo fDesdeBloq 
         Height          =   375
         Left            =   1560
         TabIndex        =   3
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _0              =   $"frmBloquearHab.frx":00B4
         _1              =   $"frmBloquearHab.frx":04BD
         _2              =   $"frmBloquearHab.frx":08C6
         _3              =   "-@A@@@)f@@@E@@@A@@@@@@@@@'@@@E@@@A@@@@@@@@@)h@@@E@@@A@@@@@@@@@,456D"
         _count          =   4
         _ver            =   2
      End
      Begin VcBndCtl.VcCalCombo fHastaBloq 
         Height          =   375
         Left            =   5280
         TabIndex        =   5
         Top             =   840
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _0              =   $"frmBloquearHab.frx":0CCF
         _1              =   $"frmBloquearHab.frx":10D8
         _2              =   $"frmBloquearHab.frx":14E1
         _3              =   "f-@@@E@@@A@@@@@@@@@'@@@E@@@A@@@@@@@@@)h@@@E@@@A@@@@@@@@@,467D"
         _count          =   4
         _ver            =   2
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "&Desde"
         Height          =   240
         Left            =   240
         TabIndex        =   2
         Top             =   900
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "&Hasta"
         Height          =   240
         Left            =   4440
         TabIndex        =   4
         Top             =   900
         Width           =   540
      End
      Begin VB.Label Label5 
         Caption         =   "&Motivo bloqueo"
         Height          =   255
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "&Observaciones"
         Height          =   240
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   1380
      End
   End
   Begin VB.Menu mnuFormulario 
      Caption         =   "&Formulario"
      Begin VB.Menu mnuFormularioAceptar 
         Caption         =   "Aceptar"
         Shortcut        =   {F12}
      End
      Begin VB.Menu div1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFormularioBloquear 
         Caption         =   "Bloquear"
         Shortcut        =   {F9}
      End
   End
End
Attribute VB_Name = "frmBloquearHab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private nrohab As Long

Private Sub Form_Load()
    
    'inicializo control data
    subInicializoControlData Me.Data1
    
    'Apariencia formulario
    mSubConfiguroFuentesControlesSistema Me
    'Apariencia grilla
    Me.MSFlexGrid1.BackColorSel = mSisColor_15FilaSeleccionada
    Me.MSFlexGrid1.ForeColorSel = mSisColor_19FilaSeleccionadaTexto
    
    'Obtengo número habitación
    nrohab = Val(frmIngHabitacion2.txtNroHab.Text)
    
    'Muestro habitación
    Me.gaHOTELtipo1.CaminoBaseDeDatos = vardir
    Me.gaHOTELtipo1.NumeroHabitacion = nrohab
    
    carga_tipo_estado_hab cboMotivoBloq, 1
    inicializo_formulario
    cargo_grilla_bloqueos_activos
    
End Sub

Private Sub botBloquear_Click()
    If valido_fechas Then
        If Not habitacion_bloqueada(nrohab, fDesdeBloq.Value, fHastaBloq.Value) Then
            If Not reservada Then
                If Not ocupada Then
                    grabo_bloqueo
                    'inicializo formulario para permitir ingreso de nuevo bloqueo
                    inicializo_formulario
                    'actualizo grilla de bloqueos
                    actualizo_grilla
                    'grabo bitacora
                    GraboBitacora "Hab. " & nrohab
                Else
                    'habitación ocupada
                    mSubMensaje 4, 1
                End If
            Else
                'habitación reservada
                mSubMensaje 4, 2
            End If
        Else
            'habitación bloqueada
            mSubMensaje 4, 3
        End If
    End If
End Sub

Private Sub inicializo_formulario()
    cboMotivoBloq.ListIndex = 0
    fDesdeBloq.Value = Null
    fHastaBloq.Value = Null
    txtObsBloq.Text = ""
End Sub

Private Sub cargo_grilla_bloqueos_activos()
    'Cargo en la grilla los bloqueos activos de la habitación
    
    Dim consulta As String
    consulta = _
        "Select hab_bloq, fdesdebloq, fhastabloq, descri, obsbloq, nrocorr_bloq " & _
        "From bloqueo_hab,tipo_estado_hab " & _
        "Where bloqueo_hab.motivobloq = tipo_estado_hab.cod " & _
        "and tipo_estado_hab.tipo_cod = 1 " & _
        "and hab_bloq = " & nrohab & _
        "and fhastabloq >= " & fechaSQL(m_FechaSis) & _
        " Order by fdesdebloq"
        
    Data1.RecordSource = consulta
    actualizo_grilla
End Sub

Private Sub grabo_bloqueo()
    Dim nroaux As Long
    nroaux = mFun_obtengo_nrocorr_bloqueo(nrohab)
    tbBLOQUEO_HAB.AddNew
        tbBLOQUEO_HAB("nrocorr_bloq") = nroaux
        tbBLOQUEO_HAB("hab_bloq") = nrohab
        tbBLOQUEO_HAB("fdesdebloq") = fDesdeBloq.Value
        tbBLOQUEO_HAB("fhastabloq") = fHastaBloq.Value
        tbBLOQUEO_HAB("motivobloq") = cboMotivoBloq.ItemData(cboMotivoBloq.ListIndex)
        tbBLOQUEO_HAB("obsbloq") = txtObsBloq.Text
    tbBLOQUEO_HAB.Update
End Sub

Private Function reservada()
    reservada = False
    If habitacion_reservada(nrohab, fDesdeBloq.Value, fHastaBloq.Value) Then
        reservada = True
    End If
End Function

Private Function ocupada()
    ocupada = False
    If Not habitacion_ocupada(nrohab, fDesdeBloq.Value) Then
        ocupada = False
    End If
End Function

Private Sub botEliminar_Click()
    elimino_bloqueo
    'actualizo grilla de bloqueos
    actualizo_grilla
End Sub

Private Sub elimino_bloqueo()
    'Elimino el bloqueo del archivo de bloqueos, tomando información de la línea seleccionada
    'en la grilla, Esta información me permite acceder al registro a borrar.
    
    'verifico que exista alguna fila seleccionada
    If MSFlexGrid1.RowSel > 0 Then
        MSFlexGrid1.col = 6
        tbBLOQUEO_HAB.Index = "pk_bloqueo_hab"
        tbBLOQUEO_HAB.Seek "=", nrohab, Val(MSFlexGrid1.Text)
        If Not tbBLOQUEO_HAB.NoMatch Then
            tbBLOQUEO_HAB.Delete
        End If
    End If
End Sub

Private Function valido_fechas()
    Dim tipo_err As Byte
    valido_fechas = False
    'ambas fechas son válidas
    If Not IsDate(fDesdeBloq.Text) Then
        tipo_err = 1
        fDesdeBloq.SetFocus
    Else
        If Not IsDate(fHastaBloq.Text) Then
            tipo_err = 1
            fHastaBloq.SetFocus
        Else
            'fecha hasta mayor a desde
            If fHastaBloq.Value < fDesdeBloq.Value Then
                tipo_err = 3
                fHastaBloq.SetFocus
            Else
                'fecha desde mayor igual a la de hoy
                If fDesdeBloq.Value < m_FechaSis Then
                    tipo_err = 2
                    fDesdeBloq.SetFocus
                Else
                    If fDesdeBloq.Value = fHastaBloq.Value Then
                        tipo_err = 4
                        fDesdeBloq.SetFocus
                    End If
                End If
            End If
        End If
    End If
    
    Select Case tipo_err
        Case 1
            'formato de fecha incorrecto
            mSubMensaje 3, 1
        Case 2
            'fecha ingresada no puede ser menor al día de hoy.
            mSubMensaje 3, 2
        Case 3
            'fecha inicial no pueder ser mayor a fecha final
            mSubMensaje 3, 3
        Case 4
            'fecha inicial y final no puede ser iguales
            mSubMensaje 3, 4
        Case Else
            valido_fechas = True
    End Select
End Function

Private Sub actualizo_grilla()
    'Actualizo grilla de bloqueos activos
    'Es llamado cada vez que se realiza uan nueva consulta
    Data1.Refresh
    MSFlexGrid1.FormatString = "      | Habitación | Desde         | Hasta         | Motivo                        | Observaciónes                                                |nrocorr "
    MSFlexGrid1.ColWidth(6) = 0
End Sub

Private Sub botConfirmar_Click()
    Unload Me
    frmIngHabitacion2.Show 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Descargo formulario de memoria
    Set frmBloquearHab = Nothing
End Sub

Private Sub mnuFormularioAceptar_Click()
    'Presionar F12 es lo mismo que presionar el boton de aceptar
    botConfirmar_Click
End Sub

Private Sub mnuFormularioBloquear_Click()
    'Cuando dentro de un formulario se realiza una operación de confirmación,
    'la cual no implica salir del formulario, ésta se puede ejecutar con la tecla F9
    'o con esta opción del menu.
    botBloquear_Click
End Sub

'******************************************************************************
'*
'*  Asistencia al usuario
'*
'******************************************************************************

Private Sub cboMotivoBloq_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 1
End Sub

Private Sub fDesdeBloq_GotFocus()
     mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 2
End Sub

Private Sub fHastaBloq_GotFocus()
     mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 3
End Sub

Private Sub txtObsBloq_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 4
End Sub

Private Sub MSFlexGrid1_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 5
End Sub

Private Sub botBloquear_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 6
End Sub

Private Sub botEliminar_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 7
End Sub

Private Sub botConfirmar_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 1, 1
End Sub

