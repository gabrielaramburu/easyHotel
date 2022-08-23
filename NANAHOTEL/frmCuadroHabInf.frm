VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmCuadroHabInf 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Información"
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5415
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   8070
      _Version        =   327680
      Style           =   1
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Reservas"
      TabPicture(0)   =   "frmCuadroHabInf.frx":0000
      Tab(0).ControlCount=   1
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      TabCaption(1)   =   "Ocupadas"
      TabPicture(1)   =   "frmCuadroHabInf.frx":001C
      Tab(1).ControlCount=   1
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      TabCaption(2)   =   "Bloqueadas"
      TabPicture(2)   =   "frmCuadroHabInf.frx":0038
      Tab(2).ControlCount=   1
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).Control(0).Enabled=   0   'False
      TabCaption(3)   =   "No asignadas"
      TabPicture(3)   =   "frmCuadroHabInf.frx":0054
      Tab(3).ControlCount=   1
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame4"
      Tab(3).Control(0).Enabled=   0   'False
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   3975
         Left            =   -74880
         TabIndex        =   40
         Top             =   480
         Width           =   4935
         Begin VB.TextBox txtObsNoAsig 
            Height          =   1095
            Left            =   0
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   48
            Top             =   2760
            Width           =   3975
         End
         Begin VB.TextBox txtTipoHabNoAsig 
            Height          =   315
            Left            =   0
            TabIndex        =   47
            Top             =   2160
            Width           =   3375
         End
         Begin VB.TextBox txtPasaNoAsig 
            Height          =   315
            Left            =   4080
            Locked          =   -1  'True
            TabIndex        =   46
            Top             =   2760
            Width           =   735
         End
         Begin VB.TextBox txtTarifaNoAsig 
            Height          =   315
            Left            =   4080
            Locked          =   -1  'True
            TabIndex        =   45
            Top             =   2160
            Width           =   735
         End
         Begin VB.TextBox txtNochesNoAsig 
            Height          =   315
            Left            =   4080
            Locked          =   -1  'True
            TabIndex        =   44
            Top             =   1560
            Width           =   735
         End
         Begin VB.TextBox txtHastaNoAsig 
            Height          =   315
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   43
            Top             =   1560
            Width           =   1215
         End
         Begin VB.TextBox txtDesdeNoAsig 
            Height          =   315
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   42
            Top             =   1560
            Width           =   1215
         End
         Begin VB.TextBox txtTitularNoAsig 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   41
            Top             =   840
            Width           =   4815
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Observaciones"
            Height          =   195
            Left            =   0
            TabIndex        =   57
            Top             =   2520
            Width           =   1065
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de habitación"
            Height          =   195
            Left            =   0
            TabIndex        =   56
            Top             =   1920
            Width           =   1320
         End
         Begin VB.Label lblReservaNoAsig 
            AutoSize        =   -1  'True
            Caption         =   "NroReservaNoAsignada"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   360
            Left            =   120
            TabIndex        =   55
            Top             =   0
            Width           =   3375
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Pasajeros"
            Height          =   195
            Left            =   4080
            TabIndex        =   54
            Top             =   2520
            Width           =   690
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Tarifa "
            Height          =   195
            Left            =   4080
            TabIndex        =   53
            Top             =   1920
            Width           =   450
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Noches"
            Height          =   195
            Left            =   4080
            TabIndex        =   52
            Top             =   1320
            Width           =   555
         End
         Begin VB.Label Label17 
            Caption         =   "Titular reserva"
            Height          =   255
            Left            =   0
            TabIndex        =   51
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Ingreso"
            Height          =   195
            Left            =   0
            TabIndex        =   50
            Top             =   1320
            Width           =   525
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Egreso"
            Height          =   195
            Left            =   1440
            TabIndex        =   49
            Top             =   1320
            Width           =   495
         End
         Begin VB.Line Line4 
            X1              =   0
            X2              =   4920
            Y1              =   480
            Y2              =   480
         End
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   4095
         Left            =   -74880
         TabIndex        =   30
         Top             =   360
         Width           =   4935
         Begin VB.TextBox txtDesdeBloq 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   38
            Top             =   1680
            Width           =   1215
         End
         Begin VB.TextBox txtHastaBloq 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   37
            Top             =   1680
            Width           =   1215
         End
         Begin VB.ComboBox cboMotivoBloq 
            Height          =   315
            Left            =   0
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   960
            Width           =   4935
         End
         Begin VB.TextBox txtObsBloq 
            Height          =   1335
            Left            =   0
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   32
            Top             =   2400
            Width           =   4935
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Observaciones"
            Height          =   195
            Left            =   0
            TabIndex        =   39
            Top             =   2160
            Width           =   1065
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Motivo bloqueo"
            Height          =   240
            Left            =   0
            TabIndex        =   35
            Top             =   720
            Width           =   1395
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            Height          =   240
            Left            =   1440
            TabIndex        =   34
            Top             =   1440
            Width           =   540
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            Height          =   240
            Left            =   0
            TabIndex        =   33
            Top             =   1440
            Width           =   615
         End
         Begin VB.Line Line3 
            X1              =   0
            X2              =   4920
            Y1              =   600
            Y2              =   600
         End
         Begin VB.Label lblBloqueo 
            AutoSize        =   -1  'True
            Caption         =   "Bloqueo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   360
            Left            =   0
            TabIndex        =   31
            Top             =   120
            Width           =   1185
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   4095
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   4935
         Begin VB.Data Data1 
            Caption         =   "Data1"
            Connect         =   "Access"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Left            =   3000
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   2  'Snapshot
            RecordSource    =   "select * from checkin,clientes"
            Top             =   120
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.TextBox txtTitularCheckin 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   500
            Left            =   0
            MultiLine       =   -1  'True
            TabIndex        =   27
            Top             =   770
            Width           =   4935
         End
         Begin VB.TextBox txtDesdeCheckin 
            Height          =   315
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   23
            Top             =   1560
            Width           =   1215
         End
         Begin VB.TextBox txtHastacheckin 
            Height          =   315
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   22
            Top             =   1560
            Width           =   1215
         End
         Begin VB.TextBox txtTarifacheckin 
            Height          =   315
            Left            =   4200
            Locked          =   -1  'True
            TabIndex        =   21
            Top             =   1560
            Width           =   735
         End
         Begin MSDBGrid.DBGrid DBGrid1 
            Bindings        =   "frmCuadroHabInf.frx":0070
            Height          =   1935
            Left            =   0
            OleObjectBlob   =   "frmCuadroHabInf.frx":0080
            TabIndex        =   29
            Top             =   2040
            Width           =   4935
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Tarifa "
            Height          =   195
            Left            =   4200
            TabIndex        =   28
            Top             =   1320
            Width           =   450
         End
         Begin VB.Label lblTitulares 
            AutoSize        =   -1  'True
            Caption         =   "Titular"
            Height          =   195
            Left            =   0
            TabIndex        =   26
            Top             =   550
            Width           =   435
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Egreso"
            Height          =   195
            Left            =   1440
            TabIndex        =   25
            Top             =   1320
            Width           =   495
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Ingreso"
            Height          =   195
            Left            =   0
            TabIndex        =   24
            Top             =   1320
            Width           =   525
         End
         Begin VB.Label lblCheckin 
            AutoSize        =   -1  'True
            Caption         =   "Check-in"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF00&
            Height          =   360
            Left            =   120
            TabIndex        =   20
            Top             =   50
            Width           =   1260
         End
         Begin VB.Line Line2 
            X1              =   0
            X2              =   4920
            Y1              =   480
            Y2              =   480
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   3975
         Left            =   -74880
         TabIndex        =   1
         Top             =   480
         Width           =   4935
         Begin VB.TextBox txtNombreTit 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   840
            Width           =   4815
         End
         Begin VB.TextBox txtdesde 
            Height          =   315
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   1560
            Width           =   1215
         End
         Begin VB.TextBox txthasta 
            Height          =   315
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   1560
            Width           =   1215
         End
         Begin VB.TextBox txtNoches 
            Height          =   315
            Left            =   4080
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   1560
            Width           =   735
         End
         Begin VB.TextBox txtTarifa 
            Height          =   315
            Left            =   4080
            Locked          =   -1  'True
            TabIndex        =   5
            Top             =   2160
            Width           =   735
         End
         Begin VB.TextBox txtCantPasa 
            Height          =   315
            Left            =   4080
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   2760
            Width           =   735
         End
         Begin VB.TextBox txtHabitacion 
            Height          =   315
            Left            =   0
            TabIndex        =   3
            Top             =   2160
            Width           =   3375
         End
         Begin VB.TextBox txtObs 
            Height          =   1095
            Left            =   0
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   2
            Top             =   2760
            Width           =   3975
         End
         Begin VB.Line Line1 
            X1              =   0
            X2              =   4920
            Y1              =   480
            Y2              =   480
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Egreso"
            Height          =   195
            Left            =   1440
            TabIndex        =   18
            Top             =   1320
            Width           =   495
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Ingreso"
            Height          =   195
            Left            =   0
            TabIndex        =   17
            Top             =   1320
            Width           =   525
         End
         Begin VB.Label labTitular 
            Caption         =   "Titular reserva"
            Height          =   255
            Left            =   0
            TabIndex        =   16
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Noches"
            Height          =   195
            Left            =   4080
            TabIndex        =   15
            Top             =   1320
            Width           =   555
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Tarifa "
            Height          =   195
            Left            =   4080
            TabIndex        =   14
            Top             =   1920
            Width           =   450
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Pasajeros"
            Height          =   195
            Left            =   4080
            TabIndex        =   13
            Top             =   2520
            Width           =   690
         End
         Begin VB.Label lblReserva 
            AutoSize        =   -1  'True
            Caption         =   "NroReserva"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   360
            Left            =   120
            TabIndex        =   12
            Top             =   0
            Width           =   1665
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Habitación"
            Height          =   195
            Left            =   0
            TabIndex        =   11
            Top             =   1920
            Width           =   765
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Observaciones"
            Height          =   195
            Left            =   0
            TabIndex        =   10
            Top             =   2520
            Width           =   1065
         End
      End
   End
End
Attribute VB_Name = "frmCuadroHabInf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
    'Si pongo estas líneas de código en el evento load
    'es sabido ya, que los tabs no muestran los controles
    Select Case tipoAccionCuadroHabInf
        Case 1  'mostrar reserva
            subDeshabilitoTabs
            Me.ssTab1.TabEnabled(0) = True
            Me.ssTab1.Tab = 0
        Case 2  'mostrar ocupada
            subDeshabilitoTabs
            Me.ssTab1.TabEnabled(1) = True
            Me.ssTab1.Tab = 1
        Case 3  'mostrar bloqueada
            subDeshabilitoTabs
            Me.ssTab1.TabEnabled(2) = True
            Me.ssTab1.Tab = 2
        Case 4  'mostrar no asignada
            subDeshabilitoTabs
            Me.ssTab1.TabEnabled(3) = True
            Me.ssTab1.Tab = 3
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    'Inicializo control data
    subInicializoControlData Me.Data1
    'Apariencia formulario
    mSubConfiguroFuentesControlesSistema Me
    mSub_bloqueo_controles_formulario Me, True
    
    'Determino para que llamo el formulario
    Select Case tipoAccionCuadroHabInf
        Case 1  'mostrar reserva
            subMuestroReserva
        Case 2  'mostrar ocupada
            subMuestroOcupada
        Case 3  'mostrar bloqueada
            subMuestroBloqueada
        Case 4  'mostrar no asignada
            subMuestroNoAsignada
    End Select
End Sub

'***************************************************
'*
'*      Muestro reserva
'*
'*
'***************************************************

Private Sub subMuestroReserva()
    'Muestro información de la reserva
    Dim reserva As Long
    Dim hab As Long
    
    'obtengo número de reserva
    reserva = Val(NroResFormato(frmCuadroHab.gHabitacion.Text))
    
    'obtengo número de habitación
    frmCuadroHab.gHabitacion.col = 0
    hab = Val(frmCuadroHab.gHabitacion.Text)
    If busco_habitaTF(hab) Then
        subMuestroDatos reserva, hab
    End If
End Sub

Private Sub subMuestroNoAsignada()
    'Muestro información reserva no asignada
    Dim reservaNoAsig As Long
    
    'obtengo número de reserva
    reservaNoAsig = Val(NroResFormato(frmCuadroHab.gHabitacion.Text))
    
    tbRESERVAS.Index = "i_reservas"
    tbRESERVAS.Seek "=", reservaNoAsig
    If Not tbRESERVAS.NoMatch Then
        'muestro reserva
        Me.lblReservaNoAsig.Caption = NroResFormato(reservaNoAsig)
        txtDesdeNoAsig.Text = tbRESERVAS("fechaing")
        txtHastaNoAsig.Text = tbRESERVAS("fechaegr")
        txtNochesNoAsig.Text = tbRESERVAS("cantnoches")
        If Not tbRESERVAS("observaciones") Then
            txtObsNoAsig.Text = tbRESERVAS("observaciones")
        End If
        'Formo nombre comleto titular
        txtTitularNoAsig.Text = tbRESERVAS("primer_ape_titular") & " " & _
        tbRESERVAS("segundo_ape_titular") & " " & _
        tbRESERVAS("primer_nom_titular") & " " & _
        tbRESERVAS("segundo_nom_titular")
        subBuscoTarifaCantPasaTipoHabDeNoAsignadas reservaNoAsig, frmCuadroHab.gHabitacion.TextMatrix(frmCuadroHab.gHabitacion.Row, 0)
    End If
End Sub

Private Sub subBuscoTarifaCantPasaTipoHabDeNoAsignadas(res As Long, corr As Long)
    'Muestro datos de la reserva no asignada: tarifa, cantidad de pasajeros y tipo de habitacion
    If mfunBuscoReservaNoAsignada(res, corr) Then
        txtTarifaNoAsig.Text = tbHAB_RESERVAS("tarifa")
        txtPasaNoAsig.Text = tbHAB_RESERVAS("pasajeros")
        txtTipoHabNoAsig.Text = mFun_BuscoDescriTipoHab(tbHAB_RESERVAS("tipohabitacion"))
    End If
End Sub

Private Sub subMuestroDatos(reserva As Long, hab As Long)
    'Muestro los datos de la reserva en el formulario busco reserva
    tbRESERVAS.Index = "i_reservas"
    tbRESERVAS.Seek "=", reserva
    If Not tbRESERVAS.NoMatch Then
        'muestro reserva
        Me.lblReserva.Caption = NroResFormato(reserva)
        txtdesde.Text = tbRESERVAS("fechaing")
        txthasta.Text = tbRESERVAS("fechaegr")
        txtNoches.Text = tbRESERVAS("cantnoches")
        If Not IsNull(tbRESERVAS("observaciones")) Then _
            Me.txtObs.Text = tbRESERVAS("observaciones")
        'muestro habitación
        txtHabitacion.Text = Str(hab) & " Suite " & busco_tipo_hab_descri(hab)
        'Formo nombre comleto titular
        txtNombreTit.Text = tbRESERVAS("primer_ape_titular") & " " & _
        tbRESERVAS("segundo_ape_titular") & " " & _
        tbRESERVAS("primer_nom_titular") & " " & _
        tbRESERVAS("segundo_nom_titular")
        subBuscoTarifaCantPasa reserva, hab
    End If
End Sub

Private Sub subBuscoTarifaCantPasa(reserva As Long, hab As Long)
    'La tarifa y la cantidad de pasajeros son datos asignados
    'a cada habitación de la reserva por eso los tengo que ir a buscar
    'a tbHAB_RESERVAS
    tbHAB_RESERVAS.Index = "ihab_reserva_hab"
    tbHAB_RESERVAS.Seek "=", reserva, hab
    If Not tbHAB_RESERVAS.NoMatch Then
        txttarifa.Text = tbHAB_RESERVAS("tarifa")
        txtCantPasa.Text = tbHAB_RESERVAS("pasajeros")
    End If
End Sub

Private Sub subDeshabilitoTabs()
    'No permito que se trabaje con otros tabs que
    'no sea el correspondiente al tipo de accion
    Me.ssTab1.TabEnabled(0) = False
    Me.ssTab1.TabEnabled(1) = False
    Me.ssTab1.TabEnabled(2) = False
    Me.ssTab1.TabEnabled(3) = False
End Sub

'***************************************************
'*
'*      Muestro ocupadas
'*
'*
'***************************************************

Private Sub subMuestroOcupada()
    'Muestro información de la ocupación
    Dim hab As Long
    

    'obtengo número de habitación
    frmCuadroHab.gHabitacion.col = 0
    hab = Val(frmCuadroHab.gHabitacion.Text)
    
    If busco_habitaTF(hab) Then
        'muestro habitacion
        lblCheckin.Caption = Str(hab) & " Suite " & busco_tipo_hab_descri(hab)
        'obtengo tarifa
        txtTarifacheckin.Text = tbHABITACIONES("tarifa")
        'muestro titular
        txtTitularCheckin.Text = funObtengoTitular(hab, tbHABITACIONES("titular_unica"))
        'obtengo datos ocupación
        If busco_habita_checkin(hab) Then
            Me.txtDesdeCheckin.Text = tbCHECKIN("fCheckDes")
            Me.txtHastacheckin.Text = tbCHECKIN("fCheckHas")
        End If
        'obtengopasajeros
        SQLpasajeros_habitacion hab, Data1
    End If
End Sub

Private Function funObtengoTitular(hab As Long, titularUnica As Long)
    'Devulevo el o los titulares de la habitación
    Dim titAux As String
    
    If titularUnica = 0 Then    'las cuentas son separadas
        titAux = "T. aloja. " & busco_titular_hab(hab, "aloja") & Chr(9)
        titAux = titAux & "T. extra " & busco_titular_hab(hab, "extra")
    Else                        'cuenta única
        titAux = "T.unica " & busco_titular_hab(hab, "unica")
    End If
    funObtengoTitular = titAux
End Function

'***************************************************
'*
'*      Muestro bloqueadas
'*
'*
'***************************************************

Private Sub subMuestroBloqueada()
    'Muestro información de bloqueo
    Dim hab As Long
    Dim nroCorrBloq As Long
    
    'obtengo número de bloqueo
    nroCorrBloq = Val(frmCuadroHab.gHabitacion.Text)
    
    'obtengo número de habitación
    frmCuadroHab.gHabitacion.col = 0
    hab = Val(frmCuadroHab.gHabitacion.Text)
    
    'cargo combo de motivos
    carga_tipo_estado_hab Me.cboMotivoBloq, 1
    
    'muestro habitacion
    lblBloqueo.Caption = Str(hab) & " Suite " & busco_tipo_hab_descri(hab)
    'busco bloqueo
    If funBuscoBloqueoTF(hab, nroCorrBloq) Then
        txtDesdeBloq.Text = tbBLOQUEO_HAB("fDesdeBloq")
        txtHastaBloq.Text = tbBLOQUEO_HAB("fHastaBloq")
        txtObsBloq.Text = tbBLOQUEO_HAB("ObsBloq")
        posiciono_combo cboMotivoBloq, tbBLOQUEO_HAB("MotivoBloq")
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Descargo formulario de memoria
    Set frmCuadroHabInf = Nothing
End Sub
