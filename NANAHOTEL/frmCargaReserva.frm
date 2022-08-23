VERSION 5.00
Object = "{EB8C3860-8603-11D0-A35D-7062E4000000}#1.1#0"; "VCBND.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCargaReserva 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Carga de reserva"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   11910
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton botImprimir 
      Height          =   375
      Left            =   3480
      Picture         =   "frmCargaReserva.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   82
      Top             =   7200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   "select * from hab_reserva"
      Top             =   0
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.Frame Frame2 
      Height          =   3255
      Left            =   120
      TabIndex        =   69
      Top             =   1440
      Width           =   11655
      Begin TabDlg.SSTab SSTab1 
         Height          =   2895
         Left            =   120
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   240
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   5106
         _Version        =   327680
         Style           =   1
         TabHeight       =   520
         ForeColor       =   12582912
         TabCaption(0)   =   "Habitación"
         TabPicture(0)   =   "frmCargaReserva.frx":0942
         Tab(0).ControlCount=   1
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame6"
         Tab(0).Control(0).Enabled=   0   'False
         TabCaption(1)   =   "Tarjeta de Crédito"
         TabPicture(1)   =   "frmCargaReserva.frx":095E
         Tab(1).ControlCount=   1
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame5"
         Tab(1).Control(0).Enabled=   0   'False
         TabCaption(2)   =   "Pre Pago / Seña"
         TabPicture(2)   =   "frmCargaReserva.frx":097A
         Tab(2).ControlCount=   1
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame3"
         Tab(2).Control(0).Enabled=   0   'False
         Begin VB.Frame Frame6 
            BorderStyle     =   0  'None
            Height          =   2415
            Left            =   120
            TabIndex        =   75
            Top             =   360
            Width           =   11175
            Begin MSFlexGridLib.MSFlexGrid dbgrid1 
               Bindings        =   "frmCargaReserva.frx":0996
               Height          =   2055
               Left            =   3360
               TabIndex        =   20
               Top             =   360
               Width           =   6495
               _ExtentX        =   11456
               _ExtentY        =   3625
               _Version        =   393216
               Cols            =   8
               AllowBigSelection=   0   'False
               FocusRect       =   2
               HighLight       =   0
               SelectionMode   =   1
               MousePointer    =   2
               FormatString    =   "      |Tipo Habitación |   Tarifa   | Cant. Pax | Nro. Hab.  |        Estado       |              "
            End
            Begin VB.CommandButton botHabitacion 
               Caption         =   "&Selección"
               Height          =   375
               Left            =   1920
               TabIndex        =   18
               Top             =   1800
               Width           =   1215
            End
            Begin VB.TextBox txtHab 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000016&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   360
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   77
               TabStop         =   0   'False
               Top             =   1800
               Width           =   735
            End
            Begin VB.ComboBox cboTipo_habitacion 
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
               Left            =   1320
               Style           =   2  'Dropdown List
               TabIndex        =   13
               Top             =   360
               Width           =   1815
            End
            Begin VB.TextBox txtTarifa 
               Alignment       =   1  'Right Justify
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
               Left            =   2280
               MaxLength       =   5
               TabIndex        =   15
               Top             =   840
               Width           =   855
            End
            Begin VB.TextBox txtPasajeros 
               Alignment       =   1  'Right Justify
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
               Left            =   2640
               MaxLength       =   2
               TabIndex        =   17
               Top             =   1320
               Width           =   495
            End
            Begin VB.CommandButton Eliminar_hab 
               Caption         =   "&Borrar"
               Enabled         =   0   'False
               Height          =   375
               Left            =   9960
               TabIndex        =   76
               Top             =   2040
               Width           =   1215
            End
            Begin VB.Label lblSignoDol 
               AutoSize        =   -1  'True
               Caption         =   "lblSignoDol"
               Height          =   240
               Index           =   1
               Left            =   720
               TabIndex        =   84
               Top             =   900
               Width           =   1050
            End
            Begin VB.Label Label26 
               Caption         =   "&Canidad de pasajeros"
               Height          =   255
               Left            =   0
               TabIndex        =   16
               Top             =   1380
               Width           =   2175
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               Caption         =   "&Tarifa"
               Height          =   240
               Left            =   0
               TabIndex        =   14
               Top             =   900
               Width           =   525
            End
            Begin VB.Label Label19 
               Caption         =   "Tipo de &Habitación"
               Height          =   495
               Left            =   0
               TabIndex        =   12
               Top             =   300
               Width           =   975
            End
            Begin VB.Label Label31 
               AutoSize        =   -1  'True
               Caption         =   "Habitación"
               Height          =   240
               Left            =   0
               TabIndex        =   78
               Top             =   1860
               Width           =   975
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Habitaciones que componen la reser&va"
               Height          =   240
               Left            =   3360
               TabIndex        =   19
               Top             =   120
               Width           =   3540
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Datos de la Tarjeta"
            Height          =   2295
            Left            =   -74760
            TabIndex        =   72
            Top             =   480
            Width           =   10935
            Begin VB.ComboBox otras_tar 
               Height          =   360
               Left            =   240
               Style           =   2  'Dropdown List
               TabIndex        =   37
               Top             =   720
               Width           =   2655
            End
            Begin VB.TextBox txtNom_tar 
               Height          =   285
               Left            =   4680
               MaxLength       =   40
               TabIndex        =   39
               Top             =   360
               Width           =   4215
            End
            Begin VB.TextBox txtApe_tar 
               Height          =   285
               Left            =   4680
               MaxLength       =   40
               TabIndex        =   41
               Top             =   720
               Width           =   4215
            End
            Begin VB.TextBox txtNro_tar 
               Height          =   285
               Left            =   4680
               MaxLength       =   50
               TabIndex        =   43
               Top             =   1080
               Width           =   2175
            End
            Begin VB.TextBox txtCod_Seg_tar 
               Height          =   285
               Left            =   4680
               MaxLength       =   6
               TabIndex        =   47
               Top             =   1920
               Width           =   975
            End
            Begin VcBndCtl.VcMask fech_vto_tar 
               Height          =   405
               Left            =   4680
               TabIndex        =   45
               Top             =   1440
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   714
               _0              =   $"frmCargaReserva.frx":09A6
               _1              =   $"frmCargaReserva.frx":0DAF
               _count          =   2
               _ver            =   2
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "&Tarejta de crédito"
               Height          =   240
               Left            =   240
               TabIndex        =   36
               Top             =   360
               Width           =   1590
            End
            Begin VB.Label Label1 
               Caption         =   "&Nombres"
               Height          =   255
               Left            =   3240
               TabIndex        =   38
               Top             =   360
               Width           =   735
            End
            Begin VB.Label Label2 
               Caption         =   "A&pellidos"
               Height          =   255
               Left            =   3240
               TabIndex        =   40
               Top             =   720
               Width           =   735
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "N&ro. de Tarjeta"
               Height          =   240
               Left            =   3240
               TabIndex        =   42
               Top             =   1080
               Width           =   1350
            End
            Begin VB.Label Label17 
               Caption         =   "&Fecha Vto."
               Height          =   255
               Left            =   3240
               TabIndex        =   44
               Top             =   1515
               Width           =   1095
            End
            Begin VB.Label Label18 
               Caption         =   "&Cod. Seg."
               Height          =   255
               Left            =   3240
               TabIndex        =   46
               Top             =   1935
               Width           =   855
            End
            Begin VB.Label Label30 
               AutoSize        =   -1  'True
               Caption         =   "mm/aaaa"
               Height          =   240
               Left            =   5760
               TabIndex        =   73
               Top             =   1560
               Width           =   900
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Datos pago de seña"
            Height          =   2295
            Left            =   -74760
            TabIndex        =   71
            Top             =   480
            Width           =   10935
            Begin VB.TextBox txtApe_Seña 
               Height          =   285
               Left            =   2160
               MaxLength       =   40
               TabIndex        =   51
               Top             =   720
               Width           =   4215
            End
            Begin VB.TextBox txtRecibo_Seña 
               Height          =   285
               Left            =   5160
               MaxLength       =   8
               TabIndex        =   58
               Top             =   1560
               Width           =   1215
            End
            Begin VB.TextBox txtImporte_Seña 
               Height          =   285
               Left            =   2160
               MaxLength       =   6
               TabIndex        =   55
               Top             =   1560
               Width           =   855
            End
            Begin VB.TextBox txtNom_Seña 
               Height          =   285
               Left            =   2160
               MaxLength       =   40
               TabIndex        =   49
               Top             =   360
               Width           =   4215
            End
            Begin VcBndCtl.VcCalCombo fecha_emi_seña 
               Height          =   375
               Left            =   2160
               TabIndex        =   53
               Top             =   1080
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   661
               _0              =   $"frmCargaReserva.frx":0EE6
               _1              =   $"frmCargaReserva.frx":12EF
               _2              =   $"frmCargaReserva.frx":16F8
               _3              =   "-@@@@%@@@C@@@@@@@D@@@A@@@)f@@@E@@@A@@@@@@@@@'@@@E@@@A@@@@@@@@@)h@@@E@@@A@@@@@@@@@,DF53"
               _count          =   4
               _ver            =   2
            End
            Begin VB.Label lblSignoDol 
               Caption         =   "lblSignoDol"
               Height          =   240
               Index           =   0
               Left            =   1080
               TabIndex        =   83
               Top             =   1575
               Width           =   1050
            End
            Begin VB.Label Label27 
               AutoSize        =   -1  'True
               Caption         =   "&Apellido"
               Height          =   240
               Left            =   240
               TabIndex        =   50
               Top             =   720
               Width           =   750
            End
            Begin VB.Label Label23 
               AutoSize        =   -1  'True
               Caption         =   "N&ro. de recibo"
               Height          =   240
               Left            =   3720
               TabIndex        =   57
               Top             =   1560
               Width           =   1275
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               Caption         =   "&Importe"
               Height          =   240
               Left            =   240
               TabIndex        =   54
               Top             =   1575
               Width           =   675
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               Caption         =   "&Fecha de emision"
               Height          =   240
               Left            =   240
               TabIndex        =   52
               Top             =   1147
               Width           =   1605
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               Caption         =   "&Nombre"
               Height          =   240
               Left            =   240
               TabIndex        =   48
               Top             =   360
               Width           =   735
            End
         End
      End
   End
   Begin Hotel_Nana.gaHOTELbarra gaHOTELbarra1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   68
      TabStop         =   0   'False
      Top             =   7605
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   582
   End
   Begin VB.Frame Frame4 
      Height          =   1400
      Left            =   120
      TabIndex        =   67
      Top             =   0
      Width           =   11655
      Begin VB.TextBox fechaegr 
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
         Left            =   9240
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox fechaing 
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
         Left            =   3720
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txt1er_Nom_titular 
         Height          =   315
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   5
         Top             =   600
         Width           =   3255
      End
      Begin VB.TextBox txt2do_Nom_titular 
         Height          =   315
         Left            =   7080
         MaxLength       =   20
         TabIndex        =   7
         Top             =   600
         Width           =   3375
      End
      Begin VB.TextBox txt1er_Ape_titular 
         Height          =   315
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   9
         Top             =   960
         Width           =   3255
      End
      Begin VB.TextBox txt2do_Ape_titular 
         Height          =   315
         Left            =   7080
         MaxLength       =   20
         TabIndex        =   11
         Top             =   960
         Width           =   3375
      End
      Begin VB.CommandButton botAyuda 
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   11160
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   600
         Width           =   375
      End
      Begin VB.Label cantidadnoches 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   240
         Left            =   11400
         TabIndex        =   80
         Top             =   270
         Width           =   105
      End
      Begin VB.Label NroReserva 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   79
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Noches"
         Height          =   240
         Left            =   10560
         TabIndex        =   74
         Top             =   277
         Width           =   705
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha de &ingreso"
         Height          =   255
         Left            =   1680
         TabIndex        =   0
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label16 
         Caption         =   "Fecha de &egreso"
         Height          =   255
         Left            =   7080
         TabIndex        =   2
         Top             =   270
         Width           =   1815
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "&1er Nombre"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   660
         Width           =   825
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "1er &Apellido"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   1020
         Width           =   825
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "&2do Nombre"
         Height          =   195
         Left            =   5400
         TabIndex        =   6
         Top             =   660
         Width           =   870
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "2&do Apellido"
         Height          =   195
         Left            =   5400
         TabIndex        =   10
         Top             =   1020
         Width           =   870
      End
   End
   Begin VB.CommandButton ResNoCorresponde 
      Height          =   375
      Left            =   6360
      Picture         =   "frmCargaReserva.frx":1B01
      Style           =   1  'Graphical
      TabIndex        =   66
      Tag             =   "Cancelar"
      Top             =   7200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton ResCorresponde 
      Caption         =   "Continuar"
      Height          =   375
      Left            =   4920
      TabIndex        =   65
      Tag             =   "Continuar"
      Top             =   7200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton NoAnular 
      Height          =   375
      Left            =   2040
      Picture         =   "frmCargaReserva.frx":23C3
      Style           =   1  'Graphical
      TabIndex        =   60
      Tag             =   "Cancela Anulación"
      Top             =   7200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton FinConsulta 
      Caption         =   "Salir"
      Height          =   375
      Left            =   7920
      TabIndex        =   61
      Top             =   7200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton AnularReserva 
      Caption         =   "Anular"
      Height          =   375
      Left            =   600
      TabIndex        =   59
      Top             =   7200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton noconfirmareserva 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   10560
      Picture         =   "frmCargaReserva.frx":2C85
      Style           =   1  'Graphical
      TabIndex        =   63
      Tag             =   "Cancela reserva"
      Top             =   7200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton confirmareserva 
      Height          =   375
      Left            =   9240
      Picture         =   "frmCargaReserva.frx":3547
      Style           =   1  'Graphical
      TabIndex        =   62
      Tag             =   "Aceptar"
      Top             =   7200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos extras de la reserva "
      ClipControls    =   0   'False
      Height          =   2295
      Left            =   120
      TabIndex        =   64
      Top             =   4800
      Width           =   11655
      Begin VB.TextBox txtExtras 
         Height          =   285
         Left            =   9720
         MaxLength       =   15
         TabIndex        =   35
         Top             =   1080
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox txtAlojamiento 
         Height          =   285
         Left            =   8160
         MaxLength       =   15
         TabIndex        =   34
         Top             =   1080
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox txtEmail 
         Height          =   285
         Left            =   8400
         MaxLength       =   40
         TabIndex        =   30
         Top             =   600
         Width           =   3135
      End
      Begin VB.CommandButton botAyudaEmp 
         Caption         =   "?"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6480
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtFax 
         Height          =   285
         Left            =   9960
         MaxLength       =   15
         TabIndex        =   25
         Top             =   240
         Width           =   1560
      End
      Begin VB.TextBox txtTelefono 
         Height          =   285
         Left            =   8400
         MaxLength       =   15
         TabIndex        =   24
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtPersona_Reserva 
         Height          =   285
         Left            =   2760
         MaxLength       =   20
         TabIndex        =   22
         Top             =   240
         Width           =   4095
      End
      Begin VB.TextBox txtNom_Agencia_Empresa 
         Height          =   285
         Left            =   2760
         MaxLength       =   40
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   600
         Width           =   3615
      End
      Begin VB.CheckBox realizadapor 
         Alignment       =   1  'Right Justify
         Caption         =   "Pe&rtenece a empresa"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox txtObservaciones 
         Height          =   975
         Left            =   120
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   32
         Top             =   1200
         Width           =   11415
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Pago alojamiento/E&xtras"
         Height          =   240
         Left            =   8280
         TabIndex        =   33
         Top             =   840
         Visible         =   0   'False
         Width           =   2235
      End
      Begin VB.Label Label7 
         Caption         =   "E-&mail"
         Height          =   255
         Left            =   7080
         TabIndex        =   29
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "&Observaciones"
         Height          =   240
         Left            =   120
         TabIndex        =   31
         Top             =   960
         Width           =   1380
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Telefono/&Fax"
         Height          =   240
         Left            =   7080
         TabIndex        =   23
         Top             =   240
         Width           =   1230
      End
      Begin VB.Label Label9 
         Caption         =   "&Persona que realizo la res."
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Label lblReservaAnulada 
      AutoSize        =   -1  'True
      Caption         =   "Reserva anulada"
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   120
      TabIndex        =   81
      Top             =   7200
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.Menu mnuFormulario 
      Caption         =   "&Formulario"
      Begin VB.Menu mnuFormularioAceptar 
         Caption         =   "Aceptar"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuFormularioCancelar 
         Caption         =   "Cancelar       "
      End
      Begin VB.Menu mnuFormularioImprimir 
         Caption         =   "Imprimir reserva"
         Shortcut        =   ^I
      End
   End
   Begin VB.Menu mnuVer 
      Caption         =   "&Ver información de ..."
      Begin VB.Menu mnuVerHabitacion 
         Caption         =   "Habitación"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuVerTarjeta 
         Caption         =   "Tarjeta de Crédito"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuVerPrePago 
         Caption         =   "Pre pago / Seña"
         Shortcut        =   {F7}
      End
   End
   Begin VB.Menu mnuIrA 
      Caption         =   "&Ir a..."
      Begin VB.Menu mnuVerCuadroSituacion 
         Caption         =   "Cuadro de situación..."
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuVerCuadroDisponibilidad 
         Caption         =   "Cuadro de disponibilidad..."
         Shortcut        =   ^D
      End
   End
   Begin VB.Menu mnuSeleccionar 
      Caption         =   "&Buscar..."
      Begin VB.Menu mnuSeleccionarClientes 
         Caption         =   "Clientes..."
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuSeleccionarEmpresas 
         Caption         =   "Empresas..."
      End
   End
End
Attribute VB_Name = "frmCargaReserva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private nrocliente As Long
Private NroCorrEmp As Long
Private habitacionAsignada As Boolean
Private fi As Date
Private fe As Date
Private fila As Integer

Private Sub Form_Load()
    'inicializo variables
    nrocliente = 0
    NroCorrEmp = 0
    
    'Apariencia formulario
    mSubConfiguroFuentesControlesSistema Me
    subMuestroBotones
    mSubBloqueoControlFormulario Me.txtNom_Agencia_Empresa, True
    mSubBloqueoControlFormulario Me.txtNom_tar, True
    mSubBloqueoControlFormulario Me.txtApe_tar, True
    'inicializo etiquetas de signo dólares
    Me.lblSignoDol(0).Caption = gblSignoDolares
    Me.lblSignoDol(1).Caption = gblSignoDolares
    
    'obtengo número de reserva
    'NOTA: esto del número de reserva esta medio entreverado.
    'a) Cuando se llama a una reserva para hacer todas las operaciones menos el alta
    'se inicializa la variable global nro_reserva
    'b) cuando realizo un alta de reserva, tengo que obtener el número
    If tipo_accion_reserva = "ALTA" Then
        nro_reserva = proxima_reserva
        frmCargaReserva.Caption = "Carga de reserva"
        'cuando realizo un alta no tengo habitaciones para mostrar, de todas maneras
        'tengo que ejecutra este procedimiento para inicializar el control data
        subMuestroHabitaciones
    End If
    
    If tipo_accion_reserva = "MODIFICAR" Then
        frmCargaReserva.Caption = "Modificación de Reserva"
        cargo_tabla_auxiliar
        'si hay habitaciones asignadas no permito modificar la fecha
        habilito_fechas False
        subMuestroHabitaciones
    End If
    
    If tipo_accion_reserva = "ANULAR" Then
        desabilito_formulario
        frmCargaReserva.Caption = "Anulación de reserva"
        subMuestroHabitaciones
    End If
    
    If tipo_accion_reserva = "CONSULTAR" Then
        desabilito_formulario
        frmCargaReserva.Caption = "Consulta de reserva"
        'verifico si la reserva a consultar esta anulada
        If consulta_reserva_anulada Then
            'muestro etiqueta que indica que la reserva esta anulada
            Me.lblReservaAnulada.Visible = True
            'modifico título del formulario
            frmCargaReserva.Caption = "Consulta de reserva (anulada)"
        End If
        subMuestroHabitaciones
        'muestro icono si es noshow
        subMuestroIconoNoShow
    End If
    
    If tipo_accion_reserva = "Check-in" Then
        desabilito_formulario
        frmCargaReserva.Caption = "Ingreso de habitaciones"
        subMuestroHabitaciones
    End If
    
    'cargo tipo habitación
    carga_tipo_hab frmCargaReserva.cboTipo_habitacion
    cboTipo_habitacion.ListIndex = -1
    'cargo_tarjetas credito
    mSubCargoComboTarjetasCredito frmCargaReserva.otras_tar
    otras_tar.ListIndex = 0
    
    If tipo_accion_reserva = "ALTA" Then
    Else
        If consulta_reserva_anulada Then
            cargo_reserva_anulada
        Else
            cargo_reserva
        End If
        'cargamos fechas para que esten disponobles para función de verificar
        fi = fechaing.Text
        fe = fechaegr.Text
    End If
    
    'muestro número de reserva en formulario
    NroReserva.Caption = mFunMuestroNroReserva(nro_reserva)
End Sub

Private Sub subMuestroHabitaciones()
    'Muestro las habitaciones que componen la reserva
    'Es ejecutado en todos los caso para los que utilizo el formulario
    'menos para el de nueva reserva
    Dim consulta As String
    If consulta_reserva_anulada Then
        consulta = "select nomtipohabitacion, tarifa, pasajeros, nrohabitacion, " & _
        "descri_estado, nrocorr, nroreserva from hab_anuladas where nroreserva = " & _
        Str(nro_reserva) & " order by nroreserva,nrocorr"

    Else
        consulta = "select nomtipohabitacion, tarifa, pasajeros, nrohabitacion, " & _
        "descri_estado, nrocorr, nroreserva, noshow from hab_reserva where nroreserva = " & _
        Str(nro_reserva) & " order by nroreserva,nrocorr"
    End If
       
    'inicializo control data
    subInicializoControlData Me.Data1
    Data1.RecordSource = consulta
    ejecuto_consulta
    dibujo_cabezales_grilla
End Sub

Private Sub subMuestroIconoNoShow()
    'Cuando consulto una reserva muestro icono de noShow en caso de que éste se
    'el estado de la reserva.
    Dim filaHab As Integer
    'recorro grilla de habitaciones
    filaHab = 1
    Do While filaHab <= DBGrid1.Rows - 1
        If DBGrid1.TextMatrix(filaHab, 8) = 1 Then  'es noShow
            'muestro ícono
            DBGrid1.col = 5
            DBGrid1.CellPictureAlignment = 1
            Set DBGrid1.CellPicture = frmMAIN.ImageList1.ListImages(8).Picture
        End If
        filaHab = filaHab + 1
    Loop
End Sub

Private Sub Form_Activate()
    'por defecto simpre muestro el tabs de habitaciones
    Me.ssTab1.Tab = 0
    
    If tipo_accion_reserva = "MODIFICAR" Then
        'si hay habitaciones asignadas no permito modificar la fecha
        marco_color_no_asignada
        If habitacionAsignada Then
            habilito_fechas False
        Else
            habilito_fechas True
        End If
        'esto es para que la grilla de habitaciones siempre se dibuje
        'Las cosa que hay que hacer dios mio!!!!!!!!
        DBGrid1.Visible = False
        DBGrid1.Visible = True
    End If
End Sub

Private Sub subMuestroBotones()
    Select Case tipo_accion_reserva
        Case "MODIFICAR"
            noconfirmareserva.Visible = True
            noconfirmareserva.Cancel = True
            confirmareserva.Visible = True
            noconfirmareserva.Left = 10560
            confirmareserva.Left = 9240
            'configuro menu general
            Me.mnuFormularioImprimir.Visible = False
            
        Case "ANULAR"
            Me.AnularReserva.TabIndex = 0
            Me.NoAnular.TabIndex = 1
            AnularReserva.Visible = True
            NoAnular.Visible = True
            NoAnular.Cancel = True
            AnularReserva.Left = 9240
            NoAnular.Left = 10560
            'configuro el menu general
            Me.mnuSeleccionar.Visible = False
            Me.mnuFormularioImprimir.Visible = False
            
        Case "CONSULTAR"
            FinConsulta.Default = True
            FinConsulta.Visible = True
            FinConsulta.Left = 10560
            FinConsulta.Cancel = True
            botImprimir.Visible = True
            botImprimir.Left = 9240
            'configuro el menu general
            Me.mnuSeleccionar.Visible = False
            Me.mnuFormularioCancelar.Visible = False
            
        Case "ALTA"
            noconfirmareserva.Visible = True
            confirmareserva.Visible = True
            
            noconfirmareserva.Left = 10560
            confirmareserva.Left = 9240
            
            noconfirmareserva.Cancel = True
            'configuro el menu general
            Me.mnuFormularioImprimir.Visible = False
            
        Case "Check-in"
            Me.ResCorresponde.Default = True
            Me.ResCorresponde.TabIndex = 0
            Me.ResNoCorresponde.TabIndex = 1
            ResCorresponde.Visible = True
            ResNoCorresponde.Visible = True
            ResCorresponde.Left = 9240
            ResNoCorresponde.Left = 10560
            ResNoCorresponde.Cancel = True
            'configuro el menu general
            Me.mnuSeleccionar.Visible = False
            Me.mnuVer.Visible = False
            Me.mnuIrA.Visible = False
            Me.mnuFormularioImprimir.Visible = False
    End Select
End Sub

Private Sub ejecuto_consulta()
    Data1.Refresh
    fila = DBGrid1.Row
End Sub

Private Sub calculo_noches()
    cantidadnoches.Caption = fe - fi
End Sub

Private Sub AnularReserva_Click()
    Dim campo As Integer
    Dim consulta As String
    'pido confirmación de usuario para anular la reserva
    If mFunMensaje(4, 52) Then
        'copio registro desde tbRESERVAS hacia tbRES_ANULADAS
        tbANULADAS.AddNew
            For campo = 0 To tbRESERVAS.Fields.Count - 1
                tbANULADAS(campo) = tbRESERVAS(campo)
            Next campo
            'verifico si hay usuarios definidos
            If tbPARAMETROS("SisAdminTF") <> 0 Then
                'hay usuarios definidos
                tbANULADAS("ultimaOprUsr") = "Anulada x " & m_UsuarioSisNom
            Else
                tbANULADAS("ultimaOprUsr") = "Anulada"
            End If
            tbANULADAS("ultimaOprFecha") = m_FechaSis
        tbANULADAS.Update
        
        'copio registros desde tbHAB_RESERVA hacia tbHAB_ANULADAS
        tbHAB_RESERVAS.Index = "ihab_reserva"
        tbHAB_RESERVAS.Seek ">=", nro_reserva, 0
        If Not tbHAB_RESERVAS.NoMatch Then
            Do While Not tbHAB_RESERVAS.EOF
                If tbHAB_RESERVAS("nroreserva") = nro_reserva Then
                    tbHAB_ANULADAS.AddNew
                        For campo = 0 To tbHAB_RESERVAS.Fields.Count - 1
                            tbHAB_ANULADAS(campo) = tbHAB_RESERVAS(campo)
                        Next campo
                    tbHAB_ANULADAS.Update
                Else
                    Exit Do
                End If
                tbHAB_RESERVAS.MoveNext
            Loop
        End If

        'borro registro de tbRESERVAS
        consulta = "DELETE FROM reservas WHERE nroreserva= " & Str(nro_reserva)
        bdHOTEL.Execute consulta
       
        'borro registros de tbHAB_RESESERVAS
        consulta = "DELETE FROM hab_reserva WHERE nroreserva = " & Str(nro_reserva)
        bdHOTEL.Execute consulta

        'grabo botacora
        GraboBitacora Mid(NroReserva, 7, 10)
        
        Unload Me
    End If
End Sub

Private Sub modificar()
    Dim res As Long, corr As Integer
    DBGrid1.Row = fila
    DBGrid1.col = 7
    res = DBGrid1.Text
    DBGrid1.col = 6
    corr = DBGrid1.Text
  
    tbHAB_RESERVAS.Index = "ihab_reserva"
    tbHAB_RESERVAS.Seek "=", res, corr
    If Not tbHAB_RESERVAS.NoMatch Then
        tbHAB_RESERVAS.Edit
            tbHAB_RESERVAS("tarifa") = Val(txttarifa.Text)
            tbHAB_RESERVAS("nrohabitacion") = Val(txtHab.Text)
            If Val(txtHab.Text) <> 0 Then
                tbHAB_RESERVAS("descri_estado") = "Asignada"
            Else
                tbHAB_RESERVAS("descri_estado") = "No Asignada"
            End If
            tbHAB_RESERVAS("pasajeros") = Val(txtPasajeros.Text)
        tbHAB_RESERVAS.Update
    End If
End Sub

Private Sub BotAyuda_Click()
    Dim nrocliaux As String
    nrocliaux = mFunBusqueda(1) 'todos los pasajeros
    If Val(nrocliaux) <> 0 Then
        nrocliente = nrocliaux
        busco_cliente
    End If
End Sub

Private Sub botAyudaEmp_Click()
    Dim nro_corr_aux As String
    nro_corr_aux = mFunBusqueda(3)  'empresas
    If Val(nro_corr_aux) <> 0 Then
        NroCorrEmp = nro_corr_aux
        If busco_empTF(NroCorrEmp) Then
            txtNom_Agencia_Empresa.Text = tbEMPRESAS("nomemp")
        End If
    End If
End Sub

Private Sub botHabitacion_Click()
    If FunValidoFechas Then
        If cboTipo_habitacion.ListIndex <> -1 Then
            tipo_accion_SeleccionHab = 1
            frmReservaSeleHab.Show 1
            If Not cancelo_seleccion_habitaciones Then
                Confirmar
                'calculo la cantidad de noches de la reserva
                calculo_noches
            End If
        Else
            'debe de seleccionar tipo de habitación
            mSubMensaje 4, 53
        End If
   End If
   'por fin!! con esto no se borra la grilla después de llamar al formulario de
   'selección de habitación
   Me.DBGrid1.Visible = False
   Me.DBGrid1.Visible = True
End Sub

Private Function FunValidoFechas()
    'Valido que las fecha de ingreso y egreso de la reserva sean correctas
    'Al ingresar la fecha es posible ingresarla sin los caracteres de de separacion (/)
    'con el fin de agilizar el proceso.
    'La función formo_fecha, introduce estos caracteres para formar una fecha válida
    fechaegr.Text = formo_fecha(fechaegr.Text)
    fechaing.Text = formo_fecha(fechaing.Text)
    FunValidoFechas = True
    'Valido formato de fechas
    If IsDate(fechaegr.Text) And IsDate(fechaing.Text) Then
        'cargo variables generales de fecha
        fi = fechaing.Text
        fe = fechaegr.Text
    Else
        'formato de fechas incorreco
        mSubMensaje 3, 1
        fechaegr.SetFocus
        FunValidoFechas = False
        Exit Function
    End If
    
    'Valido que la fecha inicial no sea menor a la del sistema
    If fi < m_FechaSis Then
        'la fecha de inicio no puede ser menor a la fecha de hoy
        mSubMensaje 3, 2
        fechaing.SetFocus
        FunValidoFechas = False
        Exit Function
    End If
    
    'Valido que la fecha de ingreso no sea mayor o igual a la fecha de egreso
    If fi >= fe Then
        'la fecha de ingreso no puede ser mayor e igual a la fecha de egreso
        mSubMensaje 4, 54
        fechaing.SetFocus
        FunValidoFechas = False
        Exit Function
    End If
End Function

Private Sub dbgrid1_KeyPress(KeyAscii As Integer)
    'Simulo que doy doble click sobre la fila.
    'De esta manera permito seleccionar registro con la tecla Enter
    If KeyAscii = 13 Then
        dbgrid1_DblClick
    End If
End Sub
    
Private Sub cboTipo_habitacion_Click()
    'Cuando se selecciona un tipo de habitación muestro la tarifa de la misma.
    Me.txttarifa.Text = mFunBuscoTarifaHab(Me.cboTipo_habitacion.ItemData(Me.cboTipo_habitacion.ListIndex))
End Sub

Private Sub dbgrid1_DblClick()
    Dim colorback As String
    Dim colorfore As String
    
    fila = DBGrid1.Row
    If DBGrid1.Rows > 1 Then
        If DBGrid1.CellBackColor = mSisColor_15FilaSeleccionada Then    'si esta marcado
            colorback = &H80000005                  'desmarco
            colorfore = &H80000008
            Eliminar_hab.Enabled = False
            Me.botHabitacion.Enabled = True
        Else                                        'marco
            colorback = mSisColor_15FilaSeleccionada
            colorfore = mSisColor_19FilaSeleccionadaTexto
            Eliminar_hab.Enabled = True
            Me.botHabitacion.Enabled = False
        End If
        'ejecuto la consulta para que se limpien todas las filas es decir para que
        'todas queden como si no estuvieran marcadas
        Data1.Refresh
        DBGrid1.Row = fila
        DBGrid1.col = 1
        DBGrid1.CellBackColor = colorback
        DBGrid1.CellForeColor = colorfore
        DBGrid1.col = 2
        DBGrid1.CellBackColor = colorback
        DBGrid1.CellForeColor = colorfore
        DBGrid1.col = 3
        DBGrid1.CellBackColor = colorback
        DBGrid1.CellForeColor = colorfore
        DBGrid1.col = 4
        DBGrid1.CellBackColor = colorback
        DBGrid1.CellForeColor = colorfore
        DBGrid1.col = 5
        DBGrid1.CellBackColor = colorback
        DBGrid1.CellForeColor = colorfore
        DBGrid1.col = 6
        DBGrid1.CellBackColor = colorback
        DBGrid1.CellForeColor = colorfore
    
        dibujo_cabezales_grilla
        marco_color_no_asignada
    End If
End Sub

Private Sub muestro_datos_hab(vacio As Boolean)
    If vacio Then
        'la tarifa no tiene porque quedar en blanco, es mejor inicializarla con la tarifa
        'del combo de tipo de habitaciones.
        txttarifa.Text = mFunBuscoTarifaHab(Me.cboTipo_habitacion.ItemData(Me.cboTipo_habitacion.ListIndex))
        txtPasajeros.Text = ""
        txtHab.Text = ""
    End If
End Sub

Private Sub Eliminar_hab_Click()
    Dim consulta As String
    DBGrid1.col = 6
    DBGrid1.Row = fila
    
    If DBGrid1.CellBackColor = mSisColor_15FilaSeleccionada Then    ' si esta marcado
        If DBGrid1.Text <> "" Then
            consulta = "DELETE FROM hab_reserva where nroreserva = " + Str(nro_reserva) + _
            "and nrocorr =" + Str(DBGrid1.Text)
            bdHOTEL.Execute consulta
            ejecuto_consulta
            dibujo_cabezales_grilla
            marco_color_no_asignada
        End If
        If habitacionAsignada Then
            habilito_fechas False
        Else
            habilito_fechas True
        End If
        Eliminar_hab.Enabled = False
        Me.botHabitacion.Enabled = True
    End If
End Sub

Private Sub fechaegr_KeyPress(KeyAscii As Integer)
    CapturoEnter KeyAscii
End Sub

Private Sub fechaing_KeyPress(KeyAscii As Integer)
    CapturoEnter KeyAscii
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Descargo formulario de memoria
    Set frmCargaReserva = Nothing
End Sub

Private Sub ResNoCorresponde_Click()
    Unload Me
End Sub

Private Sub ResCorresponde_Click()
    Unload Me
    frmCheck_in.Show 1
End Sub

Private Sub grabo_hab_reservas(hab As Integer)
    Dim corr_aux As Long
    
    If IsDate(fechaing.Text) And IsDate(fechaegr.Text) Then
        If cboTipo_habitacion.ListIndex <> -1 Then
            corr_aux = obtengo_ultimo_corr
            tbHAB_RESERVAS.AddNew
            tbHAB_RESERVAS("nroreserva") = nro_reserva
            tbHAB_RESERVAS("nrocorr") = corr_aux
            tbHAB_RESERVAS("tarifa") = Val(txttarifa.Text)
            tbHAB_RESERVAS("nrohabitacion") = hab
            If hab <> 0 Then
                tbHAB_RESERVAS("descri_estado") = "Asignada"
            Else
                tbHAB_RESERVAS("descri_estado") = "No Asignada"
            End If
            'no permito ingresar cantidad de pasajeros = 0, ya que no es lógico que se
            'realize una reserva sin pasajeros.
            If Val(txtPasajeros.Text) = 0 Then
                tbHAB_RESERVAS("pasajeros") = 1
            Else
                tbHAB_RESERVAS("pasajeros") = Val(txtPasajeros.Text)
            End If
            tbHAB_RESERVAS("tipohabitacion") = cboTipo_habitacion.ItemData(cboTipo_habitacion.ListIndex)
            tbHAB_RESERVAS("nomtipohabitacion") = cboTipo_habitacion.Text
            
            tbHAB_RESERVAS("fechaing") = fi
            tbHAB_RESERVAS("fechaegr") = fe
             
            tbHAB_RESERVAS.Update
        End If
    End If
End Sub

Private Sub confirmareserva_Click()
    Dim res_aux As String
    Dim tipoMsg As Integer
    'determino el tipo de mensaje a mostrar cuando confirmo la reserva
    If tipo_accion_reserva = "ALTA" Then  'si es un alta
        tipoMsg = 57    'aviso de cofirmación de ingreso
    Else
        If tipo_accion_reserva = "MODIFICAR" Then 'si es una modificacion
            tipoMsg = 124   'aviso de confirmación de modificación
        End If
    End If
    If valido_datos Then
        'pido confirmación de usuario antes de realizar la reserva
        If mFunMensaje(4, tipoMsg) Then
            'comienza en el lugar 6 ya que la conversión str agrega un caracter
            res_aux = Mid(Str(nro_reserva), 6, 5)
            If tipo_accion_reserva = "ALTA" Then  'si es un alta
                sumo_corr_reserva
                tbRESERVAS.AddNew
                    tbRESERVAS("nroreserva") = nro_reserva
                    grabo_reserva
                    'grabo registros de control
                    subGraboResgistrosControl "ALTA"
                tbRESERVAS.Update
                'aviso de confirmación de realizada la reserva
                mSubMensaje 4, 56, res_aux    'se realizo la reserva
                
                'imprimo nueva reserva
                subImprimirReserva "ALTA", nro_reserva, res_aux
                
                'grabo en bitacora
                GraboBitacora "Res.Nro. " & Str(res_aux)
            Else
                If tipo_accion_reserva = "MODIFICAR" Then 'si es una modificacion
                    tbRESERVAS.Edit
                        grabo_reserva
                        subGraboResgistrosControl "MODIFICAR"
                    tbRESERVAS.Update
                    'aviso de confirmación de modificación de reserva
                    mSubMensaje 4, 132, res_aux    'se modifico la reserva
                    
                    'imprimo reserva modificada
                    subImprimirReserva "MODIFICAR", nro_reserva, res_aux
                
                    'grabo en bitacora
                    GraboBitacora "Res.Nro. " & Str(res_aux)
                End If
            End If
            Unload Me
        End If
    End If
End Sub

Private Sub subGraboResgistrosControl(tipoAccion As String)
    '---------------------------------------------------------------------------------
    'Graba información de la última tarea realizada con la reserva, fecha y usuario que
    'realizó la misma.
    'Es necesario que esté instalada la aplicación perfiles y que exista al menos
    'un usuario definido, si no es así no se graba el campo nombre de usuario.
    '
    'NOTA: la fecha de realizada la operación, es la fecha del sistema.
    '----------------------------------------------------------------------------------
    'Parámetros:
    '       Entrada: "ALTA"         = estoy realizando una nueva reserva
    '                "MODIFICAR"    = estoy realizando una modificación
    '----------------------------------------------------------------------------------
    Dim oprRealizada As String
    
    'obtengo tipo de operación
    If tipoAccion = "ALTA" Then
        oprRealizada = "Realizada"
    Else
        If tipoAccion = "MODIFICAR" Then
            oprRealizada = "Modificada"
        End If
    End If
    'verifico si hay usuarios definidos
    If tbPARAMETROS("SisAdminTF") <> 0 Then
        'hay usuarios definidos
        oprRealizada = oprRealizada & " x " & m_UsuarioSisNom
    End If
    'grabo operación realizada y usuario
    tbRESERVAS("ultimaOprUsr") = oprRealizada
    tbRESERVAS("ultimaOprFecha") = Date
End Sub
       
Private Sub FinConsulta_Click()
    Unload Me
End Sub

Private Sub dibujo_cabezales_grilla()
     DBGrid1.FormatString = "   |Tipo Habitación     " & _
                            "| Tarifa   " & _
                            "| Cant. Pax " & _
                            "| Habitación " & _
                            "| Estado               " & _
                            "|Nrocorr|norreserva|noshow "
     DBGrid1.ColWidth(6) = 0    'culto la columna nrocorr
     DBGrid1.ColWidth(7) = 0    'culto la columna nroreserva
     DBGrid1.ColWidth(8) = 0    'culto la columna noShow
End Sub

Private Sub marco_color_no_asignada()
    Dim i As Integer
     DBGrid1.col = 5
     i = 1
     habitacionAsignada = False
     Do While i < DBGrid1.Rows
        DBGrid1.Row = i
        If DBGrid1.Text = "No Asignada" Then
            DBGrid1.CellForeColor = &HFF&
        Else
            habitacionAsignada = True
        End If
        i = i + 1
     Loop
End Sub

Private Sub NoAnular_Click()
    Unload Me
End Sub

Private Sub noconfirmareserva_Click()
    Dim consulta As String
    If tipo_accion_reserva = "ALTA" Then
        'pido confirmación al usuario antes de salir de la reserva
        If mFunMensaje(4, 58) Then
            'borro datos en HAB_RESERVA
            consulta = "DELETE FROM hab_reserva where nroreserva = " + Str(nro_reserva)
            bdHOTEL.Execute consulta
            ejecuto_consulta
            Unload Me
        End If
    End If
    If tipo_accion_reserva = "MODIFICAR" Then
        'pido confirmación al usuario antes de salir de la reserva
        If mFunMensaje(4, 59) Then
            retomo_datos_anteriores_habitacion
            Unload Me
        End If
    End If
End Sub

Private Sub busco_cliente()
    tbCLIENTES.Index = "iclie_nrocorr"
    tbCLIENTES.Seek "=", nrocliente
    If Not tbCLIENTES.NoMatch Then
        cargo_datos_formulario
    End If
End Sub

Private Sub cargo_datos_formulario()
    txt1er_Nom_titular.Text = tbCLIENTES("primer_nom_titular")
    txt2do_Nom_titular.Text = tbCLIENTES("segundo_nom_titular")
    txt1er_Ape_titular.Text = tbCLIENTES("primer_ape_titular")
    txt2do_Ape_titular.Text = tbCLIENTES("segundo_ape_titular")
    txtTelefono.Text = tbCLIENTES("tel_titular")
    txtFax.Text = tbCLIENTES("fax_titular")
    txtEmail.Text = tbCLIENTES("email_titular")
End Sub

Private Sub realizadapor_Click()
    If realizadapor.Value = 1 Then
        'cuando se carga el click desde la base de datos, se ejecuta este evento por lo
        'que se agregan estos controles para que no quede mal habilitado el boton de ayuda de empresas.
        If tipo_accion_reserva <> "ANULAR" And _
        tipo_accion_reserva <> "CONSULTAR" And _
        tipo_accion_reserva <> "Check-in" Then
            botAyudaEmp.Enabled = True
        End If
    Else
        NroCorrEmp = 0
        botAyudaEmp.Enabled = False
        txtNom_Agencia_Empresa.Text = ""
    End If
End Sub

Private Sub grabo_reserva()
    Dim nom_aux As String
    
    If IsDate(fecha_emi_seña.Text) Then tbRESERVAS("fechaemision") = fecha_emi_seña.Value
    
    'NOTA: acerca de tbRESERVAS("fechaVto")
    'Este campo guarda la fecha de vencimiento, su tipo es un string.
    'Como el campo se carga desde un textBox que utiliza una máscara de entrada,
    'cuando se ingresa una fecha en blanco se graba en el campo el formato de la máscara
    ' (__/___) lo que origina que la impresión de la reserva no sea muy prolija.
    'Por ese motivo cuando no se graba una fecha válida se inicializa el campo a Empty
    If IsDate(fech_vto_tar.Text) Then
        tbRESERVAS("fechavto") = fech_vto_tar.Text
    Else
        tbRESERVAS("fechavto") = Empty
    End If
    
    tbRESERVAS("fechaing") = fi
    tbRESERVAS("fechaegr") = fe
    
    If txtObservaciones.Text <> "" Then tbRESERVAS("observaciones") = txtObservaciones.Text
    
    tbRESERVAS("cantnoches") = Val(cantidadnoches)
    tbRESERVAS("personarealizoreserva") = txtPersona_Reserva.Text
    If Trim(txtNom_Agencia_Empresa.Text) = "" Then
        tbRESERVAS("agenciaempresa") = 0
        tbRESERVAS("nroagenciaempresa") = 0
    Else
        tbRESERVAS("agenciaempresa") = 1
        tbRESERVAS("nroagenciaempresa") = NroCorrEmp
    End If
    tbRESERVAS("telefono") = txtTelefono.Text
    tbRESERVAS("fax") = txtFax.Text
    tbRESERVAS("email") = txtEmail.Text
    tbRESERVAS("pagoalojamiento") = txtAlojamiento.Text
    tbRESERVAS("pagoextras") = txtExtras.Text
    tbRESERVAS("tarjetaCredito") = otras_tar.ItemData(otras_tar.ListIndex)
    tbRESERVAS("nombres") = txtNom_tar.Text
    tbRESERVAS("apellidos") = txtApe_tar.Text
    tbRESERVAS("codseg") = txtCod_Seg_tar.Text
    tbRESERVAS("nrotarjeta") = txtNro_tar.Text
    tbRESERVAS("importe") = Val(txtImporte_Seña.Text)
    tbRESERVAS("nrorecibo") = Val(txtRecibo_Seña.Text)
    tbRESERVAS("nombreseña") = txtNom_Seña.Text
    tbRESERVAS("apellidoseña") = txtApe_Seña.Text
    
    nom_aux = StrConv(txt1er_Nom_titular.Text, 2)
    tbRESERVAS("primer_nom_titular") = StrConv(nom_aux, 3)
    nom_aux = StrConv(txt2do_Nom_titular.Text, 2)
    tbRESERVAS("segundo_nom_titular") = StrConv(nom_aux, 3)
    tbRESERVAS("primer_ape_titular") = StrConv(txt1er_Ape_titular.Text, 1)
    tbRESERVAS("segundo_ape_titular") = StrConv(txt2do_Ape_titular.Text, 1)
    'grabo el número de pasajero que realizo la reserva
    'este valor se inicializa si se llamo a la ayuda
    tbRESERVAS("nroCorrCli") = nrocliente
End Sub

Private Sub cargo_reserva()
     Dim nom_aux As String
     
     If tbRESERVAS("fechaing") <> 0 Then fechaing.Text = tbRESERVAS("fechaing")
     If tbRESERVAS("fechaegr") <> 0 Then fechaegr.Text = tbRESERVAS("fechaegr")

     If Not IsNull(tbRESERVAS("fechavto")) Then fech_vto_tar.Text = tbRESERVAS("fechavto")


     If tbRESERVAS("fechaemision") <> 0 Then fecha_emi_seña.Value = tbRESERVAS("fechaemision")
     If tbRESERVAS("observaciones") <> "" Then txtObservaciones.Text = tbRESERVAS("observaciones")
     
     cantidadnoches = tbRESERVAS("cantnoches")
     txtPersona_Reserva.Text = tbRESERVAS("personarealizoreserva")
     realizadapor.Value = tbRESERVAS("agenciaempresa")
     If tbRESERVAS("agenciaempresa") = 1 Then
        If busco_empTF(tbRESERVAS("nroagenciaempresa")) Then
            txtNom_Agencia_Empresa.Text = tbEMPRESAS("nomemp")
            NroCorrEmp = tbEMPRESAS("nroCorrEmp")
        End If
     End If
     txtTelefono.Text = tbRESERVAS("telefono")
     txtFax.Text = tbRESERVAS("fax")
     txtEmail.Text = tbRESERVAS("email")
     txtAlojamiento.Text = tbRESERVAS("pagoalojamiento")
     txtExtras.Text = tbRESERVAS("pagoextras")
     
     posiciono_combo otras_tar, tbRESERVAS("tarjetaCredito")
   
     txtNom_tar.Text = tbRESERVAS("nombres")
     txtApe_tar.Text = tbRESERVAS("apellidos")
     txtCod_Seg_tar.Text = tbRESERVAS("codseg")
     txtNro_tar.Text = tbRESERVAS("nrotarjeta")
     txtImporte_Seña.Text = tbRESERVAS("importe")
     txtRecibo_Seña.Text = tbRESERVAS("nrorecibo")
     txtNom_Seña.Text = tbRESERVAS("nombreseña")
     txtApe_Seña.Text = tbRESERVAS("apellidoseña")
     
     nom_aux = StrConv(tbRESERVAS("primer_nom_titular"), 2)
     txt1er_Nom_titular.Text = StrConv(nom_aux, 3)
     
     nom_aux = StrConv(tbRESERVAS("segundo_nom_titular"), 2)
     txt2do_Nom_titular.Text = StrConv(nom_aux, 3)
     
     txt1er_Ape_titular.Text = StrConv(tbRESERVAS("primer_ape_titular"), 1)
     txt2do_Ape_titular.Text = StrConv(tbRESERVAS("segundo_ape_titular"), 1)
End Sub

Private Sub cargo_reserva_anulada()
         Dim nom_aux As String
     
     If tbANULADAS("fechaing") <> 0 Then fechaing.Text = tbANULADAS("fechaing")
     If tbANULADAS("fechaegr") <> 0 Then fechaegr.Text = tbANULADAS("fechaegr")

     fech_vto_tar.Text = tbANULADAS("fechavto")


     If tbANULADAS("fechaemision") <> 0 Then fecha_emi_seña.Value = tbANULADAS("fechaemision")
     If tbANULADAS("observaciones") <> "" Then txtObservaciones.Text = tbANULADAS("observaciones")
     
     cantidadnoches = tbANULADAS("cantnoches")
     txtPersona_Reserva.Text = tbANULADAS("personarealizoreserva")
     realizadapor.Value = tbANULADAS("agenciaempresa")
     If tbANULADAS("agenciaempresa") = 1 Then
        If busco_empTF(tbANULADAS("nroagenciaempresa")) Then
            txtNom_Agencia_Empresa.Text = tbEMPRESAS("nomemp")
        End If
     End If
     txtTelefono.Text = tbANULADAS("telefono")
     txtFax.Text = tbANULADAS("fax")
     txtEmail.Text = tbANULADAS("email")
     txtAlojamiento.Text = tbANULADAS("pagoalojamiento")
     txtExtras.Text = tbANULADAS("pagoextras")
     
     posiciono_combo otras_tar, tbANULADAS("tarjetaCredito")
   
     txtNom_tar.Text = tbANULADAS("nombres")
     txtApe_tar.Text = tbANULADAS("apellidos")
     txtCod_Seg_tar.Text = tbANULADAS("codseg")
     txtNro_tar.Text = tbANULADAS("nrotarjeta")
     txtImporte_Seña.Text = tbANULADAS("importe")
     txtRecibo_Seña.Text = tbANULADAS("nrorecibo")
     txtNom_Seña.Text = tbANULADAS("nombreseña")
     txtApe_Seña.Text = tbANULADAS("apellidoseña")
     
     nom_aux = StrConv(tbANULADAS("primer_nom_titular"), 2)
     txt1er_Nom_titular.Text = StrConv(nom_aux, 3)
     
     nom_aux = StrConv(tbANULADAS("segundo_nom_titular"), 2)
     txt2do_Nom_titular.Text = StrConv(nom_aux, 3)
     
     txt1er_Ape_titular.Text = StrConv(tbANULADAS("primer_ape_titular"), 1)
     txt2do_Ape_titular.Text = StrConv(tbANULADAS("segundo_ape_titular"), 1)

End Sub

Private Function obtengo_ultimo_corr()
    obtengo_ultimo_corr = 1
    tbHAB_RESERVAS.Index = "ihab_reserva"
    tbHAB_RESERVAS.Seek ">=", nro_reserva, 1
    If Not tbHAB_RESERVAS.NoMatch Then  'existe
        Do While Not tbHAB_RESERVAS.EOF
            If tbHAB_RESERVAS("nroreserva") = nro_reserva Then
                obtengo_ultimo_corr = tbHAB_RESERVAS("nrocorr") + 1
            Else
                Exit Do
            End If
            tbHAB_RESERVAS.MoveNext
        Loop
    End If
End Function

Private Function valido_datos()
    valido_datos = True
    If IsDate(fechaing.Text) And IsDate(fechaegr.Text) Then
        If fi > fe Then
            'La fecha de ingreso no puede ser mayor a la de egreso
            mSubMensaje 3, 3
            fechaing.SetFocus
            valido_datos = False
            Exit Function
        End If
        If fi < m_FechaSis Then
            'la fecha de ingreso no puede ser menor a la fecha de hoy
            mSubMensaje 3, 2
            fechaing.SetFocus
            valido_datos = False
            Exit Function
        End If
    Else
        If Not IsDate(fechaing.Text) Then
            'El formato de la fecha de ingreso no es válido
            mSubMensaje 3, 1
            fechaing.SetFocus
            valido_datos = False
            Exit Function
        End If
        If Not IsDate(fechaegr.Text) Then
            'El formato de la fecha de egreso no es válido
            mSubMensaje 3, 1
            fechaegr.SetFocus
            valido_datos = False
            Exit Function
        End If
    End If
    If txt1er_Ape_titular.Text = "" Then
        'debe ingresar al menos 1er Nombre y 1er apellido continuar
        mSubMensaje 4, 60
        txt1er_Ape_titular.SetFocus
        valido_datos = False
    Exit Function
    End If
    
    If DBGrid1.Rows = 1 Then
        'debe de ingresar al menos un tipo de habitación
        mSubMensaje 4, 61
        cboTipo_habitacion.SetFocus
        valido_datos = False
    End If
End Function

Private Sub cargo_tabla_auxiliar()
   borro_datos_en_tabla_auxiliar
   tbHAB_RESERVAS.Index = "ihab_reserva"
   tbHAB_RESERVAS.Seek ">=", nro_reserva, 1
   If Not tbHAB_RESERVAS.NoMatch Then  'existe
       Do While Not tbHAB_RESERVAS.EOF
           If tbHAB_RESERVAS("nroreserva") = nro_reserva Then
                tbHAB_RESERVAS_AUX.AddNew
                tbHAB_RESERVAS_AUX("nroreserva") = tbHAB_RESERVAS("nroreserva")
                tbHAB_RESERVAS_AUX("nrocorr") = tbHAB_RESERVAS("nrocorr")
                tbHAB_RESERVAS_AUX("nrohabitacion") = tbHAB_RESERVAS("nrohabitacion")
                tbHAB_RESERVAS_AUX("tarifa") = tbHAB_RESERVAS("tarifa")
                tbHAB_RESERVAS_AUX("pasajeros") = tbHAB_RESERVAS("pasajeros")
                tbHAB_RESERVAS_AUX("tipohabitacion") = tbHAB_RESERVAS("tipohabitacion")
                tbHAB_RESERVAS_AUX("nomtipohabitacion") = tbHAB_RESERVAS("nomtipohabitacion")
                tbHAB_RESERVAS_AUX("fechaing") = tbHAB_RESERVAS("fechaing")
                tbHAB_RESERVAS_AUX("fechaegr") = tbHAB_RESERVAS("fechaegr")
                tbHAB_RESERVAS_AUX("descri_estado") = tbHAB_RESERVAS("descri_estado")
                tbHAB_RESERVAS_AUX.Update
           Else
               Exit Do
           End If
           tbHAB_RESERVAS.MoveNext
       Loop
   End If
End Sub

Private Sub retomo_datos_anteriores_habitacion()
   'borro datos en tabla
   tbHAB_RESERVAS.Index = "ihab_reserva"
   tbHAB_RESERVAS.Seek ">=", nro_reserva, 1
   If Not tbHAB_RESERVAS.NoMatch Then  'existe
       Do While Not tbHAB_RESERVAS.EOF
           If tbHAB_RESERVAS("nroreserva") = nro_reserva Then
               tbHAB_RESERVAS.Delete
           Else
               Exit Do
           End If
           tbHAB_RESERVAS.MoveNext
       Loop
   End If
   
   tbHAB_RESERVAS_AUX.Index = "ihab_reserva"
   tbHAB_RESERVAS_AUX.Seek ">=", nro_reserva, 1
   If Not tbHAB_RESERVAS_AUX.NoMatch Then  'existe
       Do While Not tbHAB_RESERVAS_AUX.EOF
           If tbHAB_RESERVAS_AUX("nroreserva") = nro_reserva Then
                'verifico que no exista registro en tabla original
                tbHAB_RESERVAS.Index = "ihab_reserva"
                tbHAB_RESERVAS.Seek "=", tbHAB_RESERVAS_AUX("nroreserva"), tbHAB_RESERVAS_AUX("nrocorr")
                If Not tbHAB_RESERVAS.NoMatch Then  'existe
                    tbHAB_RESERVAS.Edit
                Else
                    tbHAB_RESERVAS.AddNew
                End If
                tbHAB_RESERVAS("nroreserva") = tbHAB_RESERVAS_AUX("nroreserva")
                tbHAB_RESERVAS("nrocorr") = tbHAB_RESERVAS_AUX("nrocorr")
                tbHAB_RESERVAS("nrohabitacion") = tbHAB_RESERVAS_AUX("nrohabitacion")
                tbHAB_RESERVAS("tarifa") = tbHAB_RESERVAS_AUX("tarifa")
                tbHAB_RESERVAS("pasajeros") = tbHAB_RESERVAS_AUX("pasajeros")
                tbHAB_RESERVAS("tipohabitacion") = tbHAB_RESERVAS_AUX("tipohabitacion")
                tbHAB_RESERVAS("nomtipohabitacion") = tbHAB_RESERVAS_AUX("nomtipohabitacion")
                tbHAB_RESERVAS("fechaing") = tbHAB_RESERVAS_AUX("fechaing")
                tbHAB_RESERVAS("fechaegr") = tbHAB_RESERVAS_AUX("fechaegr")
                tbHAB_RESERVAS("descri_estado") = tbHAB_RESERVAS_AUX("descri_estado")
                tbHAB_RESERVAS.Update
           Else
               Exit Do
           End If
           tbHAB_RESERVAS_AUX.MoveNext
       Loop
       borro_datos_en_tabla_auxiliar
   End If
End Sub

Private Sub borro_datos_en_tabla_auxiliar()
    tbHAB_RESERVAS_AUX.Index = "ihab_reserva"
    tbHAB_RESERVAS_AUX.Seek ">=", nro_reserva, 1
    If Not tbHAB_RESERVAS_AUX.NoMatch Then  'existe
        Do While Not tbHAB_RESERVAS_AUX.EOF
            If tbHAB_RESERVAS_AUX("nroreserva") = nro_reserva Then
                tbHAB_RESERVAS_AUX.Delete
            Else
                Exit Do
            End If
            tbHAB_RESERVAS_AUX.MoveNext
        Loop
    End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    'Ocurre cuando se cambia el tabs visible.
    
    'desabilito todos los tasb del control
    Me.Frame6.Enabled = False   'tabs habitación
    Me.Frame5.Enabled = False   'tabs tarjeta
    Me.Frame3.Enabled = False   'tabs pago deseña
    'determino que tabs esta visible
    Select Case ssTab1.TabCaption(ssTab1.Tab)
        Case "Habitación"
            'permito trabajar con frame habitación
            Me.Frame6.Enabled = True
        Case "Tarjeta de Crédito"
            'permito trabajar con frame tarjeta de crédito
            Me.Frame5.Enabled = True
        Case "Pre Pago / Seña"
            'permito trabajar con frame pre pago
            Me.Frame3.Enabled = True
    End Select
End Sub

Private Sub txt1er_Nom_titular_LostFocus()
    txt1er_Nom_titular.Text = mFunFormatoNombre(txt1er_Nom_titular.Text)
    'asistencia a usuarios
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub txt2do_Nom_titular_LostFocus()
    txt2do_Nom_titular.Text = mFunFormatoNombre(txt2do_Nom_titular.Text)
    'asistencia a usuarios
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub txt1er_ape_titular_LostFocus()
    txt1er_Ape_titular.Text = mFunFormatoApellido(txt1er_Ape_titular.Text)
    'asistencia a usuarios
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub txt2do_ape_titular_LostFocus()
    txt2do_Ape_titular.Text = mFunFormatoApellido(txt2do_Ape_titular.Text)
    'asistencia a usuarios
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub desabilito_formulario()
    'botones
    Me.botAyuda.Enabled = False
    Me.botAyudaEmp.Enabled = False
    Me.Eliminar_hab.Enabled = False
    Me.botHabitacion.Enabled = False
    
    'grilla
    Me.DBGrid1.BackColor = mSisColor_18ControlesNoHabilitados
    Me.DBGrid1.TabStop = False
    
    Me.fech_vto_tar.Enabled = False
    Me.fecha_emi_seña.Enabled = False
    Me.fechaegr.Enabled = False
    Me.fechaing.Enabled = False
    
    Me.realizadapor.Enabled = False
    Me.cboTipo_habitacion.Locked = True
    
    mSub_bloqueo_controles_formulario Me, True
End Sub

Private Sub habilito_fechas(x As Boolean)
    Dim opuesto As Boolean
    Dim color As String
    
    If x Then
        opuesto = False
        color = &H80000005
    Else
        opuesto = True
        color = &H80000016
    End If
    
    Me.fechaegr.Enabled = x
    Me.fechaing.Enabled = x
    
    Me.fechaing.BackColor = color
    Me.fechaegr.BackColor = color
    Me.botAyuda.Enabled = x
End Sub

Private Sub txtApe_Seña_Change()
    'Cuando cambio el contenido de este campo, copio su valor al campo apellido de la
    'tarjeta de crédito
    Me.txtApe_tar.Text = txtApe_Seña.Text
End Sub

Private Sub txtNom_Seña_Change()
    'Cuando cambio el contenido de este campo, copio su valor al campo nombre de la
    'tarjeta de crédito
    Me.txtNom_tar.Text = txtNom_Seña.Text
End Sub

Private Sub txtPasajeros_KeyPress(KeyAscii As Integer)
    ValidoNum KeyAscii, False, False
End Sub

Private Sub txttarifa_KeyPress(KeyAscii As Integer)
    ValidoNum KeyAscii, False, True
End Sub

'**************************************************************************************
'
'  Impresión de reservas
'
'***************************************************************************************

Private Sub subImprimirReserva(tipoImpresion As String, nroReservaLis As Long, _
                                resAux As String)
    '----------------------------------------------------------------------------------
    'Imprime el formulario de reserva.
    '----------------------------------------------------------------------------------
    'Parámetros.
    '   Entrada:    [tipoImpresion] determina el tipo de reporte a imprimir
    '               ALTA =          configurar el reporte para imprimir una nueva reserva
    '               MODIFICAR =     configurar el reporte para imprimir una modificación
    '               ANULAR =        configurar el listado para imprimir una anulación
    '               REIMPRESION =   confirmar el listado para reimprsión
    '
    '               [nroReservaLis] número de reserva con el cual se está trabajando
    '
    '               [resAux]        contiene el núemero de reserva con el formato corto
    '                               utilizado para mostrar mensaje al usuario.
    '
    '   NOTA: la impresión de reservas tiene una característica especial:
    '   la impresión del formulario no se realiza mediante un boton, si no que se realiza
    '   cuando se confirma la operación (anular, nueva, modificar). Por ese motivo se
    '   incluye una opción en la ficha General del cuadro de confirmación, la cual
    '   establece si se imprime o no este formulario.
    '----------------------------------------------------------------------------------
    
    Select Case tipoImpresion
        Case "ALTA"
            'valido impresión de reserva
            If tbPARAMETROS("imprimir_reserva") = 1 Then
                'se permite imprimir la reserva
                If mfunAplicoConfImp(2, 1) = 1 Then 'listado de alta reserva
                    'realizo listado
                    subArmoReporte nroReservaLis
                    'aviso de confirmación de impresión de reserva
                    mSubMensaje 4, 55, resAux    'se imprimió la reserva
                End If
            End If
            
        Case "MODIFICAR"
            'valido impresión de reserva
            If tbPARAMETROS("imprimir_reserva") = 1 Then
                'se permite imprimir la reserva
                If mfunAplicoConfImp(2, 2) = 1 Then 'listado de modificacion
                    'realizo listado
                    subArmoReporte nroReservaLis
                    'aviso de confirmación de impresión de reserva
                    mSubMensaje 4, 55, resAux    'se imprimió la reserva
                End If
            End If
    End Select
End Sub

Private Sub subArmoReporte(nroReservaLis As Long)
    '----------------------------------------------------------------------
    'Configuro el control Crystal genérico,obtengo datos y emite el listado
    '----------------------------------------------------------------------
    'Parámetros.
    '   Entrada: [nroReservaLis] el número de reserva el cual voy a imprimir
    '----------------------------------------------------------------------
    'inicializo control data
     subInicializoControlData frmMAIN.Data1CrystalReport
     
    'establesco consulta a utilizar
    'NOTA: esta consulta debe de ser la misma que se utilizo para crear el reporte,
    'por ese motivo no debe de ser modificada en aspectos tales como campos a mostrar y
    'tablas utilizadas.

    frmMAIN.Data1CrystalReport.RecordSource = _
        "select * from hab_reserva,reservas " & _
        "where hab_reserva.nroReserva = reservas.nroReserva and " & _
        "reservas.nroReserva = " & nroReservaLis
        
    'ejecuto consulta control data
    frmMAIN.Data1CrystalReport.Refresh
    
    'como de aquí en más voy a trabajar directamente con el control data
    'me aseguro de que se alla encontrado la reserva.
    If frmMAIN.Data1CrystalReport.Recordset.RecordCount > 0 Then
        'determino el nombre del reporte a utilizar
        frmMAIN.CrystalReport1.ReportFileName = vardir2 + "rptr1.rpt"

        'inicializo fórmulas
        With frmMAIN.CrystalReport1
            .Formulas(0) = "cabVersion = '" & mFunImpVersionAplicacion & "'"        'nombre y versión aplicación
            .Formulas(1) = "cabRegistro = '" & mFunImpRegistro & "'"                'hotel propietario de la aplicación
            .Formulas(2) = "cabHoraImp = ' " & CStr(Time()) & "'"                   'hora actual
            .Formulas(3) = "nroReserva = '" & mFunMuestroNroReserva(nroReservaLis) & "'"    'muestra el número de reserva
            .Formulas(4) = "titular = '" & mFunNombreTitularReserva(nroReservaLis) & "'"   'nombre completo del titular de la reserva
            'parte 3 persona que realizó la reserva
            If funVerificoImpresionParte3 Then
                .Formulas(5) = "parte3Titulo = 'Persona que realizó la reserva                                             '"
                .Formulas(6) = "parte3NombreCab = 'Nombre:'"
                .Formulas(7) = "parte3EmpCab = 'Empresa:'"
                .Formulas(8) = "parte3EmpDato = '" & mFunBuscoNombreEmpresa(frmMAIN.Data1CrystalReport.Recordset("nroAgenciaEmpresa")) & "'"
                .Formulas(9) = "parte3TelCab = 'Teléfono:'"
                .Formulas(10) = "parte3EmailCab = 'E-mail:'"
            End If
            'parte 4 pago de seña
            If funVerificoImpresionParte4 Then
                .Formulas(11) = "parte4Titulo = 'Pago de Seña                                                                      '"
                .Formulas(12) = "parte4NombreTar = '" & mFunBuscoNombreTarjetaCredito(frmMAIN.Data1CrystalReport.Recordset("tarjetaCredito")) & "'"
                .Formulas(13) = "parte4NombreCab = 'Nombre:'"
                .Formulas(14) = "parte4ApeCab = 'Apellido:'"
                .Formulas(15) = "parte4FechaEmiCab = 'Fecha emisión:'"
                .Formulas(16) = "parte4ImpCab = 'Importe " & gblSignoDolares & ":'"
                .Formulas(17) = "parte4ReciboCab = 'Recibo:'"
                .Formulas(18) = "parte4TituloTar = 'Datos tarjeta crédito'"
                .Formulas(19) = "parte4NroTCab = 'Nro.:'"
                .Formulas(20) = "parte4FechVtoCab = 'F.vto.:'"
                .Formulas(21) = "parte4CodigoCab = 'Cod.:'"
                'esta fórmula pertenece a un campo de tipo fecha, como crystal no soporta
                'fechas en blanco, la paso como parámetro.
                .Formulas(22) = "parte4FechaEmiDato = '" & mFunFormatoFecha(frmMAIN.Data1CrystalReport.Recordset("fechaEmision"), 1) & "'"
            End If
            'parte 5, observaciones
            If Not IsNull(frmMAIN.Data1CrystalReport.Recordset("observaciones")) Then
                If Trim(frmMAIN.Data1CrystalReport.Recordset("observaciones")) <> "" Then
                    .Formulas(23) = "parte5Titulo = 'Observaciones                                                                            '"
                End If
            End If
            'parte 3 habitaciones
            .Formulas(24) = "parte2SignoMonedaExtranjera = ' " & gblSignoDolares & "'"
        End With
            
        'genero listado
        frmMAIN.CrystalReport1.Action = 1
        'inicializo fórmulas
        mSubInicializoFormulas 24
    End If
End Sub

Private Function funVerificoImpresionParte3() As Boolean
    '--------------------------------------------------------------------------
    'Verifica si hay algún dato ingresado al momento de hacer la reserva,
    'correspondiente a la tercer parte del listado (persona que realizó la reserva)
    'Parámetros.
    '   Salida
    '       True =  si existe al menos un dato
    '       False = si no se ingresó ningún dato
    '--------------------------------------------------------------------------
    'por defecto asumo que no se imprime la sección
    funVerificoImpresionParte3 = False
    
    With frmMAIN.Data1CrystalReport
        'verifico si se ingresó nombre persona
        If Not IsNull(.Recordset("personaRealizoReserva")) Then
            If Trim(.Recordset("personaRealizoReserva")) <> "" Then
                'existe al menos un datos
                funVerificoImpresionParte3 = True
                Exit Function
            End If
        End If
        'verifico si se ingresó empresa
        If .Recordset("nroagenciaempresa") > 0 Then
            'existe al menos un datos
            funVerificoImpresionParte3 = True
            Exit Function
        End If
        'verifico si se ingresó teléfono
        If Not IsNull(.Recordset("telefono")) Then
            If Trim(.Recordset("telefono")) <> "" Then
                'existe al menos un datos
                funVerificoImpresionParte3 = True
                Exit Function
            End If
        End If
        'verifico se se ingresó fax
        If Not IsNull(.Recordset("fax")) Then
            If Trim(.Recordset("fax")) <> "" Then
                'existe al menos un datos
                funVerificoImpresionParte3 = True
                Exit Function
            End If
        End If
        'verifico si se ingresó e-mail
        If Not IsNull(.Recordset("email")) Then
            If Trim(.Recordset("email")) <> "" Then
                'existe al menos un datos
                funVerificoImpresionParte3 = True
                Exit Function
            End If
        End If
    End With
End Function

Private Function funVerificoImpresionParte4() As Boolean
    '--------------------------------------------------------------------------
    'Verifica si hay algún dato ingresado al momento de hacer la reserva,
    'correspondiente a la cuarta parte del listado (pago de seña)
    'Parámetros.
    '   Salida
    '       True =  si existe al menos un dato
    '       False = si no se ingresó ningún dato
    '--------------------------------------------------------------------------
    'por defecto asumo que no se imprime la sección
    funVerificoImpresionParte4 = False
    
    With frmMAIN.Data1CrystalReport
        'verifico si se ingresó nombres persona que realizó la seña
        If Not IsNull(.Recordset("nombres")) Then
            If Trim(.Recordset("nombres")) <> "" Then
                'existe al menos un datos
                funVerificoImpresionParte4 = True
                Exit Function
            End If
        End If
        'verifico si se ingresó apellidos persona
        If Not IsNull(.Recordset("apellidos")) Then
            If Trim(.Recordset("apellidos")) <> "" Then
                'existe al menos un datos
                funVerificoImpresionParte4 = True
                Exit Function
            End If
        End If
        'verifico si se ingresó fecha emisión
        If Not IsNull(.Recordset("fechaEmision")) Then
            If IsDate(.Recordset("fechaEmision")) Then
                'existe al menos un datos
                funVerificoImpresionParte4 = True
                Exit Function
            End If
        End If
        'verifico si se ingresó importe
        If .Recordset("importe") > 0 Then
            'existe al menos un datos
            funVerificoImpresionParte4 = True
            Exit Function
        End If
        'verifico si se ingresó recivo
        If .Recordset("nroRecibo") > 0 Then
            'existe al menos un datos
            funVerificoImpresionParte4 = True
            Exit Function
        End If
        'verifico si se ingresó número de tarjeta
        If Not IsNull(.Recordset("nroTarjeta")) Then
            If Trim(.Recordset("nrotarjeta")) <> "" Then
                'existe al menos un datos
                funVerificoImpresionParte4 = True
                Exit Function
            End If
        End If
        'verifico se se ingresó fecha de vencimiento
        If Not IsNull(.Recordset("fechaVto")) Then
            If IsDate(.Recordset("fechavto")) Then
                'existe al menos un datos
                funVerificoImpresionParte4 = True
                Exit Function
            End If
        End If
        'verifico si se ingresó e-mail
        If Not IsNull(.Recordset("codSeg")) Then
            If Trim(.Recordset("codSeg")) <> "" Then
                'existe al menos un datos
                funVerificoImpresionParte4 = True
                Exit Function
            End If
        End If
    End With
End Function

Private Sub botImprimir_Click()
    'Imprimo reserva desde consulta
    
    'realizo listado
    If mfunAplicoConfImp(2, 4) = 1 Then 'listado de reimpresión
        'realizo listado
        subArmoReporte nro_reserva
        'aviso de confirmación de impresión de reserva
        mSubMensaje 4, 55, Mid(Str(nro_reserva), 6, 5)   'se imprimió la reserva
    End If
End Sub

'****************************
'*
'*  Teclas de acceso
'*
'****************************

Private Sub mnuFormularioAceptar_Click()
    'Digitar F12 es lo mismo que apretar boton aceptar.
    'Este formulario es utilizado para varias tareas por lo que se ubican en el varios
    'botones que se deben ejecutar al presionar F12. Por ese motivo debo de determinar
    'que función esta cumpliendo el formulario, para así determinar el boton a ejecutar.
    Select Case tipo_accion_reserva
        Case "MODIFICAR"
            confirmareserva_Click

        Case "ANULAR"
            AnularReserva_Click
            
        Case "CONSULTAR"
            FinConsulta_Click
            
        Case "ALTA"
            confirmareserva_Click
            
        Case "Check-in"
            ResCorresponde_Click
    End Select
End Sub

Private Sub mnuFormularioCancelar_Click()
    'Cierro el formulario
    'Al igual que la opción de aceptar, debo de seleccionar el boton, dependiendo de la operación
    'que realize el formulario.
    Select Case tipo_accion_reserva
        Case "MODIFICAR"
            noconfirmareserva_Click
            
        Case "ANULAR"
            AnularReserva_Click
            
        Case "CONSULTAR"
            'para esta opción no existe boton de cancelar
            
        Case "ALTA"
            noconfirmareserva_Click
            
        Case "Check-in"
            ResCorresponde_Click
    
    End Select
End Sub

Public Sub Confirmar()
    'Este procedimiento es público ya que lo utilizo en el formulario de seleccion de habitacion
    'con la finalidad que el mismo trabaje en forma simulanea con el de reserva.
    
    'NOTA: originalmente este procedimiento se encontraba en el formulario frmCargaReserva.
    'Lo movi para este módulo con el fin de evitar posibles errores, ya que tambien es utilizado
    'en el formulario frmReservaSeleHab, lo que puede originar que el formulario frmCargaReserva
    'no se descargue de memoria. No se, creo que de todos modos no es bueno tener un procedimiento
    'público en un formulario.
    
    
    'verifico si fila esta marcada
    DBGrid1.Row = fila
    If DBGrid1.CellBackColor = &HFFFF00 Then ' si esta marcado
        modificar
    Else
        If Val(txtHab.Text) <> 0 Then
            grabo_hab_reservas (Val(txtHab.Text))
        Else
            grabo_hab_reservas (0)
        End If
    End If
    
    ejecuto_consulta
    dibujo_cabezales_grilla
    marco_color_no_asignada
    muestro_datos_hab True
    
    cboTipo_habitacion.Enabled = True
         
    If habitacionAsignada Then
        habilito_fechas False
    Else
        habilito_fechas True
    End If
End Sub

Private Sub mnuVerCuadroSituacion_Click()
    'Abre el formulario de Cuadro de situación
    OprEjecutada = 21
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        frmCuadroHab.Show 1
    End If
End Sub

Private Sub mnuVerCuadroDisponibilidad_Click()
    'Abre el formulario de Ver Disponibilidad
    OprEjecutada = 22
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        frmVerDisponibilidad.Show 1
    End If
End Sub

Private Sub mnuSeleccionarClientes_Click()
    'Busca clientes
    If Me.botAyuda.Enabled = True And _
        Me.botAyuda.Visible = True Then
        BotAyuda_Click
    End If
End Sub

Private Sub mnuSeleccionarEmpresas_Click()
    'Busca empresas
    If Me.botAyudaEmp.Enabled = True And _
    Me.botAyudaEmp.Visible = True Then
        botAyudaEmp_Click
    End If
End Sub

Private Sub mnuVerHabitacion_Click()
    'Muestro primer tabs
    ssTab1.Tab = 0
End Sub

Private Sub mnuVerTarjeta_Click()
    'Muestro segundo tabs
    ssTab1.Tab = 1
End Sub

Private Sub mnuVerPrePago_Click()
    'Muestro tercer tabs
    ssTab1.Tab = 2
End Sub

Private Sub mnuFormularioImprimir_Click()
    'Imprimo reserva
    botImprimir_Click
End Sub

'************************************************
'*
'*  Asistencia a usuarios
'*
'************************************************

Private Sub fechaing_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 53
End Sub

Private Sub fechaegr_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 54
End Sub

Private Sub txt1er_Nom_titular_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 55
End Sub

Private Sub txt2do_Nom_titular_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 56
End Sub

Private Sub txt1er_Ape_titular_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 57
End Sub

Private Sub txt2do_Ape_titular_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 58
End Sub

Private Sub cboTipo_habitacion_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 59
End Sub

Private Sub txttarifa_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 60
End Sub

Private Sub txtPasajeros_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 61
End Sub

Private Sub botHabitacion_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 62
End Sub

Private Sub dbgrid1_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 63
End Sub

Private Sub Eliminar_hab_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 64
End Sub

Private Sub txtPersona_Reserva_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 65
End Sub

Private Sub realizadapor_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 66
End Sub

Private Sub txtAlojamiento_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 67
End Sub

Private Sub txtExtras_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 68
End Sub

Private Sub txtTelefono_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 70
End Sub

Private Sub txtFax_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 71
End Sub

Private Sub txtEmail_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 72
End Sub

Private Sub txtObservaciones_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 69
End Sub

Private Sub otras_tar_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 78
End Sub

Private Sub txtNom_tar_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 73
End Sub

Private Sub txtApe_tar_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 74
End Sub

Private Sub txtNro_tar_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 75
End Sub

Private Sub fech_vto_tar_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 76
End Sub

Private Sub txtCod_Seg_tar_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 77
End Sub

Private Sub txtNom_Seña_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 79
End Sub

Private Sub txtApe_Seña_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 80
End Sub

Private Sub fecha_emi_seña_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 81
End Sub

Private Sub txtImporte_Seña_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 82
End Sub

Private Sub txtRecibo_Seña_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 83
End Sub

Private Sub NoAnular_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 1, 3
End Sub

Private Sub ResNoCorresponde_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 1, 3
End Sub

Private Sub noconfirmareserva_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 1, 3
End Sub

Private Sub AnularReserva_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 84
End Sub

Private Sub confirmareserva_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 1, 2
End Sub

Private Sub ResCorresponde_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 110
End Sub

Private Sub FinConsulta_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 1, 1
End Sub

Private Sub FinConsulta_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub cboTipo_habitacion_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub txttarifa_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub fech_vto_tar_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub confirmareserva_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub AnularReserva_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub noconfirmareserva_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub ResNoCorresponde_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub ResCorresponde_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub NoAnular_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub txtRecibo_Seña_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub txtImporte_Seña_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub fecha_emi_seña_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub txtNom_Seña_LostFocus()
    txtNom_Seña.Text = mFunFormatoNombre(txtNom_Seña.Text)
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub txtCod_Seg_tar_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub txtNro_tar_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub txtNom_tar_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub txtApe_tar_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub txtApe_Seña_LostFocus()
    txtApe_Seña.Text = mFunFormatoApellido(txtApe_Seña.Text)
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub txtObservaciones_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub otras_tar_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub txtTelefono_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub txtFax_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub txtEmail_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub Eliminar_hab_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub txtPersona_Reserva_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub realizadapor_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub txtAlojamiento_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub txtPasajeros_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub botHabitacion_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub txtExtras_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub dbgrid1_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub fechaegr_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub fechaing_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

