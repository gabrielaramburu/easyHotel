VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMAIN 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menu principal"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   11910
   Icon            =   "frmMAINver2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   11910
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FF8080&
      Height          =   5460
      Left            =   840
      ScaleHeight     =   5400
      ScaleWidth      =   1995
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1420
      Width           =   2055
      Begin VB.CommandButton botReservas 
         Caption         =   "Reservas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   0
         MaskColor       =   &H00FF0000&
         Picture         =   "frmMAINver2.frx":000C
         TabIndex        =   0
         Top             =   0
         Width           =   2000
      End
      Begin VB.CommandButton botIngresos 
         Caption         =   "Ingresos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   0
         TabIndex        =   1
         Top             =   600
         Width           =   2000
      End
      Begin VB.CommandButton botGastos 
         Caption         =   "Gastos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   0
         TabIndex        =   2
         Top             =   1200
         Width           =   2000
      End
      Begin VB.CommandButton botFacturacion 
         Caption         =   "Facturación"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   0
         TabIndex        =   3
         Top             =   1800
         Width           =   2000
      End
      Begin VB.CommandButton botInformes 
         Caption         =   "Informes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   0
         TabIndex        =   4
         Top             =   2400
         Width           =   2000
      End
      Begin VB.CommandButton botHabitacion 
         Caption         =   "Habitación"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   0
         TabIndex        =   5
         Top             =   3000
         Width           =   2000
      End
      Begin VB.CommandButton botCheckout 
         Caption         =   "Check-Out"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   0
         TabIndex        =   6
         Top             =   3600
         Width           =   2000
      End
      Begin VB.CommandButton botCierreDiario 
         Caption         =   "Cierre diario"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   0
         TabIndex        =   7
         Top             =   4200
         Width           =   2000
      End
      Begin VB.CommandButton botEstadosCuentas 
         Caption         =   "Estados de cuentas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   0
         TabIndex        =   8
         Top             =   4800
         Width           =   2000
      End
   End
   Begin Hotel_Nana.gaHOTELbarra gaHOTELbarra1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   9
      Top             =   7605
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   582
   End
   Begin VB.ListBox lstOpciones 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1305
      Left            =   2880
      TabIndex        =   11
      Top             =   1440
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      FillColor       =   &H0000C0C0&
      Height          =   8520
      Left            =   120
      Picture         =   "frmMAINver2.frx":1052
      ScaleHeight     =   8460
      ScaleWidth      =   11625
      TabIndex        =   12
      Top             =   0
      Width           =   11685
      Begin VB.Frame Frame1 
         Caption         =   "Control genérico"
         Height          =   1335
         Left            =   5040
         TabIndex        =   39
         Top             =   3720
         Visible         =   0   'False
         Width           =   2895
         Begin VB.Data Data1CrystalReport 
            Caption         =   "Data1CrystalReport"
            Connect         =   ";PWD=manyacapo;"
            DatabaseName    =   "C:\NANAHOTEL\hotel.mdb"
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   405
            Left            =   120
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   2  'Snapshot
            RecordSource    =   "select * from fac_cabezal,fac_lineas"
            Top             =   840
            Visible         =   0   'False
            Width           =   2655
         End
         Begin Crystal.CrystalReport CrystalReport1 
            Bindings        =   "frmMAINver2.frx":141974
            Left            =   240
            Top             =   360
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   262150
            WindowControlBox=   -1  'True
            WindowMaxButton =   -1  'True
            WindowMinButton =   -1  'True
            Connect         =   ";PWD=manyacapo;"
         End
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cotización"
         Height          =   240
         Left            =   3960
         TabIndex        =   38
         Top             =   5760
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ingresos:"
         Height          =   240
         Left            =   3960
         TabIndex        =   37
         Top             =   4560
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Egresos:"
         Height          =   240
         Left            =   3960
         TabIndex        =   36
         Top             =   4800
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad de pasajeros alojados:"
         Height          =   240
         Left            =   3960
         TabIndex        =   35
         Top             =   5160
         Visible         =   0   'False
         Width           =   2940
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ocupación:"
         Height          =   240
         Left            =   3960
         TabIndex        =   34
         Top             =   5400
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Label lblNroOpPrincipal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   8
         Left            =   300
         TabIndex        =   33
         Top             =   6330
         Width           =   135
      End
      Begin VB.Label lblNroOpPrincipal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   7
         Left            =   300
         TabIndex        =   32
         Top             =   5740
         Width           =   135
      End
      Begin VB.Label lblNroOpPrincipal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   300
         TabIndex        =   31
         Top             =   5140
         Width           =   135
      End
      Begin VB.Label lblNroOpPrincipal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   290
         TabIndex        =   30
         Top             =   4550
         Width           =   135
      End
      Begin VB.Label lblNroOpPrincipal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   300
         TabIndex        =   29
         Top             =   3960
         Width           =   135
      End
      Begin VB.Label lblNroOpPrincipal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   290
         TabIndex        =   28
         Top             =   3360
         Width           =   135
      End
      Begin VB.Label lblNroOpPrincipal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   300
         TabIndex        =   27
         Top             =   2760
         Width           =   135
      End
      Begin VB.Label lblNroOpPrincipal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   300
         TabIndex        =   26
         Top             =   2160
         Width           =   135
      End
      Begin VB.Label lblNroOpPrincipal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   240
         Index           =   0
         Left            =   300
         TabIndex        =   25
         Top             =   1590
         Width           =   135
      End
      Begin VB.Label lblNomHotel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lblNomHotel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   24
         Top             =   60
         Width           =   2400
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Acceso Directo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   9
         Left            =   9480
         TabIndex        =   23
         Top             =   5880
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Acceso Directo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   8
         Left            =   9480
         TabIndex        =   22
         Top             =   5400
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Acceso Directo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   7
         Left            =   9480
         TabIndex        =   21
         Top             =   4920
         Visible         =   0   'False
         Width           =   1995
      End
      Begin ComctlLib.ImageList ImageList1 
         Left            =   5040
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   15
         MaskColor       =   12632256
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   8
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMAINver2.frx":141991
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMAINver2.frx":141CB3
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMAINver2.frx":142079
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMAINver2.frx":14240B
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMAINver2.frx":1427D1
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMAINver2.frx":142AEB
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMAINver2.frx":142E05
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMAINver2.frx":143157
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label lblFechaHoy 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lblFechaHoy"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   10320
         TabIndex        =   20
         Top             =   720
         Width           =   1185
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Acceso Directo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   6
         Left            =   9480
         TabIndex        =   19
         Top             =   4440
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Acceso Directo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   5
         Left            =   9480
         TabIndex        =   18
         Top             =   3960
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Acceso Directo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   4
         Left            =   9480
         TabIndex        =   17
         Top             =   3480
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Acceso Directo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   3
         Left            =   9480
         TabIndex        =   16
         Top             =   3000
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Acceso Directo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   2
         Left            =   9480
         TabIndex        =   15
         Top             =   2520
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Acceso Directo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   1
         Left            =   9480
         TabIndex        =   14
         Top             =   2040
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Acceso Directo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   0
         Left            =   9480
         TabIndex        =   13
         Top             =   1560
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.Image Image1 
         Height          =   465
         Index           =   0
         Left            =   9000
         Picture         =   "frmMAINver2.frx":1434A9
         Top             =   1440
         Visible         =   0   'False
         Width           =   2550
      End
      Begin VB.Image Image1 
         Height          =   465
         Index           =   1
         Left            =   9000
         Picture         =   "frmMAINver2.frx":1472EB
         Top             =   1920
         Visible         =   0   'False
         Width           =   2550
      End
      Begin VB.Image Image1 
         Height          =   465
         Index           =   2
         Left            =   9000
         Picture         =   "frmMAINver2.frx":14B12D
         Top             =   2400
         Visible         =   0   'False
         Width           =   2550
      End
      Begin VB.Image Image1 
         Height          =   465
         Index           =   3
         Left            =   9000
         Picture         =   "frmMAINver2.frx":14EF6F
         Top             =   2880
         Visible         =   0   'False
         Width           =   2550
      End
      Begin VB.Image Image1 
         Height          =   465
         Index           =   4
         Left            =   9000
         Picture         =   "frmMAINver2.frx":152DB1
         Top             =   3360
         Visible         =   0   'False
         Width           =   2550
      End
      Begin VB.Image Image1 
         Height          =   465
         Index           =   5
         Left            =   9000
         Picture         =   "frmMAINver2.frx":156BF3
         Top             =   3840
         Visible         =   0   'False
         Width           =   2550
      End
      Begin VB.Image Image1 
         Height          =   465
         Index           =   6
         Left            =   9000
         Picture         =   "frmMAINver2.frx":15AA35
         Top             =   4320
         Visible         =   0   'False
         Width           =   2550
      End
      Begin VB.Image Image1 
         Height          =   465
         Index           =   7
         Left            =   9000
         Picture         =   "frmMAINver2.frx":15E877
         Top             =   4800
         Visible         =   0   'False
         Width           =   2550
      End
      Begin VB.Image Image1 
         Height          =   465
         Index           =   8
         Left            =   9000
         Picture         =   "frmMAINver2.frx":1626B9
         Top             =   5280
         Visible         =   0   'False
         Width           =   2550
      End
      Begin VB.Image Image1 
         Height          =   465
         Index           =   9
         Left            =   9000
         Picture         =   "frmMAINver2.frx":1664FB
         Top             =   5760
         Visible         =   0   'False
         Width           =   2550
      End
      Begin VB.Image Image2 
         Height          =   465
         Left            =   6000
         Picture         =   "frmMAINver2.frx":16A33D
         Top             =   2760
         Visible         =   0   'False
         Width           =   2550
      End
   End
   Begin VB.Menu mnuReservas 
      Caption         =   "&Reservas"
      Begin VB.Menu mnuReservasNueva 
         Caption         =   "Nueva"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuReservasModificar 
         Caption         =   "Modificar"
      End
      Begin VB.Menu mnuReservasConsultar 
         Caption         =   "Consultar"
      End
      Begin VB.Menu mnuReservasAnular 
         Caption         =   "Anular"
      End
   End
   Begin VB.Menu mnuIngresoPasa 
      Caption         =   "&Ingreso pasajeros"
      Begin VB.Menu mnuIngresoPasaCheckin 
         Caption         =   "Check-In"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuIngresoPasaWalkin 
         Caption         =   "Walk-In"
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuIngresoPasaWalkinHabOcupada 
         Caption         =   "Walk-In habitación ocupada"
      End
   End
   Begin VB.Menu menuGastos 
      Caption         =   "&Gastos"
      Begin VB.Menu menuGastosExtras 
         Caption         =   "Gastos extras"
         Shortcut        =   ^G
      End
      Begin VB.Menu menuGastosAlojamiento 
         Caption         =   "Gastos alojamiento"
      End
      Begin VB.Menu menuGastosResumenHabitacion 
         Caption         =   "Resumen de cuenta habitación"
      End
      Begin VB.Menu menuGastosResumenClientes 
         Caption         =   "Resumen de cuenta clientes"
      End
   End
   Begin VB.Menu mnuFacturacion 
      Caption         =   "&Facturación"
      Begin VB.Menu mnuFacturacionFacturas 
         Caption         =   "Facturas"
         Begin VB.Menu mnuFacturacionFacturasEmitir 
            Caption         =   "Emitir factura"
         End
         Begin VB.Menu mnuFacturacionFacturasConsultar 
            Caption         =   "Consultar"
         End
         Begin VB.Menu mnuFacturacionFacturasAnular 
            Caption         =   "Anular"
         End
      End
      Begin VB.Menu mnuFacturacionDevoluciones 
         Caption         =   "Devoluciones"
         Begin VB.Menu mnuFacturacionDevolucionesEmitir 
            Caption         =   "Emitir devolución"
         End
         Begin VB.Menu mnuFacturacionDevolucionesConsultar 
            Caption         =   "Consultar"
         End
      End
      Begin VB.Menu mnuRecivos 
         Caption         =   "Recibos"
         Begin VB.Menu mnuRecivosIngresar 
            Caption         =   "Ingresar recivo"
         End
         Begin VB.Menu mnuRecivosConsultar 
            Caption         =   "Consultar"
         End
         Begin VB.Menu mnuRecivosAnular 
            Caption         =   "Anular"
         End
      End
   End
   Begin VB.Menu mnuCheckOut 
      Caption         =   "&Check-Out"
   End
   Begin VB.Menu mnuInformes 
      Caption         =   "I&nformes"
      Begin VB.Menu mnuInformesCuadroSituacion 
         Caption         =   "Cuadro de situación"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuInformesDisponibilidad 
         Caption         =   "Cuadro de disponibilidad"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuInformesConsultaCompleta 
         Caption         =   "Consulta de habitaciones completa"
      End
      Begin VB.Menu mnuInformesIngresos 
         Caption         =   "Ingresos previstos"
      End
      Begin VB.Menu mnuInformesEgresos 
         Caption         =   "Egresos previstos"
      End
      Begin VB.Menu mnuInformesPasajerosHabitacion 
         Caption         =   "Pasajeros por habitación"
      End
      Begin VB.Menu mnuInformesPoblacionFlotante 
         Caption         =   "Población flotante"
      End
      Begin VB.Menu mnuInformesUbicacionPasajeros 
         Caption         =   "Ubicación de pasajeros"
      End
   End
   Begin VB.Menu mnuHabitacion 
      Caption         =   "&Habitaciones"
      Begin VB.Menu mnuHabitacionCambioTitular 
         Caption         =   "Cambio de titular"
      End
      Begin VB.Menu mnuHabitacionCambioFechaEgreso 
         Caption         =   "Cambio de fecha de egreso"
      End
      Begin VB.Menu mnuHabitacionCambioSituacion 
         Caption         =   "Cambio de situación"
      End
      Begin VB.Menu mnuHabitacionBloquear 
         Caption         =   "Bloquear "
      End
      Begin VB.Menu mnuHabitacionCambioHabitacion 
         Caption         =   "Cambio de habitación"
      End
      Begin VB.Menu mnuHabitacionCambioTarifa 
         Caption         =   "Cambio de tarifa"
      End
      Begin VB.Menu mnuHabitacionConsultaTitular 
         Caption         =   "Consulta de titular"
      End
   End
   Begin VB.Menu mnuCierreDiario 
      Caption         =   "Cierre diario"
   End
   Begin VB.Menu mnuEstadoCuenta 
      Caption         =   "&Estados de cuenta"
   End
   Begin VB.Menu mnuSis 
      Caption         =   "Siste&ma"
      Begin VB.Menu mnuSisCotizaciones 
         Caption         =   "Cotizaciones"
      End
      Begin VB.Menu mnuSisCambioUsuario 
         Caption         =   "Cambio de usuario"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuSisStandBy 
         Caption         =   "Aplicación en Stand by"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuMante 
         Caption         =   "Mantenimiento"
      End
      Begin VB.Menu mnuRecivosAuto 
         Caption         =   "Recibos automáticos"
         Begin VB.Menu mnuRecivosAutoImprimir 
            Caption         =   "Imprimir"
         End
         Begin VB.Menu mnuRecivosAutoConsultar 
            Caption         =   "Consultar"
         End
         Begin VB.Menu mnuRecivosAutoAnular 
            Caption         =   "Anular"
         End
      End
      Begin VB.Menu mnuSisCong 
         Caption         =   "Configuración del sistema"
      End
      Begin VB.Menu mnuSisEstablecerPerfil 
         Caption         =   "Establecer perfil de la aplicación"
      End
   End
   Begin VB.Menu mnuSalir 
      Caption         =   "&Salir"
   End
   Begin VB.Menu mnuAyuda 
      Caption         =   "&Ayuda ..."
      Begin VB.Menu mnuAyudaAcercaDe 
         Caption         =   "Acerca de ..."
      End
   End
End
Attribute VB_Name = "frmMAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Declaración de constantes
Private Const cColorActivo As Long = &HFFFF&    'amarillo
                                            'determina el color de los números de las opciones
                                            'principales del menú, cuando se le da el focus al
                                            'botón correspondiente.

Private AccesoPermitido As Boolean
Private PermitoDbleClik As Boolean
Private TipoOpcionesEnLista As Byte
Private continuarAplicacion As Boolean      'determina si se pudo cargar archivo de configuración
                                            'el cual contiene información inpresindible
                                            'para ejecutar el programa.

Private WithEvents PidoClave As UsuarioMuestro
Attribute PidoClave.VB_VarHelpID = -1
Private WithEvents ModoStandBy As UsuarioMuestro
Attribute ModoStandBy.VB_VarHelpID = -1
Private WithEvents PantallaConfig As ConfiguroInicial
Attribute PantallaConfig.VB_VarHelpID = -1

'Esta variable almecena la cantidad de accesos directos que muestro en el menú pricipal
'es necesaria para implementar el procedimiento que ilumina los accesos directos cuando
'el mouse se posiciona sobre ellos.Estudiando el procedimiento es facil darse cuenta del porque de su
'declaración.
Private TotAccesosDirecto As Byte

'Determina la forma como se muestra el menú de opciones
Private tipoMenu As Byte

Private Sub botCheckout_Click()
    'ejecuto opción: checkout
    mnuCheckOut_Click
End Sub

Private Sub botCierreDiario_Click()
    'Ejecuto opción cierre diario
    mnuCierreDiario_Click
End Sub

Private Sub botEstadosCuentas_Click()
    'ejecuto opción: estados de cuentas
    mnuEstadoCuenta_Click
End Sub

Private Sub botFacturacion_Click()
    subCargoLista 4
    subMuevoListaOpciones 2775, 3400
    TipoOpcionesEnLista = 4
End Sub

Private Sub botGastos_Click()
    subCargoLista 3
    subMuevoListaOpciones 4275, 2200
    TipoOpcionesEnLista = 3
End Sub

Private Sub botHabitacion_Click()
    subCargoLista 6
    subMuevoListaOpciones 4200, 3000
    TipoOpcionesEnLista = 6
End Sub

Private Sub botInformes_Click()
    subCargoLista 5
    subMuevoListaOpciones 4775, 3900
    TipoOpcionesEnLista = 5
End Sub

Private Sub botIngresos_Click()
    subCargoLista 2
    subMuevoListaOpciones 3000, 1500
    TipoOpcionesEnLista = 2
End Sub

Private Sub botReservas_Click()
    subCargoLista 1
    subMuevoListaOpciones 2275, 2200
    TipoOpcionesEnLista = 1
End Sub

Private Sub Form_Activate()
    'Cuando el formulario activo es el main muestro barra de tareas
    Me.gaHOTELbarra1.Visible = True
    
    subEstablescoPropiedadesMain
End Sub

Private Sub Form_Deactivate()
    'Muestro la barra de tareas solo en el formulario activo
    Me.gaHOTELbarra1.Visible = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'Dependiendo del dígito que se presione, simulo que se presionó
    'un botón del menú principal.
    
    Dim botonPres As Byte   'con la declaración de esta variable ahorro líneas de código
    botonPres = 6 'por defecto asumo que no presioné níngún dígito válido
    
    If KeyAscii = 27 Then
        'Cuando presiono la tecla ESC oculto (en caso de que halla alguno) el menú
        'de opciones desplegable.
        Me.lstOpciones.Visible = False
        subApagoNumerosOpcion
    End If
    
    If lstOpciones.Visible = False Then
        'cuando tengo un menú activo no permito acceder a opciones general
        'ya que los dígitos presionados corresponden a opciones dentro de cada menú
        Select Case KeyAscii
            Case 49 'reservas
                botonPres = 0
                botReservas_Click
                
            Case 50 'ingresos
                botonPres = 1
                botIngresos_Click
                
            Case 51 'gastos
                botonPres = 2
                botGastos_Click
                
            Case 52 'facturación
                botonPres = 3
                botFacturacion_Click
                
            Case 53 'informes
                botonPres = 4
                botInformes_Click
                
            Case 54 'habitación
                botonPres = 5
                botHabitacion_Click
                
            Case 55 'checkout
                'ilumino número de opción
                subApagoNumerosOpcion
                Me.lblNroOpPrincipal(6).ForeColor = cColorActivo
                botCheckout_Click
                
            Case 56 'cierre diario
                'ilumino número de opción
                subApagoNumerosOpcion
                Me.lblNroOpPrincipal(7).ForeColor = cColorActivo
                
                botCierreDiario_Click
                
            Case 57 'estados de cuenta
                'ilumino número de opción
                subApagoNumerosOpcion
                Me.lblNroOpPrincipal(8).ForeColor = cColorActivo
                botEstadosCuentas_Click
                
        End Select
    End If
    'NOTA: para las opciones 6,7 y 8 cambio el color antes de presionar el boton
    'ya que las mismas tienen la característica de que ejecutan opciones en forma
    'directa, es decir, no hay ningún menú de por medio.
    
    If botonPres < 6 Then
        'ilumino número de opción
        subApagoNumerosOpcion
        Me.lblNroOpPrincipal(botonPres).ForeColor = cColorActivo
    End If
End Sub

Private Sub Form_Load()
    'verifico si existe otra instancia de la aplicación
    'ejecutándose
    If Not funExisteOtraInstancia Then
        'obtengo base de datos de la aplicación
        subLeoArchivoConfiguracion
        
        If continuarAplicacion Then
        
            'Abro base de datos
            mSubAbroBaseDeDatos
            
            'Cargo variables desde archivo parámetros
            mSubInicioAplicacion
            
            'verifico si es una aplicación válida
            If mFunAplicacionValida Then
                'obtengo fecha sistema
                If mFunObtengoFechaSistema Then
            
                    'pido autorización para entrar al programa
                    subAutorizacion
                
                    'Avilito la utilización del procedimiento de bitacora.dll
                    Set ControlOperaciones = New GraboOperacion
                
                    If AccesoPermitido Then
                        'establesco propiedades del formulario
                        subEstablescoPropiedadesMain
                    Else
                        'no es un usuario de la aplicación.
                        terminarEjecucion = True
                    End If
                Else
                    'no se inicializó fecha del sistema
                    terminarEjecucion = True
                End If
            Else
                'no es una aplicación válida
                terminarEjecucion = True
            End If
        Else
            'no permito ejecutar dos instancias de esta
            'aplicación en una misma máquina
            terminarEjecucion = True
        End If
    Else
        'no se pudo obtener la base de datos de la aplicación
        terminarEjecucion = True
    End If
End Sub

Private Sub subEstablescoPropiedadesMain()
    'Este procedimiento se encarga de establecer las propiedades del formulario Main
    'a) Nombre del hotel (versión registrada)
    'b) Accesos directos (configurados por el usuario)
    'c) Tipo de menu (movimiento, fijo)
    'd) Fecha de la aplicación
    'e) Fecha en barra de estados
    
    'Es ejecutado al cargarse el formulario (evento load) y en el evento Activate
    'con el fin de actualizar los cambios que puden ser modificados por el usuario dentro
    'de la opción de configuración.
    
    'inicializo etiqueta lblnomHotel
    Me.lblNomHotel.Caption = funInicializoTitulo
    'muestro accesos directos
    subMuestroAccesosDirectos
    'obtengo tipo de menú
    tipoMenu = tbPARAMETROS("tipoMenu")
    'muestro fecha de la aplicación
    Me.lblFechaHoy = Format(m_FechaSis, "dddd, dd mmmm ") & "de " & _
    Format(m_FechaSis, "yyyy")
    'muestro fecha del sistema en barra de estado
    Me.gaHOTELbarra1.InicializoFecha
End Sub

Private Function funInicializoTitulo() As String
    'Obtiene el nombre del hotel (empresa) desde el archivo tbSISTEMA_LICENCIA
    'Dicha información se muestra en el formulario principal.
    '---------------------------------------------------------------------------
    'Parámetros:
    '       Salida
    '           si es una versión demo: "Versión DEMO"
    '           si es una versión registrada: el valor del campo
    '           tbSISTEMA_LICENCIA("AplicacionEmpresa")
    '-----------------------------------------------------------------------------
    Dim infAplicacion As InformacionApli
    
    If gEsUnaVersionDemo Then
        funInicializoTitulo = "Versión DEMO"
    Else
        'muestro el nombre de la empresa (hotel)
        'creo instancia
        Set infAplicacion = New InformacionApli
        funInicializoTitulo = _
        infAplicacion.mFunObtenerLicenciaApli(idApli, tbSISTEMA_LICENCIA, 2)
        'destruyo instancia
        Set infAplicacion = Nothing
    End If
End Function

Private Sub subLeoArchivoConfiguracion()
    'Estable el camino a la base de datos
    On Error GoTo errores
    Dim NumArch As Integer
    Dim linea As String
    
    continuarAplicacion = True
    NumArch = FreeFile
    Open App.Path & "\EasyHotel.txt" For Input As NumArch
    
    'leo archivo
    Line Input #NumArch, linea
    'inicializo variable global de la aplicación
    BaseDeDatosAplicacion = linea
    
    Close NumArch
errores:
    If err.Number <> 0 Then
        Select Case err.Number
            Case 53
                'Si no econtre el archivo quiere decir que la aplicación
                'se está ejecutando por primera vez, por lo tanto
                'ejecuto dll para crear nuevor archivo de configuración
                'Llamo a dll de configuración
            
                Set PantallaConfig = New ConfiguroInicial
                    PantallaConfig.MostrarPantallaConfigurar _
                    "EasyHotel.txt", _
                    App.Path, _
                    "EasyHotel"
                Set PantallaConfig = Nothing
        
                err.Number = 0
        End Select
    End If
End Sub

Private Sub subAutorizacion()
    'Determino si el usuario puede ingresar a la aplicación
    
    AccesoPermitido = False
    'Valido acceso a la aplicación
    If tbPARAMETROS("SisAdminTF") = 0 Then
        'Nunca definí perfiles de usuario, por ese motivo
        'no pido contraseña ninguna.
        AccesoPermitido = True
        'culto opciones de usuario
        mnuSisCambioUsuario.Visible = False
        Me.mnuSisStandBy.Visible = False
        'tampoco permito cambiar el usuario con dblclick
        'sobre la barra de estado
        PermitoDbleClik = False
    Else
        'Tengo definido perfiles de usuarios por lo que
        'tengo que ingresar contraseña
        
        Set PidoClave = New UsuarioMuestro
        'Ejecuto dll para pedir contraseña
        PidoClave.MuestroUsuario tbSISTEMA_USUARIOS
        Set PidoClave = Nothing
        
        PermitoDbleClik = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Este control lo realizo en este evento para mostrar también el mensaje cuando
    'cierro el formulario con Alt+ F4
    Dim i As Integer        'utilizada para descargar todos los formularios
    Dim totForm As Integer  'total de formularios cargados
    
    'si estoy ejecutando una versión demo
    If gEsUnaVersionDemo Then
        'muestro mensaje de versión demo al salir de la aplicación
        subMuestroAvisoVersionDemo
    End If

    'Una aplicación controlada por eventos se termina cuando se cierran todos sus formularios
    'y no se ejecuta ningún código. Si un formulario se mantiene oculto
    'cuando se cierre el último formulario visible, parecerá que la aplicación ha terminado (porque no hay formularios visibles),
    'pero en realidad se seguirá ejecutando hasta que se cierren todos los formularios ocultos.
    'Esta situación puede producirse porque el acceso a las propiedades o controles de un formulario no cargado
    'hace que éste se cargue de forma implícita, sin presentarlo.
    'La mejor manera de evitar este problema cuando se cierre la aplicación es asegurarse de descargar todos los formularios.
    'Si tiene más de un formulario, puede utilizar la colección Forms y la instrucción Unload.
    'Si la aplicación utiliza varios formularios, puede descargar los formularios agregando código
    'en el procedimiento de evento Unload del formulario principal.
    'Puede utilizar la colección Forms para asegurarse de buscar y cerrar todos los formularios.
    'El siguiente código usa la colección Forms para descargar todos los formularios:
        
    'obtengo el total de formulario cargados actualmente y los descatgo a todos
    totForm = Forms.Count - 1
    For i = 0 To totForm
        Unload Forms(0)
    Next
    Set ControlOperaciones = Nothing
End Sub

Private Sub gaHOTELbarra1_DblClickSobreUsuario()
    'Hacer doble click sobre la barra de estado
    'equivale a ctrol+u
    
    'Esta bandera controla que cuando el sistema esta trabajando en modo
    'libre no trabaje tampoco el dblclick sobre la barra de estado
    If PermitoDbleClik Then
        mnuSisCambioUsuario_Click
    End If
End Sub

Private Sub lstOpciones_DblClick()
    'Ejecuto la opción seleccionada en la lista de opciones
    subEjecutoOpcionLista
End Sub

Private Sub lstOpciones_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        subEjecutoOpcionLista
    End If
End Sub

Private Sub subEjecutoOpcionLista()
    'Al realizar doble clik o tipear enter sobre una de las
    'opciones de las listas de opciones, debo de ejecutar la operación correspondiente
    Select Case TipoOpcionesEnLista
        Case 1  'reservas
            Select Case lstOpciones.ListIndex
                Case 0  'nueva
                    mnuReservasNueva_Click
                Case 1  'modificar
                    mnuReservasModificar_Click
                Case 2  'consultar
                    mnuReservasConsultar_Click
                Case 3  'anular
                    mnuReservasAnular_Click
            End Select
        Case 2  'ingresos
            Select Case lstOpciones.ListIndex
                Case 0  'checkin
                    mnuIngresoPasaCheckin_Click
                Case 1  'walkin
                    mnuIngresoPasaWalkin_Click
                Case 2  'walkin ocupada
                    mnuIngresoPasaWalkinHabOcupada_Click
            End Select
        Case 3  'gastos
            Select Case lstOpciones.ListIndex
                Case 0  'ingreso gastos extras
                    menuGastosExtras_Click
                Case 1  'ingreso gastos alojamiento
                    menuGastosAlojamiento_Click
                Case 2  'resumen de gastos de habitacion
                    menuGastosResumenHabitacion_Click
                Case 3  'resumen de gastos clientes
                    menuGastosResumenClientes_Click
            End Select
        Case 4  'facturación
            Select Case lstOpciones.ListIndex
                Case 0  'emitir factura
                    mnuFacturacionFacturasEmitir_Click
                Case 1  'consultar
                    mnuFacturacionFacturasConsultar_Click
                Case 2  'anular
                    mnuFacturacionFacturasAnular_Click
                Case 3  'emitir devolucion
                    mnuFacturacionDevolucionesEmitir_Click
                Case 4  'consultar devoluciones
                    mnuFacturacionDevolucionesConsultar_Click
                Case 5  'ingresar recivos
                    mnuRecivosIngresar_Click
                Case 6  'consultar recivos
                    mnuRecivosConsultar_Click
                Case 7  'anular recivos
                    mnuRecivosAnular_Click
            End Select
        Case 5  'informes
            Select Case lstOpciones.ListIndex
                Case 0  'cuadro de situación
                    mnuInformesCuadroSituacion_Click
                Case 1  'cuadro de disponibilidad
                    mnuInformesDisponibilidad_Click
                Case 2  'consulta completa
                    mnuInformesConsultaCompleta_Click
                Case 3
                    mnuInformesIngresos_Click
                Case 4
                    mnuInformesEgresos_Click
                Case 5
                    mnuInformesPasajerosHabitacion_Click
                Case 6 'poblacion flotante
                    mnuInformesPoblacionFlotante_Click
                Case 7
                    mnuInformesUbicacionPasajeros_Click
            End Select
        Case 6  'habitaciones
            Select Case lstOpciones.ListIndex
                Case 0  'cambio de titular
                    mnuHabitacionCambioTitular_Click
                Case 1  'cambio fecha de egreso
                    mnuHabitacionCambioFechaEgreso_Click
                Case 2  'cambio de situacion
                    mnuHabitacionCambioSituacion_Click
                Case 3  'bloquear
                    mnuHabitacionBloquear_Click
                Case 4  'cambio de habitacion
                    mnuHabitacionCambioHabitacion_Click
                Case 5  'cambio de tarifa
                    mnuHabitacionCambioTarifa_Click
                Case 6  'consulta de titular
                    mnuHabitacionConsultaTitular_Click
                
            End Select
    End Select
    'Después de finalizar de ejecitar la opción correspondiente
    'desencadeno nuevamente el evento activate del formulario principal
    'para que se muestre la barra de tareas.
    Form_Activate
End Sub

Private Sub lstOpciones_LostFocus()
    lstOpciones.Visible = False
    'asistencia a usuarios
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub mnuAyudaAcercaDe_Click()
    'muestro formulario de Acerca de..
    frmAcercaDeEasyHotel.Show 1
End Sub

Private Sub mnuSisCambioUsuario_Click()
    'Cambio de usuario activo
    Set PidoClave = New UsuarioMuestro
    PidoClave.MuestroUsuario tbSISTEMA_USUARIOS
    Set PidoClave = Nothing
End Sub

Private Sub mnuSisStandBy_Click()
    'Muestra ventana de usuarios con la posibilidad
    'de salir del programa si no es un usuario

    Set ModoStandBy = New UsuarioMuestro
    ModoStandBy.MuestroUsuarioStandBy tbSISTEMA_USUARIOS
    Set ModoStandBy = Nothing
End Sub

Private Sub ModoStandBy_NotificoCliente(usuario As String, boton As Byte)
    'Este evento se ejecuta cuando hago clik
    'en algun boton del cuadro de dialogo de StandBy
    If boton = 1 Then
        'muestro usuario activo
        m_UsuarioSisNom = usuario
        Me.gaHOTELbarra1.InicializoUsuario
    Else
        'Termino con la ejecución del programa
        Unload Me
    End If
End Sub

Private Sub PantallaConfig_NotificoClientes(boton As Byte)
    'Este evento se produce al confirmar la pantalla de
    'configuración, la cual se muestra si no existe el archivo de configuración.
    
    Dim NumArch As Integer
    NumArch = FreeFile
    Dim linea As String
    
    If boton = 1 Then 'aceptar
        Open App.Path & "\EasyHotel.TXT" For Input As NumArch

        'leo archivo
        Line Input #NumArch, linea
        BaseDeDatosAplicacion = linea
        Close NumArch
    Else
        continuarAplicacion = False  'finalizo ejecución
    End If
End Sub

Private Sub PidoClave_NotificoCliente(usuario As String, boton As Byte)
    'Este evento se ejecuta cuando hago click
    'en algun boton del cuadro de díalogo de contraseña de clientes
    If boton = 1 Then   'aceptar
        'muestro usuario activo
        m_UsuarioSisNom = usuario
        Me.gaHOTELbarra1.InicializoUsuario
          
        AccesoPermitido = True
    Else
        AccesoPermitido = False
    End If
End Sub

Private Sub subCargoLista(boton As Byte)
    'Cargo las opciones en la lista de opciones
    'En la propiedad itemdata, cargo el número de mensaje que se muestra en
    'la barra de tareas, al darle el focus a la opción.
    
    'borro lista
    Me.lstOpciones.Clear
    
    Select Case boton
        Case 1  'reservas
            lstOpciones.AddItem "1. Nueva"
            lstOpciones.ItemData(lstOpciones.NewIndex) = 157
            lstOpciones.AddItem "2. Modificar"
            lstOpciones.ItemData(lstOpciones.NewIndex) = 158
            lstOpciones.AddItem "3. Consultar"
            lstOpciones.ItemData(lstOpciones.NewIndex) = 159
            lstOpciones.AddItem "4. Anular"
            lstOpciones.ItemData(lstOpciones.NewIndex) = 160
            
        Case 2  'ingresos
            lstOpciones.AddItem "1. Check-In"
            lstOpciones.ItemData(lstOpciones.NewIndex) = 161
            lstOpciones.AddItem "2. Walk-In"
            lstOpciones.ItemData(lstOpciones.NewIndex) = 162
            lstOpciones.AddItem "3. Walk-In hab. ocupada"
            lstOpciones.ItemData(lstOpciones.NewIndex) = 163
            
        Case 3  'gastos
            lstOpciones.AddItem "1. Gastos extras"
            lstOpciones.ItemData(lstOpciones.NewIndex) = 164
            lstOpciones.AddItem "2. Gastos alojamiento"
            lstOpciones.ItemData(lstOpciones.NewIndex) = 165
            lstOpciones.AddItem "3. Resumen de cuenta hab."
            lstOpciones.ItemData(lstOpciones.NewIndex) = 166
            lstOpciones.AddItem "4. Resumen de cuenta clientes"
            lstOpciones.ItemData(lstOpciones.NewIndex) = 167
            
        Case 4  'facturación
            lstOpciones.AddItem "1. Emitir documentos"
            lstOpciones.ItemData(lstOpciones.NewIndex) = 168
            lstOpciones.AddItem "2. Consultar documentos"
            lstOpciones.ItemData(lstOpciones.NewIndex) = 169
            lstOpciones.AddItem "3. Anular documentos"
            lstOpciones.ItemData(lstOpciones.NewIndex) = 170
            lstOpciones.AddItem "4. Emitir devolucion"
            lstOpciones.ItemData(lstOpciones.NewIndex) = 171
            lstOpciones.AddItem "5. Consultar devolución"
            lstOpciones.ItemData(lstOpciones.NewIndex) = 172
            lstOpciones.AddItem "6. Ingreso recibo"
            lstOpciones.ItemData(lstOpciones.NewIndex) = 173
            lstOpciones.AddItem "7. Consultar recibo"
            lstOpciones.ItemData(lstOpciones.NewIndex) = 174
            lstOpciones.AddItem "8. Anular recibo"
            lstOpciones.ItemData(lstOpciones.NewIndex) = 175
        
        Case 5  'informes
            lstOpciones.AddItem "1. Cuadro de situación"
            lstOpciones.ItemData(lstOpciones.NewIndex) = 176
            lstOpciones.AddItem "2. Cuadro de disponibildad"
            lstOpciones.ItemData(lstOpciones.NewIndex) = 177
            lstOpciones.AddItem "3. Consulta de habitaciones completa"
            lstOpciones.ItemData(lstOpciones.NewIndex) = 178
            lstOpciones.AddItem "4. Ingresos previstos"
            lstOpciones.ItemData(lstOpciones.NewIndex) = 179
            lstOpciones.AddItem "5. Egresos previstos"
            lstOpciones.ItemData(lstOpciones.NewIndex) = 180
            lstOpciones.AddItem "6. Pasajeros por habitacion"
            lstOpciones.ItemData(lstOpciones.NewIndex) = 181
            lstOpciones.AddItem "7. Población flotante"
            lstOpciones.ItemData(lstOpciones.NewIndex) = 182
            lstOpciones.AddItem "8. Ubicación de pasajeros"
            lstOpciones.ItemData(lstOpciones.NewIndex) = 183
        
        Case 6 'habitación
            lstOpciones.AddItem "1. Cambio de titular"
            lstOpciones.ItemData(lstOpciones.NewIndex) = 184
            lstOpciones.AddItem "2. Cambio de fecha de egreso"
            lstOpciones.ItemData(lstOpciones.NewIndex) = 185
            lstOpciones.AddItem "3. Cambio de situación"
            lstOpciones.ItemData(lstOpciones.NewIndex) = 186
            lstOpciones.AddItem "4. Bloquear"
            lstOpciones.ItemData(lstOpciones.NewIndex) = 187
            lstOpciones.AddItem "5. Cambio de habitación"
            lstOpciones.ItemData(lstOpciones.NewIndex) = 188
            lstOpciones.AddItem "6. Cambio de tarifa"
            lstOpciones.ItemData(lstOpciones.NewIndex) = 189
            lstOpciones.AddItem "7. Consulta de titular"
            lstOpciones.ItemData(lstOpciones.NewIndex) = 190
    End Select
End Sub

Private Sub subMuevoListaOpciones(ancho As Long, alto As Long)
    Dim i As Long
    Dim inicio As Long
    
    Me.lstOpciones.Height = alto
    Me.lstOpciones.Width = ancho
    
    'determino tipo de menú configurado por el usuario
    If tipoMenu = 0 Then
        'menu desplegable   (más vistoso)
        'a partir de esta posicion aparece a salir la ventana
        inicio = Me.Picture1.Left + Me.Picture1.Width
        Me.lstOpciones.Left = inicio - ancho 'oculto lista
        lstOpciones.Visible = True  'mostrando la lista despues de ocultarla
                                'elimino el efectode parpadeo
        Do While Me.lstOpciones.Left < inicio
            Me.lstOpciones.Left = Me.lstOpciones.Left + i
            Me.Refresh
            i = i + 2
        Loop
    Else
        'menu fijo (más rápido)
        lstOpciones.Visible = True
    End If
    lstOpciones.ListIndex = 0
    lstOpciones.SetFocus
End Sub


'******************************************************
'*
'*
'*  Accesos directos
'*
'*
'*
'******************************************************

Private Sub Label1_DblClick(Index As Integer)
    'Ejecuto opciones de acceso directo
    subEjecutoAccesosDirectos Label1(Index).Tag
End Sub

Private Sub subEjecutoAccesosDirectos(Opr As Integer)
    'Al hacer doble click sobre un acceso directo llamo a este
    'procedimiento quien se ebcarga de ejecutar la opcion correspondiente
    
    Select Case Opr
        Case 1  'nueva reserva
            mnuReservasNueva_Click
        Case 2  'modificar reserva
            mnuReservasModificar_Click
        Case 3  'consultar reserva
            mnuReservasConsultar_Click
        Case 4  'anular reserva
            mnuReservasAnular_Click
        Case 5  'checkin
            mnuIngresoPasaCheckin_Click
        Case 6  'walkin
            mnuIngresoPasaWalkin_Click
        Case 7  'walkin ocupada
            mnuIngresoPasaWalkinHabOcupada_Click
        Case 8  'ingreso gastos extras
            mnuIngresoPasaCheckin_Click
        Case 9 'ingreso gastos alojamiento
            menuGastosAlojamiento_Click
        Case 10 'resumen de gastos de habitacion
            menuGastosResumenHabitacion_Click
        Case 11 'resumen de gastos clientes
            menuGastosResumenClientes_Click
        Case 12 'emitir factura
            mnuFacturacionFacturasEmitir_Click
        Case 13  'consultar
            mnuFacturacionFacturasConsultar_Click
        Case 14  'anular
            mnuFacturacionFacturasAnular_Click
        Case 15  'emitir devolucion
            mnuFacturacionDevolucionesEmitir_Click
        Case 16  'consultar devoluciones
            mnuFacturacionDevolucionesConsultar_Click
        Case 17  'ingresar recivos
            mnuRecivosIngresar_Click
        Case 18  'consultar recivos
            mnuRecivosConsultar_Click
        Case 19  'anular recivos
            mnuRecivosAnular_Click
        Case 20     'esta opcion no existe
        Case 21  'cuadro de situación
            mnuInformesCuadroSituacion_Click
        Case 22  'cuadro de disponibilidad
            mnuInformesDisponibilidad_Click
        Case 23  'informe de situación
            'esta opción no se usa
            'mnuInformesSituacionActual_Click
        Case 24  'consulta completa
            mnuInformesConsultaCompleta_Click
        Case 25
            mnuInformesIngresos_Click
        Case 26
            mnuInformesEgresos_Click
        Case 27
            mnuInformesPasajerosHabitacion_Click
        Case 28 'poblacion flotante
            mnuInformesPoblacionFlotante_Click
        Case 29
            mnuInformesUbicacionPasajeros_Click
        Case 30  'cambio de titular
            mnuHabitacionCambioTitular_Click
        Case 31  'cambio fecha de egreso
            mnuHabitacionCambioFechaEgreso_Click
        Case 32  'cambio de situacion
            mnuHabitacionCambioSituacion_Click
        Case 33  'bloquear
            mnuHabitacionBloquear_Click
        Case 34  'cambio de habitacion
            mnuHabitacionCambioHabitacion_Click
        Case 35  'cambio de tarifa
            mnuHabitacionCambioTarifa_Click
        Case 36  'consulta de titular
            mnuHabitacionConsultaTitular_Click
    End Select
End Sub

Private Sub subMuestroAccesosDirectos()
    'Recorro el archvio tbSISTEMA_OPERACIONES y cargo las operaciones configuradas
    'con typo accesos directos
    'NOTA: el hecho de tener que crear los accesos directos cada vez que se ejecuta el
    'evento Active, no tiene ningún efecto (perseptible) con respecto al rendimiento de la
    'aplicación.
    
    subInicializoAccesosDirectos
    TotAccesosDirecto = 0
    tbSISTEMA_OPERACIONES.MoveFirst
    Do While Not tbSISTEMA_OPERACIONES.EOF
        'trabajo solo con las ya seleccionadas
        If tbSISTEMA_OPERACIONES("UsadaParaAccesoDirecto") Then
            'el número de accesos directos es limitado
            If TotAccesosDirecto < 10 Then
                'creo acceso directo
                Label1(TotAccesosDirecto).Caption = tbSISTEMA_OPERACIONES("descAccesoDirecto")
                Label1(TotAccesosDirecto).Tag = tbSISTEMA_OPERACIONES("CodOpr")
                Label1(TotAccesosDirecto).Visible = True
                Image1(TotAccesosDirecto).Visible = True
                TotAccesosDirecto = TotAccesosDirecto + 1
            Else
                Exit Do
            End If
        End If
        tbSISTEMA_OPERACIONES.MoveNext
    Loop
    'resto 1 para que su valor corresponda con la cantidad de accesos directos
    If TotAccesosDirecto > 0 Then
        TotAccesosDirecto = TotAccesosDirecto - 1
    End If
End Sub

Private Sub subInicializoAccesosDirectos()
    'Como los accesos directos pueden ser modificados por los usuarios, es necesario
    'inicializarlos para que los cambios se reflejen correctamente.
    Label1(0).Visible = False
    Image1(0).Visible = False
    Label1(1).Visible = False
    Image1(1).Visible = False
    Label1(2).Visible = False
    Image1(2).Visible = False
    Label1(3).Visible = False
    Image1(3).Visible = False
    Label1(4).Visible = False
    Image1(4).Visible = False
    Label1(5).Visible = False
    Image1(5).Visible = False
    Label1(6).Visible = False
    Image1(6).Visible = False
    Label1(7).Visible = False
    Image1(7).Visible = False
    Label1(8).Visible = False
    Image1(8).Visible = False
    Label1(9).Visible = False
    Image1(9).Visible = False
End Sub

Private Sub Image1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    'Cuando muevo el mouse sobre un acceso directo cambio (ilumino) la imagen de fondo.
    Dim i As Byte
    Image2.Top = Image1(Index).Top
    Image2.Left = Image1(Index).Left
    'muestro las etiquetas que actualmente poseen accesos directo
    For i = 0 To TotAccesosDirecto
        Image1(i).Visible = True
    Next
    Image1(Index).Visible = False
    Image2.Visible = True
    'asistencia a usuarios
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 191
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    'Cuando retiro el mouse de los accesos directos, apago la iluminación del acceso.
    Dim i As Byte
    Image2.Visible = False
    For i = 0 To TotAccesosDirecto
        Image1(i).Visible = True
    Next
    'asistencia a usuarios
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

'******************************************************
'*
'*
'*  Click del menu flotante
'*
'*
'*
'******************************************************

Private Sub mnuSisCotizaciones_Click()
    HoraIni = Time
    OprEjecutada = 62
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        frmCotizaciones.Show 1
    End If
End Sub

Private Sub mnuSisEstablecerPerfil_Click()
    OprEjecutada = 61
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        frmPerfilAplicacion.Show 1
    End If
End Sub

Private Sub menuGastosAlojamiento_Click()
    'hora de inicio de la operación
    HoraIni = Time
    OprEjecutada = 9
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        tipo_accion_inghabitacion = 4
        frmIngHabitacion.Show 1
    End If
End Sub

Private Sub menuGastosExtras_Click()
    'hora de inicio de la operación
    HoraIni = Time
    OprEjecutada = 8
    If mFunControlDeBaseDeDatos("frmIngExtras") Then
        If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
            tipo_accion_inghabitacion = 1
            frmIngHabitacion.Show 1
        End If
    Else
        frmAvisoErrorDatosMinimos.propTipoDetalle = "frmIngExtras"
        frmAvisoErrorDatosMinimos.Show 1
    End If
End Sub

Private Sub menuGastosResumenClientes_Click()
    OprEjecutada = 11
    If mFunControlDeBaseDeDatos("frmConsultaCuentas") Then
        If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
            tipo_accion_ConsultaCuentas = 2
            tipo_accion_IngEstadoCuenta = 2
            frmIngPaxEmp.Show 1
        End If
    Else
        frmAvisoErrorDatosMinimos.propTipoDetalle = "frmConsultaCuentas"
        frmAvisoErrorDatosMinimos.Show 1
    End If
End Sub

Private Sub menuGastosResumenHabitacion_Click()
    OprEjecutada = 10
    If mFunControlDeBaseDeDatos("frmConsultaCuentas") Then
        If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
            tipo_accion_ConsultaCuentas = 1
            tipo_accion_inghabitacion = 2
            frmIngHabitacion.Show 1
        End If
    Else
        frmAvisoErrorDatosMinimos.propTipoDetalle = "frmConsultaCuentas"
        frmAvisoErrorDatosMinimos.Show 1
    End If
End Sub

Private Sub mnuCheckOut_Click()
    'hora de inicio de la operación
    HoraIni = Time
    OprEjecutada = 54
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        tipo_accion_inghabitacion = 8
        frmIngHabitacion.Show 1
    End If
End Sub

Private Sub mnuCierreDiario_Click()
    'hora de inicio de la operación
    HoraIni = Time
    OprEjecutada = 57
    If mFunControlDeBaseDeDatos("frmCierreDiario") Then
        If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
            frmCierreDiario.Show 1
        End If
    Else
        frmAvisoErrorDatosMinimos.propTipoDetalle = "frmCierreDiario"
        frmAvisoErrorDatosMinimos.Show 1
    End If
End Sub

Private Sub mnuEstadoCuenta_Click()
    OprEjecutada = 58
    If mFunControlDeBaseDeDatos("frmEstadoCuentas") Then
        If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
            tipo_accion_IngEstadoCuenta = 1
            frmIngPaxEmp.Show 1
        End If
    Else
        frmAvisoErrorDatosMinimos.propTipoDetalle = "frmEstadoCuentas"
        frmAvisoErrorDatosMinimos.Show 1
    End If
End Sub

Private Sub mnuFacturacionDevolucionesConsultar_Click()
    OprEjecutada = 16
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        tipo_accion_tipodocumento = 3
        tipo_accion_devo = 3
        frmTipoDocumento.Show 1
    End If
End Sub

Private Sub mnuFacturacionDevolucionesEmitir_Click()
    'hora de inicio de la operación
    HoraIni = Time
    OprEjecutada = 15
    If mFunControlDeBaseDeDatos("frmFacturacion") Then
        If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
            tipo_accion_tipodocumento = 2
            tipo_accion_devo = 1
            frmTipoDocumento.Show 1
        End If
    Else
        frmAvisoErrorDatosMinimos.propTipoDetalle = "frmFacturacion"
        frmAvisoErrorDatosMinimos.Show 1
    End If
End Sub

Private Sub mnuFacturacionFacturasAnular_Click()
    'hora de inicio de la operación
    HoraIni = Time
    OprEjecutada = 14
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        tipo_accion_facturas = 3
        tipo_accion_tipodocumento = 1
        frmTipoDocumento.Show 1
    End If
End Sub

Private Sub mnuFacturacionFacturasConsultar_Click()
    OprEjecutada = 13
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        tipo_accion_facturas = 2
        tipo_accion_tipodocumento = 1
        frmTipoDocumento.Show 1
    End If
End Sub

Private Sub mnuFacturacionFacturasEmitir_Click()
    'hora de inicio de la operación
    HoraIni = Time
    OprEjecutada = 12
    If mFunControlDeBaseDeDatos("frmFacturacion") Then
        If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
            tipo_accion_facturas = 1
            tipo_accion_inghabitacion = 6
            frmIngHabitacion.Show 1
        End If
    Else
        frmAvisoErrorDatosMinimos.propTipoDetalle = "frmFacturacion"
        frmAvisoErrorDatosMinimos.Show 1
    End If
End Sub

Private Sub mnuHabitacionBloquear_Click()
    'hora de inicio de la operación
    HoraIni = Time

    OprEjecutada = 33
    If mFunControlDeBaseDeDatos("frmBloquearHab") Then
        If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
            tipo_accion_inghabitacion2 = 3
            frmIngHabitacion2.Show 1
        End If
    Else
        frmAvisoErrorDatosMinimos.propTipoDetalle = "frmBloquearHab"
        frmAvisoErrorDatosMinimos.Show 1
    End If
End Sub

Private Sub mnuHabitacionCambioFechaEgreso_Click()
    'hora de inicio de la operación
    HoraIni = Time
    OprEjecutada = 31
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        tipo_accion_inghabitacion = 9
        frmIngHabitacion.Show 1
    End If
End Sub

Private Sub mnuHabitacionCambioHabitacion_Click()
    'hora de inicio de la operación
    HoraIni = Time
    OprEjecutada = 34
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        frmCambioHabitacion.Show 1
    End If
End Sub

Private Sub mnuHabitacionCambioSituacion_Click()
    'hora de inicio de la operación
    HoraIni = Time
    OprEjecutada = 32
    If mFunControlDeBaseDeDatos("frmCambioSitu") Then
        If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
            tipo_accion_inghabitacion2 = 1
            frmIngHabitacion2.Show 1
        End If
    Else
        frmAvisoErrorDatosMinimos.propTipoDetalle = "frmCambioSitu"
        frmAvisoErrorDatosMinimos.Show 1
    End If
End Sub

Private Sub mnuHabitacionCambioTarifa_Click()
    'hora de inicio de la operación
    HoraIni = Time
    OprEjecutada = 35
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        tipo_accion_inghabitacion = 3
        frmIngHabitacion.Show 1
    End If
End Sub

Private Sub mnuHabitacionCambioTitular_Click()
    'hora de inicio de la operación
    HoraIni = Time
    OprEjecutada = 30
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        tipo_accion_inghabitacion = 5
        frmIngHabitacion.Show 1
    End If
End Sub

Private Sub mnuHabitacionConsultaTitular_Click()
    OprEjecutada = 36
    If mFunControlDeBaseDeDatos("frmConsultaTitular") Then
        If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
            frmConsultaTitular.Show 1
        End If
    Else
        frmAvisoErrorDatosMinimos.propTipoDetalle = "frmConsultaTitular"
        frmAvisoErrorDatosMinimos.Show 1
    End If
End Sub

Private Sub mnuInformesConsultaCompleta_Click()
    OprEjecutada = 24
    If mFunControlDeBaseDeDatos("frmConsultaCompleta") Then
        If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
            frmConsultaCompleta.Show 1
        End If
    Else
        frmAvisoErrorDatosMinimos.propTipoDetalle = "frmConsultaCompleta"
        frmAvisoErrorDatosMinimos.Show 1
    End If
End Sub

Private Sub mnuInformesCuadroSituacion_Click()
    OprEjecutada = 21
    If mFunControlDeBaseDeDatos("frmCuadroHab") Then
        If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
            frmCuadroHab.Show 1
        End If
    Else
        frmAvisoErrorDatosMinimos.propTipoDetalle = "frmCuadroHab"
        frmAvisoErrorDatosMinimos.Show 1
    End If
End Sub

Private Sub mnuInformesDisponibilidad_Click()
    OprEjecutada = 22
    If mFunControlDeBaseDeDatos("frmVerDisponibilidad") Then
        If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
            frmVerDisponibilidad.Show 1
        End If
    Else
        frmAvisoErrorDatosMinimos.propTipoDetalle = "frmVerDisponibilidad"
        frmAvisoErrorDatosMinimos.Show 1
    End If
End Sub

Private Sub mnuInformesEgresos_Click()
    OprEjecutada = 26
    If mFunControlDeBaseDeDatos("frmListadoEgresos") Then
        If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
            frmListadoEgresos.Show 1
        End If
    Else
        frmAvisoErrorDatosMinimos.propTipoDetalle = "frmListadoEgresos"
        frmAvisoErrorDatosMinimos.Show 1
    End If
End Sub

Private Sub mnuInformesIngresos_Click()
    OprEjecutada = 25
    If mFunControlDeBaseDeDatos("frmListadoIngresos") Then
        If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
            frmListadoIngresos.Show 1
        End If
    Else
        frmAvisoErrorDatosMinimos.propTipoDetalle = "frmListadoIngresos"
        frmAvisoErrorDatosMinimos.Show 1
    End If
End Sub

Private Sub mnuInformesPasajerosHabitacion_Click()
    OprEjecutada = 27
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        tipo_accion_inghabitacion = 7
        frmIngHabitacion.Show 1
    End If
End Sub

Private Sub mnuInformesUbicacionPasajeros_Click()
    OprEjecutada = 29
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        'Busca los pasajeros que están hospedados en el hotel actualmente.
        Dim cli_aux As String
        
        cli_aux = mFunBusqueda(2)
        If Val(cli_aux) <> 0 Then
            cliente_a_ubicar = cli_aux
            frmConsultaPasajeros.Show 1
        End If
    End If
End Sub

Private Sub mnuInformesPoblacionFlotante_Click()
    OprEjecutada = 28
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        frmPoblaciónFlotante.Show 1
    End If
End Sub

Private Sub mnuIngresoPasaCheckin_Click()
    'hora de inicio de la operación
    'NOTA: esta opción también se puede ejecutar desde frmListadoIngresos
    HoraIni = Time
    OprEjecutada = 5
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        tipo_accion_reserva = "Check-in"
        frmModificacionReserva.Show 1
        'por algun motivo que todavía se escapa a los alcances de mis conocimientos:
        'es necesario incluir esta línea de código para descargar el formulario y de
        'esta manera solucionar un error, que origina que no se ejecute el evento load
        'de este formulario no asignado correctamente variables importantes.
        Unload frmModificacionReserva
    End If
End Sub

Private Sub mnuIngresoPasaWalkin_Click()
    'hora de inicio de la operación
    HoraIni = Time
    OprEjecutada = 6
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        mSubCorrReservaWalkin
        tipo_accion_checkin = 2
        frmCheck_in.Show 1
    End If
End Sub

Private Sub mnuIngresoPasaWalkinHabOcupada_Click()
    'hora de inicio de la operación
    HoraIni = Time
    OprEjecutada = 7
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        mSubCorrReservaWalkin
        tipo_accion_checkin = 3
        frmCheck_in.Show 1
    End If
End Sub

Private Sub mnuMante_Click()
    'Ejecuto formulario de mantenimeinto
    frmMantenimientos.Show 1
End Sub

Private Sub mnuRecivosAnular_Click()
    'hora de inicio de la operación
    HoraIni = Time
    OprEjecutada = 19
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        tipo_accion_recivo = 6
        tipo_accion_tipodocumento = 7
        frmTipoDocumento.Show 1
    End If
End Sub

Private Sub mnuRecivosAutoAnular_Click()
    OprEjecutada = 48
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        tipo_accion_tipodocumento = 4
        tipo_accion_recivo = 5
        frmTipoDocumento.Show 1
    End If
End Sub

Private Sub mnuRecivosAutoConsultar_Click()
    OprEjecutada = 47
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        tipo_accion_tipodocumento = 4
        tipo_accion_recivo = 3
        frmTipoDocumento.Show 1
    End If
End Sub

Private Sub mnuRecivosAutoImprimir_Click()
    'hora de inicio de la operación
    HoraIni = Time
    OprEjecutada = 46
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        tipo_accion_tipodocumento = 5
        tipo_accion_recivo = 1
        frmTipoDocumento.Show 1
    End If
End Sub

Private Sub mnuRecivosConsultar_Click()
    OprEjecutada = 18
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        tipo_accion_recivo = 4
        tipo_accion_tipodocumento = 7
        frmTipoDocumento.Show 1
    End If
End Sub

Private Sub mnuRecivosIngresar_Click()
    'hora de inicio de la operación
    HoraIni = Time
    OprEjecutada = 17
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        tipo_accion_recivo = 2
        tipo_accion_tipodocumento = 6
        frmTipoDocumento.Show 1
    End If
End Sub

Private Sub mnuReservasAnular_Click()
    'hora de inicio de la operación
    HoraIni = Time
    OprEjecutada = 4
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        If mFunControlDeBaseDeDatos("frmCargaReserva") Then
            tipo_accion_reserva = "ANULAR"
            frmModificacionReserva.Show 1
        Else
            frmAvisoErrorDatosMinimos.propTipoDetalle = "frmCargaReserva"
            frmAvisoErrorDatosMinimos.Show 1
        End If
    End If
End Sub

Private Sub mnuReservasConsultar_Click()
    OprEjecutada = 3
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        If mFunControlDeBaseDeDatos("frmCargaReserva") Then
            tipo_accion_reserva = "CONSULTAR"
            frmModificacionReserva.Show 1
        Else
            frmAvisoErrorDatosMinimos.propTipoDetalle = "frmCargaReserva"
            frmAvisoErrorDatosMinimos.Show 1
        End If
    End If
End Sub

Private Sub mnuReservasModificar_Click()
    'hora de inicio de la operación
    HoraIni = Time
    OprEjecutada = 2
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        If mFunControlDeBaseDeDatos("frmCargaReserva") Then
            tipo_accion_reserva = "MODIFICAR"
            frmModificacionReserva.Show 1
        Else
            frmAvisoErrorDatosMinimos.propTipoDetalle = "frmCargaReserva"
            frmAvisoErrorDatosMinimos.Show 1
        End If
    End If
End Sub

Private Sub mnuReservasNueva_Click()
    'hora de inicio de la operación
    HoraIni = Time
    OprEjecutada = 1
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        If mFunControlDeBaseDeDatos("frmCargaReserva") Then
            tipo_accion_reserva = "ALTA"
            frmCargaReserva.Show 1
        Else
            frmAvisoErrorDatosMinimos.propTipoDetalle = "frmCargaReserva"
            frmAvisoErrorDatosMinimos.Show 1
        End If
    End If
End Sub

Private Sub mnuSisCong_Click()
    'hora de inicio de la operación
    HoraIni = Time

    OprEjecutada = 49
    If funUsuarioAutorizo(m_UsuarioSisNom, OprEjecutada) Then
        frmSisConfig.Show 1
    End If
End Sub

Private Sub mnuSalir_Click()
    'En el evento unload muestro aviso de versión demo si corresponde
    Unload Me
End Sub

Private Sub subApagoNumerosOpcion()
    'Cambio el boton de todos los números de opciones a color negro,
    'para establecer el color, de la etiqueta correspondiente al boton activo.
    Dim color As Long
    color = &H80000012  'negro
    Me.lblNroOpPrincipal(0).ForeColor = color
    Me.lblNroOpPrincipal(1).ForeColor = color
    Me.lblNroOpPrincipal(2).ForeColor = color
    Me.lblNroOpPrincipal(3).ForeColor = color
    Me.lblNroOpPrincipal(4).ForeColor = color
    Me.lblNroOpPrincipal(5).ForeColor = color
    Me.lblNroOpPrincipal(6).ForeColor = color
    Me.lblNroOpPrincipal(7).ForeColor = color
    Me.lblNroOpPrincipal(8).ForeColor = color
End Sub

'**************************************************************
'*
'*  Asistencia a usuarios
'*
'**************************************************************

Private Sub lstOpciones_Click()
    'Muestro en la barra de tareas, el mensaje correspondiente a la opción
    'seleccionada actualmente.
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, lstOpciones.ItemData(lstOpciones.ListIndex)
End Sub

Private Sub lstOpciones_GotFocus()
    'Al darle el focus desencadeno el evento click para que se muestre la
    'descripción del primer elemento de la lista en la barra de tareas.
    lstOpciones_Click
End Sub

Private Sub botReservas_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 148
    'ilumino número de opción
    subApagoNumerosOpcion
    Me.lblNroOpPrincipal(0).ForeColor = cColorActivo
End Sub

Private Sub botIngresos_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 149
    'ilumino número de opción
    subApagoNumerosOpcion
    Me.lblNroOpPrincipal(1).ForeColor = cColorActivo
End Sub

Private Sub botGastos_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 150
    'ilumino número de opción
    subApagoNumerosOpcion
    Me.lblNroOpPrincipal(2).ForeColor = cColorActivo
End Sub

Private Sub botFacturacion_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 151
    'ilumino número de opción
    subApagoNumerosOpcion
    Me.lblNroOpPrincipal(3).ForeColor = cColorActivo
End Sub

Private Sub botInformes_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 152
    'ilumino número de opción
    subApagoNumerosOpcion
    Me.lblNroOpPrincipal(4).ForeColor = cColorActivo
End Sub

Private Sub botHabitacion_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 153
    'ilumino número de opción
    subApagoNumerosOpcion
    Me.lblNroOpPrincipal(5).ForeColor = cColorActivo
End Sub

Private Sub botCheckout_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 154
    'ilumino número de opción
    subApagoNumerosOpcion
    Me.lblNroOpPrincipal(6).ForeColor = cColorActivo
End Sub

Private Sub botCierreDiario_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 155
    'ilumino número de opción
    subApagoNumerosOpcion
    Me.lblNroOpPrincipal(7).ForeColor = cColorActivo
End Sub

Private Sub botEstadosCuentas_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 156
    'ilumino número de opción
    subApagoNumerosOpcion
    Me.lblNroOpPrincipal(8).ForeColor = cColorActivo
End Sub

Private Sub botEstadosCuentas_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub botCierreDiario_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub botCheckout_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub botHabitacion_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub botInformes_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub botFacturacion_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub botGastos_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub botIngresos_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub botReservas_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

