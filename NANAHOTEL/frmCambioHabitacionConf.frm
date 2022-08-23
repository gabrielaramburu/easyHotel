VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmCambioHabitacionConf 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Resumen de cambios de habitación"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   6480
      TabIndex        =   26
      Text            =   "Text2"
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   5400
      TabIndex        =   25
      Text            =   "Text1"
      Top             =   3840
      Width           =   1215
   End
   Begin Hotel_Nana.gaHOTELbarra gaHOTELbarra1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   24
      Top             =   7890
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   582
   End
   Begin VB.CommandButton botAceptar 
      Height          =   375
      Left            =   10560
      Picture         =   "frmCambioHabitacionConf.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Tag             =   "Aceptar"
      Top             =   7440
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Situación actual de habitación HACIA "
      Height          =   3495
      Left            =   120
      TabIndex        =   1
      Top             =   3840
      Width           =   11655
      Begin VB.Data Data2 
         Caption         =   "Data2"
         Connect         =   "Access"
         DatabaseName    =   "C:\NANAHOTEL\hotel.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   8640
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select * from checkin,clientes"
         Top             =   480
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox txtTit1H 
         Height          =   285
         Left            =   1560
         TabIndex        =   14
         Top             =   360
         Width           =   4575
      End
      Begin VB.TextBox txtTit2H 
         Height          =   285
         Left            =   1560
         TabIndex        =   13
         Top             =   720
         Width           =   4575
      End
      Begin VB.Frame Frame6 
         Caption         =   "Situación"
         Height          =   975
         Left            =   6840
         TabIndex        =   10
         Top             =   2160
         Width           =   3975
         Begin VB.PictureBox Picture4 
            Height          =   615
            Left            =   240
            ScaleHeight     =   555
            ScaleWidth      =   1035
            TabIndex        =   19
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label LSituHacia 
            Caption         =   "LSituHacia"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1560
            TabIndex        =   22
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Estado"
         Height          =   975
         Left            =   6840
         TabIndex        =   9
         Top             =   1080
         Width           =   3975
         Begin VB.PictureBox Picture3 
            Height          =   615
            Left            =   240
            ScaleHeight     =   555
            ScaleWidth      =   1035
            TabIndex        =   18
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label LEstadoHacia 
            Caption         =   "LEstadoHacia"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1560
            TabIndex        =   21
            Top             =   360
            Width           =   1455
         End
      End
      Begin MSDBGrid.DBGrid gPasaHacia 
         Bindings        =   "frmCambioHabitacionConf.frx":08B6
         Height          =   2055
         Left            =   240
         OleObjectBlob   =   "frmCambioHabitacionConf.frx":08C6
         TabIndex        =   4
         Top             =   1200
         Width           =   6135
      End
      Begin VB.Label LTit1Hacia 
         Caption         =   "LTit1Hacia"
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label LTit2Hacia 
         Caption         =   "LTit2Hacia"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Situación actual de habitación DESDE"
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   11655
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   "C:\NANAHOTEL\hotel.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   8760
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select * from checkin,clientes"
         Top             =   240
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox txtTit2D 
         Height          =   285
         Left            =   1800
         TabIndex        =   12
         Top             =   720
         Width           =   4575
      End
      Begin VB.TextBox txtTit1D 
         ForeColor       =   &H80000012&
         Height          =   285
         Left            =   1800
         TabIndex        =   11
         Top             =   360
         Width           =   4575
      End
      Begin VB.Frame Frame5 
         Caption         =   "Situación"
         Height          =   975
         Left            =   6840
         TabIndex        =   6
         Top             =   2160
         Width           =   3975
         Begin VB.Image Image1 
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Left            =   240
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label LSituDesde 
            Caption         =   "LSituDesde"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1560
            TabIndex        =   23
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Estado"
         Height          =   975
         Left            =   6840
         TabIndex        =   5
         Top             =   1080
         Width           =   3975
         Begin VB.PictureBox Picture1 
            Height          =   615
            Left            =   240
            ScaleHeight     =   555
            ScaleWidth      =   1035
            TabIndex        =   17
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label LEstadoDesde 
            Caption         =   "LEstadoDesde"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1560
            TabIndex        =   20
            Top             =   480
            Width           =   1815
         End
      End
      Begin MSDBGrid.DBGrid gPasaDesde 
         Bindings        =   "frmCambioHabitacionConf.frx":129B
         Height          =   2055
         Left            =   240
         OleObjectBlob   =   "frmCambioHabitacionConf.frx":12AB
         TabIndex        =   3
         Top             =   1200
         Width           =   6135
      End
      Begin VB.Label LTit2desde 
         Caption         =   "LTit2desde"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label LTit1desde 
         Caption         =   "LTit1desde"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmCambioHabitacionConf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private hab_desde As Long
Private hab_hacia As Long

Private Sub botAceptar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    'Apariencia formulario
    mSubConfiguroFuentesControlesSistema Me
    
    mSub_bloqueo_controles_formulario Me, True
    hab_desde = frmCambioHabitacion.txtHabDesde.Text
    hab_hacia = frmCambioHabitacion.txtHabHacia.Text
    
    muestro_datos_habitaciones
    SQLpasajeros_habitacion hab_desde, Data1
    SQLpasajeros_habitacion hab_hacia, Data2
End Sub

Private Sub muestro_datos_habitaciones()
    'Habitación desde
    Me.Frame1.Caption = frmCambioHabitacion.LhabDesde.Caption
    
    If busco_habitaTF(hab_desde) Then
        If busco_estado_habTF(2, tbHABITACIONES("situacionhab")) Then
            LSituDesde.Caption = tbTIPO_ESTADO_HAB("descri")
        End If
        'si la habitación queda ocupada: muestro titulares
        If busco_habita_checkin(hab_desde) Then
            LEstadoDesde.Caption = "Ocupada"
            
            muestro_titular LTit1desde, LTit2desde, txtTit1D, txtTit2D
        Else    'si la habitacion esta libre
            LEstadoDesde.Caption = "Libre"
            
            LTit1desde.Visible = False
            LTit2desde.Visible = False
            txtTit1D.Visible = False
            txtTit2D.Visible = False
        End If
    End If
            
    If busco_habitaTF(hab_hacia) Then
        If busco_estado_habTF(2, tbHABITACIONES("situacionhab")) Then
            LSituHacia.Caption = tbTIPO_ESTADO_HAB("descri")
        End If
        
        LEstadoHacia.Caption = "Ocupada"
        'Habitación hacia
        Me.Frame2.Caption = frmCambioHabitacion.LhabHacia.Caption    'muestro titulares
        If busco_habitaTF(hab_hacia) Then
            muestro_titular LTit1Hacia, LTit2Hacia, txtTit1H, txtTit2H
        End If
    End If
End Sub

Private Sub muestro_titular(l1 As label, l2 As label, t1 As TextBox, t2 As TextBox)
    If tbHABITACIONES("titular_unica") <> 0 Then 'unico titular
        l1.Caption = "Titular único"
        t1.Text = busco_titular_hab(tbHABITACIONES("nrohab"), "unica")
        t2.Visible = False
        l2.Visible = False
    Else
        l1.Caption = "Titular aloja. "
        l2.Caption = "Titular extras"
        t1.Text = busco_titular_hab(tbHABITACIONES("nrohab"), "aloja")
        t2.Text = busco_titular_hab(tbHABITACIONES("nrohab"), "extra")
    End If
End Sub


