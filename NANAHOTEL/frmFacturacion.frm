VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmFacturacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Facturación"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   11910
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin Hotel_Nana.gaHOTELtitular gaHOTELtitular1 
      Height          =   1335
      Left            =   120
      TabIndex        =   72
      TabStop         =   0   'False
      Top             =   0
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   2355
      BackColor       =   -2147483633
   End
   Begin TabDlg.SSTab sstab1 
      Height          =   6135
      Left            =   120
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   1440
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   10821
      _Version        =   327680
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmFacturacion.frx":0000
      Tab(0).ControlCount=   20
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "labDocu(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblImpMinimo(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblImpBasico(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblImpExento(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblIVAm(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblTotalGral(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblIVAb(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblNroDocu(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblFacturar(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblSignoMon(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "dbgrid1(0)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "botSalir"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "botAnular"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "botCancelar(0)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "botImprimir(0)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "msfgTotales(0)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Frame5(0)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "botSalirConsulta"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Frame1(0)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "cboFacturar(0)"
      Tab(0).Control(19).Enabled=   0   'False
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmFacturacion.frx":001C
      Tab(1).ControlCount=   18
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "labDocu(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblImpMinimo(1)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lblImpBasico(1)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lblImpExento(1)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lblIVAm(1)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "lblTotalGral(1)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "lblIVAb(1)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "lblNroDocu(1)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "lblFacturar(1)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "lblSignoMon(1)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "msfgTotales(1)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "dbgrid1(1)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "txtNroCli(1)"
      Tab(1).Control(12).Enabled=   -1  'True
      Tab(1).Control(13)=   "botCancelar(1)"
      Tab(1).Control(13).Enabled=   -1  'True
      Tab(1).Control(14)=   "botImprimir(1)"
      Tab(1).Control(14).Enabled=   -1  'True
      Tab(1).Control(15)=   "Frame5(1)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Frame1(1)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "cboFacturar(1)"
      Tab(1).Control(17).Enabled=   -1  'True
      Begin VB.ComboBox cboFacturar 
         Height          =   360
         Index           =   1
         Left            =   -74880
         Style           =   2  'Dropdown List
         TabIndex        =   76
         Top             =   5280
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.ComboBox cboFacturar 
         Height          =   360
         Index           =   0
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   74
         Top             =   5280
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Frame Frame1 
         Caption         =   "T&ipos de documentos"
         Height          =   2895
         Index           =   1
         Left            =   -68280
         TabIndex        =   3
         Top             =   2640
         Visible         =   0   'False
         Width           =   4695
         Begin VB.CommandButton botConfirmar 
            Caption         =   "&Confirmar"
            Height          =   375
            Index           =   1
            Left            =   3360
            TabIndex        =   5
            Tag             =   "Aceptar"
            Top             =   2280
            Width           =   1215
         End
         Begin VB.ListBox LstTipoDoc 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2160
            Index           =   1
            Left            =   240
            TabIndex        =   4
            Top             =   360
            Width           =   2655
         End
         Begin VB.Image Image1 
            Height          =   1215
            Index           =   1
            Left            =   3120
            Stretch         =   -1  'True
            Top             =   360
            Visible         =   0   'False
            Width           =   1335
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "&Tipos de documentos"
         Height          =   2775
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   2640
         Visible         =   0   'False
         Width           =   4695
         Begin VB.CommandButton botConfirmar 
            Caption         =   "&Confirmar"
            Height          =   375
            Index           =   0
            Left            =   3360
            TabIndex        =   2
            Tag             =   "Aceptar"
            Top             =   2160
            Width           =   1215
         End
         Begin VB.ListBox LstTipoDoc 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2160
            Index           =   0
            Left            =   240
            TabIndex        =   1
            Top             =   360
            Width           =   2655
         End
         Begin VB.Image Image1 
            Height          =   1215
            Index           =   0
            Left            =   3120
            Stretch         =   -1  'True
            Top             =   360
            Visible         =   0   'False
            Width           =   1335
         End
      End
      Begin VB.CommandButton botSalirConsulta 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   4920
         TabIndex        =   71
         Top             =   5690
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Frame Frame5 
         Caption         =   "Cabezal"
         Height          =   2175
         Index           =   0
         Left            =   120
         TabIndex        =   44
         Top             =   360
         Visible         =   0   'False
         Width           =   11295
         Begin VB.TextBox txtNom 
            Height          =   375
            Index           =   0
            Left            =   2040
            TabIndex        =   9
            Top             =   720
            Width           =   5055
         End
         Begin VB.CommandButton botModificar 
            Caption         =   "&Modificar"
            Height          =   375
            Index           =   0
            Left            =   9660
            TabIndex        =   47
            Top             =   1665
            Width           =   1215
         End
         Begin VB.TextBox fechaemi 
            Height          =   375
            Index           =   0
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox txtDir 
            Height          =   375
            Index           =   0
            Left            =   2040
            TabIndex        =   11
            Top             =   1200
            Width           =   5055
         End
         Begin VB.TextBox txtLoc 
            Height          =   375
            Index           =   0
            Left            =   2040
            TabIndex        =   13
            Top             =   1680
            Width           =   2055
         End
         Begin VB.ComboBox cboPais 
            Height          =   360
            Index           =   0
            Left            =   5520
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   1680
            Width           =   1575
         End
         Begin VB.TextBox txtCP 
            Alignment       =   1  'Right Justify
            Height          =   375
            Index           =   0
            Left            =   8280
            TabIndex        =   19
            Top             =   1200
            Width           =   1215
         End
         Begin VB.TextBox txtRuc 
            Alignment       =   1  'Right Justify
            Height          =   375
            Index           =   0
            Left            =   8280
            MaxLength       =   15
            TabIndex        =   17
            Top             =   720
            Width           =   2055
         End
         Begin VB.CommandButton botConsultar 
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
            Height          =   375
            Index           =   0
            Left            =   10440
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   720
            Width           =   435
         End
         Begin VB.TextBox txtNroCli 
            Height          =   285
            Index           =   0
            Left            =   4200
            TabIndex        =   45
            Text            =   "0"
            Top             =   360
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label lblCotizacion 
            Caption         =   "lblCotizacion"
            Height          =   255
            Index           =   0
            Left            =   8280
            TabIndex        =   77
            Top             =   300
            Width           =   2055
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "F&echa"
            Height          =   240
            Index           =   0
            Left            =   240
            TabIndex        =   6
            Top             =   300
            Width           =   570
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "&Nombre completo"
            Height          =   240
            Index           =   0
            Left            =   240
            TabIndex        =   8
            Top             =   780
            Width           =   1620
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "&Dirección"
            Height          =   240
            Index           =   0
            Left            =   240
            TabIndex        =   10
            Top             =   1260
            Width           =   855
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "&Localidad"
            Height          =   240
            Index           =   0
            Left            =   240
            TabIndex        =   12
            Top             =   1740
            Width           =   900
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "&Pais"
            Height          =   240
            Index           =   0
            Left            =   4800
            TabIndex        =   14
            Top             =   1740
            Width           =   405
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "&C.P."
            Height          =   240
            Index           =   0
            Left            =   7560
            TabIndex        =   18
            Top             =   1260
            Width           =   360
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "&R.U.C"
            Height          =   240
            Index           =   0
            Left            =   7560
            TabIndex        =   16
            Top             =   780
            Width           =   525
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Cabezal"
         Height          =   2175
         Index           =   1
         Left            =   -74880
         TabIndex        =   54
         Top             =   360
         Visible         =   0   'False
         Width           =   11295
         Begin VB.TextBox txtNom 
            Height          =   375
            Index           =   1
            Left            =   2040
            TabIndex        =   23
            Top             =   720
            Width           =   5055
         End
         Begin VB.CommandButton botModificar 
            Caption         =   "&Modificar"
            Height          =   375
            Index           =   1
            Left            =   9720
            TabIndex        =   56
            Top             =   1620
            Width           =   1215
         End
         Begin VB.TextBox fechaemi 
            Height          =   375
            Index           =   1
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   21
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox txtDir 
            Height          =   375
            Index           =   1
            Left            =   2040
            TabIndex        =   25
            Top             =   1200
            Width           =   5055
         End
         Begin VB.TextBox txtLoc 
            Height          =   375
            Index           =   1
            Left            =   2040
            TabIndex        =   27
            Top             =   1680
            Width           =   2055
         End
         Begin VB.ComboBox cboPais 
            Height          =   360
            Index           =   1
            Left            =   5520
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   1680
            Width           =   1575
         End
         Begin VB.TextBox txtCP 
            Alignment       =   1  'Right Justify
            Height          =   375
            Index           =   1
            Left            =   8280
            TabIndex        =   33
            Top             =   1200
            Width           =   1215
         End
         Begin VB.TextBox txtRuc 
            Alignment       =   1  'Right Justify
            Height          =   375
            Index           =   1
            Left            =   8280
            TabIndex        =   31
            Top             =   720
            Width           =   2055
         End
         Begin VB.CommandButton botConsultar 
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
            Height          =   375
            Index           =   1
            Left            =   10440
            TabIndex        =   55
            TabStop         =   0   'False
            Top             =   720
            Width           =   435
         End
         Begin VB.Label lblCotizacion 
            Caption         =   "lblCotizacion"
            Height          =   255
            Index           =   1
            Left            =   8280
            TabIndex        =   78
            Top             =   300
            Width           =   2055
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "F&echa"
            Height          =   240
            Index           =   1
            Left            =   240
            TabIndex        =   20
            Top             =   300
            Width           =   570
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "&Nombre completo"
            Height          =   240
            Index           =   1
            Left            =   240
            TabIndex        =   22
            Top             =   780
            Width           =   1620
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "&Dirección"
            Height          =   240
            Index           =   1
            Left            =   240
            TabIndex        =   24
            Top             =   1260
            Width           =   855
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "&Localidad"
            Height          =   240
            Index           =   1
            Left            =   240
            TabIndex        =   26
            Top             =   1740
            Width           =   900
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "&Pais"
            Height          =   240
            Index           =   1
            Left            =   4800
            TabIndex        =   28
            Top             =   1740
            Width           =   405
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "&C.P."
            Height          =   240
            Index           =   1
            Left            =   7560
            TabIndex        =   32
            Top             =   1260
            Width           =   360
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "&R.U.C"
            Height          =   240
            Index           =   1
            Left            =   7560
            TabIndex        =   30
            Top             =   780
            Width           =   525
         End
      End
      Begin MSFlexGridLib.MSFlexGrid msfgTotales 
         Height          =   615
         Index           =   0
         Left            =   4680
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   5040
         Visible         =   0   'False
         Width           =   6750
         _ExtentX        =   11906
         _ExtentY        =   1085
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
         ScrollBars      =   0
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CommandButton botImprimir 
         Height          =   375
         Index           =   0
         Left            =   8880
         Picture         =   "frmFacturacion.frx":0038
         Style           =   1  'Graphical
         TabIndex        =   43
         Tag             =   "Imprimir"
         Top             =   5690
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton botImprimir 
         Height          =   375
         Index           =   1
         Left            =   -66120
         Picture         =   "frmFacturacion.frx":097A
         Style           =   1  'Graphical
         TabIndex        =   42
         Tag             =   "Imprimir"
         Top             =   5690
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton botCancelar 
         Height          =   375
         Index           =   1
         Left            =   -64800
         Picture         =   "frmFacturacion.frx":12BC
         Style           =   1  'Graphical
         TabIndex        =   41
         Tag             =   "Cancelar"
         Top             =   5690
         Width           =   1215
      End
      Begin VB.CommandButton botCancelar 
         Cancel          =   -1  'True
         Height          =   375
         Index           =   0
         Left            =   10200
         Picture         =   "frmFacturacion.frx":1B7E
         Style           =   1  'Graphical
         TabIndex        =   40
         Tag             =   "Cancelar"
         Top             =   5690
         Width           =   1215
      End
      Begin VB.CommandButton botAnular 
         Caption         =   "&Anular"
         Height          =   375
         Left            =   6240
         TabIndex        =   39
         Top             =   5690
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtNroCli 
         Height          =   360
         Index           =   1
         Left            =   -71040
         TabIndex        =   38
         Text            =   "0"
         Top             =   720
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CommandButton botSalir 
         Height          =   375
         Left            =   7560
         Picture         =   "frmFacturacion.frx":2440
         Style           =   1  'Graphical
         TabIndex        =   37
         Tag             =   "Cancelar"
         Top             =   5690
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSFlexGridLib.MSFlexGrid dbgrid1 
         Height          =   2370
         Index           =   0
         Left            =   120
         TabIndex        =   48
         Top             =   2640
         Visible         =   0   'False
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   4180
         _Version        =   393216
         Cols            =   0
         FixedCols       =   0
         AllowBigSelection=   0   'False
         FocusRect       =   2
         HighLight       =   0
         ScrollBars      =   2
         AllowUserResizing=   1
         MousePointer    =   2
      End
      Begin MSFlexGridLib.MSFlexGrid dbgrid1 
         Height          =   2370
         Index           =   1
         Left            =   -74880
         TabIndex        =   49
         Top             =   2640
         Visible         =   0   'False
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   4180
         _Version        =   393216
         Cols            =   0
         FixedCols       =   0
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         ScrollBars      =   2
         AllowUserResizing=   1
         MousePointer    =   2
      End
      Begin MSFlexGridLib.MSFlexGrid msfgTotales 
         Height          =   615
         Index           =   1
         Left            =   -70320
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   5040
         Visible         =   0   'False
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   1085
         _Version        =   393216
         FixedCols       =   0
         ScrollBars      =   0
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblSignoMon 
         Caption         =   "lblSignoMon(1)"
         Height          =   255
         Index           =   1
         Left            =   -69840
         TabIndex        =   80
         Top             =   5760
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lblSignoMon 
         Caption         =   "lblSignoMon(0)"
         Height          =   255
         Index           =   0
         Left            =   3480
         TabIndex        =   79
         Top             =   5760
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lblFacturar 
         AutoSize        =   -1  'True
         Caption         =   "Fac&turar"
         Height          =   240
         Index           =   1
         Left            =   -74880
         TabIndex        =   75
         Top             =   5040
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lblFacturar 
         AutoSize        =   -1  'True
         Caption         =   "Fac&turar"
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   73
         Top             =   5040
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lblNroDocu 
         Caption         =   "0"
         Height          =   375
         Index           =   1
         Left            =   -70800
         TabIndex        =   70
         Top             =   5220
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblNroDocu 
         Caption         =   "0"
         Height          =   375
         Index           =   0
         Left            =   4080
         TabIndex        =   69
         Top             =   4995
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblIVAb 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Index           =   1
         Left            =   -72240
         TabIndex        =   68
         Top             =   5160
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label lblTotalGral 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Index           =   1
         Left            =   -71760
         TabIndex        =   67
         Top             =   5160
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label lblIVAm 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Index           =   1
         Left            =   -72840
         TabIndex        =   66
         Top             =   5160
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label lblImpExento 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Index           =   1
         Left            =   -73440
         TabIndex        =   65
         Top             =   5160
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label lblImpBasico 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Index           =   1
         Left            =   -74040
         TabIndex        =   64
         Top             =   5160
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label lblImpMinimo 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Index           =   1
         Left            =   -74640
         TabIndex        =   63
         Top             =   5160
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label lblIVAb 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Index           =   0
         Left            =   2760
         TabIndex        =   62
         Top             =   4995
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label lblTotalGral 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Index           =   0
         Left            =   3240
         TabIndex        =   61
         Top             =   4995
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label lblIVAm 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Index           =   0
         Left            =   2160
         TabIndex        =   60
         Top             =   4995
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label lblImpExento 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Index           =   0
         Left            =   1560
         TabIndex        =   59
         Top             =   4995
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label lblImpBasico 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Index           =   0
         Left            =   960
         TabIndex        =   58
         Top             =   4995
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label lblImpMinimo 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   57
         Top             =   4995
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label labDocu 
         AutoSize        =   -1  'True
         Caption         =   "labDocu(0)"
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
         Height          =   300
         Index           =   0
         Left            =   120
         TabIndex        =   51
         Top             =   5760
         Visible         =   0   'False
         Width           =   1410
      End
      Begin VB.Label labDocu 
         Caption         =   "labDocu(1)"
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
         Height          =   255
         Index           =   1
         Left            =   -74880
         TabIndex        =   50
         Top             =   5760
         Visible         =   0   'False
         Width           =   4575
      End
   End
   Begin Hotel_Nana.gaHOTELbarra gaHOTELbarra1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   7605
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   582
   End
   Begin Hotel_Nana.gaHOTELcli gaHOTELcli1 
      Height          =   735
      Left            =   4560
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   0
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   1296
      BackColor       =   -2147483633
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   5040
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   7
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFacturacion.frx":2D02
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFacturacion.frx":301C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFacturacion.frx":3336
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFacturacion.frx":3650
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFacturacion.frx":396A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFacturacion.frx":3C84
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFacturacion.frx":3F9E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFormulario 
      Caption         =   "&Formulario"
      Begin VB.Menu mnuFormularioAnular 
         Caption         =   "Anular           F12"
      End
      Begin VB.Menu mnuFormularioCancelarAnulacion 
         Caption         =   "Cancelar"
      End
      Begin VB.Menu mnuFormularioSalirConsulta 
         Caption         =   "Salir          F12"
      End
      Begin VB.Menu mnuFormularioImprimir 
         Caption         =   "Imprimir documento        "
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuFormularioCancelar 
         Caption         =   "Cancelar"
      End
   End
   Begin VB.Menu mnuVer 
      Caption         =   "&Ver información de..."
      Begin VB.Menu mnuVerExtras 
         Caption         =   "Extras"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuVerAloja 
         Caption         =   "Alojamiento"
         Shortcut        =   {F6}
      End
   End
End
Attribute VB_Name = "frmFacturacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private hab_cuenta As Long
Private tipocuenta As String
Private cotizacion As Single
Private m_Nro_Documento As Long
Private m_Tipo_Documento As Byte
Private EstoyCargandoCombo As Boolean   'controla que no se ejecute el procedimiento que
                                        'recalcula el total de la factura cuando se
                                        'inicializa el combo de selección.

Private Sub Form_Load()
    'Apariencia formulario
    mSubConfiguroFuentesControlesSistema Me
    
    'configuro menu de opciones
    subConfiguroMenuOpciones tipo_accion_facturas
    
    'configuro cabezal de grillas, estas líenas tiene que estar aunque solo utilize una sola
    'grilla ya que el procedimiento, grillas formato que está al final de este evento,
    'asume que las dos grillas tienen columnas.
    subCabezalGrilla Me.DBGrid1(0)
    subCabezalGrilla Me.DBGrid1(1)
            
    'cargo combos de facturar
    subCargoComboFacturar 0
    subCargoComboFacturar 1
    
    Select Case tipo_accion_facturas
        Case 1  'nueva
            hab_cuenta = Val(frmIngHabitacion.txtNroHab.Text)
            cotizacion = busco_cotiza
            'muestro cotización actual en formulario
            Me.lblCotizacion(0).Caption = "Cotización " & gblSignoMonedaNacional & " " & cotizacion
            Me.lblCotizacion(1).Caption = "Cotización " & gblSignoMonedaNacional & " " & cotizacion
            cabezal_formulario_habitacion
            
            SubSeleccionoDocumento
            
        Case 2  'consulta
            
            obtengo_datos_documento
            verifico_tipo_cuenta_cliente
        
            cabezal_formulario_cliente
            inicializo_form_cliente
            
            mSub_muestro_lineas_documento m_Tipo_Documento, m_Nro_Documento, frmFacturacion
          
            'Muestro grilla de totales
            mSubMuestro_Totales msfgTotales(0), _
                        Me.lblImpExento(0).Caption, _
                        Me.lblImpMinimo(0).Caption, _
                        Me.lblIVAm(0).Caption, _
                        Me.lblImpBasico(0).Caption, _
                        Me.lblIVAb(0).Caption, _
                        Me.lblTotalGral(0).Caption
                    
            'muestro cotización con que se realizó el documento en formulario
            Me.lblCotizacion(0).Caption = "Cotización " & gblSignoMonedaNacional & _
            " " & mFunObtengoCotiDocu(m_Tipo_Documento, m_Nro_Documento)
            
            'ordeno lineas grillas por fecha
            subOrdenoGrillaMSFlex DBGrid1(0), 1
            
        Case 3  'anulación
            
            obtengo_datos_documento
            verifico_tipo_cuenta_cliente
            
            cabezal_formulario_cliente
            inicializo_form_cliente
            
            mSub_muestro_lineas_documento m_Tipo_Documento, m_Nro_Documento, frmFacturacion
            
           
            'Muestro grilla de totales
            mSubMuestro_Totales msfgTotales(0), _
                        Me.lblImpExento(0).Caption, _
                        Me.lblImpMinimo(0).Caption, _
                        Me.lblIVAm(0).Caption, _
                        Me.lblImpBasico(0).Caption, _
                        Me.lblIVAb(0).Caption, _
                        Me.lblTotalGral(0).Caption
           
            
            'muestro cotización con que se realizó el documento en formulario
            Me.lblCotizacion(0).Caption = "Cotización " & gblSignoMonedaNacional & _
            " " & mFunObtengoCotiDocu(m_Tipo_Documento, m_Nro_Documento)
            
            'ordeno lineas grillas por fecha
            subOrdenoGrillaMSFlex DBGrid1(0), 1
    End Select
    
    grillas_formato 'oculto columnas
End Sub
    
Private Sub subCargoComboFacturar(i As Byte)
    'El combo de facturar contiene tres opciones que determinan cuales de
    'las líneas de la grilla se toman en cuenta para realizar el documento.
    EstoyCargandoCombo = True
    Me.cboFacturar(i).AddItem "Todas la líneas"
    Me.cboFacturar(i).AddItem "Solo las seleccionadas"
    Me.cboFacturar(i).AddItem "Solo las no seleccionadas"
    'por defecto selecciono la primera opción
    Me.cboFacturar(i).ListIndex = 0
    EstoyCargandoCombo = False
End Sub

Private Sub botConsultar_Click(Index As Integer)
    Dim nro_corr_aux As String
    Dim nrocliente As Long
    nro_corr_aux = mFunBusqueda(1) 'todos los clientes
    If Val(nro_corr_aux) <> 0 Then
        nrocliente = nro_corr_aux
        cargo_cabezal Index, nrocliente
    End If
End Sub

Private Sub botImprimir_Click(Index As Integer)
    Dim tipodocu As Byte
    
    'aviso de confirmación de impresión
    If mfunAplicoConfImp(1, 8) = 1 Then
        tipodocu = obtengo_digito(Index)
        Me.lblNroDocu(Index).Caption = mFun_obtengo_proximo_documento(tipodocu)
        If mFun_realizo_lineas(tipodocu, Me, Index) > 0 Then
                           
            'Grabo cabezal de la factura
            mSub_grabo_cabezal_documento _
                tipodocu, _
                Val(lblNroDocu(Index).Caption), _
                fechaemi(Index).Text, _
                txtNom(Index).Text, _
                txtDir(Index).Text, _
                txtLoc(Index).Text, _
                txtRuc(Index).Text, _
                txtCP(Index).Text, _
                cboPais(Index).ItemData(cboPais(Index).ListIndex), _
                txtNroCli(Index).Text, _
                lblTotalGral(Index).Caption, _
                0, _
                cotizacion, _
                lblIVAb(Index).Caption, _
                lblIVAm(Index).Caption, _
                lblImpExento(Index).Caption, _
                lblImpBasico(Index).Caption, _
                lblImpMinimo(Index).Caption, _
                mFun_PorIva(2), _
                mFun_PorIva(1), _
                mFun_PorIva(3), _
                cabFechaEntrada, _
                cabFechaSalida, _
                cabCantPax, cabTipoHab
                

                '[nro_fact_docu]
                'cuando es una factura lleva cero.
                'Este dato es utilizado para las devoluciones e indica la factura
                'a la cual estará asociada.
                               
            'actualizo estado de cuentas
            mSub_grabo_estado_cuentas _
                            tipodocu, _
                            Val(lblNroDocu(Index).Caption), _
                            Val(txtNroCli(Index).Text), _
                            Val(lblTotalGral(Index).Caption), _
                            fechaemi(Index).Text
            mSubArmoReporteFactura tipodocu, Val(lblNroDocu(Index).Caption)
            'aviso de emisión de documento
            mSubMensaje 4, 46, CStr(lblNroDocu(Index).Caption)
            
            'grabo bitácora
            GraboBitacora "Docu " & lblNroDocu(Index).Caption
            If Index = 0 Then
                If ssTab1.TabEnabled(1) = False Or ssTab1.TabVisible(1) = False Then
                    Unload Me
                    frmIngHabitacion.Show 1
                    Exit Sub
                End If
                ssTab1.Tab = 1
                ssTab1.TabEnabled(0) = False
            Else
                If ssTab1.TabEnabled(0) = False Then
                    Unload Me
                    frmIngHabitacion.Show 1
                End If
                ssTab1.Tab = 0
                ssTab1.TabEnabled(1) = False
            End If
            ssTab1.TabCaption(Index) = "Facturado: " & lblNroDocu(Index).Caption
        Else
            'no se seleccionaron gastos a imprimir
            mSubMensaje 4, 47
        End If
    End If
End Sub

Private Function obtengo_digito(i As Integer)
    obtengo_digito = Me.LstTipoDoc(i).ItemData(Me.LstTipoDoc(i).ListIndex)
End Function

Private Sub botModificar_Click(Index As Integer)
    If botModificar(Index).Caption = "&Modificar" Then
        mSub_cambio_cabezal True, Index, Me
        botModificar(Index).Caption = "C&onfirmar"
        botConsultar(Index).Enabled = True
        botImprimir(Index).Enabled = False
    Else
        mSub_cambio_cabezal False, Index, Me
        botModificar(Index).Caption = "&Modificar"
        botConsultar(Index).Enabled = False
        botImprimir(Index).Enabled = True
    End If
End Sub

Private Sub dbgrid1_DblClick(Index As Integer)
    marco_grilla DBGrid1(Index), 1, DBGrid1(Index).Cols - 1
    'Calculo nuevo total de la factura
    subRecalculoTotal (Index)
End Sub

Private Sub cboFacturar_Click(Index As Integer)
    'Calculo nuevo total de la factura siempre y cuando no este cargado el combo,
    'ya que este evento se ejecuta también cuando se carga el mismo, al iniciar
    'el formulario.
    If EstoyCargandoCombo = False Then
        subRecalculoTotal (Index)
    End If
End Sub

Private Sub dbgrid1_KeyPress(Index As Integer, KeyAscii As Integer)
    'Permite la selección de filas de la grilla con la tecla enter
    If KeyAscii = vbKeyReturn Then
        marco_grilla DBGrid1(Index), 1, DBGrid1(Index).Cols - 1
        'Calculo nuevo total de la factura
        subRecalculoTotal (Index)
    End If
End Sub

Private Sub grillas_formato()
    DBGrid1(0).ColWidth(10) = 0 'habitación del gasto
    DBGrid1(0).ColWidth(11) = 0 'nrocorr del gasto
    DBGrid1(0).ColWidth(12) = 0 'tipo
    
    DBGrid1(1).ColWidth(10) = 0 'habitación del gasto
    DBGrid1(1).ColWidth(11) = 0 'nrocorr del gasto
    DBGrid1(1).ColWidth(12) = 0 'tipo
    
End Sub

Private Sub SubSeleccionoDocumento()
    If tbHABITACIONES("titular_unica") <> 0 Then 'cuenta unica
        If Not FunBuscoGastosExtras(tbHABITACIONES("titular_unica")) And _
        Not FunBuscoGastosAloja(tbHABITACIONES("titular_unica")) Then
            ssTab1.TabVisible(1) = False
            ssTab1.TabCaption(0) = "No tiene gastos para facturar."
        Else
            ssTab1.TabVisible(1) = False
            ssTab1.TabCaption(0) = "Extras y alojamiento"
            subMuestroSeleccionDocumento 0
            subCargoDocumentosEnLista 0
            'Muestro primer documento por defecto
            LstTipoDoc(0).ListIndex = 0
        End If
    Else
        'permito cambiar fichas con teclas predeterminadas
        Me.mnuVer.Enabled = True
        If Not FunBuscoGastosExtras(tbHABITACIONES("titular_extra")) Then
            ssTab1.TabCaption(0) = "No tiene gastos extras"
        Else
            ssTab1.TabCaption(0) = "Extras"
            subMuestroSeleccionDocumento 0
            subCargoDocumentosEnLista 0
            'Muestro primer documento por defecto
            LstTipoDoc(0).ListIndex = 0
        End If
        If Not FunBuscoGastosAloja(tbHABITACIONES("titular_aloja")) Then
            ssTab1.TabCaption(1) = "No tiene gastos alojamiento"
        Else
            ssTab1.TabCaption(1) = "Alojamiento"
            subMuestroSeleccionDocumento 1
            subCargoDocumentosEnLista 1
            'Muestro primer documento por defecto
            LstTipoDoc(1).ListIndex = 0
        End If
    End If
    
End Sub

Private Sub subMuestroSeleccionDocumento(i As Byte)
    'Muestro frame de seleccion de documentos
    Frame1(i).Visible = True
    'Posiciono frame
    Frame1(i).Top = 360
    Frame1(i).Height = 5175
    botConfirmar(i).Top = 4680
End Sub

Private Sub botConfirmar_Click(Index As Integer)
    If Index = 0 Then
        subPrimeraFactura
    Else
        subSegundaFactura
    End If
End Sub

Private Sub subPrimeraFactura()
    'Realizo primera factura
    'Puede ser de extras solo o de estras y alojamientos, dependiendo
    'del tipo de cuenta de la habitación.
    
    mSub_cambio_cabezal False, 0, Me
    carga_tipo_pais frmFacturacion.cboPais(0)
    MuestroControles 0
    
    If tbHABITACIONES("titular_unica") <> 0 Then 'cuenta unica
        cargo_cabezal 0, tbHABITACIONES("titular_unica")
        'muestro líneas de extras y alojamiento en la misma factura
        cargo_grilla 2
    Else
        cargo_cabezal 0, tbHABITACIONES("titular_extra")
        'solo muestro líneas de extras
        cargo_grilla 1
    End If
    
   
    'Muestro grilla de totales
    mSubMuestro_Totales msfgTotales(0), _
        Me.lblImpExento(0).Caption, _
        Me.lblImpMinimo(0).Caption, _
        Me.lblIVAm(0).Caption, _
        Me.lblImpBasico(0).Caption, _
        Me.lblIVAb(0).Caption, _
        Me.lblTotalGral(0).Caption
   
        
    'ordeno lineas grillas por fecha
    subOrdenoGrillaMSFlex DBGrid1(0), 1
End Sub

Private Sub subSegundaFactura()
    'Realizo segunda factura,
    'siempre será de gastos de alojamiento.
    mSub_cambio_cabezal False, 1, Me
    carga_tipo_pais frmFacturacion.cboPais(1)
    cargo_cabezal 1, tbHABITACIONES("titular_aloja")
    MuestroControles 1
    
    'Muestro solo líneas de alojamiento
    cargo_grilla 3
    
   
    'Muestro grilla de totales
    mSubMuestro_Totales msfgTotales(1), _
        Me.lblImpExento(1).Caption, _
        Me.lblImpMinimo(1).Caption, _
        Me.lblIVAm(1).Caption, _
        Me.lblImpBasico(1).Caption, _
        Me.lblIVAb(1).Caption, _
        Me.lblTotalGral(1).Caption
   
        
    'ordeno lineas grillas por fecha
    subOrdenoGrillaMSFlex DBGrid1(1), 1
End Sub

Private Function FunBuscoGastosExtras(tit As Long)
    'Busco si el titular tiene gastos extras
    FunBuscoGastosExtras = False
    tbCUENTAS.Index = "i_titular"
    tbCUENTAS.Seek ">=", 0, tit, 0
    If Not tbCUENTAS.NoMatch Then   'si se posiciona
      If tbCUENTAS("titular_cuenta") = tit And _
            tbCUENTAS("facturado") = 0 Then
            FunBuscoGastosExtras = True 'tiene gastos extras
        End If
    End If
End Function

Private Function FunBuscoGastosAloja(tit As Long)
    'Busco si el titular tiene gastos alojamiento
    FunBuscoGastosAloja = False
    tbCUENTAS_ALOJA.Index = "i_titular"
    tbCUENTAS_ALOJA.Seek ">=", 0, tit, 0, 0
    If Not tbCUENTAS_ALOJA.NoMatch Then
        If tbCUENTAS_ALOJA("titular_aloja") = tit And _
        tbCUENTAS_ALOJA("facturado") = 0 Then
            FunBuscoGastosAloja = True  'tiene gastos alojamiento
        End If
    End If
End Function

Private Sub MuestroControles(i As Byte)
    'Prepara los controles del formulario para mostrar documnetos tanto en
    'el tabs de extras como en el tabs de alojamiento.
    'También configura las opciones del menú.
    
    'Muestro combo de facturar
    Me.lblFacturar(i).Visible = True
    Me.cboFacturar(i).Visible = True
    
    'Oculto frame de seleccion de documentos
    Frame1(i).Visible = False
    
    'Cambio título del formulario
    Me.Caption = "Facturación: nuevo documento"
    
    'Muestro cabezal de factuta
    Frame5(i).Visible = True
    
    'Muestro grillas lineas y totales
    DBGrid1(i).Visible = True
    msfgTotales(i).Visible = True
    
    labDocu(i).Caption = Me.LstTipoDoc(i).Text
    labDocu(i).Visible = True
    
    'Muestro los botones de imprimir
    Me.botImprimir(i).Visible = True
    
    'habilito opciones del menú
    Me.mnuFormularioImprimir.Enabled = True
End Sub

Private Sub verifico_tipo_cuenta_cliente()
    'A diferencia del procedimiento verifico_tipo_cuenta_habitación,
    'éste no necesita verificar el tipo de cuenta de una habitación,
    'ya que solo se mostrará en pantalla una factura,
    'y no nos interesa si pertenece a una facturación de alojamiento o extras.
    'Nota: simpre trabajaremos con un solo tabs(el 0)
    
    mSub_cambio_cabezal False, 0, Me
    carga_tipo_pais frmFacturacion.cboPais(0)
    'muestro solo un ficha
    ssTab1.TabVisible(1) = False
    MuestroControles 0
    
    'Obtengo el número de factura y la descripción del tipo de documento
    'para mostrarlo
    ssTab1.TabCaption(0) = frmTipoDocumento.LstTipoDoc.Text & " " & m_Nro_Documento
    
    
    'Los datos del cabezal se cargan con los datos que tengo en la factura
    mSub_cargo_cabezal_desde_documento m_Tipo_Documento, m_Nro_Documento, frmFacturacion, 0
End Sub

Private Sub cargo_grilla(tipo As Byte)
    'tipo=1 solo proceso gastos extras  (cuenta separadas)
    'tipo=2 proceso gastos extras y de alojamiento, los muestro en una misma factura (cuenta única)
    'tipo=3 solo proceso gastos alojamiento (cuenta separadas)
    
    Dim fechaAloja As Date
    'Cuando realizo en la factura la línea correspondiente al importe del alojamiento,
    'la misma, pertenecerá al total de todos los días hospedados en el hotel de cada habitación.
    'Por ese motivo no tiene mucho sentido incluir un campo fecha en la línea del total de alojamiento.
    'Sin embargo, para no dejar el campo en blanco, la línea de alojamiento siempre tendrá
    'la misma fecha que la factura, es decir la fecha atual del sistema.
    'Para los demás tipos de alojamiento (medio día, descuentos, correciones y otros) no ocurre lo mismo
    'Es necesario mantener esta información tal cual esta graba en el archivo de gastos para
    'poder acceder al mismo para eliminarlo.
                            
    Dim titular, hab_ant As Long
    Dim total_hab As Double
    
    'inicializo fecha para línea de alojamiento
    fechaAloja = m_FechaSis
    
    If tipo = 1 Or tipo = 2 Then
        'cargo extras
        'obtengo número de titular extras
        titular = busco_titular_hab2(hab_cuenta, "extra")
        
        tbCUENTAS.Index = "i_titular"
        tbCUENTAS.Seek ">=", 0, titular, 0
        If Not tbCUENTAS.NoMatch Then   'si se posiciona
            Do While Not tbCUENTAS.EOF
                If tbCUENTAS("titular_cuenta") = titular And _
                tbCUENTAS("facturado") = 0 Then
                    linea_factura
                    tbCUENTAS.MoveNext
                Else
                    Exit Do
                End If
            Loop
        End If
    End If
    
    If tipo = 2 Or tipo = 3 Then
        'obtengo número de titular aloja
        titular = busco_titular_hab2(hab_cuenta, "aloja")
    
        tbCUENTAS_ALOJA.Index = "i_TipoGastos"
        tbCUENTAS_ALOJA.Seek ">=", 0, titular, 0, 0
        If Not tbCUENTAS_ALOJA.NoMatch Then
            Do While Not tbCUENTAS_ALOJA.EOF
                If tbCUENTAS_ALOJA("titular_aloja") = titular And _
                tbCUENTAS_ALOJA("facturado") = 0 And tbCUENTAS_ALOJA("tipoAloja") <> 1 And _
                tbCUENTAS_ALOJA("tipoAloja") <> 4 Then
                    If tipo = 2 Then
                        'Muestro gastos alojamientos con gastos extras (cuenta única)
                        linea_factura2 tbCUENTAS_ALOJA("fecha"), _
                        tbCUENTAS_ALOJA("tarifa"), _
                        tbCUENTAS_ALOJA("habitacion_cuenta_aloja"), _
                        tbCUENTAS_ALOJA("habitacion_cuenta_aloja"), _
                        tbCUENTAS_ALOJA("nrocorr_cuenta_aloja"), _
                        tbCUENTAS_ALOJA("tipoAloja"), _
                        tbCUENTAS_ALOJA("obsAloja")
                    End If
                    If tipo = 3 Then
                        'Muestro gastos alojamiento en grilla de alojamiento    (cuenta separadas)
                        linea_factura_aloja tbCUENTAS_ALOJA("fecha"), _
                        tbCUENTAS_ALOJA("tarifa"), _
                        tbCUENTAS_ALOJA("habitacion_cuenta_aloja"), _
                        tbCUENTAS_ALOJA("habitacion_cuenta_aloja"), _
                        tbCUENTAS_ALOJA("nrocorr_cuenta_aloja"), _
                        tbCUENTAS_ALOJA("tipoAloja"), _
                        tbCUENTAS_ALOJA("obsAloja")
                    End If
                End If
                tbCUENTAS_ALOJA.MoveNext
            Loop
        End If
        
        'recorro nuevamente para mostrar agrupados los gastos de alojamiento.
        'tipo automáticos y corrección por habitación
        tbCUENTAS_ALOJA.Index = "i_TipoGastos"
        tbCUENTAS_ALOJA.Seek ">=", 0, titular, 0, 0
        If Not tbCUENTAS_ALOJA.NoMatch Then
            'recorro todos los gastos de alojamiento de titular
            Do While Not tbCUENTAS_ALOJA.EOF
                If tbCUENTAS_ALOJA("facturado") = 0 And tbCUENTAS_ALOJA("titular_aloja") = titular Then
                    'cargo variables de inicio del corte por hbitación
                    total_hab = 0: hab_ant = tbCUENTAS_ALOJA("habitacion_cuenta_aloja")
                    'realizo corte por habitación discriminando solo los gastos de tipo 1 y 4
                    Do While Not tbCUENTAS_ALOJA.EOF
                        If tbCUENTAS_ALOJA("facturado") = 0 And tbCUENTAS_ALOJA("titular_aloja") = titular And tbCUENTAS_ALOJA("habitacion_cuenta_aloja") = hab_ant Then
                            If (tbCUENTAS_ALOJA("tipoAloja") = 1 Or tbCUENTAS_ALOJA("tipoAloja") = 4) Then
                                total_hab = total_hab + tbCUENTAS_ALOJA("tarifa")
                            End If
                        Else
                            Exit Do
                        End If
                        tbCUENTAS_ALOJA.MoveNext
                    Loop
                    'muestro total de alojamiento para esa habitación
                    'Muestro gastos alojamientos con gastos extras (cuenta única)
                    If tipo = 2 Then linea_factura2 fechaAloja, total_hab, hab_ant, hab_ant, 0, 1, "Total Período alojamiento"
                    'Muestro gastos alojamiento en grilla de alojamiento    (cuenta separadas)
                    If tipo = 3 Then linea_factura_aloja fechaAloja, total_hab, hab_ant, hab_ant, 0, 1, "Total Período alojamiento"
                    'saldo del recorrido ya que los demás gastos no me interesan
                Else
                    Exit Do
                End If
            Loop
        End If
    End If
End Sub

Private Sub linea_factura_aloja(fecha As Date, _
                                tot_aloja As Double, _
                                hab As Long, _
                                k1 As Long, _
                                k2 As Long, _
                                descriAloja As Integer, _
                                obsAloja As String)
    'Creo lineas de gastos alojamiento en factura de alojamiento (cuentas separadas)
    Dim linea_factura As String
    Dim tipo As String
    Dim total_conver As Double
    Dim nacCliFactura As Integer

    'calculo total convertido de dolares a la moneda de la factura
    total_conver = total_convertido(tot_aloja, 1)
    
    'discrimino iva e imponible dependiendo del tipo de iva del alojamiento
    nacCliFactura = Me.cboPais(1).ItemData(Me.cboPais(1).ListIndex)
    SubCalculoIva mFunTipoIvaALoja(nacCliFactura, 1), total_conver, 1

    'calculo total general factura
    Me.lblTotalGral(1).Caption = Val(Me.lblTotalGral(1).Caption) + total_conver
        
    tipo = gblSignoDolares  'la moneda del alojamiento es siempre en dólares
    
    linea_factura = _
    Chr(9) & _
    fecha & _
    Chr(9) & _
    hab & _
    Chr(9) & _
    descriAloja & _
    Chr(9) & _
    funObtengoDescSisConstantes(1, descriAloja) & " " & obsAloja & _
    Chr(9) & _
    tipo & _
    Chr(9) & _
    Chr(9) & _
    Chr(9) & _
    Format(tot_aloja, "#####0.00") & _
    Chr(9) & _
    Format(total_conver, "#####0.00") & _
    Chr(9) & _
    k1 & _
    Chr(9) & _
    k2 & _
    Chr(9) & _
    "a"

    DBGrid1(1).AddItem linea_factura
End Sub

Private Sub linea_factura2(fecha As Date, _
                            total_aloja As Double, _
                            hab As Long, _
                            k1 As Long, _
                            k2 As Long, _
                            descriAloja As Integer, _
                            obsAloja As String)
    'Creo lineas de gastos alojamiento en grilla compartida con gastos extras (cuenta unica)
    Dim linea_factura As String
    Dim tipo As String
    Dim total_conver As Double
    Dim nacCliFactura As Integer
    
    'calculo total convertido de dolares a la moneda de la factura
    total_conver = total_convertido(total_aloja, 1)
    
    'discrimino iva e imponible dependiendo del tipo de iva del alojamiento
    nacCliFactura = Me.cboPais(1).ItemData(Me.cboPais(1).ListIndex)
    SubCalculoIva mFunTipoIvaALoja(nacCliFactura, 1), total_conver, 0

    'calculo total general factura
    Me.lblTotalGral(0).Caption = Val(Me.lblTotalGral(0).Caption) + total_conver
    
    tipo = gblSignoDolares  'la moneda del alojamiento es siempre en dólares
    
    linea_factura = _
    Chr(9) & _
    fecha & _
    Chr(9) & _
    hab & _
    Chr(9) & _
    descriAloja & _
    Chr(9) & _
    funObtengoDescSisConstantes(1, descriAloja) & " " & obsAloja & _
    Chr(9) & _
    tipo & _
    Chr(9) & _
    Chr(9) & _
    Chr(9) & _
    Format(total_aloja, "#####0.00") & _
    Chr(9) & _
    Format(total_conver, "#####0.00") & _
    Chr(9) & _
    k1 & _
    Chr(9) & _
    k2 & _
    Chr(9) & _
    "a"

    DBGrid1(0).AddItem linea_factura
End Sub

Private Sub linea_factura()
    'Creo lineas de gastos extras
    'Grilla correspondiente a cuenta única o separadas.
    Dim linea_factura
    Dim descri_art As String
    Dim tipo As String
    Dim total_art As Double
    Dim total_art_conv As Double
    Dim importe_art As Double
    Dim CodIvaArt As Byte
    
    'obtengo descripción articulo
    If busco_articuloTF(tbCUENTAS("articulo_cuenta")) Then
        descri_art = tbARTICULOS("descriarticulo")
        CodIvaArt = tbARTICULOS("CodIvaArticulo")
    End If
    
    'obtengo importe extra y total
    If tbCUENTAS("moneda_cuenta") = 0 Then 'm/n
        tipo = gblSignoMonedaNacional
        importe_art = tbCUENTAS("importe_mnacional_cuenta")
        total_art = tbCUENTAS("total_mnacional_cuenta")
    Else
        tipo = gblSignoDolares
        importe_art = tbCUENTAS("importe_dolares_cuenta")
        total_art = tbCUENTAS("total_dolares_cuenta")
    End If
    
    'calculo total convertido
    total_art_conv = total_convertido(total_art, tbCUENTAS("moneda_cuenta"))
    
    'discrimino iva e imponible
    SubCalculoIva CodIvaArt, total_art_conv, 0

    'calculo total general factura
    Me.lblTotalGral(0).Caption = Val(Me.lblTotalGral(0).Caption) + total_art_conv
                    
    
    linea_factura = _
    Chr(9) & _
    tbCUENTAS("fechagasto_cuenta") & _
    Chr(9) & _
    tbCUENTAS("habitacion_cuenta") & _
    Chr(9) & _
    tbCUENTAS("articulo_cuenta") & _
    Chr(9) & _
    descri_art & _
    Chr(9) & _
    tipo & _
    Chr(9) & _
    tbCUENTAS("cantidad_cuenta") & _
    Chr(9) & _
    Format(importe_art, "#####0.00") & _
    Chr(9) & _
    Format(total_art, "#####0.00") & _
    Chr(9) & _
    Format(total_art_conv, "#####0.00") & _
    Chr(9) & _
    tbCUENTAS("habitacion_cuenta") & _
    Chr(9) & _
    tbCUENTAS("nrocorr_cuenta") & _
    Chr(9) & _
    "e"

    DBGrid1(0).AddItem linea_factura
End Sub

Private Sub SubCalculoIva(tipoIva As Byte, totalArt As Double, i As Integer)
    Dim iva As Double
    Dim imponible As Single
    Dim porcentajeIva As Single
    
    porcentajeIva = mFun_PorIva(tipoIva)
        
    If tipoIva <> 3 Then    ' si es exento no puedo dividir por 0
        imponible = (totalArt / ((porcentajeIva / 100) + 1)) 'calculo imponible
        iva = totalArt - imponible                      'calculo iva
    End If
    
    If tipoIva = 1 Then 'basico
        Me.lblImpBasico(i).Caption = Val(lblImpBasico(i).Caption) + imponible
        Me.lblIVAb(i).Caption = Val(Me.lblIVAb(i).Caption) + iva
    Else
        If tipoIva = 2 Then 'minimo
            Me.lblImpMinimo(i).Caption = Val(Me.lblImpMinimo(i).Caption) + imponible
            Me.lblIVAm(i).Caption = Val(Me.lblIVAm(i).Caption) + iva
        Else
            If tipoIva = 3 Then 'exento
                Me.lblImpExento(i).Caption = Val(Me.lblImpExento(i).Caption) + totalArt
            End If
        End If
    End If
End Sub

Private Sub subRecalculoTotal(indGrilla As Integer)
    '----------------------------------------------------------------------------------
    'Calculo el nuevo total de la factura, cuando se seleccionan líneas de la
    'misma, es deir no se realiza la facturación total de los gastos.
    '----------------------------------------------------------------------------------
    'Parámetros.
    '   Entrada [indGrilla] grilla de donde voy a obtener la información.
    '                       0 = grilla de gastos exyras o               (c. separadas)
    '                       0 = grilla de gastos extras y alojamiento   (c. única)
    '                       1 = grilla de gastos alojamiento            (c. separadas)
    '----------------------------------------------------------------------------------
    Dim j As Integer
    Dim tipodocu As Byte
    Dim impGasto As Double
    Dim codArt As Long
    Dim tipoIva As Byte
    
    'inicializo etiquetas de totales, antes de calcular los nuevos valores
    subInicializoEtiquetasDeTotales (indGrilla)
    'inicializo contador de filas recorridas en grilla de gastos
    j = 2
    'obtengo el tipo de documento que estoy realizando.
    tipodocu = obtengo_digito(indGrilla)
    
    'recorro todas las líneas de la grilla (todos los gastos a facturar)
    Do While j < Me.DBGrid1(indGrilla).Rows
        Me.DBGrid1(indGrilla).row = j
        'aplico filtro de líneas
        If valido_linea(indGrilla, Me, tipodocu) Then
                'obtengo tipo de iva del gasto
                'tengo que verificar que el gasto no sea un alojamiento
                If mFunctionValorCeldaMSFGrid(Me.DBGrid1(indGrilla), 12, _
                    Me.DBGrid1(indGrilla).row) = "a" Then
                    'es un gasto alojamiento
                    'el iva del alojamiento se obtiene del archivo de parámetros
                    tipoIva = mFunTipoIvaALoja(Me.cboPais(1).ItemData(Me.cboPais(1).ListIndex), 1)
                Else
                    'es un gasto extra
                    codArt = mFunctionValorCeldaMSFGrid(Me.DBGrid1(indGrilla), 3, _
                            Me.DBGrid1(indGrilla).row)
                    tipoIva = mFunObtengoCodIvaArticulo(codArt)
                End If
            'obtengo importe del gasto
            impGasto = mFunctionValorCeldaMSFGrid(Me.DBGrid1(indGrilla), 9, _
                       Me.DBGrid1(indGrilla).row)
            
            'calculo nuevos importes (iva, imponibles y total exento)
            SubCalculoIva tipoIva, impGasto, indGrilla
            'calculo total general.
            Me.lblTotalGral(indGrilla).Caption = _
            Val(Me.lblTotalGral(indGrilla).Caption) + impGasto
        End If
        j = j + 1
    Loop
    'muestro nuevos totales en pantalla
    mSubMuestro_Totales msfgTotales(indGrilla), _
        Me.lblImpExento(indGrilla).Caption, _
        Me.lblImpMinimo(indGrilla).Caption, _
        Me.lblIVAm(indGrilla).Caption, _
        Me.lblImpBasico(indGrilla).Caption, _
        Me.lblIVAb(indGrilla).Caption, _
        Me.lblTotalGral(indGrilla).Caption
       
End Sub

Private Sub subInicializoEtiquetasDeTotales(ind As Integer)
    '----------------------------------------------------------------------
    'Antes de cargar estas variables con el nuevo total, es necesario
    'inicializarlas a 0.
    '----------------------------------------------------------------------
    Me.lblImpExento(ind).Caption = 0
    Me.lblImpMinimo(ind).Caption = 0
    Me.lblIVAm(ind).Caption = 0
    Me.lblImpBasico(ind).Caption = 0
    Me.lblIVAb(ind).Caption = 0
    Me.lblTotalGral(ind).Caption = 0
End Sub
Private Sub cargo_cabezal(i As Integer, cli As Long)
    'Cargo el cabezal con los datos del cliente a facturar (nueva factura)
    If busco_clienteTF(cli) Then
        txtNom(i).Text = tbCLIENTES("nombre_completo_titular")
        txtDir(i).Text = tbCLIENTES("direccion_titular")
        txtLoc(i).Text = tbCLIENTES("localidad_titular")
        txtRuc(i).Text = tbCLIENTES("ruc_titular")
        txtCP(i).Text = tbCLIENTES("cod_postal_titular")
        txtNroCli(i).Text = cli
        posiciono_combo cboPais(i), tbCLIENTES("pais_reside_titular")
        fechaemi(i).Text = m_FechaSis
    End If
End Sub

Private Function total_convertido(total_a_convertir As Double, moneda_gasto As Byte)
    Dim moneda_a_convertir As Byte
    Dim digito As Byte
    
    'obtengo la moneda del documento que estoy trabajando
    digito = obtengo_digito(Me.ssTab1.Tab)
    If digito = 1 Or digito = 3 Then
        moneda_a_convertir = 0  'm/n
    Else
        moneda_a_convertir = 1  'dol
    End If
    
    'El gasto esta en M/N y tengo que pasar a Dol
    If moneda_gasto = 0 Then 'm/n
        If moneda_a_convertir = 1 Then  'dol
            total_convertido = total_a_convertir / cotizacion
        Else
            total_convertido = total_a_convertir
        End If
    Else
    'El gasto esta en Dol y tengo que pasar a M/N
        If moneda_a_convertir = 0 Then  'm/n
            total_convertido = total_a_convertir * cotizacion
        Else
            total_convertido = total_a_convertir
        End If
    End If
End Function

Private Sub cabezal_formulario_habitacion()
    'Oculto cabezal de cliente
    Me.gaHOTELcli1.Visible = False
    
    Me.gaHOTELtitular1.Width = 11535
    Me.gaHOTELtitular1.CaminoBaseDeDatos = vardir
    Me.gaHOTELtitular1.NumeroHabitacion = hab_cuenta
End Sub

Private Sub cabezal_formulario_cliente()
    'Oculto cabezal de habitación
    Me.gaHOTELtitular1.Visible = False
    
    'Posiciono y cambio tamaño
    Me.gaHOTELcli1.Left = 120
    Me.gaHOTELcli1.Width = 11535
    
    Me.gaHOTELcli1.CaminoBaseDeDatos = vardir
    Me.gaHOTELcli1.CodigoCliente = txtNroCli(0).Text
End Sub

Private Sub inicializo_form_cliente()
    'No muestro el combo de facturar
    Me.cboFacturar(0).Visible = False
    Me.lblFacturar(0).Visible = False
    Me.cboFacturar(1).Visible = False
    Me.lblFacturar(1).Visible = False
    
    'No muesrto la etiqueta de seleccion de documento
    labDocu(0).Visible = False
    labDocu(1).Visible = False
    
    'No muestro los botones del cabezal de factura
    botConsultar(0).Enabled = False
    botConsultar(1).Enabled = False
    botModificar(0).Enabled = False
    botModificar(1).Enabled = False
    
    'No muesrto los botones de imprimir y cancelar
    botImprimir(0).Visible = False
    botImprimir(1).Visible = False
    botCancelar(0).Visible = False
    botCancelar(1).Visible = False
    
    'Como el cabezal de clientes es más angosto que el de habitaciones
    'subo un poco el cuerpo de la factura
    Me.ssTab1.Top = 960
    
    'Cambio los botones de imprimir y cancelar por
    'anular  y salir o
    'salir unicamente
    
    If tipo_accion_facturas = 2 Then    'consulto
        'Cambio titulo del formulario
        Me.Caption = "Facturación: consulto documento"
        botAnular.Visible = False
        botSalirConsulta.Visible = True
        botSalirConsulta.Top = 5690
        botSalirConsulta.Left = 10200
    End If

    If tipo_accion_facturas = 3 Then    'anulo
        'Cambio titulo del formulario
        Me.Caption = "Facturación: anulo documento"
        
        'aparece más volcado a la izquierda que el de salir
        botAnular.Visible = True
        botSalir.Visible = True
        botSalir.Top = 5690
        botSalir.Left = 10200
        botAnular.Top = 5690
        botAnular.Left = 8880
    End If
    
    'No permito trabajar con la grilla de lineas
    DBGrid1(0).Enabled = False
End Sub

Private Sub obtengo_datos_documento()
    m_Nro_Documento = frmTipoDocumento.txtNroDoc.Text
    m_Tipo_Documento = frmTipoDocumento.LstTipoDoc.ItemData(frmTipoDocumento.LstTipoDoc.ListIndex)
End Sub

Private Sub botSalir_Click()
    Unload Me
    frmTipoDocumento.Show 1
End Sub

Private Sub botAnular_Click()
    'aviso de confirmación de anulación de documneto
    
    If mFunMensaje(4, 48) Then

        mSub_creo_gastos_nuevamente m_Tipo_Documento, m_Nro_Documento, Me.txtNroCli(0)
    
        'elimino documento anulado
        mSub_Elimino_Documento m_Tipo_Documento, m_Nro_Documento
        
        'elimino documento del estado de cuentas, solo si es crédito.
        mSub_Elimino_Documento_EstadoCuenta m_Tipo_Documento, m_Nro_Documento
        
        'aviso de confirmación de anulación
        mSubMensaje 4, 49
        
        'grabo bitácora
        GraboBitacora "Docu " & m_Nro_Documento
        Unload Me
        frmTipoDocumento.Show 1
    End If
End Sub

Private Sub subCargoDocumentosEnLista(i As Integer)
    'contado moneda nacional
    LstTipoDoc(i).AddItem "Contado " & gblSignoMonedaNacional
    LstTipoDoc(i).ItemData(LstTipoDoc(i).NewIndex) = 1
    'contado dólares
    LstTipoDoc(i).AddItem "Contado " & gblSignoDolares
    LstTipoDoc(i).ItemData(LstTipoDoc(i).NewIndex) = 2
    'factura moneda nacional
    LstTipoDoc(i).AddItem "Factura " & gblSignoMonedaNacional
    LstTipoDoc(i).ItemData(LstTipoDoc(i).NewIndex) = 3
    'factura dólares
    LstTipoDoc(i).AddItem "Factura " & gblSignoDolares
    LstTipoDoc(i).ItemData(LstTipoDoc(i).NewIndex) = 4
End Sub

Private Sub botCancelar_Click(Index As Integer)
    Unload Me
    frmIngHabitacion.Show 1
End Sub

Private Sub botSalirConsulta_Click()
    Unload Me
    frmTipoDocumento.Show 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Descargo formulario de memoria
    Set frmFacturacion = Nothing
End Sub

Private Sub LstTipoDoc_Click(Index As Integer)
    'Cargo la imagen correspondiente al documento seleccionado
    Image1(Index).Picture = _
    ImageList1.ListImages.Item(LstTipoDoc(Index).ItemData(LstTipoDoc(Index).ListIndex)).Picture

End Sub

Private Sub LstTipoDoc_DblClick(Index As Integer)
    botConfirmar_Click (Index)
End Sub

Private Sub mnuFormularioAnular_Click()
    'Equivale a presionar el boton de anular, cuando estoy anulando un documento,
    'o la tecla F12.
    botAnular_Click
End Sub

Private Sub mnuFormularioCancelar_Click()
    'Equivale a presionar el boton de cancelar o la tecla ESC.
    
    'Como existe una matriz de controles tengo que identificar, que boton de la matriz
    'tengo que presionar.
    botCancelar_Click (Me.ssTab1.Tab)
End Sub

Private Sub mnuFormularioCancelarAnulacion_Click()
    'Equivale a presionar el boton de cancelar, cuando estoy anulando un documento
    botSalir_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 27
            'Dependiendo de la tarea del formulario, al cerrar el mismo debo de regresar
            'al fomulario de ingreso de habitación o al de tipo de documento.
            If tipo_accion_facturas = 1 Then    'nueva
                'regreso a ingreso de habitación
                botCancelar_Click (0)
            End If
            
            If tipo_accion_facturas = 2 Then    'cosultar
                'regreso a tipo de documento
                botSalirConsulta_Click
            End If
            
            If tipo_accion_facturas = 3 Then    'anular
                'regreso a tipo de documento
                botSalir_Click
            End If
            
        Case vbKeyF9
            If tipo_accion_facturas = 1 Then    'solo trabaja cuando estoy creando un nuevo documento
                'verifico que opción esta activa
                If botConfirmar(Me.ssTab1.Tab).Visible = True Then
                    'confirmo documento
                    botConfirmar_Click (Me.ssTab1.Tab)
                Else
                    If botModificar(Me.ssTab1.Tab).Visible = True Then
                        'confirmo nuevos datos del cliente en el documento
                        botModificar_Click (Me.ssTab1.Tab)
                    End If
                End If
            End If
            
        Case vbKeyF12
            If tipo_accion_facturas = 2 Then        'consultar documento
                'cierro la consulta
                botSalirConsulta_Click
            Else
                If tipo_accion_facturas = 3 Then    'anular documento
                    'anulo el documento
                    botAnular_Click
                End If
            End If
        
        Case vbKeyF1
            If tipo_accion_facturas = 1 Then    'solo trabaja cuando estoy creando un nuevo documento
                If botConsultar(Me.ssTab1.Tab).Enabled = True Then
                    botConsultar_Click (Me.ssTab1.Tab)
                End If
            End If
    End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'Intercepto la tecla F9 y F12 las cuales no son interceptadas por el evento keypress
    If KeyCode = vbKeyF9 Or _
        KeyCode = vbKeyF12 Or _
        KeyCode = vbKeyF1 Then
        Form_KeyPress (KeyCode)
    End If
End Sub

Private Sub mnuFormularioImprimir_Click()
    'Equivale a presionar el boton de imprimir o la tecla Ctrol+I
    botImprimir_Click (Me.ssTab1.Tab)
End Sub

Private Sub mnuFormularioSalirConsulta_Click()
    'Equivale a presionar el boton de salir, cuando se consulta un documento, o
    'la tecla F12.
    botSalirConsulta_Click
End Sub

Private Sub subConfiguroMenuOpciones(tipoMenu As Byte)
    'Configuro las opciones que se muestran en el menú de opciones,
    'dependiendo de la tarea que realiza el formulario.
    Select Case tipoMenu
        Case 1  'nuevo documento
            'op. de anular
            Me.mnuFormularioAnular.Visible = False
            Me.mnuFormularioCancelarAnulacion.Visible = False
            'op. de consultar
            Me.mnuFormularioSalirConsulta.Visible = False
            'configuro estado incial de opciones
            Me.mnuVer.Enabled = False
            Me.mnuFormularioImprimir.Enabled = False
            
        Case 2  'consultar documento
            'op. de anular
            Me.mnuFormularioAnular.Visible = False
            Me.mnuFormularioCancelarAnulacion.Visible = False
            'op. de nueo documento
            Me.mnuVer.Visible = False
            Me.mnuFormularioImprimir.Visible = False
            Me.mnuFormularioCancelar.Visible = False
            
        Case 3  'anular documento
            'op. de consultar
            Me.mnuFormularioSalirConsulta.Visible = False
            'op. de nueo documento
            Me.mnuVer.Visible = False
            Me.mnuFormularioImprimir.Visible = False
            Me.mnuFormularioCancelar.Visible = False
    End Select
End Sub

Private Sub mnuVerAloja_Click()
    'Cambio de tab visible
    If Me.ssTab1.TabEnabled(1) = True Then
        Me.ssTab1.Tab = 1
    End If
End Sub

Private Sub mnuVerExtras_Click()
    'Cambio de tab visible
    If Me.ssTab1.TabEnabled(0) = True Then
        Me.ssTab1.Tab = 0
    End If
End Sub

'************************************************************
'*
'*  Asistencia a usuario
'*
'************************************************************

Private Sub txtNom_GotFocus(Index As Integer)
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 122
End Sub

Private Sub txtDir_GotFocus(Index As Integer)
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 123
End Sub

Private Sub txtLoc_GotFocus(Index As Integer)
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 124
End Sub

Private Sub cboPais_GotFocus(Index As Integer)
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 133
End Sub

Private Sub txtRuc_GotFocus(Index As Integer)
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 125
End Sub

Private Sub txtCP_GotFocus(Index As Integer)
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 126
End Sub

Private Sub botModificar_GotFocus(Index As Integer)
    'Este boton tiene dos usos por lo que es necesario identificar
    'que uso le estoy dando al momento de darle el focus.
    If Me.botModificar(Index).Caption = "&Modificar" Then
        'permito modificar los datos del cabezal
        mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 127
    Else
        mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 128
        'confirmo los datos del cabezal
    End If
End Sub

Private Sub cboFacturar_GotFocus(Index As Integer)
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 129
End Sub

Private Sub botConfirmar_GotFocus(Index As Integer)
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 121
End Sub

Private Sub LstTipoDoc_GotFocus(Index As Integer)
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 120
End Sub

Private Sub botSalirConsulta_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 1, 1
End Sub

Private Sub botAnular_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 130
End Sub

Private Sub botSalir_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 1, 1
End Sub

Private Sub botImprimir_GotFocus(Index As Integer)
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 131
End Sub

Private Sub botCancelar_GotFocus(Index As Integer)
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 1, 3
End Sub

Private Sub dbgrid1_GotFocus(Index As Integer)
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 132
End Sub

Private Sub LstTipoDoc_LostFocus(Index As Integer)
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub dbgrid1_LostFocus(Index As Integer)
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub botCancelar_LostFocus(Index As Integer)
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub cboFacturar_LostFocus(Index As Integer)
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub botConfirmar_LostFocus(Index As Integer)
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub botSalirConsulta_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub botAnular_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub botSalir_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub botImprimir_LostFocus(Index As Integer)
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub cboPais_LostFocus(Index As Integer)
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub txtRuc_LostFocus(Index As Integer)
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub txtCP_LostFocus(Index As Integer)
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub botModificar_LostFocus(Index As Integer)
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub txtNom_LostFocus(Index As Integer)
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub txtDir_LostFocus(Index As Integer)
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub txtLoc_LostFocus(Index As Integer)
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

