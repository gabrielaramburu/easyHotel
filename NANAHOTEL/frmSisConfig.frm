VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmSisConfig 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuarci�n del sistema"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7455
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5775
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   10186
      _Version        =   327680
      Style           =   1
      Tabs            =   8
      Tab             =   7
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "&Colores"
      TabPicture(0)   =   "frmSisConfig.frx":0000
      Tab(0).ControlCount=   2
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame8"
      Tab(0).Control(1).Enabled=   0   'False
      TabCaption(1)   =   "&Fuentes"
      TabPicture(1)   =   "frmSisConfig.frx":001C
      Tab(1).ControlCount=   1
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame9"
      Tab(1).Control(0).Enabled=   0   'False
      TabCaption(2)   =   "&General"
      TabPicture(2)   =   "frmSisConfig.frx":0038
      Tab(2).ControlCount=   1
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame10"
      Tab(2).Control(0).Enabled=   0   'False
      TabCaption(3)   =   "Acc&esos directos"
      TabPicture(3)   =   "frmSisConfig.frx":0054
      Tab(3).ControlCount=   1
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame2"
      Tab(3).Control(0).Enabled=   0   'False
      TabCaption(4)   =   "C&uadro de habitaciones"
      TabPicture(4)   =   "frmSisConfig.frx":0070
      Tab(4).ControlCount=   1
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame3"
      Tab(4).Control(0).Enabled=   0   'False
      TabCaption(5)   =   "Cua&dro de disponibilidad"
      TabPicture(5)   =   "frmSisConfig.frx":008C
      Tab(5).ControlCount=   1
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame11"
      Tab(5).Control(0).Enabled=   0   'False
      TabCaption(6)   =   "&Listados"
      TabPicture(6)   =   "frmSisConfig.frx":00A8
      Tab(6).ControlCount=   1
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame16"
      Tab(6).Control(0).Enabled=   0   'False
      TabCaption(7)   =   "F&acturaci�n"
      TabPicture(7)   =   "frmSisConfig.frx":00C4
      Tab(7).ControlCount=   1
      Tab(7).ControlEnabled=   -1  'True
      Tab(7).Control(0)=   "Frame20"
      Tab(7).Control(0).Enabled=   0   'False
      Begin VB.Frame Frame20 
         Height          =   4935
         Left            =   120
         TabIndex        =   125
         Top             =   720
         Width           =   6855
         Begin VB.ComboBox cboNacionalidades 
            Height          =   360
            Left            =   3960
            Style           =   2  'Dropdown List
            TabIndex        =   135
            Top             =   1800
            Width           =   2775
         End
         Begin VB.TextBox txtCantViasDocu 
            Height          =   375
            Left            =   360
            MaxLength       =   2
            TabIndex        =   133
            Top             =   2760
            Width           =   975
         End
         Begin VB.ComboBox cboImpAlojaExtranjero 
            Height          =   360
            Left            =   360
            Style           =   2  'Dropdown List
            TabIndex        =   131
            Top             =   1800
            Width           =   3015
         End
         Begin VB.CheckBox chkImpAlojaExtranjeros 
            Caption         =   "&Aplicar diferente tipo de impuesto en alojamiento a extranjeros"
            Height          =   375
            Left            =   360
            TabIndex        =   129
            Top             =   1200
            Width           =   6135
         End
         Begin VB.ComboBox cboImpAloja 
            Height          =   360
            Left            =   360
            Style           =   2  'Dropdown List
            TabIndex        =   128
            Top             =   600
            Width           =   3015
         End
         Begin VB.CheckBox chkMostarTotales 
            Caption         =   "&Mostar totales de documentos de forma resumida."
            Height          =   255
            Left            =   480
            TabIndex        =   126
            Top             =   4080
            Width           =   4935
         End
         Begin VB.Label lblCantMinVias 
            AutoSize        =   -1  'True
            Caption         =   "lblCantMinVias"
            Height          =   240
            Left            =   1680
            TabIndex        =   136
            Top             =   2820
            Width           =   1335
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            Caption         =   "&Nacionalidad local"
            Height          =   240
            Left            =   3960
            TabIndex        =   134
            Top             =   1560
            Width           =   1695
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            Caption         =   "Cantidad de &v�as documentos"
            Height          =   240
            Left            =   360
            TabIndex        =   132
            Top             =   2520
            Width           =   2670
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            Caption         =   "T&ipo impuesto alojamiento extranjero"
            Height          =   240
            Left            =   360
            TabIndex        =   130
            Top             =   1560
            Width           =   3315
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            Caption         =   "&Tipo impuesto alojamiento"
            Height          =   240
            Left            =   360
            TabIndex        =   127
            Top             =   360
            Width           =   2385
         End
      End
      Begin VB.Frame Frame11 
         BorderStyle     =   0  'None
         Height          =   4935
         Left            =   -74880
         TabIndex        =   79
         Top             =   720
         Width           =   6855
         Begin VB.Frame Frame18 
            Caption         =   "Iluminar"
            Height          =   2895
            Left            =   2760
            TabIndex        =   103
            Top             =   0
            Width           =   4095
            Begin VB.CheckBox chkPpioA�oDis 
               Caption         =   "Principio de a�o"
               Height          =   195
               Left            =   240
               TabIndex        =   110
               Top             =   1125
               Width           =   1935
            End
            Begin VB.CheckBox chkPpioMesDis 
               Caption         =   "Principio de mes"
               Height          =   255
               Left            =   240
               TabIndex        =   109
               Top             =   495
               Width           =   1815
            End
            Begin VB.CommandButton botColorMesDis 
               Height          =   285
               Left            =   3045
               Picture         =   "frmSisConfig.frx":00E0
               Style           =   1  'Graphical
               TabIndex        =   108
               Top             =   480
               Visible         =   0   'False
               Width           =   300
            End
            Begin VB.CommandButton botColorA�oDis 
               Height          =   285
               Left            =   3045
               Picture         =   "frmSisConfig.frx":038A
               Style           =   1  'Graphical
               TabIndex        =   107
               Top             =   1080
               Visible         =   0   'False
               Width           =   300
            End
            Begin VB.CommandButton botColor1SemanaDis 
               Height          =   285
               Left            =   2685
               Picture         =   "frmSisConfig.frx":0634
               Style           =   1  'Graphical
               TabIndex        =   106
               Top             =   1800
               Visible         =   0   'False
               Width           =   300
            End
            Begin VB.CheckBox chkSemanaDis 
               Caption         =   "Cada semana"
               Height          =   195
               Left            =   240
               TabIndex        =   105
               Top             =   1800
               Width           =   1575
            End
            Begin VB.CommandButton botColor2SemanaDis 
               Height          =   285
               Left            =   2685
               Picture         =   "frmSisConfig.frx":08DE
               Style           =   1  'Graphical
               TabIndex        =   104
               Top             =   2400
               Visible         =   0   'False
               Width           =   300
            End
            Begin VB.Label lblVentanaColorMesDis 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   285
               Left            =   2280
               TabIndex        =   118
               Top             =   480
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.Label lblColorMesDis 
               AutoSize        =   -1  'True
               Caption         =   "Color mes"
               Height          =   195
               Left            =   2280
               TabIndex        =   117
               Top             =   240
               Visible         =   0   'False
               Width           =   690
            End
            Begin VB.Label lblVentanaColorA�oDis 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   285
               Left            =   2280
               TabIndex        =   116
               Top             =   1080
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.Label lblColorA�oDis 
               AutoSize        =   -1  'True
               Caption         =   "Color a�o"
               Height          =   195
               Left            =   2280
               TabIndex        =   115
               Top             =   840
               Visible         =   0   'False
               Width           =   675
            End
            Begin VB.Label lblVentanaColor1SemanaDis 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   285
               Left            =   1920
               TabIndex        =   114
               Top             =   1800
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.Label lblColor1SemanaDis 
               AutoSize        =   -1  'True
               Caption         =   "Color 1era. semana"
               Height          =   195
               Left            =   1920
               TabIndex        =   113
               Top             =   1560
               Visible         =   0   'False
               Width           =   1365
            End
            Begin VB.Label lblVentanaColor2SemanaDis 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   285
               Left            =   1920
               TabIndex        =   112
               Top             =   2400
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.Label lblColor2SemanaDis 
               AutoSize        =   -1  'True
               Caption         =   "Color 2da. semana"
               Height          =   195
               Left            =   1920
               TabIndex        =   111
               Top             =   2160
               Visible         =   0   'False
               Width           =   1320
            End
         End
         Begin VB.Frame Frame13 
            Caption         =   "Tama�o de celdas"
            Height          =   1335
            Left            =   0
            TabIndex        =   96
            Top             =   0
            Width           =   2655
            Begin VB.TextBox txtAnchoCeldaDis 
               Height          =   315
               Left            =   840
               MaxLength       =   3
               TabIndex        =   98
               Top             =   360
               Width           =   495
            End
            Begin VB.TextBox txtLargoCeldaDis 
               Height          =   315
               Left            =   840
               MaxLength       =   2
               TabIndex        =   97
               Top             =   840
               Width           =   495
            End
            Begin VB.Label Label28 
               AutoSize        =   -1  'True
               Caption         =   "Ancho"
               Height          =   195
               Left            =   120
               TabIndex        =   102
               Top             =   420
               Width           =   465
            End
            Begin VB.Label Label27 
               AutoSize        =   -1  'True
               Caption         =   "Largo"
               Height          =   195
               Left            =   120
               TabIndex        =   101
               Top             =   900
               Width           =   405
            End
            Begin VB.Label Label26 
               AutoSize        =   -1  'True
               Caption         =   "caracteres"
               Height          =   240
               Left            =   1440
               TabIndex        =   100
               Top             =   870
               Width           =   960
            End
            Begin VB.Label Label25 
               AutoSize        =   -1  'True
               Caption         =   "pixeles"
               Height          =   195
               Left            =   1440
               TabIndex        =   99
               Top             =   420
               Width           =   480
            End
         End
         Begin VB.Frame Frame12 
            Caption         =   "Rango de fechas predeterminado"
            Height          =   975
            Left            =   2760
            TabIndex        =   93
            Top             =   3000
            Width           =   4095
            Begin VB.TextBox txtCantDiasDis 
               Height          =   315
               Left            =   120
               TabIndex        =   94
               Top             =   360
               Width           =   615
            End
            Begin VB.Label Label29 
               AutoSize        =   -1  'True
               Caption         =   "d�as a partir de la fecha del sistema"
               Height          =   195
               Left            =   840
               TabIndex        =   95
               Top             =   420
               Width           =   2505
            End
         End
         Begin VB.Frame Frame14 
            Caption         =   "Mostrar"
            Height          =   1095
            Left            =   0
            TabIndex        =   89
            Top             =   1440
            Width           =   2655
            Begin VB.CheckBox chkMuestroIconoOcupada 
               Caption         =   "Icono de ocupadas"
               Height          =   255
               Left            =   120
               TabIndex        =   91
               Top             =   240
               Width           =   2295
            End
            Begin VB.ComboBox cboAlinIcono 
               Height          =   360
               ItemData        =   "frmSisConfig.frx":0B88
               Left            =   1200
               List            =   "frmSisConfig.frx":0B95
               Style           =   2  'Dropdown List
               TabIndex        =   90
               Top             =   600
               Width           =   1215
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               Caption         =   "Alineaci�n"
               Height          =   195
               Left            =   120
               TabIndex        =   92
               Top             =   660
               Width           =   735
            End
         End
         Begin VB.Frame Frame15 
            Caption         =   "Fuentes"
            Height          =   2175
            Left            =   0
            TabIndex        =   80
            Top             =   2640
            Width           =   2655
            Begin VB.ComboBox cboTama�oDigitosDis 
               Height          =   360
               ItemData        =   "frmSisConfig.frx":0BB5
               Left            =   1680
               List            =   "frmSisConfig.frx":0BC5
               Style           =   2  'Dropdown List
               TabIndex        =   83
               Top             =   480
               Width           =   735
            End
            Begin VB.ComboBox cboTama�oLetrasDis 
               Height          =   360
               ItemData        =   "frmSisConfig.frx":0BD8
               Left            =   1680
               List            =   "frmSisConfig.frx":0BE8
               Style           =   2  'Dropdown List
               TabIndex        =   82
               Top             =   1200
               Width           =   735
            End
            Begin VB.ComboBox cboAlinFuente 
               Height          =   360
               ItemData        =   "frmSisConfig.frx":0BFB
               Left            =   1200
               List            =   "frmSisConfig.frx":0C08
               Style           =   2  'Dropdown List
               TabIndex        =   81
               Top             =   1680
               Width           =   1215
            End
            Begin VB.Label Label31 
               AutoSize        =   -1  'True
               Caption         =   "Tama�o"
               Height          =   240
               Left            =   1560
               TabIndex        =   88
               Top             =   240
               Width           =   765
            End
            Begin VB.Label Label30 
               AutoSize        =   -1  'True
               Caption         =   "Tama�o d�gitos"
               Height          =   195
               Left            =   120
               TabIndex        =   87
               Top             =   480
               Width           =   1110
            End
            Begin VB.Label Label32 
               AutoSize        =   -1  'True
               Caption         =   "Tama�o letras"
               Height          =   195
               Left            =   120
               TabIndex        =   86
               Top             =   1200
               Width           =   1005
            End
            Begin VB.Label Label33 
               AutoSize        =   -1  'True
               Caption         =   "Tama�o"
               Height          =   240
               Left            =   1560
               TabIndex        =   85
               Top             =   960
               Width           =   765
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               Caption         =   "Alineaci�n"
               Height          =   195
               Left            =   120
               TabIndex        =   84
               Top             =   1740
               Width           =   735
            End
         End
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   4935
         Left            =   -74880
         TabIndex        =   46
         Top             =   720
         Width           =   6855
         Begin VB.Frame Frame17 
            Caption         =   "Tama�o de celdas"
            Height          =   1335
            Left            =   0
            TabIndex        =   72
            Top             =   0
            Width           =   2535
            Begin VB.TextBox txtAnchoCelda 
               Height          =   315
               Left            =   840
               MaxLength       =   3
               TabIndex        =   74
               Top             =   360
               Width           =   495
            End
            Begin VB.TextBox txtLargoCelda 
               Height          =   315
               Left            =   840
               MaxLength       =   2
               TabIndex        =   73
               Top             =   840
               Width           =   495
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "Ancho"
               Height          =   195
               Left            =   120
               TabIndex        =   78
               Top             =   420
               Width           =   465
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "Largo"
               Height          =   195
               Left            =   120
               TabIndex        =   77
               Top             =   900
               Width           =   405
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               Caption         =   "caracteres"
               Height          =   240
               Left            =   1440
               TabIndex        =   76
               Top             =   840
               Width           =   960
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               Caption         =   "pixeles"
               Height          =   195
               Left            =   1440
               TabIndex        =   75
               Top             =   420
               Width           =   480
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Rango de fechas predeterminado"
            Height          =   975
            Left            =   2640
            TabIndex        =   69
            Top             =   3240
            Width           =   4095
            Begin VB.TextBox txtCantDias 
               Height          =   315
               Left            =   120
               TabIndex        =   70
               Top             =   360
               Width           =   615
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               Caption         =   "d�as a partir de la fecha del sistema"
               Height          =   195
               Left            =   840
               TabIndex        =   71
               Top             =   420
               Width           =   2505
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "No asignadas"
            Height          =   1095
            Left            =   0
            TabIndex        =   66
            Top             =   1440
            Width           =   2535
            Begin VB.ComboBox cboMostrar 
               Height          =   360
               ItemData        =   "frmSisConfig.frx":0C28
               Left            =   1080
               List            =   "frmSisConfig.frx":0C32
               Style           =   2  'Dropdown List
               TabIndex        =   67
               Top             =   360
               Width           =   1095
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               Caption         =   "Mostrar"
               Height          =   195
               Left            =   240
               TabIndex        =   68
               Top             =   360
               Width           =   525
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Iluminar"
            Height          =   3135
            Left            =   2640
            TabIndex        =   50
            Top             =   0
            Width           =   4095
            Begin VB.CheckBox chkPpioA�o 
               Caption         =   "Principio de a�o"
               Height          =   195
               Left            =   240
               TabIndex        =   57
               Top             =   1245
               Width           =   1815
            End
            Begin VB.CheckBox chkPpioMes 
               Caption         =   "Principio de mes"
               Height          =   255
               Left            =   240
               TabIndex        =   56
               Top             =   480
               Width           =   1815
            End
            Begin VB.CommandButton botColorMes 
               Height          =   285
               Left            =   2925
               Picture         =   "frmSisConfig.frx":0C4E
               Style           =   1  'Graphical
               TabIndex        =   55
               Top             =   480
               Visible         =   0   'False
               Width           =   300
            End
            Begin VB.CommandButton botColorA�o 
               Height          =   285
               Left            =   2925
               Picture         =   "frmSisConfig.frx":0EF8
               Style           =   1  'Graphical
               TabIndex        =   54
               Top             =   1200
               Visible         =   0   'False
               Width           =   300
            End
            Begin VB.CommandButton botColor1Semana 
               Height          =   285
               Left            =   2925
               Picture         =   "frmSisConfig.frx":11A2
               Style           =   1  'Graphical
               TabIndex        =   53
               Top             =   1920
               Visible         =   0   'False
               Width           =   300
            End
            Begin VB.CheckBox chkCadaSemana 
               Caption         =   "Cada semana"
               Height          =   195
               Left            =   240
               TabIndex        =   52
               Top             =   1920
               Width           =   1815
            End
            Begin VB.CommandButton botColor2Semana 
               Height          =   285
               Left            =   2925
               Picture         =   "frmSisConfig.frx":144C
               Style           =   1  'Graphical
               TabIndex        =   51
               Top             =   2640
               Visible         =   0   'False
               Width           =   300
            End
            Begin VB.Label lblVentanaColorMes 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   285
               Left            =   2160
               TabIndex        =   65
               Top             =   480
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.Label lblColorMes 
               AutoSize        =   -1  'True
               Caption         =   "Color mes"
               Height          =   195
               Left            =   2160
               TabIndex        =   64
               Top             =   240
               Visible         =   0   'False
               Width           =   690
            End
            Begin VB.Label lblVentanaColorA�o 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   285
               Left            =   2160
               TabIndex        =   63
               Top             =   1200
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.Label lblColora�o 
               AutoSize        =   -1  'True
               Caption         =   "Color a�o"
               Height          =   195
               Left            =   2160
               TabIndex        =   62
               Top             =   960
               Visible         =   0   'False
               Width           =   675
            End
            Begin VB.Label lblVentanaColor1Semana 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   285
               Left            =   2160
               TabIndex        =   61
               Top             =   1920
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.Label lbl1Semana 
               AutoSize        =   -1  'True
               Caption         =   "Color 1era. semana"
               Height          =   240
               Left            =   2160
               TabIndex        =   60
               Top             =   1680
               Visible         =   0   'False
               Width           =   1755
            End
            Begin VB.Label lblVentanaColor2Semana 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   285
               Left            =   2160
               TabIndex        =   59
               Top             =   2640
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.Label lbl2Semana 
               AutoSize        =   -1  'True
               Caption         =   "Color 2da. semana"
               Height          =   195
               Left            =   2160
               TabIndex        =   58
               Top             =   2400
               Visible         =   0   'False
               Width           =   1320
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Mostrar"
            Height          =   1575
            Left            =   0
            TabIndex        =   47
            Top             =   2640
            Width           =   2535
            Begin VB.CheckBox chkLineasDivisorias 
               Caption         =   "L�neas divisorias"
               Height          =   255
               Left            =   120
               TabIndex        =   49
               Top             =   360
               Width           =   1935
            End
            Begin VB.CheckBox chkIndicadorPpioFin 
               Caption         =   "Indicador de principio y fin "
               Height          =   495
               Left            =   120
               TabIndex        =   48
               Top             =   720
               Width           =   2055
            End
         End
      End
      Begin VB.Frame Frame16 
         Caption         =   "Configuraci�n de listados del sistema"
         Height          =   4935
         Left            =   -74880
         TabIndex        =   41
         Top             =   720
         Width           =   6855
         Begin MSFlexGridLib.MSFlexGrid gListados 
            Height          =   2655
            Left            =   240
            TabIndex        =   1
            Top             =   600
            Width           =   6375
            _ExtentX        =   11245
            _ExtentY        =   4683
            _Version        =   393216
            SelectionMode   =   1
         End
         Begin VB.CheckBox clickMostrarVistaPrevia 
            Caption         =   "Mostrar vista previa"
            Height          =   255
            Left            =   240
            TabIndex        =   45
            Top             =   4515
            Width           =   2295
         End
         Begin VB.CheckBox clickSeleccionarImpre 
            Caption         =   "Seleccionar impresora"
            Height          =   255
            Left            =   240
            TabIndex        =   44
            Top             =   4200
            Width           =   2415
         End
         Begin VB.CheckBox clickMostrarConfirmacion 
            Caption         =   "Mostrar mensaje confirmaci�n"
            Height          =   255
            Left            =   3240
            TabIndex        =   43
            Top             =   4200
            Width           =   3015
         End
         Begin VB.CheckBox clickImprimirLogo 
            Caption         =   "Imprimir logo en encabezado"
            Height          =   255
            Left            =   3240
            TabIndex        =   42
            Top             =   4515
            Visible         =   0   'False
            Width           =   3015
         End
         Begin VB.ComboBox cboImpresorasSis 
            Height          =   360
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   3600
            Width           =   3855
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "I&mpresora predeterminada del listado"
            Height          =   240
            Left            =   240
            TabIndex        =   2
            Top             =   3360
            Width           =   3375
         End
         Begin VB.Label Label21 
            Caption         =   "L&istados existentes"
            Height          =   255
            Left            =   240
            TabIndex        =   0
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame Frame10 
         Height          =   4935
         Left            =   -74880
         TabIndex        =   26
         Top             =   720
         Width           =   6855
         Begin VB.Frame Frame19 
            Caption         =   "S�mbolos de monedas "
            Height          =   975
            Left            =   240
            TabIndex        =   120
            Top             =   360
            Width           =   4935
            Begin VB.TextBox txtSimboloMonedaNacional 
               Height          =   360
               Left            =   2280
               MaxLength       =   3
               TabIndex        =   123
               Top             =   367
               Width           =   495
            End
            Begin VB.TextBox txtSimboloDolares 
               Height          =   360
               Left            =   240
               MaxLength       =   3
               TabIndex        =   121
               Top             =   360
               Width           =   495
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               Caption         =   "Moneda nacional"
               Height          =   240
               Left            =   3000
               TabIndex        =   124
               Top             =   420
               Width           =   1560
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               Caption         =   "D�lares"
               Height          =   240
               Left            =   840
               TabIndex        =   122
               Top             =   420
               Width           =   720
            End
         End
         Begin VB.CheckBox chkImprimirReserva 
            Caption         =   "&Imprimir reservas al realizar una nueva, modicar o anular."
            Height          =   255
            Left            =   240
            TabIndex        =   119
            Top             =   1680
            Width           =   5415
         End
         Begin VB.CheckBox chkMenuFijo 
            Caption         =   "M&ostrar men� de opciones sin movimiento."
            Height          =   255
            Left            =   240
            TabIndex        =   40
            Top             =   2160
            Width           =   4335
         End
      End
      Begin VB.Frame Frame9 
         Height          =   4935
         Left            =   -74880
         TabIndex        =   16
         Top             =   720
         Width           =   6855
         Begin VB.ComboBox cboFuentesEtiquetas 
            Height          =   360
            ItemData        =   "frmSisConfig.frx":16F6
            Left            =   240
            List            =   "frmSisConfig.frx":1700
            Style           =   2  'Dropdown List
            TabIndex        =   39
            Top             =   3480
            Width           =   2055
         End
         Begin VB.ComboBox cboTipoFuente 
            Height          =   315
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   1320
            Width           =   3495
         End
         Begin VB.ComboBox cboTama�oFuente 
            Height          =   315
            ItemData        =   "frmSisConfig.frx":1717
            Left            =   3960
            List            =   "frmSisConfig.frx":1724
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   1320
            Width           =   735
         End
         Begin VB.CommandButton botMuestra 
            Caption         =   "Muestra"
            Height          =   375
            Left            =   4560
            TabIndex        =   18
            Top             =   2160
            Width           =   1215
         End
         Begin VB.ComboBox cboElemento 
            Height          =   315
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   600
            Width           =   3495
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Tama�o de los fuentes de las etiquetas"
            Height          =   240
            Left            =   240
            TabIndex        =   38
            Top             =   3120
            Width           =   3525
         End
         Begin VB.Line Line1 
            X1              =   120
            X2              =   6720
            Y1              =   2880
            Y2              =   2880
         End
         Begin VB.Label Label1 
            Caption         =   "Fuente"
            Height          =   255
            Left            =   240
            TabIndex        =   25
            Top             =   1080
            Width           =   615
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Tama�o"
            Height          =   240
            Left            =   3960
            TabIndex        =   24
            Top             =   1080
            Width           =   765
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Muestra"
            Height          =   240
            Left            =   240
            TabIndex        =   23
            Top             =   1920
            Width           =   720
         End
         Begin VB.Label lblMuestra 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Esto es un texto de pruebas para muestra."
            Height          =   375
            Left            =   240
            TabIndex        =   22
            Top             =   2160
            Width           =   3495
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Elementos"
            Height          =   240
            Left            =   240
            TabIndex        =   21
            Top             =   360
            Width           =   960
         End
      End
      Begin VB.Frame Frame8 
         Height          =   4935
         Left            =   -74880
         TabIndex        =   10
         Top             =   720
         Width           =   6855
         Begin VB.TextBox txtDescEle 
            Height          =   1815
            Left            =   240
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   37
            Text            =   "frmSisConfig.frx":1733
            Top             =   2760
            Width           =   5655
         End
         Begin VB.CommandButton botColor 
            Height          =   285
            Left            =   5595
            Picture         =   "frmSisConfig.frx":1739
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   600
            Width           =   300
         End
         Begin VB.ListBox lstElementos 
            Height          =   1740
            Left            =   240
            TabIndex        =   11
            Top             =   600
            Width           =   4335
         End
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   5400
            Top             =   1080
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   327680
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Descripci�n del elemento seleccionado."
            Height          =   240
            Left            =   240
            TabIndex        =   36
            Top             =   2520
            Width           =   3615
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Color"
            Height          =   195
            Left            =   4800
            TabIndex        =   15
            Top             =   360
            Width           =   360
         End
         Begin VB.Label Label3 
            Caption         =   "Elementos"
            Height          =   255
            Left            =   240
            TabIndex        =   13
            Top             =   360
            Width           =   855
         End
         Begin VB.Label lblColor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   4800
            TabIndex        =   12
            Top             =   600
            Width           =   735
         End
      End
      Begin VB.Frame Frame2 
         Height          =   4935
         Left            =   -74880
         TabIndex        =   9
         Top             =   720
         Width           =   6855
         Begin VB.ListBox lstOperaciones 
            Height          =   4140
            Left            =   120
            TabIndex        =   31
            Top             =   480
            Width           =   2295
         End
         Begin VB.CommandButton botAgrego 
            Caption         =   ">"
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
            Left            =   2640
            TabIndex        =   30
            Top             =   720
            Width           =   855
         End
         Begin VB.CommandButton botSaco 
            Caption         =   "<"
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
            Left            =   2640
            TabIndex        =   29
            Top             =   1440
            Width           =   855
         End
         Begin VB.ListBox lstOperacionesAccesoDirecto 
            Height          =   1500
            Left            =   3720
            TabIndex        =   28
            Top             =   480
            Width           =   3015
         End
         Begin VB.TextBox txtDescOpciones 
            BackColor       =   &H80000016&
            Height          =   2055
            Left            =   2640
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   27
            Top             =   2580
            Width           =   4095
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "(M�ximo 10 )"
            Height          =   240
            Left            =   3720
            TabIndex        =   35
            Top             =   2040
            Width           =   1125
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Descripci�n de la opci�n"
            Height          =   195
            Left            =   2640
            TabIndex        =   34
            Top             =   2330
            Width           =   1755
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Opciones existentes"
            Height          =   195
            Left            =   120
            TabIndex        =   33
            Top             =   240
            Width           =   1425
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Opciones seleccionadas"
            Height          =   240
            Left            =   3720
            TabIndex        =   32
            Top             =   240
            Width           =   2250
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   4335
         Left            =   -74760
         TabIndex        =   8
         Top             =   720
         Width           =   5895
      End
   End
   Begin VB.CommandButton botAplicar 
      Caption         =   "&Aplicar"
      Height          =   375
      Left            =   6120
      TabIndex        =   6
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton botAceptar 
      Height          =   375
      Left            =   3480
      Picture         =   "frmSisConfig.frx":19E3
      Style           =   1  'Graphical
      TabIndex        =   5
      Tag             =   "Aceptar"
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton botCancelar 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   4800
      Picture         =   "frmSisConfig.frx":2299
      Style           =   1  'Graphical
      TabIndex        =   4
      Tag             =   "Cancelar"
      Top             =   6000
      Width           =   1215
   End
End
Attribute VB_Name = "frmSisConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private color As OLE_COLOR
Private primera_vez_fuentes As Boolean
Private primera_vez_accesos As Boolean
Private primera_vez_cuadroHab As Boolean
Private primera_vez_cuadroDis As Boolean
Private primera_vez_colores As Boolean
Private primera_vez_listados As Boolean
Private primera_vez_facturacion As Boolean

'Constantes que determinan valores m�nimo de configuraci�n de grillas
'de cuadro de habitaciones
Private Const LargoMaximoCelda As Integer = 40 'caracteres
Private Const LargoMinimoCelda As Integer = 3   'caracteres
Private Const AnchoMinimoCelda As Integer = 200    'pixeles
Private Const CantDiasMinimo As Integer = 1

'Constante que determina valor m�nimo de v�as en cuadro de configuraci�n de facturas
Private Const cantMinimoVias As Byte = 1

Private Sub Form_Load()
    primera_vez_colores = True
    primera_vez_fuentes = True
    primera_vez_accesos = True
    primera_vez_cuadroHab = True
    primera_vez_cuadroDis = True
    primera_vez_listados = True
    primera_vez_facturacion = True
    SSTab1_Click (0)
End Sub

Private Sub botAceptar_Click()
    botAplicar_Click
    Unload Me
End Sub

Private Sub botAplicar_Click()
    Select Case Me.SSTab1.TabCaption(SSTab1.Tab)
        Case "&Colores"  'colores
            'Grabo color del elemento actual seleccionado
            grabo_color _
                    lstElementos.ItemData(lstElementos.ListIndex), _
                    color
            mSub_Inicialixo_colores_sistema
        Case "&Fuentes"  'fuentes
            grabo_fuente _
                    cboElemento.ItemData(cboElemento.ListIndex), _
                    cboTipoFuente.Text, Val(cboTama�oFuente.Text)
            mSub_Inicializo_fuentes_sistema
        Case "&General"  'general
            grabo_general
            
        Case "Acc&esos directos"
            'no es necesario confirmar nada ya que las actualizaciones
            'en la base de datos se realizan en el momento que se trabaja en la ficha
            
        Case "C&uadro de habitaciones"   'cuadro de habitaciones
            subGraboCuadroDeHabitaciones
            
        Case "Cua&dro de disponibilidad" 'cuadro de disponibilidad
            subGraboCuadroDeDisponibilidad
            
        Case "&Listados"                'Configuraci�n de listados
            subGraboListados
            
        Case "F&acturaci�n"
            subGraboFacturacion
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Descargo formulario de memoria
    Set frmSisConfig = Nothing
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    '-----------------------------------------------------------------------
    'Opciones a realizar cada vez que doy clik en una ficha
    '-----------------------------------------------------------------------
    'Es necesario bloquer todos los frames que no estan activos para que el
    'cursor no se ubique en controles no visibles.
    Frame2.Enabled = False  'accesos
    Frame3.Enabled = False  'cuadro hab
    Frame8.Enabled = False  'colores
    Frame9.Enabled = False  'fuentes
    Frame10.Enabled = False 'general
    Frame11.Enabled = False 'cuadro disponibilidad
    Frame16.Enabled = False 'listados
    Frame20.Enabled = False 'facturaci�n
    
    If SSTab1.TabCaption(SSTab1.Tab) = "Habitaci�n" Then
        'permito trabajar con frame
        Frame8.Enabled = True  'colores
        'esta ficha utiliza este boton
        botAplicar.Enabled = True
        If primera_vez_colores Then
            primera_vez_colores = False
            cargo_elementos
            'muestro primer elemento
            If Me.lstElementos.ListCount > 0 Then
                Me.lstElementos.ListIndex = 0
            End If
            mSubBloqueoControlFormulario Me.txtDescEle, True
        End If
    End If
    
    If SSTab1.TabCaption(SSTab1.Tab) = "&Fuentes" Then
        'permito trabajar con fuentes
        Frame9.Enabled = True  'fuentes
        'esta ficha utiliza este boton
        botAplicar.Enabled = True
        'si prmera vez que doy click en el tab fuentes
        If primera_vez_fuentes Then
            'las proxima vez que de click en el tab de fechas no
            'ejecuto procedimientos de inicializaci�n
            primera_vez_fuentes = False
            cargo_elementos_fuentes
            mSubCargoCombosFuentes Me.cboTipoFuente
            If cboElemento.ListCount > 0 Then
                cboElemento.ListIndex = 0
            End If
        End If
    End If
    
    If SSTab1.TabCaption(SSTab1.Tab) = "&General" Then
        'permito trabajar con frame
        Frame10.Enabled = True 'general
        'esta ficha si utiliza este boton
        botAplicar.Enabled = True
        Me.chkMenuFijo.Value = tbPARAMETROS("tipoMenu")
        Me.chkImprimirReserva.Value = tbPARAMETROS("imprimir_Reserva")
        If Not IsNull(tbPARAMETROS("simboloDolares")) Then Me.txtSimboloDolares.Text = tbPARAMETROS("simboloDolares")
        If Not IsNull(tbPARAMETROS("simboloMonedaNacional")) Then Me.txtSimboloMonedaNacional.Text = tbPARAMETROS("simboloMonedaNacional")
        'inicializo variables globales de signo de moneda
        gblSignoMonedaNacional = mFunObtengoSignoMoneda(0)
        gblSignoDolares = gblSignoMonedaNacional = mFunObtengoSignoMoneda(1)
    End If
    
    If SSTab1.TabCaption(SSTab1.Tab) = "Acc&esos directos" Then
        'permito trabajar con frame
        Frame2.Enabled = True  'accesos
        'esta ficha no utiliza este boton
        botAplicar.Enabled = False
        'ejecuto solo la primera vez que doy click en la ficha
        If primera_vez_accesos Then
            primera_vez_accesos = False
            subCargoListaOperaciones
            subCargoListaOperacionesAccesoDirecto
        End If
    End If
    If SSTab1.TabCaption(SSTab1.Tab) = "C&uadro de habitaciones" Then
        'permito trabajar con frame
        Frame3.Enabled = True  'cuadro hab
        'esta ficha si utiliza este boton
        botAplicar.Enabled = True
        'ejecuto solo la primera vez que doy click en la ficha
        If primera_vez_cuadroHab Then
            primera_vez_cuadroHab = False
            subCargoDatosCuadroHab
        End If
    End If
    
    If SSTab1.TabCaption(SSTab1.Tab) = "Cua&dro de disponibilidad" Then
        'permito trabajar con frame
        Frame11.Enabled = True 'cuadro disponibilidad
        'esta ficha si utiliza este boton
        botAplicar.Enabled = True
        'ejecuto solo la primera vez que doy click en la ficha
        If primera_vez_cuadroDis Then
            primera_vez_cuadroDis = False
            subCargoDatosCuadroDisponibilidad
        End If
    End If
    
    If SSTab1.TabCaption(SSTab1.Tab) = "&Listados" Then
        'permito trabajar con frame
        Frame16.Enabled = True 'listados
        'esta ficha si utiliza este boton
        botAplicar.Enabled = True
        'ejecuto solo la primera vez que doy click en la ficha
        If primera_vez_listados Then
            primera_vez_listados = False
            subCargoListados
            mSubCargoImpresorasInstaladas Me.cboImpresorasSis
        End If
    End If
    
    If SSTab1.TabCaption(SSTab1.Tab) = "F&acturaci�n" Then
        'permito trabajar con frame
        Frame20.Enabled = True 'facturaci�n
        'esta ficha si utiliza este boton
        botAplicar.Enabled = True
        'ejecuto solo la primera vez que doy click en la ficha
        If primera_vez_facturacion Then
            primera_vez_facturacion = False
            'cargo combo de nacionalidades
            carga_tipo_nacionalidad Me.cboNacionalidades
            'cargo combos de tipo de impuestos
            carga_tipoIVA Me.cboImpAloja
            carga_tipoIVA Me.cboImpAlojaExtranjero
            subCargoFacturacion
        End If
    End If
    
End Sub

Private Sub botCancelar_Click()
    Unload Me
End Sub

'************************************************************************
'*
'*      C�digo para colores
'*
'************************************************************************
Private Sub botColor_Click()
    ' Establece Cancel a True.
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    ' Establece la propiedad Flags.
    CommonDialog1.Flags = cdlCCRGBInit
    ' Presenta el cuadro de di�logo Color.
    CommonDialog1.ShowColor
    ' Establece el color de fondo del formulario al
    ' color seleccionado.
    color = CommonDialog1.color
    muestro_color_etiqueta color
    Exit Sub

ErrHandler:
    ' El usuario hizo clic en el bot�n Cancelar.
    Exit Sub
End Sub

Private Sub cargo_elementos()
    'Recorro archivo de colores y cargo los elementos configurables del sistema.
    'Puede haber elementos que el usuario no pueda seleccionarlos, como por ejemplo el
    'fondo de los controles que estan bloqueados.
    'Este tipo de elementos no se cargan.
    tbSIS_COLORES.MoveFirst
    Do While Not tbSIS_COLORES.EOF
        'verifico que se a un elemento configurable por el usuario.
        If tbSIS_COLORES("muestroAUsuario") = True Then
            'agrego elemento a la lista que aparece en el formulario.
            lstElementos.AddItem tbSIS_COLORES("descapa")
            lstElementos.ItemData(lstElementos.NewIndex) = tbSIS_COLORES("codapa")
        End If
        tbSIS_COLORES.MoveNext
    Loop
End Sub

Private Sub lstElementos_Click()
    Dim color As OLE_COLOR
    Dim elemento As Integer
    'Cada vez que cambio de elemento muestro la descripci�n del elemento seleccioado
    'y el color actual configurado.
    elemento = lstElementos.ItemData(lstElementos.ListIndex)
    color = obtengo_color_elemento(elemento)
    'muestro color actual configurado
    muestro_color_etiqueta color
    'muestro descripci�n larga del elemento
    subMuestroDescLarga elemento
End Sub

Private Sub muestro_color_etiqueta(color As OLE_COLOR)
    lblColor.BackColor = color
End Sub

Private Sub subMuestroDescLarga(elemento As Integer)
    'Muestra la descripci�n del elemento seleccionado, para que el usuario tenga
    'una referencia de con que elemento est� trabajando.
    If busco_SisColorTF(elemento) Then
        Me.txtDescEle.Text = tbSIS_COLORES("descLargaApa")
    End If
End Sub

Private Function obtengo_color_elemento(elemento As Integer)
    'Obtengo el color del elemento seleccionado
    obtengo_color_elemento = 0
    If busco_SisColorTF(elemento) Then
        obtengo_color_elemento = tbSIS_COLORES("colorapa")
    End If
End Function

Private Sub grabo_color(elemento As Integer, color As OLE_COLOR)
    'Grabo color
    If busco_SisColorTF(elemento) Then
        tbSIS_COLORES.Edit
            tbSIS_COLORES("colorapa") = color
        tbSIS_COLORES.Update
    End If
End Sub

Private Function busco_SisColorTF(elemento As Integer)
    'Busca un elemento de configuraci�n en el archivo SISTEMA_COLORES
    busco_SisColorTF = False
    tbSIS_COLORES.Index = "pk_colores"
    tbSIS_COLORES.Seek "=", elemento
    If Not tbSIS_COLORES.NoMatch Then
        busco_SisColorTF = True
    End If
End Function

'************************************************************************
'*
'*      C�digo para fuentes
'*
'************************************************************************
Private Sub cargo_elementos_fuentes()
    'Recorro archivo de fuentes y cargo todos los elementos configurables del sistema
    tbSIS_FUENTES.MoveFirst
    Do While Not tbSIS_FUENTES.EOF
        'agrego elemento a la lista que aparece en el formulario
        cboElemento.AddItem tbSIS_FUENTES("DescApaFuente")
        cboElemento.ItemData(cboElemento.NewIndex) = tbSIS_FUENTES("CodApaFuente")
        tbSIS_FUENTES.MoveNext
    Loop
End Sub

Private Sub cboElemento_Click()
    On Error GoTo error
    'Cada vez que cambia un elemento, muestro el tipo de letra y tama�o
    'que tiene  configurado. Tambien Cambio la etiqueta de muestra
    
    'Obtengo elemento seleccionado y busco fuente para ese elemento
    If busco_SisFuenteTF(cboElemento.ItemData(cboElemento.ListIndex)) Then
        cboTipoFuente.Text = tbSIS_FUENTES("TipoApaFuente")
        cboTama�oFuente.Text = tbSIS_FUENTES("TamApaFuente")
        botMuestra_Click
    End If
    Exit Sub
error:
    MsgBox "No existe en el sistema el tipo de letra (fuente) " & _
    tbSIS_FUENTES("TipoApaFuente")
End Sub
    
Private Function busco_SisFuenteTF(elemento As Integer)
    'Busca un elemento de configuraci�n en el archivo SISTEMA_FUENTES
    busco_SisFuenteTF = False
    tbSIS_FUENTES.Index = "pk_sistema_fuentes"
    tbSIS_FUENTES.Seek "=", elemento
    If Not tbSIS_FUENTES.NoMatch Then
        busco_SisFuenteTF = True
    End If
End Function

Private Sub botMuestra_Click()
    'Muestra en la etiqueta de muestra, el tipo de letra seleccionado
    'con el tama�o tambi�n seleccionado
    lblMuestra.Font.Name = cboTipoFuente.Text
    lblMuestra.Font.Size = Val(cboTama�oFuente.Text)
End Sub

Private Sub grabo_fuente(elemento As Integer, fuente As String, tam As Byte)
    'Grabo color
    If busco_SisFuenteTF(elemento) Then
        tbSIS_FUENTES.Edit
            tbSIS_FUENTES("TipoApafuente") = fuente
            tbSIS_FUENTES("TamApafuente") = tam
        tbSIS_FUENTES.Update
    End If
End Sub


'************************************************************************
'*
'*      C�digo para general
'*
'*
'************************************************************************

Private Sub grabo_general()
    tbPARAMETROS.Edit
        tbPARAMETROS("tipomenu") = chkMenuFijo.Value
        tbPARAMETROS("imprimir_reserva") = chkImprimirReserva.Value
        tbPARAMETROS("simboloMonedaNacional") = Trim(txtSimboloMonedaNacional.Text)
        tbPARAMETROS("simboloDolares") = Trim(txtSimboloDolares.Text)
    tbPARAMETROS.Update
End Sub


'************************************************************************
'*
'*      C�digo para Acceso Directos
'*
'*
'************************************************************************

Private Sub subCargoListaOperaciones()
    'Recorro archvio SISTEMA_OPERACIONES y cargo en lista de operaciones
    lstOperaciones.Clear
    tbSISTEMA_OPERACIONES.MoveFirst
    Do While Not tbSISTEMA_OPERACIONES.EOF
        'no trabajo con las ya seleccionadas
        If Not tbSISTEMA_OPERACIONES("UsadaParaAccesoDirecto") Then
            'no trabajo con las opciones generales
            If tbSISTEMA_OPERACIONES("TipoOpr") = 2 Then
                'agrego elemento a la lista de operaciones
                lstOperaciones.AddItem tbSISTEMA_OPERACIONES("DescOpr")
                lstOperaciones.ItemData(lstOperaciones.NewIndex) = tbSISTEMA_OPERACIONES("CodOpr")
            End If
        End If
        tbSISTEMA_OPERACIONES.MoveNext
    Loop
    msubPosicionoListasAlPrincipio Me.lstOperaciones
End Sub

Private Sub subCargoListaOperacionesAccesoDirecto()
    'Recorro archivo SISTEMA_OPERACIONES y cargo las listas
    'de tipo acceso directo
    lstOperacionesAccesoDirecto.Clear
    tbSISTEMA_OPERACIONES.MoveFirst
    Do While Not tbSISTEMA_OPERACIONES.EOF
        'trabajo solo con las ya seleccionadas
        If tbSISTEMA_OPERACIONES("UsadaParaAccesoDirecto") Then
            'agrego elemento a la lista de operaciones
            lstOperacionesAccesoDirecto.AddItem tbSISTEMA_OPERACIONES("DescAccesoDirecto")
            lstOperacionesAccesoDirecto.ItemData(lstOperacionesAccesoDirecto.NewIndex) = tbSISTEMA_OPERACIONES("CodOpr")
        End If
        tbSISTEMA_OPERACIONES.MoveNext
    Loop
    msubPosicionoListasAlPrincipio lstOperacionesAccesoDirecto
End Sub

Private Sub lstOperaciones_Click()
    'Cada vez que me pocisiono sobre una opci�n muestro descripci�n
    If busco_operacion(lstOperaciones.ItemData(lstOperaciones.ListIndex)) Then
        Me.txtDescOpciones.Text = tbSISTEMA_OPERACIONES("InfOpr")
    End If
End Sub

Private Sub botAgrego_Click()
    'Creo un nuevo acceso directo
    
    'Si no hay ningun elemento seleccionado no hago nada
    If lstOperaciones.ListIndex >= 0 Then
        subMarcoOperaciones lstOperaciones.ItemData(lstOperaciones.ListIndex), True
        'Actualizo listas
        subCargoListaOperaciones
        subCargoListaOperacionesAccesoDirecto
    End If
End Sub

Private Sub botSaco_Click()
    'Elimino un acceso directo
    
    'Si no hay ningun elemento seleccionado no hago nada
    If lstOperacionesAccesoDirecto.ListIndex >= 0 Then
        subMarcoOperaciones lstOperacionesAccesoDirecto.ItemData(lstOperacionesAccesoDirecto.ListIndex), False
        'Actualizo listas
        subCargoListaOperaciones
        subCargoListaOperacionesAccesoDirecto
    End If
End Sub

Private Sub subMarcoOperaciones(Opr As Integer, marca As Boolean)
    'Crea o elimina un acceso directo
    If busco_operacion(Opr) Then
        tbSISTEMA_OPERACIONES.Edit
            tbSISTEMA_OPERACIONES("UsadaParaAccesoDirecto") = marca
        tbSISTEMA_OPERACIONES.Update
    End If
End Sub

'************************************************************************
'*
'*      C�digo para cuadro de habitaciones
'*
'*
'************************************************************************

Private Sub subCargoDatosCuadroHab()
    'Leo archivo de configuraci�n y muestro datos actuales
    tbSISTEMA_CONF_FORMULARIOS.Index = "pk_CodFormulario"
    tbSISTEMA_CONF_FORMULARIOS.Seek ">=", 1, 1
    '1 = c�digo del formulario ed cuadro de habitaciones
    ',1 = primera variable de configuraci�n
    Do While Not tbSISTEMA_CONF_FORMULARIOS.EOF
        If tbSISTEMA_CONF_FORMULARIOS("CodFormulario") = 1 Then
            'mientras recorro varibles del formulario de cuadro de habitaci�n: proceso
            Select Case tbSISTEMA_CONF_FORMULARIOS("codConfiguracion")
                Case 1  'ancho de celda
                    Me.txtAnchoCelda.Text = tbSISTEMA_CONF_FORMULARIOS("1ValorNumerico")
                Case 2  'largo de celda
                    Me.txtLargoCelda.Text = tbSISTEMA_CONF_FORMULARIOS("1ValorNumerico")
                Case 3  'mostrar habitaci�nes no asignadas
                    Me.cboMostrar.ListIndex = tbSISTEMA_CONF_FORMULARIOS("1ValorNumerico")
                Case 4  'mostrar l�neas divisorias
                    Me.chkLineasDivisorias.Value = tbSISTEMA_CONF_FORMULARIOS("1ValorBol")
                Case 5  'cant. d�as predeterminados
                    Me.txtCantDias.Text = tbSISTEMA_CONF_FORMULARIOS("1ValorNumerico")
                Case 6  'iluminar mes
                    Me.chkPpioMes.Value = tbSISTEMA_CONF_FORMULARIOS("1Valorbol")
                    subCargoColorVentaDesdeArchivo tbSISTEMA_CONF_FORMULARIOS("1ValorNumerico"), lblVentanaColorMes
                Case 7  'iluminar a�o
                    Me.chkPpioA�o.Value = tbSISTEMA_CONF_FORMULARIOS("1Valorbol")
                    subCargoColorVentaDesdeArchivo tbSISTEMA_CONF_FORMULARIOS("1ValorNumerico"), lblVentanaColorA�o
                Case 8  'iluminar semana
                    Me.chkCadaSemana.Value = tbSISTEMA_CONF_FORMULARIOS("1Valorbol")
                    subCargoColorVentaDesdeArchivo tbSISTEMA_CONF_FORMULARIOS("1ValorNumerico"), lblVentanaColor1Semana
                    subCargoColorVentaDesdeArchivo tbSISTEMA_CONF_FORMULARIOS("2ValorNumerico"), lblVentanaColor2Semana
                Case 9 'indicador de ppio y fin de lineas
                    Me.chkIndicadorPpioFin.Value = tbSISTEMA_CONF_FORMULARIOS("1Valorbol")
            End Select
            tbSISTEMA_CONF_FORMULARIOS.MoveNext
        Else
            Exit Do
        End If
    Loop
End Sub

Private Sub subGraboCuadroDeHabitaciones()
    'Una vez confirmado los cambios los grabo
    If mFunPosicionoParaGrabar(1, 1) Then   'ancho de celda
        tbSISTEMA_CONF_FORMULARIOS.Edit
            tbSISTEMA_CONF_FORMULARIOS("1ValorNumerico") = Me.txtAnchoCelda.Text
        tbSISTEMA_CONF_FORMULARIOS.Update
    End If
    If mFunPosicionoParaGrabar(1, 2) Then   'largo de celda
        tbSISTEMA_CONF_FORMULARIOS.Edit
            tbSISTEMA_CONF_FORMULARIOS("1ValorNumerico") = Me.txtLargoCelda.Text
        tbSISTEMA_CONF_FORMULARIOS.Update
    End If
    If mFunPosicionoParaGrabar(1, 3) Then   'mostrar habitaciones no asignadas
        tbSISTEMA_CONF_FORMULARIOS.Edit
            tbSISTEMA_CONF_FORMULARIOS("1ValorNumerico") = Me.cboMostrar.ListIndex
        tbSISTEMA_CONF_FORMULARIOS.Update
    End If
    If mFunPosicionoParaGrabar(1, 4) Then   'mostrar lineas divisorias
        tbSISTEMA_CONF_FORMULARIOS.Edit
            tbSISTEMA_CONF_FORMULARIOS("1Valorbol") = Me.chkLineasDivisorias.Value
        tbSISTEMA_CONF_FORMULARIOS.Update
    End If
    If mFunPosicionoParaGrabar(1, 5) Then   'cantidad de d�as
        tbSISTEMA_CONF_FORMULARIOS.Edit
            tbSISTEMA_CONF_FORMULARIOS("1ValorNumerico") = Me.txtCantDias.Text
        tbSISTEMA_CONF_FORMULARIOS.Update
    End If
    If mFunPosicionoParaGrabar(1, 6) Then   'iluminar mes
        tbSISTEMA_CONF_FORMULARIOS.Edit
            tbSISTEMA_CONF_FORMULARIOS("1ValorBol") = Me.chkPpioMes.Value
            tbSISTEMA_CONF_FORMULARIOS("1ValorNumerico") = Me.lblVentanaColorMes.BackColor
        tbSISTEMA_CONF_FORMULARIOS.Update
    End If
    If mFunPosicionoParaGrabar(1, 7) Then   'iluminar a�o
        tbSISTEMA_CONF_FORMULARIOS.Edit
            tbSISTEMA_CONF_FORMULARIOS("1ValorBol") = Me.chkPpioA�o.Value
            tbSISTEMA_CONF_FORMULARIOS("1ValorNumerico") = Me.lblVentanaColorA�o.BackColor
        tbSISTEMA_CONF_FORMULARIOS.Update
    End If
    If mFunPosicionoParaGrabar(1, 8) Then   'iluminar semanas
        tbSISTEMA_CONF_FORMULARIOS.Edit
            tbSISTEMA_CONF_FORMULARIOS("1ValorBol") = Me.chkCadaSemana.Value
            tbSISTEMA_CONF_FORMULARIOS("1ValorNumerico") = Me.lblVentanaColor1Semana.BackColor
            tbSISTEMA_CONF_FORMULARIOS("2ValorNumerico") = Me.lblVentanaColor2Semana.BackColor
        tbSISTEMA_CONF_FORMULARIOS.Update
    End If
    If mFunPosicionoParaGrabar(1, 9) Then   'indicador de ppio y fin de archivo
        tbSISTEMA_CONF_FORMULARIOS.Edit
            tbSISTEMA_CONF_FORMULARIOS("1ValorBol") = Me.chkIndicadorPpioFin.Value
        tbSISTEMA_CONF_FORMULARIOS.Update
    End If
End Sub

Private Sub chkCadaSemana_Click()
    'Muestro o no, la opci�n de cambio de color por semana
    botColor1Semana.Visible = chkCadaSemana.Value
    lbl1Semana.Visible = chkCadaSemana.Value
    lblVentanaColor1Semana.Visible = chkCadaSemana.Value
    botColor2Semana.Visible = chkCadaSemana.Value
    lbl2Semana.Visible = chkCadaSemana.Value
    lblVentanaColor2Semana.Visible = chkCadaSemana.Value
End Sub

Private Sub chkPpioA�o_Click()
    'Muestro o no, la opci�n de cambio de color a�o.
    botColorA�o.Visible = chkPpioA�o.Value
    lblColora�o.Visible = chkPpioA�o.Value
    lblVentanaColorA�o.Visible = chkPpioA�o.Value
End Sub

Private Sub chkPpioMes_Click()
    'Muestro o no, la opci�n de cambio de color mes.
    botColorMes.Visible = chkPpioMes.Value
    lblColorMes.Visible = chkPpioMes.Value
    lblVentanaColorMes.Visible = chkPpioMes.Value
End Sub

Private Sub botColorMes_Click()
    'Cambio color preestablecido para mes
    muestro_color_Ventana lblVentanaColorMes
End Sub

Private Sub botColor1Semana_Click()
    'Cambio color preestablecido para 1 semana
    muestro_color_Ventana lblVentanaColor1Semana
End Sub

Private Sub botColor2Semana_Click()
    'Cambio color preestablecido para 2 semana
    muestro_color_Ventana lblVentanaColor2Semana
End Sub

Private Sub botColorA�o_Click()
    'Cambio color preestablecido para a�o
    muestro_color_Ventana lblVentanaColorA�o
End Sub

Private Sub muestro_color_Ventana(ventana As Label)
    'Llamo al formulario de colores

    ' Establece Cancel a True.
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    ' Establece la propiedad Flags.
    CommonDialog1.Flags = cdlCCRGBInit
    ' Presenta el cuadro de di�logo Color.
    CommonDialog1.ShowColor
    ' Establece el color de fondo del formulario al
    ' color seleccionado.
    color = CommonDialog1.color
    ventana.BackColor = color
    Exit Sub

ErrHandler:
    ' El usuario hizo clic en el bot�n Cancelar.
    Exit Sub
End Sub

Private Sub subCargoColorVentaDesdeArchivo(color As OLE_COLOR, ventana As Label)
    ventana.BackColor = color
End Sub

Private Sub txtAnchoCelda_KeyPress(KeyAscii As Integer)
    'Solo permito ingreso de n�meros
    ValidoNum KeyAscii, False, False
End Sub

Private Sub txtCantDias_KeyPress(KeyAscii As Integer)
    'Solo permito ingreso de n�meros
    ValidoNum KeyAscii, False, False
End Sub

Private Sub txtAnchoCelda_LostFocus()
    'controlo tama�o m�nimo de celda
    If Val(txtAnchoCelda.Text) < AnchoMinimoCelda Then
        txtAnchoCelda.Text = AnchoMinimoCelda
    End If
End Sub

Private Sub txtCantDias_LostFocus()
    'controlo cantidad m�nima de d�as
    If Val(txtCantDias.Text) < CantDiasMinimo Then
        txtCantDias.Text = CantDiasMinimo
    End If
End Sub

Private Sub txtLargoCelda_KeyPress(KeyAscii As Integer)
    'Solo permito ingreso de n�meros
    ValidoNum KeyAscii, False, False
End Sub

Private Sub txtLargoCelda_LostFocus()
    'controlo tama�o m�nimo de celda
    If CInt(txtLargoCelda.Text) < LargoMinimoCelda Then
        txtLargoCelda.Text = LargoMinimoCelda
    End If
    'controlo tama�o m�ximo de celda
    If CInt(txtLargoCelda.Text) > LargoMaximoCelda Then
        txtLargoCelda.Text = LargoMaximoCelda
    End If
End Sub

'************************************************************************
'*
'*      C�digo para cuadro de disponibilidad
'*
'************************************************************************

Private Sub subCargoDatosCuadroDisponibilidad()
    'Leo archivo de configuraci�n y muestro datos actuales
    tbSISTEMA_CONF_FORMULARIOS.Index = "pk_CodFormulario"
    tbSISTEMA_CONF_FORMULARIOS.Seek ">=", 2, 1
    '2 = c�digo del formulario ed cuadro de disponibilidad
    ',1 = primera variable de configuraci�n
    Do While Not tbSISTEMA_CONF_FORMULARIOS.EOF
        If tbSISTEMA_CONF_FORMULARIOS("CodFormulario") = 2 Then
            'mientras recorro varibles del formulario de cuadro de disponibilidad: proceso
            Select Case tbSISTEMA_CONF_FORMULARIOS("codConfiguracion")
                Case 1  'ancho de celda
                    Me.txtAnchoCeldaDis.Text = tbSISTEMA_CONF_FORMULARIOS("1ValorNumerico")
                Case 2  'largo de celda
                    Me.txtLargoCeldaDis.Text = tbSISTEMA_CONF_FORMULARIOS("1ValorNumerico")
                Case 3  'cant. d�as predeterminados
                    Me.txtCantDiasDis.Text = tbSISTEMA_CONF_FORMULARIOS("1ValorNumerico")
                Case 4  'iluminar mes
                    Me.chkPpioMesDis.Value = tbSISTEMA_CONF_FORMULARIOS("1Valorbol")
                    subCargoColorVentaDesdeArchivo tbSISTEMA_CONF_FORMULARIOS("1ValorNumerico"), lblVentanaColorMesDis
                Case 5  'iluminar a�o
                    Me.chkPpioA�oDis.Value = tbSISTEMA_CONF_FORMULARIOS("1Valorbol")
                    subCargoColorVentaDesdeArchivo tbSISTEMA_CONF_FORMULARIOS("1ValorNumerico"), lblVentanaColorA�oDis
                Case 6  'iluminar semana
                    Me.chkSemanaDis.Value = tbSISTEMA_CONF_FORMULARIOS("1Valorbol")
                    subCargoColorVentaDesdeArchivo tbSISTEMA_CONF_FORMULARIOS("1ValorNumerico"), lblVentanaColor1SemanaDis
                    subCargoColorVentaDesdeArchivo tbSISTEMA_CONF_FORMULARIOS("2ValorNumerico"), lblVentanaColor2SemanaDis
                Case 7 'muestro icono ocupada
                    Me.chkMuestroIconoOcupada.Value = tbSISTEMA_CONF_FORMULARIOS("1Valorbol")
                Case 8  'tama�o fuente d�gitos
                    Me.cboTama�oDigitosDis.Text = tbSISTEMA_CONF_FORMULARIOS("1ValorNumerico")
                Case 9  'tama�o fuente caracteres
                    Me.cboTama�oLetrasDis.Text = tbSISTEMA_CONF_FORMULARIOS("1ValorNumerico")
                Case 10 'alineaci�n icono
                    Me.cboAlinIcono.ListIndex = tbSISTEMA_CONF_FORMULARIOS("1ValorNumerico")
                Case 11 'alineaci�n fuente
                    Me.cboAlinFuente.ListIndex = tbSISTEMA_CONF_FORMULARIOS("1ValorNumerico")
            End Select
            tbSISTEMA_CONF_FORMULARIOS.MoveNext
        Else
            Exit Do
        End If
    Loop
End Sub

Private Sub txtAnchoCeldaDis_KeyPress(KeyAscii As Integer)
    'Solo permito ingreso de n�meros
    ValidoNum KeyAscii, False, False
End Sub

Private Sub txtAnchoCeldaDis_LostFocus()
    'controlo tama�o m�nimo de celda
    If Val(txtAnchoCeldaDis.Text) < AnchoMinimoCelda Then
        txtAnchoCeldaDis.Text = AnchoMinimoCelda
    End If
End Sub

Private Sub txtLargoCeldaDis_KeyPress(KeyAscii As Integer)
    'Solo permito ingreso de n�meros
    ValidoNum KeyAscii, False, False
End Sub

Private Sub txtLargoCeldaDis_LostFocus()
        'controlo tama�o m�nimo de celda
    If Val(txtLargoCeldaDis.Text) < LargoMinimoCelda Then
        txtLargoCeldaDis.Text = LargoMinimoCelda
    End If
    'controlo tama�o m�ximo de celda
    If Val(txtLargoCeldaDis.Text) > LargoMaximoCelda Then
        txtLargoCeldaDis.Text = LargoMaximoCelda
    End If
End Sub

Private Sub botColor1SemanaDis_Click()
    'Cambio color preestablecido para primer semana
    muestro_color_Ventana lblVentanaColor1SemanaDis
End Sub

Private Sub botColor2SemanaDis_Click()
    'Cambio color preestablecido para segunda semana
    muestro_color_Ventana lblVentanaColor2SemanaDis
End Sub

Private Sub botColorA�oDis_Click()
    'Cambio color preestablecido para a�o
    muestro_color_Ventana lblVentanaColorA�oDis
End Sub

Private Sub botColorMesDis_Click()
    'Cambio color preestablecido para mes
    muestro_color_Ventana lblVentanaColorMesDis
End Sub

Private Sub txtCantDiasDis_KeyPress(KeyAscii As Integer)
    'Solo permito ingreso de n�meros
    ValidoNum KeyAscii, False, False
End Sub

Private Sub txtCantDiasDis_LostFocus()
    'controlo cantidad m�nima de d�as
    If Val(txtCantDiasDis.Text) < CantDiasMinimo Then
        txtCantDiasDis.Text = CantDiasMinimo
    End If
End Sub

Private Sub chkPpioA�oDis_Click()
    'Muestro o no, la opci�n de cambio de color a�o.
    botColorA�oDis.Visible = chkPpioA�oDis.Value
    lblColorA�oDis.Visible = chkPpioA�oDis.Value
    lblVentanaColorA�oDis.Visible = chkPpioA�oDis.Value
End Sub

Private Sub chkPpioMesDis_Click()
    'Muestro o no, la opci�n de cambio de color  mes
    botColorMesDis.Visible = chkPpioMesDis.Value
    lblColorMesDis.Visible = chkPpioMesDis.Value
    lblVentanaColorMesDis.Visible = chkPpioMesDis.Value
End Sub

Private Sub chkSemanaDis_Click()
    'Muestro o no, la opci�n de cambio de color semanal.
    botColor1SemanaDis.Visible = chkSemanaDis.Value
    lblColor1SemanaDis.Visible = chkSemanaDis.Value
    lblVentanaColor1SemanaDis.Visible = chkSemanaDis.Value
    
    botColor2SemanaDis.Visible = chkSemanaDis.Value
    lblColor2SemanaDis.Visible = chkSemanaDis.Value
    lblVentanaColor2SemanaDis.Visible = chkSemanaDis.Value
End Sub

Private Sub subGraboCuadroDeDisponibilidad()
    'Una vez confirmado los cambios los grabo
    If mFunPosicionoParaGrabar(2, 1) Then   'ancho de celda
        tbSISTEMA_CONF_FORMULARIOS.Edit
            tbSISTEMA_CONF_FORMULARIOS("1ValorNumerico") = Me.txtAnchoCeldaDis.Text
        tbSISTEMA_CONF_FORMULARIOS.Update
    End If
    If mFunPosicionoParaGrabar(2, 2) Then   'largo de celda
        tbSISTEMA_CONF_FORMULARIOS.Edit
            tbSISTEMA_CONF_FORMULARIOS("1ValorNumerico") = Me.txtLargoCeldaDis.Text
        tbSISTEMA_CONF_FORMULARIOS.Update
    End If
    If mFunPosicionoParaGrabar(2, 3) Then   'cantidad de d�as
        tbSISTEMA_CONF_FORMULARIOS.Edit
            tbSISTEMA_CONF_FORMULARIOS("1ValorNumerico") = Me.txtCantDiasDis.Text
        tbSISTEMA_CONF_FORMULARIOS.Update
    End If
    If mFunPosicionoParaGrabar(2, 4) Then   'iluminar mes
        tbSISTEMA_CONF_FORMULARIOS.Edit
            tbSISTEMA_CONF_FORMULARIOS("1ValorBol") = Me.chkPpioMesDis.Value
            tbSISTEMA_CONF_FORMULARIOS("1ValorNumerico") = Me.lblVentanaColorMesDis.BackColor
        tbSISTEMA_CONF_FORMULARIOS.Update
    End If
    If mFunPosicionoParaGrabar(2, 5) Then   'iluminar a�o
        tbSISTEMA_CONF_FORMULARIOS.Edit
            tbSISTEMA_CONF_FORMULARIOS("1ValorBol") = Me.chkPpioA�oDis.Value
            tbSISTEMA_CONF_FORMULARIOS("1ValorNumerico") = Me.lblVentanaColorA�oDis.BackColor
        tbSISTEMA_CONF_FORMULARIOS.Update
    End If
    If mFunPosicionoParaGrabar(2, 6) Then   'iluminar semanas
        tbSISTEMA_CONF_FORMULARIOS.Edit
            tbSISTEMA_CONF_FORMULARIOS("1ValorBol") = Me.chkSemanaDis.Value
            tbSISTEMA_CONF_FORMULARIOS("1ValorNumerico") = Me.lblVentanaColor1SemanaDis.BackColor
            tbSISTEMA_CONF_FORMULARIOS("2ValorNumerico") = Me.lblVentanaColor2SemanaDis.BackColor
        tbSISTEMA_CONF_FORMULARIOS.Update
    End If
    If mFunPosicionoParaGrabar(2, 7) Then   'muestro icono ocupada
        tbSISTEMA_CONF_FORMULARIOS.Edit
            tbSISTEMA_CONF_FORMULARIOS("1ValorBol") = Me.chkMuestroIconoOcupada.Value
        tbSISTEMA_CONF_FORMULARIOS.Update
    End If
    If mFunPosicionoParaGrabar(2, 8) Then   'tama�o fuente d�gitos
        tbSISTEMA_CONF_FORMULARIOS.Edit
            tbSISTEMA_CONF_FORMULARIOS("1ValorNumerico") = Val(Me.cboTama�oDigitosDis.Text)
        tbSISTEMA_CONF_FORMULARIOS.Update
    End If
    If mFunPosicionoParaGrabar(2, 9) Then   'tama�o fuente caracteres
        tbSISTEMA_CONF_FORMULARIOS.Edit
            tbSISTEMA_CONF_FORMULARIOS("1ValorNumerico") = Val(Me.cboTama�oLetrasDis.Text)
        tbSISTEMA_CONF_FORMULARIOS.Update
    End If
    If mFunPosicionoParaGrabar(2, 10) Then   'alineaci�n icono
        tbSISTEMA_CONF_FORMULARIOS.Edit
            tbSISTEMA_CONF_FORMULARIOS("1ValorNumerico") = Me.cboAlinIcono.ListIndex
        tbSISTEMA_CONF_FORMULARIOS.Update
    End If
    If mFunPosicionoParaGrabar(2, 11) Then   'alineaci�n fuente
        tbSISTEMA_CONF_FORMULARIOS.Edit
            tbSISTEMA_CONF_FORMULARIOS("1ValorNumerico") = Me.cboAlinFuente.ListIndex
        tbSISTEMA_CONF_FORMULARIOS.Update
    End If
End Sub

'***********************************************************************************
'*
'*      C�digo para configuraci�n de listados
'*
'************************************************************************************
'   NOTA: por limitaciones de CrystalReport no se implementa la posibilidad de incluir
'   o no la impresi�n del logo del hotel, en los cabezales de los reportes.
'************************************************************************************

Private Sub subCargoListados()
    '--------------------------------------------------------------------------
    'Carga en la grilla los listados existentes.
    '--------------------------------------------------------------------------
    'declaro variables locales para utilizar archivo tbSISTEMA_LISTADOS
    Dim tablaSisLis As Recordset
    Set tablaSisLis = tbSISTEMA_LISTADOS
    
    'configuro cabezal grilla
    gListados.FormatString = "|Nombre                                           " & _
                            "|Descripci�n                                                                    " & _
                            "|TipoLis|CodLis"
    'oculto columnas no visibles
    gListados.ColWidth(3) = 0
    gListados.ColWidth(4) = 0
        
    'recorro el archivo de listados y cargo en grilla
    tablaSisLis.Index = "pk_listados"
    tablaSisLis.Seek ">=", 0, 0
    If Not tablaSisLis.NoMatch Then
        Do While Not tablaSisLis.EOF
            'agrego linea al la grilla
            gListados.AddItem Chr(9) & tablaSisLis("nomLis") & _
                                Chr(9) & tablaSisLis("descLis") & _
                                Chr(9) & tablaSisLis("tipoLis") & _
                                Chr(9) & tablaSisLis("codLis")
            tablaSisLis.MoveNext
        Loop
    End If
End Sub

Private Sub gListados_EnterCell()
    '------------------------------------------------------------------------------------
    'Cada vez que se cambia la celda activa, se muestran los nuevos valores del listado
    'seleccioado.
    '------------------------------------------------------------------------------------
    On Error Resume Next
    
    'declaraci�n de variable para utilizar biblioteca impresion.dll
    Dim biblioImpresion As ImpresionGeneral
    Set biblioImpresion = New ImpresionGeneral
    
    'declaro variables locales para utilizar archivo tbSISTEMA_LISTADOS
    Dim tablaSisLis As Recordset
    Set tablaSisLis = tbSISTEMA_LISTADOS
    
    Dim tipoLis As Byte         'tipo del listado
    Dim codLis As Integer       'c�digo del listado
    Dim impSis As String
            
    'obtengo tipo y c�digo del listado
    tipoLis = CByte(gListados.TextMatrix(gListados.row, 3))
    codLis = CInt(gListados.TextMatrix(gListados.row, 4))
    
    'busco listado seleccionado en grilla
    tablaSisLis.Index = "pk_listados"
    tablaSisLis.Seek "=", tipoLis, codLis
    If Not tablaSisLis.NoMatch Then
        'existe listado
        'muestro informacion listados
        Me.clickSeleccionarImpre.Value = tablaSisLis("seleccionarImpLis")
        Me.clickImprimirLogo.Value = tablaSisLis("imprimirLogoLis")
        Me.clickMostrarConfirmacion.Value = tablaSisLis("mensajeConfLis")
        'verifico si esta opci�n es posible modificarla
        If tablaSisLis("tipoLis") <> 1 Then
            'permito mostrar vista previa
            Me.clickMostrarVistaPrevia.Value = tablaSisLis("mostrarVistaPrevia")
            Me.clickMostrarVistaPrevia.Enabled = True
        Else
            'es un listado de tipo documento, por lo que no puedo mostrar vista
            'previa ya que el documento tiene que ir directamente a la impresora.
            Me.clickMostrarVistaPrevia.Value = 0
            Me.clickMostrarVistaPrevia.Enabled = False
        End If
        
        'verifico si el listado tiene asignada una impresora
        If IsNull(tablaSisLis("impreLis")) Then
            impSis = ""
        Else
            impSis = tablaSisLis("impreLis")
        End If
        
        'verifico si la impresora es una impresora del sistema
        If biblioImpresion.mFunExisteImpresoraInstalada(impSis) Then
            'la impresora del listado esta instalada en el sistema
            'muestro impresora en combo
            Me.cboImpresorasSis.Text = tablaSisLis("impreLis")
        Else
            'la impresora no esta instalda
            'muestro entonces la impresora del sistema
            Me.cboImpresorasSis.Text = Printer.DeviceName
        End If
    End If
    
    Set tablaSisLis = Nothing
    Set biblioImpresion = Nothing
End Sub

Private Sub subGraboListados()
    '---------------------------------------------------------------------
    'Grabo la configuraci�n del listado actualmente seleccionado
    '---------------------------------------------------------------------
    On Error Resume Next
    
    'declaro variables locales para utilizar archivo tbSISTEMA_LISTADOS
    Dim tablaSisLis As Recordset
    Set tablaSisLis = tbSISTEMA_LISTADOS
    
    Dim tipoLis As Byte         'tipo del listado
    Dim codLis As Integer       'c�digo del listado
    
    'obtengo tipo y c�digo del listado
    tipoLis = CByte(gListados.TextMatrix(gListados.row, 3))
    codLis = CInt(gListados.TextMatrix(gListados.row, 4))
    
    'busco listado seleccionado en grilla
    tablaSisLis.Index = "pk_listados"
    tablaSisLis.Seek "=", tipoLis, codLis
    If Not tablaSisLis.NoMatch Then
        'existe listado
        'grabo datos
        tablaSisLis.Edit
            tablaSisLis("seleccionarImpLis") = Me.clickSeleccionarImpre.Value
            tablaSisLis("imprimirLogoLis") = Me.clickImprimirLogo.Value
            tablaSisLis("mensajeConfLis") = Me.clickMostrarConfirmacion.Value
            tablaSisLis("mostrarVistaPrevia") = Me.clickMostrarVistaPrevia.Value
            tablaSisLis("impreLis") = Me.cboImpresorasSis.Text
        tablaSisLis.Update
    End If
    Set tablaSisLis = Nothing
End Sub

'****************************************************************
'*
'*  C�digo para Facturaci�n
'*
'*****************************************************************

Private Sub subGraboFacturacion()
    '----------------------------------------------------------------
    'Grabo propiedades de Facturaci�n.
    '----------------------------------------------------------------
    tbPARAMETROS.Edit
        tbPARAMETROS("SisMostrarTotalesResumidos") = Me.chkMostarTotales.Value
        tbPARAMETROS("factCantViasImpresas") = Me.txtCantViasDocu.Text
        tbPARAMETROS("factDiferenciarImpAlojaExt") = Me.chkImpAlojaExtranjeros.Value
        tbPARAMETROS("factNacionalidadLocal") = Me.cboNacionalidades.ItemData(Me.cboNacionalidades.ListIndex)
        tbPARAMETROS("factTipoImpAlojaExt") = Me.cboImpAlojaExtranjero.ItemData(Me.cboImpAlojaExtranjero.ListIndex)
        tbPARAMETROS("TipoIvaAloja") = Me.cboImpAloja.ItemData(Me.cboImpAloja.ListIndex)
    tbPARAMETROS.Update
End Sub

Private Sub subCargoFacturacion()
    '----------------------------------------------------------------
    'Obtengo propiedades de Facturaci�n
    '----------------------------------------------------------------
    chkMostarTotales.Value = tbPARAMETROS("SisMostrarTotalesResumidos")
    Me.txtCantViasDocu.Text = tbPARAMETROS("factCantViasImpresas")
    Me.chkImpAlojaExtranjeros.Value = tbPARAMETROS("factDiferenciarImpAlojaExt")
    'actualizo los controles seg�n valor asignado al control check
    subActualizoControlesImpExt
    posiciono_combo Me.cboNacionalidades, tbPARAMETROS("factNacionalidadLocal")
    posiciono_combo Me.cboImpAlojaExtranjero, tbPARAMETROS("factTipoImpAlojaExt")
    posiciono_combo Me.cboImpAloja, tbPARAMETROS("TipoIvaAloja")
    
    'inicializo etiqueta que mustra cantidad m�nimas de v�as
    Me.lblCantMinVias.Caption = "(cantidad m�nima = " & cantMinimoVias & ")"
End Sub

Private Sub chkImpAlojaExtranjeros_Click()
    'actualizo controles al cambiar el estado del control
    subActualizoControlesImpExt
End Sub

Private Sub subActualizoControlesImpExt()
    '--------------------------------------------------------------------------------
    'Actualiza las propiedades de los cotroles que configuran el impuesto a los ext.
    'dependiendo si esta activada la opci�n o no.
    '---------------------------------------------------------------------------------
    If chkImpAlojaExtranjeros.Value = 0 Then
        'no aplico diferencia entre extranjeros y pax nacionales
        mSubBloqueoControlFormulario Me.cboImpAlojaExtranjero, True
        mSubBloqueoControlFormulario Me.cboNacionalidades, True
    Else
        mSubBloqueoControlFormulario Me.cboImpAlojaExtranjero, False
        mSubBloqueoControlFormulario Me.cboNacionalidades, False
    End If
End Sub

Private Sub txtCantViasDocu_KeyPress(KeyAscii As Integer)
    'Solo permito ingreso de n�meros
    ValidoNum KeyAscii, False, False
End Sub

Private Sub txtCantViasDocu_LostFocus()
    'Controlo que la cantidad de v�a impresas sea como m�nimo 1
    If Val(txtCantViasDocu.Text) < cantMinimoVias Then
        txtCantViasDocu.Text = cantMinimoVias
    End If
    Me.lblCantMinVias.Caption = "(cantidad m�nima = " & cantMinimoVias & ")"
End Sub

