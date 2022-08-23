VERSION 5.00
Begin VB.Form frmInicioSesion 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   4215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7845
   LinkTopic       =   "Form1"
   ScaleHeight     =   4112.195
   ScaleMode       =   0  'Usuario
   ScaleWidth      =   7845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Versión 1.0"
      Height          =   240
      Left            =   5520
      TabIndex        =   7
      Top             =   2040
      Width           =   2230
   End
   Begin VB.Label lblRevision 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Revisión"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   5520
      TabIndex        =   6
      Top             =   2280
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   3615
      Left            =   120
      Top             =   120
      Width           =   5175
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Producto protegido por las leyes internacionales como se describe en Acerca de, menu de Ayuda."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   240
      TabIndex        =   5
      Top             =   3840
      Width           =   7125
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Copyright (c) 2000 - 2002"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   5880
      TabIndex        =   4
      Top             =   600
      Width           =   1860
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Para Windows de 32 bits"
      Height          =   240
      Left            =   5520
      TabIndex        =   3
      Top             =   1560
      Width           =   2205
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Marcos Bernini"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   5880
      TabIndex        =   2
      Top             =   1080
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Gabriel Aramburu"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   5880
      TabIndex        =   1
      Top             =   840
      Width           =   1275
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "EasyHotel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5880
      TabIndex        =   0
      Top             =   120
      Width           =   1755
   End
   Begin VB.Line Line4 
      X1              =   7800
      X2              =   7800
      Y1              =   0
      Y2              =   4097.561
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   7800
      Y1              =   4097.561
      Y2              =   4097.561
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   0
      Y1              =   4097.561
      Y2              =   0
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   7800
      Y1              =   0
      Y2              =   0
   End
End
Attribute VB_Name = "frmInicioSesion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    On Error Resume Next
    'muestro revisión
    lblRevision.Caption = "Revisión " & App.Major & "." & App.Minor & "." & App.Revision
    'obtengo imagen
    Set Image1 = LoadPicture(App.Path & "\logohotel.bmp")
End Sub
