VERSION 5.00
Begin VB.Form frmAcercaDe 
   Caption         =   "Acerca de..."
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5550
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   5550
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Changoski"
      Height          =   255
      Left            =   3960
      TabIndex        =   5
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Matilindo"
      Height          =   255
      Left            =   2520
      TabIndex        =   4
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "¿Quién es  mejor jugador de ajedrez?"
      Height          =   240
      Left            =   2040
      TabIndex        =   3
      Top             =   2400
      Width           =   3345
   End
   Begin VB.Image Image1 
      Height          =   3675
      Left            =   240
      Picture         =   "frmAcercaDe.frx":0000
      Top             =   120
      Width           =   3825
   End
   Begin VB.Label Label3 
      Caption         =   "Setiembre 2003"
      Height          =   255
      Left            =   2640
      TabIndex        =   2
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Creado por Gabriel Aramburu"
      Height          =   240
      Left            =   2640
      TabIndex        =   1
      Top             =   600
      Width           =   2640
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Listas 1.0"
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
      Left            =   2640
      TabIndex        =   0
      Top             =   240
      Width           =   990
   End
End
Attribute VB_Name = "frmAcercaDe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    MsgBox "Respuesta equivocada", vbCritical
End Sub

Private Sub Command2_Click()
    MsgBox "Respuesta correcta."
    Unload Me
End Sub
