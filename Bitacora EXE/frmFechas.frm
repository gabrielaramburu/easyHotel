VERSION 5.00
Object = "{EB8C3860-8603-11D0-A35D-7062E4000000}#1.1#0"; "VCBND.OCX"
Begin VB.Form frmFechas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingrese fecha"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Ingrese fecha de selección"
      Height          =   3015
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4455
      Begin VcBndCtl.VcCalCombo fFechaIni 
         Height          =   375
         Left            =   1560
         TabIndex        =   0
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _0              =   $"frmFechas.frx":0000
         _1              =   $"frmFechas.frx":0409
         _2              =   $"frmFechas.frx":0812
         _3              =   "-E@@@A@@@@@@@@@'@@@E@@@A@@@@@@@@@)h@@@E@@@A@@@@@@@@@,4E7F"
         _count          =   4
         _ver            =   2
      End
      Begin VB.CommandButton botCancelar 
         Cancel          =   -1  'True
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   3120
         TabIndex        =   3
         Top             =   2520
         Width           =   1215
      End
      Begin VB.CommandButton botAceptar 
         Caption         =   "&Aceptar"
         Default         =   -1  'True
         Height          =   375
         Left            =   1800
         TabIndex        =   2
         Top             =   2520
         Width           =   1215
      End
      Begin VcBndCtl.VcCalCombo fFechaFin 
         Height          =   375
         Left            =   1560
         TabIndex        =   1
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _0              =   $"frmFechas.frx":0C1B
         _1              =   $"frmFechas.frx":1024
         _2              =   $"frmFechas.frx":142D
         _3              =   "-@@@E@@@A@@@@@@@@@'@@@E@@@A@@@@@@@@@)h@@@E@@@A@@@@@@@@@,4E7F"
         _count          =   4
         _ver            =   2
      End
      Begin VB.Label lblFechaFin 
         Caption         =   "lblFechaFin"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1140
         Width           =   1095
      End
      Begin VB.Label lblFechaIni 
         Caption         =   "lblFechaIni"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   540
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmFechas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    If tipo_accion_fechas = 1 Then  'fecha única
        'oculta segunda fechas
        lblFechaFin.Visible = False
        fFechaFin.Visible = False
        'cambio titulo a etiqueta de fecha inicial
        lblFechaIni.Caption = "Fecha "
    End If
    
    If tipo_accion_fechas = 2 Then  'rango de fechas
        'cambio titulo a etiqueta de fecha inicial
        lblFechaIni.Caption = "Fecha inicial"
        'cambio titulo a fecha final
        lblFechaFin.Caption = "Fecha final"
    End If
End Sub

Private Sub botAceptar_Click()
    If funValidoFechas Then
        Me.Hide
    End If
End Sub

Private Function funValidoFechas()
    'Valido que las fechas esten correctas
    funValidoFechas = True
    If IsDate(fFechaIni.Text) Then
        If tipo_accion_fechas = 2 Then  'rango de fechas
            If IsDate(fFechaFin.Text) Then
                If fFechaIni.Value > fFechaFin.Value Then
                    MsgBox "La fecha inicial debe de ser menor igual a la fecha final", vbExclamation
                    funValidoFechas = False
                    Exit Function
                End If
            Else
                MsgBox "La fecha final no es correcta"
                funValidoFechas = False
                Exit Function
            End If
        End If
    Else
        MsgBox "La fecha inicial no es correcta"
        funValidoFechas = False
        Exit Function
    End If
End Function

Private Sub botCancelar_Click()
    CanceloIngresoFechas = True
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Descargo formulario de memoria
    Set frmFechas = Nothing
End Sub
