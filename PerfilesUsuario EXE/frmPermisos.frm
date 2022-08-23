VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmPermisos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Permisos de usuario"
   ClientHeight    =   7755
   ClientLeft      =   1170
   ClientTop       =   690
   ClientWidth     =   9690
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   9690
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton botCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   8400
      TabIndex        =   10
      Top             =   7320
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Permisos de usuario "
      Height          =   7215
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   9495
      Begin VB.ComboBox cboOpr1Nivel 
         Height          =   360
         Left            =   3000
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   600
         Width           =   3615
      End
      Begin VB.CommandButton botNoAutorizo 
         Caption         =   "&No autorizo >>"
         Height          =   375
         Left            =   3960
         TabIndex        =   9
         Top             =   3480
         Width           =   1455
      End
      Begin VB.CommandButton botAutorizo 
         Caption         =   "<<   &Autorizo"
         Height          =   375
         Left            =   3960
         TabIndex        =   8
         Top             =   2520
         Width           =   1455
      End
      Begin VB.ComboBox cboUsr 
         Height          =   360
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   600
         Width           =   2535
      End
      Begin ComctlLib.ListView lwAuto 
         Height          =   5490
         Left            =   240
         TabIndex        =   5
         Top             =   1560
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   9684
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin ComctlLib.ListView lwNoAuto 
         Height          =   5490
         Left            =   5520
         TabIndex        =   7
         Top             =   1560
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   9684
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Image Image1 
         Height          =   105
         Left            =   240
         Picture         =   "frmPermisos.frx":0000
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   9090
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "O&peraciones 1er. nivel"
         Height          =   240
         Left            =   3000
         TabIndex        =   2
         Top             =   360
         Width           =   2010
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "&Denegadas"
         Height          =   240
         Left            =   5520
         TabIndex        =   6
         Top             =   1320
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "&Opciones autorizadas"
         Height          =   240
         Left            =   240
         TabIndex        =   4
         Top             =   1320
         Width           =   1965
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&Usuarios"
         Height          =   240
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Width           =   810
      End
   End
   Begin VB.Menu mnuFormulario 
      Caption         =   "&Formulario"
      Begin VB.Menu mnuFormularioSalir 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "frmPermisos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim permitoProcesar As Boolean  'controla que no se ejecute el evento click de los combos
                                'de usuarios y operaciones cuando se cargan.

Private Sub Form_Load()
    permitoProcesar = False
    Me.lwAuto.SmallIcons = frmMain.ImageList1
    Me.lwNoAuto.SmallIcons = frmMain.ImageList1
    'Cargo combos con operaciones de primer nivel
    mSubCargoComboOpr Me.cboOpr1Nivel, True, 1
    'Cargo combo con usuarios
    mSubCargoComboUsr Me.cboUsr, False
    permitoProcesar = True
    cboUsr_Click    'inicializo listas
End Sub

Private Sub cboOpr1Nivel_Click()
    'Cuando cambio el valor del combo actualizo las listas
    If permitoProcesar Then subProceso
End Sub

Private Sub cboUsr_Click()
    'Cuando cambio el valor del combo actualizo las listas
    If permitoProcesar Then subProceso
End Sub

Private Sub subProceso()
    'Antes de cargar los datos borro las listas
    subLimpioListas lwNoAuto
    subLimpioListas lwAuto
    
    'Cada vez que cambio de usuario muestro sus permisos
    subMuestroOpciones
    
    'Luego de cargar los datos los muestro
    'en forma de listas
    lwNoAuto.View = lvwList
    lwAuto.View = lvwList
End Sub

Private Sub botAutorizo_Click()
    'Autorizo las operaciones seleccionadas en la lista
    'de no autorizadas.
    Dim i As Integer
    Dim Opr As Integer
    Dim Cambio As Boolean
    i = 1
    
    Cambio = False
    'recorro todos los elementos no autorizados
    Do While i <= lwNoAuto.ListItems.Count
        'Si esta seleccionado
        If lwNoAuto.ListItems.Item(i).Selected = True Then
            'Busco el tipo de operacion
            Opr = Mid(lwNoAuto.ListItems.Item(i).Key, 4)
            'creo nuevo registro en archivo perfiles
            tbSISTEMA_PERFILES.AddNew
                tbSISTEMA_PERFILES("CodOpr") = Opr
                tbSISTEMA_PERFILES("NomUsr") = cboUsr.Text
            tbSISTEMA_PERFILES.Update
            Cambio = True
        End If
        i = i + 1
    Loop
    If Cambio Then
        'actualizo la informacion de las listas
        subProceso
    End If
End Sub

Private Sub botNoAutorizo_Click()
    'No autorizo las operaciones seleccionadas en la lista
    'de  autorizadas.
    Dim i As Integer
    Dim Opr As Integer
    Dim Cambio As Boolean
    i = 1
    Cambio = False
    'recorro todos los elementos autorizados
    Do While i <= lwAuto.ListItems.Count
        'Si esta seleccionado
        If lwAuto.ListItems.Item(i).Selected = True Then
            'Busco el tipo de operacion
            Opr = Mid(lwAuto.ListItems.Item(i).Key, 4)
            'borro registro en archivo perfiles
            tbSISTEMA_PERFILES.Index = "pk_perfiles"
            tbSISTEMA_PERFILES.Seek "=", Opr, cboUsr.Text
            If Not tbSISTEMA_PERFILES.NoMatch Then 'existe
                tbSISTEMA_PERFILES.Delete
            End If
            Cambio = True
        End If
        i = i + 1
    Loop
    
    If Cambio Then
        'actualizo la informacion de las listas
        subProceso
    End If
End Sub

Private Sub subMuestroOpciones()
    'Recorro el archivo de operaciones y determino
    'para el usuario actual, el estado de cada operación
    
    tbSISTEMA_OPERACIONES.Index = "pk_operaciones"
    tbSISTEMA_OPERACIONES.MoveFirst
    'recorro las opciones
    Do While Not tbSISTEMA_OPERACIONES.EOF
        'aplico filtro de nivel de operaciones
        If (Me.cboOpr1Nivel.ItemData(Me.cboOpr1Nivel.ListIndex) = _
        tbSISTEMA_OPERACIONES("perteneceA")) Or Me.cboOpr1Nivel.Text = "(Todas)" Then
            'Busco si esa operación esta permitida para el usuario
            If funOpcionAutorizada Then
                'Cargo en lista de autorizados
                subCargoOpciones Me.lwAuto, _
                    tbSISTEMA_OPERACIONES("TipoOpr"), _
                    tbSISTEMA_OPERACIONES("DescOpr"), _
                    tbSISTEMA_OPERACIONES("CodOpr")
            Else
                'Cargo en lista de no autorizadas
                subCargoOpciones Me.lwNoAuto, _
                    tbSISTEMA_OPERACIONES("TipoOpr"), _
                    tbSISTEMA_OPERACIONES("DescOpr"), _
                    tbSISTEMA_OPERACIONES("CodOpr")
            End If
        End If
        tbSISTEMA_OPERACIONES.MoveNext
    Loop
End Sub

Private Function funOpcionAutorizada()
    'Determina si una operacion esta autorizada para un determinado usuario.
    tbSISTEMA_PERFILES.Index = "pk_perfiles"
    tbSISTEMA_PERFILES.Seek "=", tbSISTEMA_OPERACIONES("CodOpr"), cboUsr.Text
    If Not tbSISTEMA_PERFILES.NoMatch Then  'autorizado por que existe
        funOpcionAutorizada = True
    Else                                'no autorizado
        funOpcionAutorizada = False
    End If
End Function

Private Sub subCargoOpciones(lista As ListView, _
                                TipoOpr As Byte, _
                                DescOpr As String, _
                                CodOpr As Integer)
    'Carga la lista de opciones a la lista pasada como parámetro
    Dim imagen As Byte
    
    If TipoOpr = 1 Then
        imagen = 2
    Else
        imagen = 3
    End If
    lista.ListItems.Add , "Opr" & Str(CodOpr), DescOpr, , imagen
    'Es necesario cargarle la palabra Opr antes de cada Codigo de operacion
    'para que funcione bien.
End Sub

Private Sub subLimpioListas(lista As ListView)
    'Limpio las listas
    lista.ColumnHeaders.Clear
    lista.ListItems.Clear
End Sub
    
Private Sub botCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Descargo formulario de memoria
    Set frmPermisos = Nothing
End Sub

Private Sub mnuFormularioSalir_Click()
    botCancelar_Click
End Sub
