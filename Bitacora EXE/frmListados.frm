VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmListados 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Asistente de creación de listados"
   ClientHeight    =   7650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8085
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7650
   ScaleWidth      =   8085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab sstab1 
      Height          =   7455
      Left            =   120
      TabIndex        =   41
      Top             =   120
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   13150
      _Version        =   327680
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Paso &1: Listado"
      TabPicture(0)   =   "frmListados.frx":0000
      Tab(0).ControlCount=   1
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame5"
      Tab(0).Control(0).Enabled=   0   'False
      TabCaption(1)   =   "Paso &2: Selección"
      TabPicture(1)   =   "frmListados.frx":001C
      Tab(1).ControlCount=   1
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      TabCaption(2)   =   "Paso &3: Columnas"
      TabPicture(2)   =   "frmListados.frx":0038
      Tab(2).ControlCount=   1
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame6"
      Tab(2).Control(0).Enabled=   0   'False
      TabCaption(3)   =   "Paso &4: Orden"
      TabPicture(3)   =   "frmListados.frx":0054
      Tab(3).ControlCount=   1
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame4"
      Tab(3).Control(0).Enabled=   0   'False
      Begin VB.Frame Frame6 
         Caption         =   "&Columnas a mostrar"
         Height          =   6735
         Left            =   -74880
         TabIndex        =   22
         Top             =   480
         Width           =   7455
         Begin VB.CommandButton botTodosCol 
            Caption         =   "Todos"
            Height          =   375
            Left            =   4080
            TabIndex        =   26
            Top             =   2880
            Width           =   1215
         End
         Begin VB.CommandButton botCancelarP3 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   120
            TabIndex        =   29
            Top             =   6240
            Width           =   1215
         End
         Begin VB.CommandButton botSiguienteP3 
            Caption         =   "Siguiente >"
            Height          =   375
            Left            =   6120
            TabIndex        =   28
            Top             =   6240
            Width           =   1215
         End
         Begin VB.CommandButton botAtrasP3 
            Caption         =   "< Atrás"
            Height          =   375
            Left            =   4800
            TabIndex        =   27
            Top             =   6240
            Width           =   1215
         End
         Begin VB.CommandButton botBajar 
            Caption         =   "Bajar"
            Height          =   375
            Left            =   4080
            TabIndex        =   25
            Top             =   960
            Width           =   1215
         End
         Begin VB.CommandButton botSubir 
            Caption         =   "Subir"
            Height          =   375
            Left            =   4080
            TabIndex        =   24
            Top             =   480
            Width           =   1215
         End
         Begin VB.ListBox lstCol 
            Height          =   2760
            Left            =   240
            Style           =   1  'Checkbox
            TabIndex        =   23
            Top             =   480
            Width           =   3615
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "&Tipo de listado"
         Height          =   6735
         Left            =   -74880
         TabIndex        =   30
         Top             =   480
         Width           =   7455
         Begin VB.CheckBox chkCorte 
            Caption         =   "Realizar corte de control"
            Height          =   375
            Left            =   240
            TabIndex        =   36
            Top             =   3480
            Width           =   3855
         End
         Begin VB.CheckBox chkSaltoPag 
            Caption         =   "Realizar salto de página para cada corte"
            Enabled         =   0   'False
            Height          =   255
            Left            =   240
            TabIndex        =   37
            Top             =   3840
            Width           =   4815
         End
         Begin VB.CommandButton botDesc 
            Caption         =   "Descendente"
            Height          =   375
            Left            =   4080
            TabIndex        =   35
            Top             =   2865
            Width           =   1455
         End
         Begin VB.CommandButton botAsc 
            Caption         =   "Ascendente"
            Height          =   375
            Left            =   4080
            TabIndex        =   34
            Top             =   2400
            Width           =   1455
         End
         Begin VB.CommandButton botCancelarP4 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   120
            TabIndex        =   40
            Top             =   6240
            Width           =   1215
         End
         Begin VB.CommandButton botFinalizar 
            Caption         =   "Finalizar"
            Height          =   375
            Left            =   6120
            TabIndex        =   39
            Top             =   6240
            Width           =   1215
         End
         Begin VB.CommandButton botAtrasP4 
            Caption         =   "< Atrás"
            Height          =   375
            Left            =   4800
            TabIndex        =   38
            Top             =   6240
            Width           =   1215
         End
         Begin VB.ListBox lstOrden 
            Height          =   2760
            Left            =   240
            Style           =   1  'Checkbox
            TabIndex        =   31
            Top             =   480
            Width           =   3615
         End
         Begin VB.CommandButton botSubirOrden 
            Caption         =   "Subir"
            Height          =   375
            Left            =   4080
            TabIndex        =   32
            Top             =   480
            Width           =   1215
         End
         Begin VB.CommandButton botBajarOrden 
            Caption         =   "Bajar"
            Height          =   375
            Left            =   4080
            TabIndex        =   33
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Ordenado por:"
            Height          =   195
            Left            =   240
            TabIndex        =   44
            Top             =   480
            Width           =   1020
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Nuevo listado"
         Height          =   6735
         Left            =   120
         TabIndex        =   43
         Top             =   480
         Width           =   7455
         Begin VB.CommandButton botCancelarP1 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   120
            TabIndex        =   7
            Top             =   6240
            Width           =   1215
         End
         Begin VB.CommandButton botContinuarP1 
            Caption         =   "Siguiente >"
            Height          =   375
            Left            =   6120
            TabIndex        =   6
            Top             =   6240
            Width           =   1215
         End
         Begin VB.TextBox txtNombreListado 
            Height          =   375
            Left            =   240
            MaxLength       =   50
            TabIndex        =   1
            Top             =   600
            Width           =   2655
         End
         Begin VB.TextBox txtInfListado 
            Height          =   2295
            Left            =   240
            MultiLine       =   -1  'True
            TabIndex        =   5
            Top             =   2040
            Width           =   4815
         End
         Begin VB.TextBox txtDescListado 
            Height          =   375
            Left            =   240
            MaxLength       =   100
            TabIndex        =   3
            Top             =   1320
            Width           =   4815
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "&Nombre"
            Height          =   240
            Left            =   240
            TabIndex        =   0
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "&Información"
            Height          =   240
            Left            =   240
            TabIndex        =   4
            Top             =   1800
            Width           =   1035
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "&Descripción corta"
            Height          =   240
            Left            =   240
            TabIndex        =   2
            Top             =   1080
            Width           =   1575
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Selección de operaciones a listar"
         Height          =   6735
         Left            =   -74880
         TabIndex        =   42
         Top             =   480
         Width           =   7455
         Begin VB.Frame Frame2 
            Caption         =   "Trabajar con fecha de:"
            Height          =   1575
            Left            =   360
            TabIndex        =   45
            Top             =   4320
            Width           =   4335
            Begin VB.OptionButton opTodas 
               Caption         =   "Todas"
               Height          =   255
               Left            =   120
               TabIndex        =   14
               Top             =   360
               Value           =   -1  'True
               Width           =   1575
            End
            Begin VB.OptionButton opFechaRango 
               Caption         =   "Pedir rango "
               Height          =   255
               Left            =   1920
               TabIndex        =   18
               Top             =   720
               Width           =   1455
            End
            Begin VB.OptionButton opFechaSola 
               Caption         =   "Pedir solo una fecha"
               Height          =   255
               Left            =   1920
               TabIndex        =   17
               Top             =   360
               Width           =   2295
            End
            Begin VB.OptionButton opSistema 
               Caption         =   "Sistema"
               Height          =   255
               Left            =   120
               TabIndex        =   15
               Top             =   720
               Width           =   1695
            End
            Begin VB.OptionButton opAplicacion 
               Caption         =   "Aplicación"
               Height          =   255
               Left            =   120
               TabIndex        =   16
               Top             =   1080
               Width           =   1935
            End
         End
         Begin VB.CommandButton botTodosOpr 
            Caption         =   "Todos"
            Height          =   375
            Left            =   5040
            TabIndex        =   13
            Top             =   2520
            Width           =   1215
         End
         Begin VB.CommandButton botTodosUsr 
            Caption         =   "Todos"
            Height          =   375
            Left            =   5040
            TabIndex        =   12
            Top             =   480
            Width           =   1215
         End
         Begin VB.CommandButton botCancelarP2 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   360
            TabIndex        =   21
            Top             =   6240
            Width           =   1215
         End
         Begin VB.CommandButton botSiguienteP2 
            Caption         =   "Siguiente >"
            Height          =   375
            Left            =   6120
            TabIndex        =   20
            Top             =   6240
            Width           =   1215
         End
         Begin VB.CommandButton botAtrasP2 
            Caption         =   "< Atrás"
            Height          =   375
            Left            =   4800
            TabIndex        =   19
            Top             =   6240
            Width           =   1215
         End
         Begin VB.ListBox lstUsr 
            Height          =   1410
            Left            =   360
            Style           =   1  'Checkbox
            TabIndex        =   9
            Top             =   480
            Width           =   4335
         End
         Begin VB.ListBox lstOpr 
            Height          =   1410
            Left            =   360
            Style           =   1  'Checkbox
            TabIndex        =   11
            Top             =   2520
            Width           =   4335
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "&Usuarios"
            Height          =   240
            Left            =   360
            TabIndex        =   8
            Top             =   240
            Width           =   810
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "&Operaciones"
            Height          =   240
            Left            =   360
            TabIndex        =   10
            Top             =   2280
            Width           =   1170
         End
      End
   End
End
Attribute VB_Name = "frmListados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'USUARIO "Usuario por defecto."
    'Cuando la aplicación principal no tiene definidos perfiles de usuarios, se muestra
    'en la lista de usuarios, el usuario denominado "Usuario por defecto."
    
    'Cuando se ejecuta en la aplicación general, el procedimiento que crea un nuevo registro
    'en la tabla de bitácora, se le pasa como parámetro el nombre del usuario actual de la aplicación.
    'Si la aplicación no tiene definidos perfiles de usuarios, el valor de dicho parámetro
    'será Empty. La rutina que se encuntra en bitacora.DLL reconoce este valor y asigna al campo
    'usuario del registro del archvio de bitácora, el valor "Usuario por defecto."
    
    'Al utilizar la aplicación bitacora y crear un nuevo listado, solo se podrá trabajar con el usuario
    'por defecto (ya que no existe ningún otro definido).

'SE DEFINEN USUARIOS
    'Si luego se definen 1 o más perfiles de usuarios, la aplicación general dejará de enviar
    'el campo usuario = Empty, pasándole el nombre del usuario que realizó la operación.
    
    'En la aplicación bitácora, por otra parte, al realizar un nuevo listado, ya no aparecerá
    'más el usuario por defecto en la lista de usuarios del paso 2 , suplantándose por todos los
    'usuarios definidos.
    'Esto implica que ya no se podrán crear (si ejecutar los ya creados) nuevos listados
    'para el usuario por defecto.(es lógico que así fuese, ya que no exitirán operacione nuevas realizadas por
    'este usuario)
    'Del mismo modos si se eliman todos los perfiles de usuario establecidos, nuevamente
    'aparecerá el usuario por defecto, en la lista de usuarios, no pudiéndose crear
    'nuevos listados que no sean para este tipo de usuario.
    

Private PrimeraVezP2 As Boolean
Private PrimeraVezP3 As Boolean



Private Sub chkCorte_Click()
    If chkCorte.Value = 1 Then  'realizo corte de control
        Me.chkSaltoPag.Enabled = True
    Else
        Me.chkSaltoPag.Value = 0
        Me.chkSaltoPag.Enabled = False
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    PrimeraVezP2 = True
    PrimeraVezP3 = True
End Sub

Private Sub botAsc_Click()
    'El campo seleccionado se ordena en forma ascendente
                                        
    lstOrden.List(lstOrden.ListIndex) = "ASC" & _
                Mid(lstOrden.List(lstOrden.ListIndex), 6)
End Sub

Private Sub botDesc_Click()
    'El campo seleccionado se ordena en forma descendente
    
    lstOrden.List(lstOrden.ListIndex) = "DESC" & _
                Mid(lstOrden.List(lstOrden.ListIndex), 6)
End Sub

Private Sub botAtrasP2_Click()
    Me.sstab1.TabEnabled(1) = False
    'Retrocedo al paso 1
    Me.sstab1.Tab = 0
    Me.sstab1.TabEnabled(0) = True
End Sub

Private Sub botAtrasP3_Click()
    Me.sstab1.TabEnabled(2) = False
    'Retrocedo al paso 2
    Me.sstab1.Tab = 1
    Me.sstab1.TabEnabled(1) = True
End Sub

Private Sub botAtrasP4_Click()
    Me.sstab1.TabEnabled(3) = False
    'Retrocedo al paso 3
    Me.sstab1.Tab = 2
    Me.sstab1.TabEnabled(2) = True
End Sub

Private Sub botCancelarP1_Click()
    Unload Me
End Sub

Private Sub botCancelarP2_Click()
    Unload Me
End Sub

Private Sub botCancelarP3_Click()
    Unload Me
End Sub

Private Sub botCancelarP4_Click()
    Unload Me
End Sub

Private Sub botContinuarP1_Click()
    If funValidoSigP1 Then
        'paso siguiente tab (paso 2)
        Me.sstab1.TabEnabled(1) = True
        Me.sstab1.Tab = 1
        Me.sstab1.TabEnabled(0) = False
        
        If PrimeraVezP2 Then
            'carga usuarios a la lista de usuarios
            subCargoUsuarios
            
            'cargo operaciones a la lista de operaciones
            subCargoOperaciones
            
            PrimeraVezP2 = False
        End If
    End If
End Sub

Private Sub botFinalizar_Click()
    'grabo listado
    subGraboListado
End Sub

Private Sub botSiguienteP2_Click()
    If funValidoSigP2 Then
        'paso siguiente tab (paso 3)
        Me.sstab1.TabEnabled(2) = True
        Me.sstab1.Tab = 2
        Me.sstab1.TabEnabled(1) = False
        
        If PrimeraVezP3 Then
            'cargo columnas a listar
            subCargoColumnas
            PrimeraVezP3 = False
        End If
    End If
End Sub

Private Sub subCargoColumnas()
    'Muestro en la lista los campos del archivo de bitacora
    'IMPORTANTE: el orden de los campos en el archivo de bitacora
    'debe de coincidir con el siguiente indicado en el itemdata
    Me.lstCol.AddItem "Fecha"
    Me.lstCol.ItemData(Me.lstCol.NewIndex) = 0
    Me.lstCol.AddItem "Operación"
    Me.lstCol.ItemData(Me.lstCol.NewIndex) = 2
    Me.lstCol.AddItem "Usuario"
    Me.lstCol.ItemData(Me.lstCol.NewIndex) = 3
    Me.lstCol.AddItem "Hora inicio"
    Me.lstCol.ItemData(Me.lstCol.NewIndex) = 4
    Me.lstCol.AddItem "Hora fin"
    Me.lstCol.ItemData(Me.lstCol.NewIndex) = 5
    Me.lstCol.AddItem "Observaciones"
    Me.lstCol.ItemData(Me.lstCol.NewIndex) = 6
End Sub

Private Sub botSiguienteP3_Click()
    'No necesito controlar que se la primera vez que
    'ingreso ya que cada vez, debo de cargar la lista
    'con los campo seleccionado en el paso anterior
    
    If funValidoSigP3 Then
        'paso siguiente tab (paso 4)
        Me.sstab1.TabEnabled(3) = True
        Me.sstab1.Tab = 3
        Me.sstab1.TabEnabled(2) = False
        
        'Limpio lista
        Me.lstOrden.Clear
        'Cargo campos por los que se pueda ordenar
        subCargoCamposOrdenar
        
    End If
End Sub

Private Sub subCargoCamposOrdenar()
    'Muestro los campos que hayan sido seleccionados
    Dim i As Integer
    Dim linea As String
    Dim espacios As String
    
    espacios = "        "
    'recorro la lista de columnas a mostrar del paso 3
    i = 0
    Do While i < Me.lstCol.ListCount
        ' si la columna está seleccionada
        If lstCol.Selected(i) = True Then
            linea = "ASC" & espacios & Me.lstCol.List(i)
            Me.lstOrden.AddItem linea
            'Tambien cargo el itemdata
            Me.lstOrden.ItemData(Me.lstOrden.NewIndex) = Me.lstCol.ItemData(i)
        End If
        i = i + 1
    Loop
End Sub

Private Sub botSubirOrden_Click()
    subSubir lstOrden
End Sub

Private Sub botSubir_Click()
    subSubir lstCol
End Sub

Private Sub botBajarOrden_Click()
    subBajar lstOrden
End Sub

Private Sub botBajar_Click()
    subBajar lstCol
End Sub

Private Sub subSubir(lst As ListBox)
    'Cada vez que apreto el boton subir subo el elemento
    '(nombre columna) , una posición.
    Dim subir As String
    Dim subirItemData As Integer
    Dim estadoSubir As Boolean
    Dim bajar As String
    Dim bajarItemData As String
    
    Dim estadoBajar As Boolean
    Dim indice As Integer
    
    'si es el primero de la lista no hago nada
    If lst.ListIndex > 0 Then
        indice = lst.ListIndex  'posición actual
        'almaceno información del elemento seleccionado
        subir = lst.List(indice)                'texto
        subirItemData = lst.ItemData(indice)    'itemdata
        estadoSubir = lst.Selected(indice)      'selección
        
        'almaceno información del elemento que está arriba
        bajar = lst.List(indice - 1)            'texto
        bajarItemData = lst.ItemData(indice - 1)  'itemdata
        estadoBajar = lst.Selected(indice - 1)  'selección
        
        'intercambio texto, marcas y itemdata
        lst.List(indice) = bajar
        lst.ItemData(indice) = bajarItemData
        lst.Selected(indice) = estadoBajar
        
        lst.List(indice - 1) = subir
        lst.ItemData(indice - 1) = subirItemData
        lst.Selected(indice - 1) = estadoSubir
        
        'me posiciono un elemento más arriba
        lst.ListIndex = indice - 1
    End If
End Sub

Private Sub subBajar(lst As ListBox)
    'Cada vez que apreto el boton bajar, bajo el elemento
    '(nombre columna) una posición.
    Dim subir As String
    Dim subirItemData As Integer
    Dim estadoSubir As Boolean
    Dim bajar As String
    Dim bajarItemData As String
    Dim estadoBajar As Boolean
    Dim indice As Integer
    
    'si es el primero de la lista no hago nada
    If lst.ListIndex < lst.ListCount - 1 Then
        indice = lst.ListIndex  'posición actual
        'almaceno información del elemento que esta abajo
        subir = lst.List(indice + 1)              'texto
        subirItemData = lst.ItemData(indice + 1)  'itemdata
        estadoSubir = lst.Selected(indice + 1)    'selección
        
        'almaceno información del elemento seleccionado
        bajar = lst.List(indice)                'texto
        bajarItemData = lst.ItemData(indice)    'itemdata
        estadoBajar = lst.Selected(indice)      'selección
        
        'intercambio texto, marcas y itemdata
        lst.List(indice) = subir
        lst.ItemData(indice) = subirItemData
        lst.Selected(indice) = estadoSubir
        
        lst.List(indice + 1) = bajar
        lst.ItemData(indice + 1) = bajarItemData
        lst.Selected(indice + 1) = estadoBajar
        
        'me posiciono un elemento más abajo
        lst.ListIndex = indice + 1
    End If
End Sub

Private Sub botTodosCol_Click()
    'Selecciono todas las columnas
    mSubSeleccionoTodos frmListados.lstCol
End Sub

Private Sub botTodosOpr_Click()
    'Selecciono todos las operaciones
    mSubSeleccionoTodos frmListados.lstOpr
End Sub

Private Sub botTodosUsr_Click()
    'Selecciono todos los usuarios
    mSubSeleccionoTodos frmListados.lstUsr
End Sub

Private Sub Form_Activate()
    'Por defecto muestro tab 1
    Me.sstab1.Tab = 0
    
    'No permito trabajar con los demás tabs
    Me.sstab1.TabEnabled(1) = False
    Me.sstab1.TabEnabled(2) = False
    Me.sstab1.TabEnabled(3) = False
End Sub

Private Function funValidoSigP1()
    'Valida que se pueda pasar al paso2.
    funValidoSigP1 = True
    If txtNombreListado.Text = Empty Then
        MsgBox "Debe de ingresar nombre de listado, para continuar", vbExclamation
        txtNombreListado.SetFocus
        funValidoSigP1 = False
    End If
End Function

Private Function funValidoSigP2()
    'Valida que se pueda pasar al paso 3
    funValidoSigP2 = True
    'Es obligación seleccionar algún usuario
    If lstUsr.SelCount = 0 Then
        MsgBox "Debe de seleccionar usuarios", vbExclamation
        funValidoSigP2 = False
        Exit Function
    End If
    'Es obligación seleccionar alguna operación
    If lstOpr.SelCount = 0 Then
        MsgBox "Debe de seleccionar operaciones", vbExclamation
        funValidoSigP2 = False
        Exit Function
    End If
End Function

Private Function funValidoSigP3()
    'Valida que se pueda pasar al paso 4
    funValidoSigP3 = True
    If lstCol.SelCount = 0 Then
        MsgBox "Debe de seleccionar por lo menos una columna", vbExclamation
        funValidoSigP3 = False
        Exit Function
    End If
End Function

Private Sub subCargoUsuarios()
    'Recorro los usuarios de la aplicación
    'y los mustro en la lista.
    'Si no existen perfiles de usuarios definidos no se muestra
    'el usuario por defecto y no se permite trabajar con el control lista.
    
    'verifico si tengo usuarios definidos
    If tbSISTEMA_USUARIOS.RecordCount > 0 Then
        'existen usuarios
        tbSISTEMA_USUARIOS.Index = "iclaves"
        tbSISTEMA_USUARIOS.MoveFirst
        If Not tbSISTEMA_USUARIOS.NoMatch Then 'si hay usuarios
            Do While Not tbSISTEMA_USUARIOS.EOF
                lstUsr.AddItem tbSISTEMA_USUARIOS("NomUsr")
                tbSISTEMA_USUARIOS.MoveNext
            Loop
        End If
    Else
        'no existen usuario, por lo que creo el usuario por defecto
        lstUsr.AddItem "Usuario por defecto."
        Me.lstUsr.Enabled = False
        Me.botTodosUsr.Enabled = False
        'selecciono el usuario
        Me.lstUsr.Selected(0) = True
    End If
End Sub

Private Sub subCargoOperaciones()
    'Recorro las operaciones del sistema
    'y las muestra en la lista de operaciones
    
    tbSISTEMA_OPERACIONES.Index = "i_DescOpr"
    tbSISTEMA_OPERACIONES.MoveFirst
    If Not tbSISTEMA_OPERACIONES.NoMatch Then   'si hay operaciones
        Do While Not tbSISTEMA_OPERACIONES.EOF
            lstOpr.AddItem tbSISTEMA_OPERACIONES("DescOpr")
            lstOpr.ItemData(lstOpr.NewIndex) = tbSISTEMA_OPERACIONES("CodOpr")
            tbSISTEMA_OPERACIONES.MoveNext
        Loop
    End If
End Sub

Private Sub subGraboListado()
    'crea un nuevo registro en el archivo sistema_bitacora_listado
    'que con información hacerca del listado creado.

    tbSISTEMA_BITACORAlistados.Index = "pk_listado"
    tbSISTEMA_BITACORAlistados.Seek "=", Me.txtNombreListado.Text
    If tbSISTEMA_BITACORAlistados.NoMatch Then  'si no existe
        tbSISTEMA_BITACORAlistados.AddNew
            tbSISTEMA_BITACORAlistados("NomLst") = Me.txtNombreListado.Text
            tbSISTEMA_BITACORAlistados("DescLst") = Me.txtDescListado.Text
            tbSISTEMA_BITACORAlistados("InfLst") = Me.txtInfListado.Text
            tbSISTEMA_BITACORAlistados("WhereOprLst") = funObtengoOpr
            tbSISTEMA_BITACORAlistados("WhereUsrLst") = funObtengoUsr
            tbSISTEMA_BITACORAlistados("WhereTipoFechaLst") = funObtengoTipoFecha
            tbSISTEMA_BITACORAlistados("ColumnLst") = funObtengoColumnas
            tbSISTEMA_BITACORAlistados("ColumnDescLst") = funObtengoColumnasDesc
            tbSISTEMA_BITACORAlistados("OrdenLst") = funObtengoOrden
            tbSISTEMA_BITACORAlistados("CampoCorteLst") = funObtengoCampoCorte
            tbSISTEMA_BITACORAlistados("RealizoCorte") = funObtengoTipoCorte
        tbSISTEMA_BITACORAlistados.Update
        Unload Me
        frmFinAsistente.Show 1
    Else
        'si existe tiene que cambiarle el nombre
        MsgBox "El nombre del listado que ingreso ya existe.", vbExclamation
        sstab1.Tab = 0
        txtNombreListado.SetFocus
    End If
End Sub

Private Function funObtengoOpr()
    'Crea un string con información hacerca,
    'de las operaciones seleccionadas
    
    Dim i As Integer
    Dim linea As String
    'Recorro la lista de operaciones y selecciono las marcadas
    i = 0
    Do While i < Me.lstOpr.ListCount
        If Me.lstOpr.Selected(i) Then
            'si el elemento esta selccionado
            linea = linea & " " & "CodOprBit = " & Me.lstOpr.ItemData(i) & " or "
        End If
        i = i + 1
    Loop
    'saco de la lista el último or
    linea = Mid(linea, 1, Len(linea) - 3)
    funObtengoOpr = linea
End Function

Private Function funObtengoUsr()
    'Crea un string con información hacerca
    'de los usuarios seleccionados
    
    Dim i As Integer
    Dim linea As String
    i = 0
    Do While i < Me.lstUsr.ListCount
        If Me.lstUsr.Selected(i) Then
            'si el elemento esta selccionado
            linea = linea & " " & "NomUsrBit = '" & Me.lstUsr.List(i) & "' or "
        End If
        i = i + 1
    Loop
    
    'saco de la lista el último or
    linea = Mid(linea, 1, Len(linea) - 3)
    funObtengoUsr = linea
End Function

Private Function funObtengoTipoFecha()
    'Determina la forma de trabajar con las fechas
    '1= fecha sistema
    '2= fecha aplicación  (archivo parámetros)
    '3= pido fecha
    '4= pido rango de fechas
    If Me.opSistema.Value = True Then
        funObtengoTipoFecha = 1
    Else
        If Me.opAplicacion = True Then
            funObtengoTipoFecha = 2
        Else
            If Me.opFechaSola = True Then
                funObtengoTipoFecha = 3
            Else
                If Me.opFechaRango = True Then
                    funObtengoTipoFecha = 4
                Else
                    If Me.opTodas = True Then
                        funObtengoTipoFecha = 5
                    End If
                End If
            End If
        End If
    End If
End Function

Private Function funObtengoColumnas()
    'Recorro las columnas seleccionadas
    'y creo un string con las mismas
    
    Dim i As Integer
    Dim linea As String
    i = 0
    Do While i < Me.lstCol.ListCount
        If Me.lstCol.Selected(i) Then
            'si el elemento esta selccionado
            linea = linea & " " & _
            funObtengoNomCol(Me.lstCol.ItemData(i)) & ","
        End If
        i = i + 1
    Loop
    'saco de la lista la última coma
    linea = Mid(linea, 1, Len(linea) - 1)
    funObtengoColumnas = linea
End Function

Private Function funObtengoNomCol(i As Integer)
    'Devuleve el nombre del campo del archivo sistema_bitacora
    'que corresponda con el parametro
    funObtengoNomCol = tbSISTEMA_BITACORA.Fields(i).Name
End Function

Private Function funObtengoColumnasDesc()
    'Grabo la descripción que se mostrará en la pantalla
    'correspondientes a los campos que se seleccionaron
    'Esto es muy importante ya que me servirá para crear el cabezal
    'del comtrol listview
    
    Dim i As Integer
    Dim linea As String
    i = 0
    'Recorro las columnas seleccionadas
    Do While i < Me.lstCol.ListCount
        If Me.lstCol.Selected(i) Then
            'si el elemento esta selccionado
            linea = linea & Me.lstCol.List(i) & ","
        End If
        i = i + 1
    Loop
    'saco de la lista la última coma
    linea = Mid(linea, 1, Len(linea) - 1)
    funObtengoColumnasDesc = linea
End Function

Private Function funObtengoOrden()
    'Grabo los nombres de los campos por los cuales
    'se ordenará el listado y si lo harán, en forma
    'ascendente o descendente.
    
    Dim i As Integer
    Dim linea As String
    Dim campo As String
    Dim TipoOrden As String
    Dim comando As String
    i = 0
    comando = ""
    funObtengoOrden = ""
    'Recorro los campos seleccionados
    Do While i < Me.lstOrden.ListCount
        If Me.lstOrden.Selected(i) Then
            'si esta seleccionado
            comando = " ORDER BY "
            TipoOrden = Mid(Me.lstOrden.List(i), 1, 4)  'obtengo "ASC" o "DESC"
            campo = funObtengoNomCol(Me.lstOrden.ItemData(i))
            linea = linea & campo & " " & TipoOrden & ","
        End If
        i = i + 1
    Loop
    'saco de la lista la última coma, solo si seleccione algún campo
    If Len(linea) > 0 Then
        linea = Mid(linea, 1, Len(linea) - 1)
        funObtengoOrden = comando & linea
    End If
End Function

Private Function funObtengoCampoCorte()
    'Obtengo el campo por el cual se realiza el corte de control
    
    Dim i As Integer
    Dim CampoCorte As String
    
    If Me.chkCorte.Value = 1 Then
        i = 0
        'Recorro los campos de orden seleccionados
        Do While i < Me.lstOrden.ListCount
            If Me.lstOrden.Selected(i) Then
                'si esta seleccionado, este será el campo por el cual realizo el corte
                CampoCorte = LTrim(Mid(Me.lstOrden.List(i), 4))
                Exit Do 'solo presiso el primer campo seleccionado
            End If
            i = i + 1
        Loop
        If Len(CampoCorte) > 0 Then
            'si se seleccionó algun campo para ordennar
            funObtengoCampoCorte = CampoCorte
        Else
            'obtengo primera columna seleccionado
            i = 0
            'Recorro las columnas seleccionadas
            Do While i < Me.lstCol.ListCount
                If Me.lstCol.Selected(i) Then
                    'si el elemento esta selccionado,
                    'este  será el campo del corte de control
                    funObtengoCampoCorte = LTrim(Mid(Me.lstCol.List(i), 4))
                    Exit Do
                End If
                i = i + 1
            Loop
        End If
    Else
        'no realizo corte de control
        funObtengoCampoCorte = ""
    End If
End Function

Private Function funObtengoTipoCorte()
    'Determina si realizo corte de control y si es así determina
    'si salto de página con cada sección del corte
    '1=realizo corte de control comun
    '2=realizo corte de control con salto de página
    '3=no realizo corte de control
    If frmListados.chkCorte.Value = 1 Then
        funObtengoTipoCorte = 1
        If frmListados.chkSaltoPag.Value = 1 Then
            funObtengoTipoCorte = 2
        End If
    Else
        funObtengoTipoCorte = 3
    End If
End Function

Private Sub Form_Unload(Cancel As Integer)
    'Descargo formulario de memoria
    Set frmListados = Nothing
End Sub
