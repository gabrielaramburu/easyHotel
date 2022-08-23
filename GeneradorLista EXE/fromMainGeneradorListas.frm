VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmGeneral 
   Caption         =   "Generador de listas"
   ClientHeight    =   4005
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8640
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4005
   ScaleWidth      =   8640
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7920
      Top             =   -120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FileName        =   "*.txt"
      Filter          =   "*.txt"
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   6588
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Actualizar Base de datos"
      TabPicture(0)   =   "fromMainGeneradorListas.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Nuevas listas"
      TabPicture(1)   =   "fromMainGeneradorListas.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Información Base de Datos"
      TabPicture(2)   =   "fromMainGeneradorListas.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   3135
         Left            =   -74880
         TabIndex        =   9
         Top             =   480
         Width           =   8175
         Begin VB.Label lblHoraUltimaActual 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   3600
            TabIndex        =   15
            Top             =   1200
            Width           =   1935
         End
         Begin VB.Label lblCantArchivos 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   3600
            TabIndex        =   14
            Top             =   1680
            Width           =   1935
         End
         Begin VB.Label lblUltimaActualizacion 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   3600
            TabIndex        =   13
            Top             =   720
            Width           =   1935
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Cantidad de archivos MP3"
            Height          =   240
            Left            =   240
            TabIndex        =   12
            Top             =   1680
            Width           =   2355
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Última hora de actualización"
            Height          =   240
            Left            =   240
            TabIndex        =   11
            Top             =   1200
            Width           =   2520
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Última fecha de actualización"
            Height          =   240
            Left            =   240
            TabIndex        =   10
            Top             =   720
            Width           =   2850
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   3135
         Left            =   -74760
         TabIndex        =   2
         Top             =   480
         Width           =   7935
         Begin VB.CommandButton botGenerarNuevaLista 
            Caption         =   "&Generar"
            Height          =   375
            Left            =   6120
            TabIndex        =   24
            Top             =   2040
            Width           =   1575
         End
         Begin ComctlLib.ProgressBar ProgressBar2 
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   2880
            Visible         =   0   'False
            Width           =   7455
            _ExtentX        =   13150
            _ExtentY        =   450
            _Version        =   327682
            Appearance      =   1
         End
         Begin VB.CommandButton botGuardarComoSalida 
            Caption         =   "&Guardar como"
            Height          =   375
            Left            =   6120
            TabIndex        =   21
            Top             =   1320
            Width           =   1575
         End
         Begin VB.TextBox txtListaSalida 
            Height          =   360
            Left            =   0
            TabIndex        =   20
            Top             =   1320
            Width           =   5775
         End
         Begin VB.TextBox txtListaEntrada 
            Height          =   360
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   18
            Top             =   480
            Width           =   5775
         End
         Begin VB.CommandButton botExaminarEntrada 
            Caption         =   "&Examinar"
            Height          =   375
            Left            =   6120
            TabIndex        =   17
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label lblBarraProgreso2 
            AutoSize        =   -1  'True
            Caption         =   "Generando nueva lista"
            Height          =   240
            Left            =   120
            TabIndex        =   23
            Top             =   2520
            Visible         =   0   'False
            Width           =   2025
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "I&ngrese nuevo nombre de la lista (SALIDA)"
            Height          =   240
            Left            =   0
            TabIndex        =   19
            Top             =   1080
            Width           =   3780
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "&Ingrese lista a corregir (ENTRADA)"
            Height          =   240
            Left            =   0
            TabIndex        =   16
            Top             =   240
            Width           =   3135
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   3135
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   7935
         Begin VB.Data Data1 
            Caption         =   "Data1"
            Connect         =   "Access"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   405
            Left            =   360
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   ""
            Top             =   1320
            Visible         =   0   'False
            Width           =   2415
         End
         Begin ComctlLib.ProgressBar ProgressBar1 
            Height          =   255
            Left            =   360
            TabIndex        =   7
            Top             =   2640
            Visible         =   0   'False
            Width           =   7455
            _ExtentX        =   13150
            _ExtentY        =   450
            _Version        =   327682
            Appearance      =   1
         End
         Begin VB.CommandButton botGenerarBaseDeDatos 
            Caption         =   "&Generar"
            Height          =   375
            Left            =   6240
            TabIndex        =   6
            Top             =   1320
            Width           =   1575
         End
         Begin VB.CommandButton botExaminar 
            Caption         =   "&Examinar"
            Height          =   375
            Left            =   6240
            TabIndex        =   5
            Top             =   480
            Width           =   1575
         End
         Begin VB.TextBox txtArchivoOrigen 
            Height          =   360
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   480
            Width           =   5775
         End
         Begin VB.Label lblBarraProgreso 
            AutoSize        =   -1  'True
            Caption         =   "Actualizando base de datos"
            Height          =   240
            Left            =   360
            TabIndex        =   8
            Top             =   2400
            Visible         =   0   'False
            Width           =   2505
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "A&rchivo origen"
            Height          =   240
            Left            =   240
            TabIndex        =   3
            Top             =   240
            Width           =   1290
         End
      End
   End
   Begin VB.Menu mnuFormulario 
      Caption         =   "&Formulario"
      Begin VB.Menu mnuFormularioSalir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu mnuAcercaDe 
      Caption         =   "&Acerca de ..."
   End
End
Attribute VB_Name = "frmGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub botExaminar_Click()
     CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    'Busco archivos
    Me.CommonDialog1.ShowOpen
    Me.txtArchivoOrigen.Text = Me.CommonDialog1.filename
ErrHandler:
    ' El usuario ha hecho clic en el botón Cancelar
    Exit Sub

End Sub

Private Sub botExaminarEntrada_Click()
     CommonDialog1.CancelError = True
    On Error GoTo ErrHandler

    'Busco archivos
    Me.CommonDialog1.ShowOpen
    Me.txtListaEntrada.Text = Me.CommonDialog1.filename
    
ErrHandler:
    ' El usuario ha hecho clic en el botón Cancelar
    Exit Sub
End Sub

Private Sub botGenerarBaseDeDatos_Click()
    Dim miArchivo As String
    'determino si se seleccionó un archivo
    If Len(txtArchivoOrigen) > 0 Then
        'tengo que determinar si el archivo existe
        miArchivo = Dir(txtArchivoOrigen.Text)
        If miArchivo <> "" Then
            'tengo que verificar si el archivo es de tipo texto
           ' If funTipoTexto(txtArchivoOrigen) Then
                'permito generar lista
                If mFunMensaje(4, 4) Then
                    subConfiguroBarraProgreso
                    subGeneroBaseDeDatos miArchivo
                    subGraboInformacionProceso
                    mSubMensaje 4, 5
                    Me.lblBarraProgreso.Visible = False
                    Me.ProgressBar1.Visible = False
                End If
            'Else
                'El archivo no es de tipo texto
            '    mSubMensaje 4, 3
            'End If
        Else
            'El archivo no existe
            mSubMensaje 4, 2
        End If
    Else
        'Debe de ingresar un archivo de origen
        mSubMensaje 4, 1
    End If

End Sub

Private Sub subGeneroBaseDeDatos(archivo As String)
    'Recorro el archivo y genero base de datos
    Dim LineaTexto As String
    Dim tot As Integer
    Open archivo For Input As #1   ' Abre el archivo.
    tot = 0
    Do While Not EOF(1) ' Repite el bucle hasta el final del archivo.
        Line Input #1, LineaTexto   ' Lee el carácter en la variable.
        subProcesoInformacion LineaTexto
        tot = tot + 1
        subIncrementoBarraProgreso tot
    Loop
    Close #1    ' Cierra el archivo.
End Sub

Private Sub subProcesoInformacion(linea As String)
    'Descompongo la linea en dos partes:
    '1 ruta de acceso al archivo
    '2 nombre de archivo incluyendo extención
    Dim nombreArchivo As String
    Dim rutaAcceso As String
    
    Dim archivoListas As Recordset
    Set archivoListas = tbARCHIVOS
    
    nombreArchivo = funObtengoNombreArchivo(linea)
    rutaAcceso = funObtengoRutaAcceso(linea)
    'grabo información en archivo
    archivoListas.Index = "pkArchivo"
    archivoListas.Seek "=", nombreArchivo
    If archivoListas.NoMatch Then
        'creo un nuevo registro
        archivoListas.AddNew
            archivoListas("nomArchivo") = nombreArchivo
            archivoListas("ubicacionArchivo") = rutaAcceso
        archivoListas.Update
    Else
        'actualizo
        archivoListas.Edit
            archivoListas("ubicacionArchivo") = rutaAcceso
        archivoListas.Update
    End If
    Set archivoListas = Nothing
End Sub

Private Function funObtengoRutaAcceso(linea As String) As String
    'Obtengo la ruta del del archvivo
    'Para eso comienzo a recorrer la linea desde la izquirda hasta encontrar el
    'caracter de \, el cual indica que comienza la ruta de acceso.
    Dim aux As String
    Dim cursor As Integer
    Dim comienzaRuta As Boolean
    comienzaRuta = False
    
    cursor = Len(linea)
    Do While cursor > 0
        caracter = Mid(linea, cursor, 1)
        If caracter = "\" Then
            comienzaRuta = True
        End If
        If comienzaRuta Then
            aux = caracter & aux
        End If
        cursor = cursor - 1
    Loop
    funObtengoRutaAcceso = aux
End Function

Private Function funObtengoNombreArchivo(linea As String) As String
    'Obtengo el nombre del archvivo
    'Para eso comienzo a recorrer la linea desde la izquirda hasta encontrar el
    'caracter de \, el cual indica que comienza la ruta de acceso.
    Dim aux As String
    Dim cursor As Integer
    
    cursor = Len(linea)
    Do While cursor > 0
        caracter = Mid(linea, cursor, 1)
        If caracter <> "\" Then
            aux = caracter & aux
        Else
            Exit Do
        End If
        cursor = cursor - 1
    Loop
    funObtengoNombreArchivo = aux
End Function

Private Sub subConfiguroBarraProgreso()
    'configuro control de barra de progreso
    
    'aviso de que el proceso va a demorar
    Me.ProgressBar1.Min = 0
    Me.ProgressBar1.Max = funObtengoCantidadDeLineas(Me.txtArchivoOrigen)
    Me.ProgressBar1.Value = 0
    Me.ProgressBar1.Visible = True
    Me.lblBarraProgreso.Visible = True
End Sub

Private Sub subConfiguroBarraProgresoListas()
    'configuro control de barra de progreso
    
    'aviso de que el proceso va a demorar
    Me.ProgressBar2.Min = 0
    Me.ProgressBar2.Max = funObtengoCantidadDeLineas(Me.txtListaEntrada)
    Me.ProgressBar2.Value = 0
    Me.ProgressBar2.Visible = True
    Me.lblBarraProgreso2.Visible = True
End Sub


Private Sub subIncrementoBarraProgreso(valor As Integer)
    'Incremento la barra de progreso
    Me.ProgressBar1.Value = valor
End Sub

Private Function funObtengoCantidadDeLineas(archivo As String) As Integer
    'Recorro el archivo de texto y cuento las cantidad de lineas
    Dim tot As Integer
    Dim LíneaTexto
    Open archivo For Input As #1   ' Abre el archivo.
    tot = 0
    Do While Not EOF(1) ' Repite el bucle hasta el final del archivo.
        Line Input #1, LíneaTexto   ' Lee el carácter en la variable.
        tot = tot + 1
    Loop
    Close #1    ' Cierra el archivo.
    funObtengoCantidadDeLineas = tot
End Function

Private Sub botGenerarNuevaLista_Click()
    Dim miArchivoEntrada As String
    Dim miArchivoSalida As String
    'determino si se seleccionó un archivo
    If Len(txtListaEntrada.Text) > 0 Then
        'tengo que determinar si el archivo existe
        miArchivoEntrada = Dir(txtListaEntrada.Text)
        If miArchivoEntrada = "" Then
            'El archivo no existe
            mSubMensaje 4, 2
        Else
            If Len(Me.txtListaSalida.Text) > 0 Then
                miArchivoSalida = Dir(Me.txtListaSalida)
                'determino si el archivo salida ya existe
                If miArchivoSalida <> "" Then
                    If mFunMensaje(4, 7) Then
                        'pregunto si quiero continuar
                        If mFunMensaje(4, 8) Then
                            subGeneroNuevaLista miArchivoEntrada, miArchivoSalida
                        End If
                    End If
                Else
                    'pregunto si quiero continuar
                    If mFunMensaje(4, 8) Then
                        subGeneroNuevaLista miArchivoEntrada, miArchivoSalida
                    End If
                End If
                
            Else
                'Debe de ingresar archivo de salida
                mSubMensaje 4, 6
            End If
        End If
    Else
        'Debe de ingresar un archivo de origen
        mSubMensaje 4, 1
    End If
End Sub

Private Sub subGeneroNuevaLista(archEntrada As String, archSalida As String)
    
    'Reccorro lista inicial
    'Obtengo nombre archivo
    'Busco ubicación actual
    'Genero nueva linea de archivo salida
    Dim LineaTexto As String
    Dim nombreMp3 As String
    Dim tot As Integer
    
    subConfiguroBarraProgresoListas
    
    Open archEntrada For Input As #1   ' Abre el archivo entrada.
    Open archSalida For Output As #2     'Abre el archivo salida
    tot = 0
    Do While Not EOF(1) ' Repite el bucle hasta el final del archivo.
        Line Input #1, LineaTexto   ' Lee el carácter en la variable
        nombreMp3 = funObtengoNombreArchivo(LineaTexto)
        subGeneroLinea nombreMp3
        tot = tot + 1
    Loop
    Close #1    ' Cierra el archivo.
    Close #2
    
    mSubMensaje 4, 9
    Me.lblBarraProgreso2.Visible = False
    Me.ProgressBar2.Visible = False

End Sub

Private Sub subGeneroLinea(archivo As String)
    'busco archivo
    Dim tbArch As Recordset
    Set tbArch = tbARCHIVOS
    tbArch.Index = "pkArchivo"
    tbArch.Seek "=", archivo
    If Not tbArch.NoMatch Then
        'el archivo existe en la base de datos
        'grabo archivo
        subGraboArchivo tbArch("NomArchivo"), tbArch("ubicacionArchivo")
    End If
    Set tbArch = Nothing
End Sub

Private Sub subGraboArchivo(arch As String, ruta As String)
    Dim linea As String
    linea = ruta & arch
    Print #2, linea
End Sub

Private Sub botGuardarComoSalida_Click()
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    'Busco archivos
    Me.CommonDialog1.ShowSave
    Me.txtListaSalida = Me.CommonDialog1.filename
ErrHandler:
    ' El usuario ha hecho clic en el botón Cancelar
    Exit Sub

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Quito de memoria
    Set frmGeneral = Nothing
End Sub

Private Sub mnuAcercaDe_Click()
    frmAcercaDe.Show 1
End Sub

Private Sub mnuFormularioSalir_Click()
    'Cierro formulario
    Unload Me
End Sub

Private Sub subMuestroInfoBaseDeDatos()
    'Muestro información acerca de la base de datos
    
    Dim tbInf As Recordset
    Set tbInf = tbINFO
    
    tbInf.Index = "pkClave"
    tbInf.Seek "=", 1
    If Not tbInf.NoMatch Then
        Me.lblUltimaActualizacion = tbInf(1)
        Me.lblHoraUltimaActual = tbInf(2)
        Me.lblCantArchivos = tbInf(3)
    End If
    Set tbInf = Nothing
End Sub

Private Sub subGraboInformacionProceso()
    'Grabo información que sirve para consultar datos acerca de l
    'base de datos
    Dim tbInf As Recordset
    Dim tbArch As Recordset
    Set tbInf = tbINFO
    Set tbArch = tbARCHIVOS
    
    tbInf.Index = "pkClave"
    tbInf.Seek "=", 1
    If tbInf.NoMatch Then
        tbInf.AddNew
        tbInf(0) = 1
    Else
        tbInf.Edit
    End If
        tbInf(1) = Date
        tbInf(2) = Str(Time)
        tbArch.MoveLast
        tbInf(3) = tbArch.RecordCount
    tbInf.Update
    Set tbInf = Nothing
    Set tbArch = Nothing
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    If SSTab1.Tab = 2 Then
        subMuestroInfoBaseDeDatos
    End If

End Sub
