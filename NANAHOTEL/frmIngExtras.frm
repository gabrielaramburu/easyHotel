VERSION 5.00
Object = "{EB8C3860-8603-11D0-A35D-7062E4000000}#1.1#0"; "VCBND.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmIngExtras 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ingreso de extras"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   11910
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin Hotel_Nana.gaHOTELtitular gaHOTELtitular1 
      Height          =   1335
      Left            =   120
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   0
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   2355
      BackColor       =   -2147483633
   End
   Begin VB.Frame Frame3 
      Caption         =   "Datos del gasto"
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   11655
      Begin VB.TextBox txtCantidad 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   1440
         MaxLength       =   5
         TabIndex        =   10
         Top             =   1770
         Width           =   1215
      End
      Begin VB.TextBox txtBoleta 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   1440
         MaxLength       =   5
         TabIndex        =   4
         Top             =   795
         Width           =   1215
      End
      Begin VB.ComboBox cboMoneda 
         Height          =   360
         ItemData        =   "frmIngExtras.frx":0000
         Left            =   5760
         List            =   "frmIngExtras.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   330
         Width           =   3135
      End
      Begin VB.TextBox txtCodArt 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   1440
         MaxLength       =   5
         TabIndex        =   6
         Top             =   1275
         Width           =   1215
      End
      Begin VB.CommandButton botAyudaArt 
         Caption         =   "?"
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
         Left            =   9120
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1275
         Width           =   495
      End
      Begin VB.TextBox txtDescripcionArt 
         Height          =   375
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1275
         Width           =   6135
      End
      Begin VB.CommandButton botEliminarGasto 
         Caption         =   "Bo&rrar"
         Height          =   375
         Left            =   9960
         TabIndex        =   17
         Top             =   360
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton botCancelarMod 
         Caption         =   "Cancelar mo&d."
         Height          =   375
         Left            =   9960
         TabIndex        =   18
         Top             =   840
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton botConfirmarGasto 
         Caption         =   "Con&firmar "
         Height          =   375
         Left            =   9960
         TabIndex        =   19
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox txtImporte 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   6720
         MaxLength       =   5
         TabIndex        =   16
         Top             =   1763
         Width           =   975
      End
      Begin VB.ComboBox cboPuntoVenta 
         Height          =   360
         ItemData        =   "frmIngExtras.frx":0004
         Left            =   5760
         List            =   "frmIngExtras.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   825
         Width           =   3135
      End
      Begin VcBndCtl.VcCalCombo FechaGasto 
         Height          =   375
         Left            =   1440
         TabIndex        =   2
         Top             =   300
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _0              =   $"frmIngExtras.frx":0008
         _1              =   $"frmIngExtras.frx":0411
         _2              =   $"frmIngExtras.frx":081A
         _3              =   "-A@@@)f@@@E@@@A@@@@@@@@@'@@@E@@@A@@@@@@@@@)h@@@E@@@A@@@@@@@@@,456D"
         _count          =   4
         _ver            =   2
      End
      Begin VB.Label lblSignoMon 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "lblSignoMon"
         Height          =   240
         Left            =   5400
         TabIndex        =   26
         Top             =   1830
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "F. &gasto"
         Height          =   240
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   720
      End
      Begin VB.Label Mo 
         AutoSize        =   -1  'True
         Caption         =   "&Moneda"
         Height          =   240
         Left            =   4560
         TabIndex        =   11
         Top             =   360
         Width           =   750
      End
      Begin VB.Label Label7 
         Caption         =   "B&oleta"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   855
         Width           =   615
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "&Cantidad"
         Height          =   240
         Left            =   240
         TabIndex        =   9
         Top             =   1830
         Width           =   810
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "&Importe"
         Height          =   240
         Left            =   4560
         TabIndex        =   15
         Top             =   1830
         Width           =   675
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "&Articulo"
         Height          =   240
         Left            =   240
         TabIndex        =   5
         Top             =   1335
         Width           =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "&Punto venta"
         Height          =   240
         Left            =   4560
         TabIndex        =   13
         Top             =   855
         Width           =   1050
      End
   End
   Begin Hotel_Nana.gaHOTELbarra gaHOTELbarra1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   7605
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   582
   End
   Begin VB.CommandButton botGrabar 
      Height          =   375
      Left            =   9240
      Picture         =   "frmIngExtras.frx":0C23
      Style           =   1  'Graphical
      TabIndex        =   22
      Tag             =   "Aceptar"
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton botCancelar 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   10560
      Picture         =   "frmIngExtras.frx":14D9
      Style           =   1  'Graphical
      TabIndex        =   24
      Tag             =   "Cancelar"
      Top             =   7200
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid gmsGastos 
      Height          =   3135
      Left            =   120
      TabIndex        =   21
      Top             =   3960
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   5530
      _Version        =   393216
      Cols            =   10
      Enabled         =   -1  'True
      FocusRect       =   2
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      FormatString    =   $"frmIngExtras.frx":1D9B
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Gas&tos ingresados "
      Height          =   240
      Left            =   120
      TabIndex        =   20
      Top             =   3720
      Width           =   1755
   End
   Begin VB.Menu mnuFormulario 
      Caption         =   "&Formulario"
      Begin VB.Menu mnuFormularioAceptar 
         Caption         =   "Aceptar"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuFormularioCancelar 
         Caption         =   "Cancelar"
      End
      Begin VB.Menu div1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFormularioConfirmar 
         Caption         =   "Confirmar"
         Shortcut        =   {F9}
      End
   End
   Begin VB.Menu mnuBuscar 
      Caption         =   "&Buscar"
      Begin VB.Menu mnuBuscarArticulos 
         Caption         =   "Artículos..."
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmIngExtras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private hab_cuenta As Long
Private modifico_gasto As Boolean

Private Sub botAyudaArt_Click()
    Me.txtCodArt.Text = mFunBusqueda(6)
End Sub

Private Sub botGrabar_Click()
    If gmsGastos.Rows > 2 Then
        'pregunta si se confirma la operción
        If mFunMensaje(4, 9) Then
            grabo_cuentas
            'grabo bitacoras
            GraboBitacora "Hab. " & hab_cuenta
            Unload Me
            frmIngHabitacion.Show 1
        End If
    Else
        'no ha ingresado gastos
        mSubMensaje 4, 10
    End If
End Sub

Private Sub cboMoneda_Click()
    'Muestro el tipo de moneda con la cual estoy trabajando.
    If Me.cboMoneda.ListIndex > -1 Then
        Me.lblSignoMon.Caption = mFunObtengoSignoMoneda(Me.cboMoneda.ItemData(Me.cboMoneda.ListIndex))
        Me.lblSignoMon.Visible = True
    Else
        Me.lblSignoMon.Visible = False
    End If
End Sub

Private Sub Form_Load()
    'Apariencia formulario
    mSubConfiguroFuentesControlesSistema Me
    
    modifico_gasto = False
        
    'cargo combos
    carga_tipo_moneda cboMoneda
    cboMoneda.ListIndex = -1
    carga_punto_venta cboPuntoVenta
    cboPuntoVenta.ListIndex = -1
    
    'obtengo habitacion
    hab_cuenta = Val(frmIngHabitacion.txtNroHab.Text)
    'muestro cabezal
    Me.gaHOTELtitular1.CaminoBaseDeDatos = vardir
    Me.gaHOTELtitular1.NumeroHabitacion = hab_cuenta
    'armo cabezal grilla
    cabezal_grilla
    'el control de descripción de articulo no esta disponible
    mSubBloqueoControlFormulario Me.txtDescripcionArt, True
End Sub

Private Sub botCancelar_Click()
    Dim res As Integer
    If gmsGastos.Rows > 2 Then  'si hay gastos
        'confirmación de que realmente desea salir
        If mFunMensaje(4, 11) Then
            Unload Me
            frmIngHabitacion.Show 1
        End If
    Else
        Unload Me
        frmIngHabitacion.Show 1
    End If
End Sub

Private Sub botCancelarMod_Click()
    mSub_limpio_controles_formulario frmIngExtras
    subLimpioGrilla
    cambio_botones
End Sub

Private Sub botEliminarGasto_Click()
    mSub_limpio_controles_formulario frmIngExtras
    cambio_botones
    gmsGastos.RemoveItem gmsGastos.Row
End Sub

Private Sub botConfirmarGasto_Click()
    If valido_datos Then
        muestro_gasto_en_grilla
        mSub_limpio_controles_formulario frmIngExtras
        cambio_botones
    End If
End Sub

Private Sub cambio_botones()
    botCancelarMod.Visible = False
    botEliminarGasto.Visible = False
    gmsGastos.Enabled = True
    FechaGasto.SetFocus
    modifico_gasto = False
    Me.botConfirmarGasto.Caption = "Con&firmar"
End Sub

Private Sub muestro_gasto_en_grilla()
    Dim cadena As String
    Dim cadena_importes As String
    
    Dim total As Double
    
    'calculo totales para desplegar
    If Val(txtCantidad.Text) <> 0 Then
        total = Val(txtImporte.Text) * Val(txtCantidad.Text)
    Else
        total = Val(txtImporte.Text)
    End If
    
    If cboMoneda.ItemData(cboMoneda.ListIndex) = 0 Then
        'moneda nacional
        cadena_importes = _
        txtImporte.Text & _
        Chr(9) & _
        total & _
        Chr(9) & _
        Chr(9) & _
        Chr(9)
    Else
        'dolares
        cadena_importes = _
        Chr(9) & _
        Chr(9) & _
        txtImporte.Text & _
        Chr(9) & _
        total & _
        Chr(9)
    End If
        
    cadena = Chr(9) & _
    FechaGasto.Value & _
    Chr(9) & _
    txtCodArt.Text & _
    Chr(9) & _
    txtDescripcionArt.Text & _
    Chr(9) & _
    txtCantidad.Text & _
    Chr(9) & _
    cadena_importes & _
    cboPuntoVenta.ItemData(cboPuntoVenta.ListIndex) & _
    Chr(9) & _
    txtBoleta.Text & _
    Chr(9) & _
    cboMoneda.ItemData(cboMoneda.ListIndex)

    If modifico_gasto = False Then
        gmsGastos.AddItem cadena
    Else
        gmsGastos.RemoveItem gmsGastos.Row
        gmsGastos.AddItem cadena
    End If
End Sub

Private Sub grabo_cuentas()
    'Recorro la grilla y grabo gastos, asignándole un número correlativo a cada
    'gasto.
    
    Dim total As Double
    Dim nrocorr As Long
    Dim i As Integer
    Dim moneda As Byte
    Dim cantidad As Double
    Dim importe As Double
    
    i = 2
    Do While i < gmsGastos.Rows
        gmsGastos.Row = i
        nrocorr = obtengo_proximo_gasto(FechaGasto.Value)
        
        tbCUENTAS.AddNew
            'Clave primaria
            tbCUENTAS("habitacion_cuenta") = hab_cuenta
            tbCUENTAS("nrocorr_cuenta") = nrocorr
            ColGri 1    'fecha
                tbCUENTAS("fechagasto_cuenta") = gmsGastos.Text
        
        
            ColGri 2    'artículo
                tbCUENTAS("articulo_cuenta") = gmsGastos.Text
        
            ColGri 4    'cantidad
                tbCUENTAS("cantidad_cuenta") = gmsGastos.Text
                cantidad = gmsGastos.Text
                
            ColGri 5    'm/n
                If Val(gmsGastos.Text) <> 0 Then
                    importe = gmsGastos.Text
                End If
        
            ColGri 7    'dolares
                If Val(gmsGastos.Text) <> 0 Then
                    importe = gmsGastos.Text
                End If
        
            ColGri 9    'punto de venta
                tbCUENTAS("puntoventa_cuenta") = gmsGastos.Text
        
            ColGri 10   'boleta
                tbCUENTAS("boleta_cuenta") = gmsGastos.Text
        
            ColGri 11   'moneda
                tbCUENTAS("moneda_cuenta") = gmsGastos.Text
                moneda = gmsGastos.Text
            
            If cantidad <> 0 Then   'cantidad
                total = importe * cantidad
            Else
                total = importe
            End If
                
            'moneda nacional
            If moneda = 0 Then
                tbCUENTAS("importe_dolares_cuenta") = 0
                tbCUENTAS("total_dolares_cuenta") = 0
                
                tbCUENTAS("importe_mnacional_cuenta") = importe
                tbCUENTAS("total_mnacional_cuenta") = total
            Else
            'dolares
                tbCUENTAS("importe_mnacional_cuenta") = 0
                tbCUENTAS("total_mnacional_cuenta") = 0
                
                tbCUENTAS("importe_dolares_cuenta") = importe
                tbCUENTAS("total_dolares_cuenta") = total
            End If
            
            tbCUENTAS("titular_cuenta") = busco_titular_hab2(hab_cuenta, "extra")
        tbCUENTAS.Update
        i = i + 1
    Loop
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Descargo formulario de memoria
    Set frmIngExtras = Nothing
End Sub

Private Sub gmsGastos_DblClick()
    Dim filaSeleccionada As Integer
    Dim colSeleccionada As Integer
       
    'Cuando ejecuto el procedimiento marco_celdas_grilla pierdo la referencia de la fila
    'donde se realizo el doble click, por ese motivo almaceno el valor de la fila en la
    'variable filaSeleccionada. (la variable colSeleccionada es para mantener choerencia).
    filaSeleccionada = gmsGastos.Row
    colSeleccionada = 1

    'verifico que no sea la celda vacía
    gmsGastos.col = 1
    If gmsGastos.Text <> "" Then
        'habilito eliminacion y modificación
        botCancelarMod.Visible = True
        botEliminarGasto.Visible = True
        botConfirmarGasto.Caption = "Con&firmar mod."
        
        'como solo puede haber un fila seleccionada por vez, es necesario desmarcar todas
        'las filas de la grilla antes de marcar un nueva.
        subLimpioGrilla
        
        'posiciono lugar original donde hice el doble click para poder marcar la fila
        gmsGastos.Row = filaSeleccionada
        gmsGastos.col = 1
        'marco fila seleccionada
        marco_celdas_grilla Me.gmsGastos, 1, Me.gmsGastos.Cols - 1, Me.gmsGastos.Row, Me.gmsGastos.Row
        gmsGastos.CellBackColor = mSisColor_15FilaSeleccionada
        gmsGastos.CellForeColor = mSisColor_19FilaSeleccionadaTexto
        
        'posiciono lugar original donde hice el doble click para obtener datos de la grilla
        'los cuales se mostraran en el formulario.
        gmsGastos.Row = filaSeleccionada
        gmsGastos.col = 1
        cargo_datos_formulario
        modifico_gasto = True
        
        'no permito utilizar la grilla si estoy modificando gastos
        'gmsGastos.Enabled = False
        
        'le doy el focus a el primer control de ingreso de datos
        Me.FechaGasto.SetFocus
    End If
End Sub

Private Sub gmsGastos_KeyPress(KeyAscii As Integer)
    'Tambien permito seleccionar un elemento de la grilla con la tecla enter
    If KeyAscii = vbKeyReturn Then
        'simulo que se hizo un doble click con el mouse
        gmsGastos_DblClick
    End If
End Sub

Private Sub cargo_datos_formulario()
    'Cargo los datos desde la grilla al formulario para poderlos modificar o borrar.
    
    ColGri 1    'fecha
        Me.FechaGasto.Text = gmsGastos.Text
        
    ColGri 2    'artículo
        txtCodArt.Text = gmsGastos.Text
        
    ColGri 3    'descripcion
        If busco_articuloTF(Val(txtCodArt.Text)) Then
            txtDescripcionArt.Text = tbARTICULOS("descriarticulo")
        End If
    ColGri 4    'cantidad
        txtCantidad.Text = gmsGastos.Text
        
    ColGri 5    'm/n
        If Val(gmsGastos.Text) <> 0 Then
            txtImporte.Text = gmsGastos.Text
        End If
        
    ColGri 7    'dolares
        If Val(gmsGastos.Text) <> 0 Then
            txtImporte.Text = gmsGastos.Text
        End If
        
    ColGri 9    'punto de venta
        posiciono_combo cboPuntoVenta, gmsGastos.Text
        
    ColGri 10   'boleta
        txtBoleta = gmsGastos.Text
        
    ColGri 11   'moneda
        posiciono_combo cboMoneda, gmsGastos.Text
    
End Sub

Private Sub ColGri(col As Byte)
    gmsGastos.col = col
End Sub

Private Sub cabezal_grilla()
     
    gmsGastos.FormatString = _
    " |Fecha Gasto" & _
    "| Código" & _
    "| Decripción Artículo                                                  " & _
    "| Cant.  " & _
    "| P.u. " & gblSignoMonedaNacional & "     " & _
    "| Total             " & _
    "| P.u. " & gblSignoDolares & "     " & _
    "| Total            " & _
    "| PtoVta " & _
    "| Bolteta" & _
    "| Moneda"
    
    'Las últimas tres columnas no se muestran
    gmsGastos.ColWidth(9) = 0
    gmsGastos.ColWidth(10) = 0
    gmsGastos.ColWidth(11) = 0
End Sub

Private Sub subLimpioGrilla()
    'Desmarco la fila seleccionada. Como no se cual fila es, inicializo toda la
    'grilla a sus propiedades iniciales.
    marco_celdas_grilla gmsGastos, 1, gmsGastos.Cols - 1, 1, gmsGastos.Rows - 1
    gmsGastos.CellBackColor = gmsGastos.BackColor
    gmsGastos.CellForeColor = gmsGastos.ForeColor
End Sub

Private Sub txtBoleta_KeyPress(KeyAscii As Integer)
    ValidoNum KeyAscii, False, False
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
    ValidoNum KeyAscii, True, True
End Sub

Private Sub txtCodArt_Change()
    'Busco artículo
    If busco_articuloTF(Val(txtCodArt.Text)) Then
        txtDescripcionArt.Text = tbARTICULOS("descriarticulo")
        posiciono_combo cboMoneda, tbARTICULOS("MonedaArticulo")
        posiciono_combo cboPuntoVenta, tbARTICULOS("PuntoVentaArticulo")
        txtImporte.Text = Val(tbARTICULOS("PrecioArticulo"))
    Else
        'inicializo para nuevo artículo
        txtDescripcionArt.Text = Empty
        txtImporte.Text = Empty
        cboMoneda.ListIndex = -1
        cboPuntoVenta.ListIndex = -1
    End If
End Sub

Private Sub txtCodArt_KeyPress(KeyAscii As Integer)
    ValidoNum KeyAscii, False, False
End Sub

Private Sub txtCodArt_LostFocus()
    'asistencia a usuarios
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub txtImporte_KeyPress(KeyAscii As Integer)
    ValidoNum KeyAscii, True, True
End Sub

Private Function valido_datos()
    valido_datos = True
    If Not IsDate(FechaGasto.Text) Then
        'formato de fecha erronea
        mSubMensaje 3, 1
        FechaGasto.SetFocus
        valido_datos = False
        Exit Function
    End If

    'Me parece que puede ser necesario configurar esa opción    gabriel
    If Val(txtBoleta.Text) = 0 Then
        'boleta no puede estar vacía
        mSubMensaje 4, 12
        txtBoleta.SetFocus
        valido_datos = False
        Exit Function
    End If
    
    If busco_articuloTF(Val(txtCodArt.Text)) = False Then
        'el articulo no existe
        mSubMensaje 4, 13
        txtCodArt.SetFocus
        valido_datos = False
        Exit Function
    End If
    
    If Val(txtImporte.Text) = 0 Then
        'el importe no puede ser 0
        mSubMensaje 4, 14
        txtImporte.SetFocus
        valido_datos = False
        Exit Function
    End If
    
    If Val(txtCantidad.Text) = 0 Then
        'la cantidad no puede ser 0
        mSubMensaje 4, 15
        txtCantidad.SetFocus
        valido_datos = False
        Exit Function
    End If
End Function

Private Sub mnuBuscarArticulos_Click()
    'Esta opción equivale a presionar el boton de ayuda
    botAyudaArt_Click
End Sub

Private Sub mnuFormularioAceptar_Click()
    'Esta opción equivale a presionar el botón de aceptar o a la tecla F12
    botGrabar_Click
End Sub

Private Sub mnuFormularioCancelar_Click()
    'Esta opción equivale al botón cancelar
    botCancelar_Click
End Sub

Private Sub mnuFormularioConfirmar_Click()
    'Esta opción equivale al boton Confirmar gasto o a la tecla F9
    botConfirmarGasto_Click
End Sub

'****************************************************
'*
'*  Asistencia a usuarios
'*
'****************************************************

Private Sub FechaGasto_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 34
End Sub

Private Sub txtBoleta_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 35
End Sub

Private Sub txtCodArt_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 36
End Sub

Private Sub txtCantidad_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 37
End Sub

Private Sub cboMoneda_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 38
End Sub

Private Sub cboPuntoVenta_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 39
End Sub

Private Sub txtImporte_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 40
End Sub

Private Sub botEliminarGasto_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 41
End Sub

Private Sub botCancelarMod_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 42
End Sub

Private Sub botConfirmarGasto_GotFocus()
    'Como el boton de ConfirmarGasto realiza dos operaciones
    'tengo que determinar que operación está realizando en el momento de recivir el focus
    If Trim(Me.botConfirmarGasto.Caption) = "Con&firmar" Then
        'Ingreso un nuevo gasto
        mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 43
    Else
        'Modifico un gasto seleccionado
        mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 44
    End If
End Sub

Private Sub gmsGastos_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 2, 45
End Sub

Private Sub botGrabar_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 1, 2
End Sub

Private Sub botCancelar_GotFocus()
    mSubMuestroInformacionEnLinea Me.gaHOTELbarra1, 1, 3
End Sub

Private Sub botCancelar_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub botGrabar_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub gmsGastos_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub botConfirmarGasto_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub FechaGasto_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub botCancelarMod_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub botEliminarGasto_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub cboPuntoVenta_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub cboMoneda_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub txtCantidad_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub txtImporte_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub

Private Sub txtBoleta_LostFocus()
    mSubLimpioInformacionEnLinea Me.gaHOTELbarra1
End Sub


