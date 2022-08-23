Attribute VB_Name = "Facturación"
Option Explicit
                    'Estas variables se cargan con información que se muestra al imprimir la facturar
Public cabFechaEntrada As String
Public cabFechaSalida As String
Public cabCantPax As String
Public cabTipoHab As String


Public Sub mSub_cargo_cabezal_desde_documento(tipo_docu As Byte, _
                                            nro_docu As Long, _
                                            f As Form, _
                                            i As Byte)
    'Carga el cabezal del documento.
    'Se utiliza para mostrar una factura (anulación o cunsulta)
    'Se utiliza para mostrar una devolución (anulación o consulta)
    'Se utiliza para crear el cabezal de una nueva devolución a partir del cabezal
    'de una factura.
    
    If busco_documentoTF(tipo_docu, nro_docu) Then
        f.txtNom(i).Text = tbCABEZAL("nom_docu")
        f.txtDir(i).Text = tbCABEZAL("dir_docu")
        f.txtLoc(i).Text = tbCABEZAL("loc_docu")
        f.txtRuc(i).Text = tbCABEZAL("ruc_docu")
        f.txtCP(i).Text = tbCABEZAL("cp_docu")
        f.txtNroCli(i).Text = tbCABEZAL("nrocorr_docu")
        posiciono_combo f.cboPais(i), tbCABEZAL("pais_docu")
        f.fechaemi(i).Text = tbCABEZAL("fecha_docu")
        f.lblTotalGral(i).Caption = tbCABEZAL("tot_docu")
        f.lblIVAb(i).Caption = tbCABEZAL("tot_iva_basico")
        f.lblIVAm(i).Caption = tbCABEZAL("tot_iva_minimo")
        f.lblImpExento(i).Caption = tbCABEZAL("tot_exento")
        f.lblImpBasico(i).Caption = tbCABEZAL("tot_impbasico")
        f.lblImpMinimo(i).Caption = tbCABEZAL("tot_impminimo")
    End If
End Sub

Public Sub mSub_muestro_lineas_documento(tipo_docu As Byte, nro_docu As Long, f As Form)
    'Muestra las lineas de un documento.
    'En el caso de que el documento sea una factura: se utiliza para mostrar anular o consultar
    'En el caso de que el documento sea una devolución existente: se utiliza,
    'para mostrar las líneas de la devolución (anulación o consulta)
    'En el caso de que sea una nueva devolución: se utiliza para mostrar las líneas de la
    'factura a la cual se le quiere efectuar la devolución. Si se realiza la misma, estas líneas pasaran
    'a formar parte de la nueva devolución. ¿capisco?
    
    
    'Recorro las líneas del documento y las cargo a la grilla
    tbLINEAS.Index = "i_lineas"
    tbLINEAS.Seek ">=", tipo_docu, nro_docu, 1
    If Not tbLINEAS.NoMatch Then
        Do While Not tbLINEAS.EOF
            If tbLINEAS("tipo_linea") = tipo_docu _
            And tbLINEAS("nro_factura") = nro_docu Then
                f.DBGrid1(0).AddItem creo_linea
                tbLINEAS.MoveNext
            Else
                Exit Do
            End If
        Loop
    End If
End Sub

Private Function creo_linea()
    Dim linea_factura As String
    linea_factura = _
    Chr(9) & _
    tbLINEAS("fecha_linea") & _
    Chr(9) & _
    tbLINEAS("hab_linea") & _
    Chr(9) & _
    tbLINEAS("artcod_linea") & _
    Chr(9) & _
    tbLINEAS("artdes_linea") & _
    Chr(9) & _
    paso_moneda_a_desc(tbLINEAS("tipom_linea")) & _
    Chr(9) & _
    Format(tbLINEAS("artcant_linea"), "####0;;#") & _
    Chr(9) & _
    Format(tbLINEAS("artpu_linea"), "####0.00;;#") & _
    Chr(9) & _
    Format(tbLINEAS("arttotal_linea"), "#####0.00") & _
    Chr(9) & _
    Format(tbLINEAS("totalconv_linea"), "####0.00") & _
    Chr(9) & _
    tbLINEAS("hab_linea") & _
    Chr(9) & _
    tbLINEAS("nro_linea")
    creo_linea = linea_factura
End Function

Public Function mFun_obtengo_proximo_documento(digitodocu As Byte)
    Dim aux As Long
    Dim prox As Long
    Dim serie As Byte
    Dim digito As Byte
    Dim i As Byte
    Dim i2 As Byte
    Dim i3 As Byte

    Select Case digitodocu
        'Dependiendo del documento que voy a realizar, depende el campo del archivo
        'parametros. Este campo determina el próximo número de documento.
        'i= campo del numero de documento
        'i2= campo del número de serie
        'i3= digito que corresponde a cada tipo de documento
        Case 1  'contado m/n
            i = 5
            i2 = 6
            i3 = 7
                
        Case 2  'contado U$S
            i = 8
            i2 = 9
            i3 = 10
            
        Case 3  'factura m/n
            i = 11
            i2 = 12
            i3 = 13
            
        Case 4  'factura U$S
            i = 14
            i2 = 15
            i3 = 16
            
        Case 5  'dev.cdo. m/n
            i = 17
            i2 = 18
            i3 = 19
            
        Case 6  'dev.cdo. U$S
            i = 20
            i2 = 21
            i3 = 23
            
        Case 7  'dev.cre. m/n
            i = 24
            i2 = 25
            i3 = 26
            
        Case 8  'dev.cre. U$S
            i = 27
            i2 = 28
            i3 = 29
    End Select
        
    'obtengo proximo número
    prox = tbPARAMETROS(i)
    serie = tbPARAMETROS(i2)
    
    '-------------------------------------------------------------------------------------
    'NOTA: los campos donde se almacenan los dígitos de cada tipo de documento en
    'el archivo parámetros, debe de estar inicializado al momento de instalar la aplicación
    'También los campos donde se almacena el próximo número de documento, el cual debe de
    'inicializarse a 1.
    '--------------------------------------------------------------------------------------
    digito = tbPARAMETROS(i3)
            
    'formo el número de documento.
    aux = (serie * 1000000) + (digito * 100000) + prox
    mFun_obtengo_proximo_documento = aux
    
    tbPARAMETROS.Edit
        'sumo 1 al campo correspondiente, este número será utilizado por el próximo
        'documento del mismo tipo que se realize.
        tbPARAMETROS(i) = tbPARAMETROS(i) + 1
        
        'cada vez que se el número de documento llegue a 100000, empiezo nuevamente
        'desde 0, pero ahumentando la serie.
        If tbPARAMETROS(i) = 100000 Then   'ahumento serie
            tbPARAMETROS(i2) = tbPARAMETROS(i2) + 1
            tbPARAMETROS(i) = 0
        End If
    tbPARAMETROS.Update
End Function

Public Sub mSub_grabo_estado_cuentas(tipoDoc As Byte, _
                                    nrodocu As Long, _
                                    nrocli As Long, _
                                    totaldocu As Double, _
                                    fechadocu As Date)
    'Grabo la factura o devolución a la cuenta del titular de la factura
    
    'Tomo en cuenta solo las facturas o las devoluciones crédito
    If tipoDoc = 3 Or tipoDoc = 4 Or tipoDoc = 7 Or tipoDoc = 8 Then
        tbESTADO_CUENTAS.AddNew
            tbESTADO_CUENTAS("tipodoc") = tipoDoc
            tbESTADO_CUENTAS("nrodoc") = nrodocu
            'Grabo titular factura
            tbESTADO_CUENTAS("nrocli") = nrocli
            If tipoDoc = 7 Or tipoDoc = 8 Then          'devoluciones
                tbESTADO_CUENTAS("debe") = 0
                tbESTADO_CUENTAS("haber") = totaldocu
            Else
                tbESTADO_CUENTAS("debe") = totaldocu    'facturas
                tbESTADO_CUENTAS("haber") = 0
            End If
            
            tbESTADO_CUENTAS("fecha") = fechadocu
            If tipoDoc = 3 Or tipoDoc = 7 Then 'm/n
                tbESTADO_CUENTAS("moneda") = 0
            End If
            If tipoDoc = 4 Or tipoDoc = 8 Then 'dol
                tbESTADO_CUENTAS("moneda") = 1
            End If
            
        tbESTADO_CUENTAS.Update
    End If
End Sub

Public Sub mSub_cambio_cabezal(marca As Boolean, i As Integer, f As Form)
    'Es utilizado en frmFacturacion y frmDevolución.
    
    Dim color As String
    Dim x As Boolean
    
    If marca Then
        color = &H80000005  'blanco
        x = False
    Else
        color = mSisColor_18ControlesNoHabilitados  'color bloqueado
        x = True
    End If
        
    f.txtNom(i).Locked = x
    f.txtDir(i).Locked = x
    f.txtLoc(i).Locked = x
    f.txtCP(i).Locked = x
    f.txtRuc(i).Locked = x
    
    
    f.txtNom(i).TabStop = marca
    f.txtDir(i).TabStop = marca
    f.txtLoc(i).TabStop = marca
    f.txtCP(i).TabStop = marca
    f.txtRuc(i).TabStop = marca
    
    f.txtNom(i).BackColor = color
    f.txtDir(i).BackColor = color
    f.txtLoc(i).BackColor = color
    f.txtCP(i).BackColor = color
    f.txtRuc(i).BackColor = color
    
    'Este control siempre esta bloqueado
    f.fechaemi(i).BackColor = mSisColor_18ControlesNoHabilitados
    'No permito cambiar el país del cliente ya, que esto influye directmente sobre
    'la factura, debido a que el impuesto del alojamiento depende de si el
    'cliente es extranjero o nacional.
    f.cboPais(i).Locked = True
    f.cboPais(i).BackColor = mSisColor_18ControlesNoHabilitados
    f.cboPais(i).TabStop = False
End Sub

Public Sub mSub_grabo_cabezal_documento(tipo_docu As Byte, _
                                        nro_docu As Long, _
                                        fecha_docu As Date, _
                                        nom_docu As String, _
                                        dir_docu As String, _
                                        loc_docu As String, _
                                        ruc_docu As String, _
                                        cp_docu As String, _
                                        pais_docu As Byte, _
                                        nrocorr_docu As Long, _
                                        tot_docu As Double, _
                                        nro_fact_docu As Long, _
                                        cotizacion As Single, _
                                        tot_iva_basico As Double, _
                                        tot_iva_minimo As Double, _
                                        tot_exento As Double, _
                                        tot_impBasico As Double, _
                                        tot_impMinimo As Double, _
                                        porIvaMin As Single, _
                                        porIvaBas As Single, _
                                        porIvaExe As Single, _
                                        fechaEntrada As String, _
                                        fechaSalida As String, _
                                        cantPaxAlojados As String, _
                                        habAlojadoTitularFact As String)
                                        
    '----------------------------------------------------------------------------------
    'NOTA: los últimos cuatro campos de tipo string, se almacenan en campos de tipo
    'memo ya que este tipo de campo, da la facilidad de poder ingresar información
    'en varias líneas, lo que facilita la impresión de esta información al momento
    'de realizar la factura.
    'Básicamente el problema de diseño radica, en que una factura pude ser realizada
    'Esta fue la manera más rápida de contemplar esto en el momento de impresión.
    '-----------------------------------------------------------------------------------
    
    'Graba el cabezal de los documento de tipo 1 hasta 8.
    tbCABEZAL.AddNew
        tbCABEZAL("tipo_docu") = tipo_docu
        tbCABEZAL("nro_docu") = nro_docu
        tbCABEZAL("fecha_docu") = fecha_docu
        tbCABEZAL("nom_docu") = nom_docu
        tbCABEZAL("dir_docu") = dir_docu
        tbCABEZAL("loc_docu") = loc_docu
        tbCABEZAL("ruc_docu") = ruc_docu
        tbCABEZAL("cp_docu") = cp_docu
        tbCABEZAL("pais_docu") = pais_docu
        tbCABEZAL("nrocorr_docu") = nrocorr_docu
        tbCABEZAL("tot_docu") = tot_docu
        tbCABEZAL("nro_fact_docu") = nro_fact_docu
        tbCABEZAL("cotizacion") = cotizacion
        tbCABEZAL("tot_iva_basico") = tot_iva_basico
        tbCABEZAL("tot_iva_minimo") = tot_iva_minimo
        tbCABEZAL("tot_exento") = tot_exento
        tbCABEZAL("tot_impbasico") = tot_impBasico
        tbCABEZAL("tot_impminimo") = tot_impMinimo
        tbCABEZAL("porIvaMin") = porIvaMin
        tbCABEZAL("porIvaBas") = porIvaBas
        tbCABEZAL("porIvaExe") = porIvaExe
        tbCABEZAL("HabFechaEntrada") = fechaEntrada
        tbCABEZAL("HabFechaSalida") = fechaSalida
        tbCABEZAL("HabCantPaxAlojados") = cantPaxAlojados
        tbCABEZAL("habTipo") = habAlojadoTitularFact
    tbCABEZAL.Update
End Sub

Public Function mFun_realizo_lineas(tipodocu As Byte, f As Form, i As Integer)
    'Grabo lineas del documento. Puede ser una factura o una devolución
    'f= frmFacturacion o frmDevolucion
    
    Dim j As Integer
    Dim k1 As Long
    Dim k2 As Long
    Dim tipo As String
    Dim totl As Integer
    Dim FechaGasto As Date
    
    j = 2
    'inicializo variables de impresión
    subInicializoVariablesimpresion
    mFun_realizo_lineas = 0
    'mientras tengo lineas en el documento las grabo
    Do While j < f.DBGrid1(i).Rows
        f.DBGrid1(i).row = j
        If valido_linea(i, f, tipodocu) Then
            mFun_realizo_lineas = mFun_realizo_lineas + 1
            'obtengo tipo de linea
            If i = 0 Then   'primera factura:extras o extras+alojamiento
                f.DBGrid1(i).col = 12
                If f.DBGrid1(i).Text = "e" Then
                    tipo = "e"
                Else
                    tipo = "a"
                End If
            Else            'segunda factura: solo alojamiento
                tipo = "a"
            End If
                    
            tbLINEAS.AddNew
                tbLINEAS("tipo_linea") = tipodocu
                tbLINEAS("nro_factura") = f.lblNroDocu(i).Caption
                tbLINEAS("nro_linea") = j - 1
            
                'grabo el tipo de gasto
                tbLINEAS("tipogasto_linea") = tipo
                
                f.DBGrid1(i).col = 1
                tbLINEAS("fecha_linea") = f.DBGrid1(i).Text
                FechaGasto = f.DBGrid1(i).Text
                
                f.DBGrid1(i).col = 2
                tbLINEAS("hab_linea") = f.DBGrid1(i).Text
                'grabo información para imprimir
                If tipo = "a" Then
                    subCargoVariblesDeImpresion tbLINEAS("hab_linea"), 0
                End If
                
                f.DBGrid1(i).col = 3
                tbLINEAS("artcod_linea") = f.DBGrid1(i).Text
                
                
                'Grabo porcentaje de iva asigando al gasto extra
                If tipo = "e" Then  'extras
                    tbLINEAS("ArtPorIva") = mFun_PorIvaArt(Val(f.DBGrid1(i).Text))
                Else
                    'Grabo porcentaje de iva asigando al gasto alojamiento
                    tbLINEAS("ArtPorIva") = mFunTipoIvaALoja(f.cboPais(i).ItemData(f.cboPais(i).ListIndex), 2)
                End If
                
                f.DBGrid1(i).col = 4
                tbLINEAS("artdes_linea") = f.DBGrid1(i).Text
                
                f.DBGrid1(i).col = 5
                
                tbLINEAS("tipom_linea") = paso_moneda_a_codigo(f.DBGrid1(i).Text)
                
                f.DBGrid1(i).col = 6
                If tipo = "e" Then
                    tbLINEAS("artcant_linea") = f.DBGrid1(i).Text
                End If
                        
                f.DBGrid1(i).col = 7
                If tipo = "e" Then
                    tbLINEAS("artpu_linea") = f.DBGrid1(i).Text
                End If
                            
                f.DBGrid1(i).col = 8
                If tipo = "e" Then
                    tbLINEAS("arttotal_linea") = f.DBGrid1(i).Text
                End If
                
                f.DBGrid1(i).col = 9
                tbLINEAS("totalconv_linea") = f.DBGrid1(i).Text
                    
                f.DBGrid1(i).col = 10 'habitación
                k1 = f.DBGrid1(i).Text
                    
                f.DBGrid1(i).col = 11
                'nrocorr del gasto, lo necesito para eliminar los gastos después de facturar
                k2 = f.DBGrid1(i).Text 'tipo gasto
                
                If EsDevolucion(tipodocu) Then
                    'creo nuevamente los gastos.
                Else
                    'elimino los gastos facturados
                    'solo en caso que este creando las lineas de una factura
                    If tipo = "e" Then
                        elimino_gastos_extras FechaGasto, k2
                    Else
                        'verifico que tipo de línea de alojamiento quiero borrar
                        If tbLINEAS("artcod_linea") = 1 Then
                            'estoy procesando una línea de resumen
                            elimino_gastos_aloja True, k1, FechaGasto, 0
                        Else
                            'estoy procesando un alínea perteneciente a un descuento, medio día, otro
                            elimino_gastos_aloja False, k1, FechaGasto, k2
                        End If
                        
                    End If
                End If
                tbLINEAS.Update
        End If
        j = j + 1
    Loop
End Function

Private Sub subCargoVariblesDeImpresion(hab As Long, origenLlamado As Byte, _
                                        Optional tipoDoc As Byte, _
                                        Optional NroDoc As Long)
    '----------------------------------------------------------------------------------------------------------------
    'El motivo de existencia de estos campos en la base de datos, radica en varios puntos.:
    '1) para cada tipo de gasto a imprimir en el detalle de la factura tengo que incluir diferentes
    '   tipo y cantidad de información, por ejemplo al imprimir la línea de gastos de alojamiento
    '   sería correcto incluir información acerca del período que no tengo por que incluir en una
    '   línea de gastos extras.
    '2) una factura puede tener varias habitaciónes, cada una de ellas con información diferente.
    '   Esto origina que no pueda tener un cabezal en la factura, con información de solo una habitación.
    '3) la utilización de crystal report para emitir la factura me impide realizar dos tipos diferentes
    '   de recorridos dentro de un mismo listado, cos que sería mucho más facil de hacer si se implementara
    '   un rutina de impresión. En ese caso no habría problema de realizar dos cabezales diferentes.
    '   pero perdería las ventajas de realizar la factura mediante crystal.
    '
    '   Para obtener una mayor flexibilidad al momneto de decidir sobre el formato de una factura
    '   se incuye esta información, que facilita la creación de "dos cabezales" en un informe de crysal report.
    '-----------------------------------------------------------------------------------------------------------------
    'Parámetros.
    '   [origenLlamado] detetermina desde donde estoy llamado al procedimiento
    '                   0 = desde facturación
    '                   1 = desde devoluciones.
    '   Cuando realizo una factura estos datos se cargan dependiendo de la iformación
    '   que contengan las líneas de alojamiento de la factura.
    '   Pero cuando realizo una devolución lo correcto es pasar los datos de la factura
    '   a la devolución ya que no tengo no debo de obtener nuevamente los datos.
    '--------------------------------------------------------------------------------------------------------------------
    If origenLlamado = 0 Then   'desde facturación
        cabFechaEntrada = cabFechaEntrada & mFunObtengoFechaAlojaHab(hab, 0) & Chr(10)   'obtengo fecha entrada
        cabFechaSalida = cabFechaSalida & mFunObtengoFechaAlojaHab(hab, 1) & Chr(10)     'obtengo fecha salida
        cabCantPax = cabCantPax & mFunObtengoTotPaxAlojadosHab(hab) & Chr(10)            'obtengo total de pasajeros alojados en la habitación
        cabTipoHab = cabTipoHab & hab & " Suite " & busco_tipo_hab_descri(hab) & Chr(10)  'obtengo descripción hab
    End If
End Sub

Private Sub subInicializoVariablesimpresion()
    cabFechaEntrada = Empty
    cabFechaSalida = Empty
    cabCantPax = Empty
    cabTipoHab = Empty
End Sub

Public Function valido_linea(i As Integer, f As Form, tipodocu As Byte)
    'Cuando recorro la grilla de gastos, solo proceso
    'aquellas lineas que esten habilitadas para facturar. Esto ocurre en los documentos de
    'tipo factura. Si el documento es una devolución tomo todas las lineas en cuenta.
    
    'También se utiliza para recalcular el total de la factura, cuando se cambia el criterio
    'de selección de líneas. Por ese motivo se declara como pública.
    Dim color As Long
    color = mSisColor_15FilaSeleccionada
    
    valido_linea = False
    If EsDevolucion(tipodocu) = True Then
        valido_linea = True
        Exit Function
    End If
    
    If f.cboFacturar(i).ListIndex = 0 Then
        'Si es una factura y esta seleccionado el boton de todas las lineas o
        valido_linea = True
    Else
        If f.cboFacturar(i).ListIndex = 1 Then   'solo marcados
            If CLng(f.DBGrid1(i).CellBackColor) = color Then
                valido_linea = True
            End If
        Else
            If f.cboFacturar(i).ListIndex = 2 Then 'solo no marcados
                If CLng(f.DBGrid1(i).CellBackColor) <> color Then
                    valido_linea = True
                End If
            End If
        End If
    End If
End Function

Private Function EsDevolucion(tipo As Byte)
    'Determina si el tipo de documento es una devolución o no.
    EsDevolucion = False
    If tipo = 5 Or tipo = 6 Or tipo = 7 Or tipo = 8 Then
        EsDevolucion = True
    End If
End Function

Private Sub elimino_gastos_extras(FechaGasto As Date, nrocorr As Long)
    'Cuando realizo una factura, tengo que eliminar los gastos extras.
    'De esta manera no aparecerán más en el resumen de cuenta.
    
    'Al contrario de los gastos alojamiento, los gastos extras se corresponden con un registro
    'en el archivo tbCUENTAS, por ese motivo accedo directamente por clave primaria al gasto.
    tbCUENTAS.Index = "i_cuentas"
    tbCUENTAS.Seek "=", FechaGasto, nrocorr
    If Not tbCUENTAS.NoMatch Then   'existe
        tbCUENTAS.Delete
    End If
End Sub

Private Sub elimino_gastos_aloja(gastoAlojaResumen As Boolean, hab As Long, fecha As Date, _
                                nroCorrGasto As Long)
    '------------------------------------------------------------------------------------------------------
    'Cuando realizo una factura, tengo que eliminar los gastos alojamiento.
    'De esta manera no aparecerán más en el resumen de cuenta.
    '------------------------------------------------------------------------------------------------------
    'Parámetros.
    '   Entrada:    [gastoAlojaResumen]
    '               True:   línea de factura resumida
    '               False:  línea de factura no resumida
    '               Existen dos tipos de líneas de alojamiento, las líneas que corresponden
    '               a gastos de tipo 2,3 y 5 las cuales no estan resumidas, es decir estan en
    '               relación con una línea en el archivo CUENTAS_ALOJA y las líneas que
    '               corresponden a gastos de tipo 1 y 4 las cuales son producto de varias líneas
    '               de dicho archivo, es decir están resumidas.
    '
    '               [Hab]   habitación del gasto
    '
    '               [fecha] fecha del gasto.
    '               Si la línea esta resumida este valor no interesa, ya que por defecto
    '               la fecha de la línea (o líneas si hay más de una habitación por titular)
    '               se inicializa con la fecha de facturación.
    '
    '               [nroCorrGasto]  número correlativo por el cual accedo (conjuntamente con la fecha)
    '               al gasto en el archivo CUENTAS_ALOJA. Solo lo utilizo si no es una línea de resumen.
    '--------------------------------------------------------------------------------------------------------
    Dim titular As Long

    
    If gastoAlojaResumen = True Then
        'obtengo número de titular aloja
        titular = busco_titular_hab2(hab, "aloja")
        'el procedimiento es el mismo que utilizo para crear la línea,
        'con la diferencia de que acá solo recorro los gastos para una sola habitación
        
        tbCUENTAS_ALOJA.Index = "i_TipoGastos"
        tbCUENTAS_ALOJA.Seek ">=", 0, titular, hab, 0
        If Not tbCUENTAS_ALOJA.NoMatch Then
            'recorro todos los gastos de alojamiento de titular para la habitación
            Do While Not tbCUENTAS_ALOJA.EOF
                If tbCUENTAS_ALOJA("facturado") = 0 And tbCUENTAS_ALOJA("titular_aloja") = titular And tbCUENTAS_ALOJA("habitacion_cuenta_aloja") = hab Then
                    'solo trajajo con los gastos de tipo 1 y 4
                    If (tbCUENTAS_ALOJA("tipoAloja") = 1 Or tbCUENTAS_ALOJA("tipoAloja") = 4) Then
                        'borro el gasto
                        tbCUENTAS_ALOJA.Delete
                    End If
                    tbCUENTAS_ALOJA.MoveNext
                Else
                    Exit Do
                End If
            Loop
        End If
    Else
        'acceso directamente al regisrto
        tbCUENTAS_ALOJA.Index = "pi_cuentas_aloja"
        tbCUENTAS_ALOJA.Seek "=", fecha, nroCorrGasto
        If Not tbCUENTAS_ALOJA.NoMatch Then
            'localizé el gasto y lo borro
            tbCUENTAS_ALOJA.Delete
        End If
    End If

End Sub

Public Sub mSub_Elimino_Documento(tipo_docu As Byte, nro_docu As Long)
    'Elimino documento anulado.
    
    If busco_documentoTF(tipo_docu, nro_docu) Then
        'borro lineas
        tbLINEAS.Index = "i_lineas"
        tbLINEAS.Seek ">=", tipo_docu, nro_docu, 1
        If Not tbLINEAS.NoMatch Then
            Do While Not tbLINEAS.EOF
                If tbLINEAS("tipo_linea") = tipo_docu _
                And tbLINEAS("nro_factura") = nro_docu Then
                    tbLINEAS.Delete
                    tbLINEAS.MoveNext
                Else
                    Exit Do
                End If
            Loop
        End If

        'borro cabezal
        tbCABEZAL.Delete
    End If
End Sub

Public Sub mSub_Elimino_Documento_EstadoCuenta(tipo_docu As Byte, nro_docu As Long)
    'si el documento es crédito lo elimino del estado de cuentas.
    If tipo_docu = 3 Or tipo_docu = 4 Or tipo_docu = 7 Or tipo_docu = 8 Then
        tbESTADO_CUENTAS.Index = "pi_estado_cuentas"
        tbESTADO_CUENTAS.Seek "=", tipo_docu, nro_docu
        If Not tbESTADO_CUENTAS.NoMatch Then
            tbESTADO_CUENTAS.Delete
        End If
    End If
End Sub

Public Sub mSub_creo_gastos_nuevamente(tipo_docu As Byte, nro_docu As Long, cli As Long)
    'Recorro el documento y vuelvo a cargar los gastos extras y/o alojamineto
    'al archivo correspondiente
    'Es utilizado cuando: anulo una factura o
    'Creo una nueva devolución
    
    'Como recorro las lineas del documento es necesario
    'pasarle como parametro el cliente al cual pertenecen los gastos
    
    tbLINEAS.Index = "i_lineas"
    tbLINEAS.Seek ">=", tipo_docu, nro_docu, 1
    If Not tbLINEAS.NoMatch Then
        Do While Not tbLINEAS.EOF
            If tbLINEAS("tipo_linea") = tipo_docu _
            And tbLINEAS("nro_factura") = nro_docu Then
                'identifico a los tipos de gastos
                If tbLINEAS("tipogasto_linea") = "a" Then   'alojamineto
                    creo_gastos_aloja cli
                Else
                    creo_gasto_extra cli                       'extras
                End If
                tbLINEAS.MoveNext
            Else
                Exit Do
            End If
        Loop
    End If
End Sub

Private Sub creo_gasto_extra(cli As Long)
    'cli= cliente a quien pertenecen los gastos
    Dim prox As Long
    prox = obtengo_proximo_gasto(tbLINEAS("fecha_linea"))
    tbCUENTAS.AddNew
        tbCUENTAS("fechagasto_cuenta") = tbLINEAS("fecha_linea")
        tbCUENTAS("nrocorr_cuenta") = prox
        
        tbCUENTAS("habitacion_cuenta") = tbLINEAS("hab_linea")
        tbCUENTAS("moneda_cuenta") = tbLINEAS("tipom_linea")
        'quedan vacíos
        'tbCUENTAS ("boleta_cuenta")
        'tbCUENTAS ("puntoventa_cuenta")
        '*
        tbCUENTAS("articulo_cuenta") = tbLINEAS("artcod_linea")
        tbCUENTAS("cantidad_cuenta") = tbLINEAS("artcant_linea")
        If tbLINEAS("tipom_linea") = 0 Then 'M/N
            tbCUENTAS("importe_dolares_cuenta") = 0
            tbCUENTAS("total_dolares_cuenta") = 0
            tbCUENTAS("importe_mnacional_cuenta") = tbLINEAS("artpu_linea")
            tbCUENTAS("total_mnacional_cuenta") = tbLINEAS("arttotal_linea")
        Else                                'U$S
            tbCUENTAS("importe_dolares_cuenta") = tbLINEAS("artpu_linea")
            tbCUENTAS("total_dolares_cuenta") = tbLINEAS("arttotal_linea")
            tbCUENTAS("importe_mnacional_cuenta") = 0
            tbCUENTAS("total_mnacional_cuenta") = 0
        End If
        tbCUENTAS("titular_cuenta") = cli
        tbCUENTAS("facturado") = 0
        tbCUENTAS("documento") = 0
    tbCUENTAS.Update
End Sub

Private Sub creo_gastos_aloja(cli As Long)
    'cli= cliente a quien pertenecen los gastos
    Dim prox As Long
    prox = obtengo_ultimo_corr_aloja(tbLINEAS("fecha_linea"))
    tbCUENTAS_ALOJA.AddNew
        tbCUENTAS_ALOJA("habitacion_cuenta_aloja") = tbLINEAS("hab_linea")
        tbCUENTAS_ALOJA("nrocorr_cuenta_aloja") = prox
        tbCUENTAS_ALOJA("fecha") = tbLINEAS("fecha_linea")
        tbCUENTAS_ALOJA("tarifa") = tbLINEAS("totalconv_linea") 'el importe de los gastos de alojamiento se guarda en este campo
        tbCUENTAS_ALOJA("titular_aloja") = cli
        tbCUENTAS_ALOJA("facturado") = 0
        tbCUENTAS_ALOJA("documento") = 0
        tbCUENTAS_ALOJA("tipoAloja") = tbLINEAS("ArtCod_linea")
        tbCUENTAS_ALOJA("obsAloja") = tbLINEAS("ArtDes_linea")
    tbCUENTAS_ALOJA.Update
End Sub

Public Function mFunDescripcionTipoDocu(tipodocu As Byte)
    'Debuelve la descripción de cada tipo de documento
    Select Case tipodocu
        Case 1  'contado moneda nacional
            mFunDescripcionTipoDocu = "Contado " & gblSignoMonedaNacional
        Case 2  'contado dólares
            mFunDescripcionTipoDocu = "Contado " & gblSignoDolares
        Case 3  'crédito moneda nacional
            mFunDescripcionTipoDocu = "Crédito " & gblSignoMonedaNacional
        Case 4  'crédito dólares
            mFunDescripcionTipoDocu = "Crédito " & gblSignoDolares
        Case 5  'dev. cdo. moneda nacional
            mFunDescripcionTipoDocu = "Dev. Cdo. " & gblSignoMonedaNacional
        Case 6  'dev. cdo. dólares
            mFunDescripcionTipoDocu = "Dev. Cdo. " & gblSignoDolares
        Case 7  'dev. cre. moneda nacional
            mFunDescripcionTipoDocu = "Dev. Cre. " & gblSignoMonedaNacional
        Case 8  'dev. cre. dólares
            mFunDescripcionTipoDocu = "Dev. Cre. " & gblSignoDolares
    End Select
End Function

Public Function mFun_PorIva(CodIva As Byte)
    'Debuelve el porcentaje de iva asociado a un código de iva
    If mFun_BuscoIvaTF(CodIva) Then
        mFun_PorIva = tbIVA("ValorIva")
    End If
End Function

Public Function mFun_PorIvaArt(art As Long)
    'Debuelve el porcentaje de iva de un artículo
    mFun_PorIvaArt = 0
    'Busco artículo
    If busco_articuloTF(art) Then
        'Busco porcentaje iva
        If mFun_BuscoIvaTF(tbARTICULOS("CodIvaArticulo")) Then
            mFun_PorIvaArt = tbIVA("ValorIva")
        End If
    End If
End Function

Public Function mFunTipoIvaALoja(nacionalidadCli As Integer, tipoDatoDev As Byte) As Variant
    '-------------------------------------------------------------------------------------
    'Devuelve el tipo de iva del alojamiento, dependiendo de la nacionalidad
    'del cliente.
    '-------------------------------------------------------------------------------------
    'Parámetros.
    '   Entrada:    [nacionalidadCli]   nacionalidad del cliente con el que se está
    '                                   trabajando.
    '               [tipoDatoDev]       1 = devuelvo el tipo de iva (1,2 o 3)
    '                                   2 = devuelvo el porcentaje de iva
    '                                   (archivo IVA)
    '   Salida:     tipo de iva alojamiento.
    '               Si esta seleccionado en la pantalla de configuración del sistema,
    '               la opción de aplicar un tipo de impuesto distinto a los extranjeros,
    '               y el cliente es extranjero, entonces se devulve el valor almacenado
    '               en el campo tbPARAMETROS("factTipoImpAlojaExt")
    '               Si no esta seleccionada esta opción entonces devuelvo el valor
    '               almacenado en el campo tbPARAMETROS("tipoIvaAloja")
    '
    '               Independientemente del valor devuelto, el mismo siempre corresponderá
    '               a un código de iva existente. 1 = Básico,2 = Mínimo, 3= exento
    '----------------------------------------------------------------------------------------------
    
    Dim tipoIva As Byte
    If tbPARAMETROS("factDiferenciarImpAlojaExt") = 1 Then
        'aplico diferente tipo de impuesto al alojamiento, dependiendo de la nacionalidad del pax
        If nacionalidadCli = tbPARAMETROS("factNacionalidadLocal") Then
            'el cliente es nacional
            tipoIva = tbPARAMETROS("tipoIvaAloja")
        Else
            'el cliente es extranjero
            tipoIva = tbPARAMETROS("factTipoImpAlojaExt")
        End If
    Else
        'el impuesto de alojamiento es el mismo para todos los pax
        tipoIva = tbPARAMETROS("tipoIvaAloja")
    End If
    'determino que tipo de dato devuelvo
    Select Case tipoDatoDev
        Case 1  'tipo de iva
            mFunTipoIvaALoja = tipoIva
        Case 2  'porcentaje asociado
            mFunTipoIvaALoja = mFunObtengoPorcentajeIva(tipoIva)
    End Select
End Function


Public Function mFunObtengoCodIvaArticulo(codArt As Long) As Byte
    '-----------------------------------------------------------------------
    'Devuelve el código de iva de un artículo determinado.
    '------------------------------------------------------------------------
    'Parámetros.
    '   Entrada.
    '       [codArt]    artículo del cual quiero obtener el IVA
    '
    '   Salida  código de Iva   0 = básico
    '                           1 = mínimo
    '                           2 = exento
    '                           99 = artículo con código de Iva inválido
    '-------------------------------------------------------------------------
    'declaro variables para utilizar tabla de artículos
    Dim tbArt As Recordset
    Set tbArt = tbARTICULOS
    tbArt.Index = "i_articulo"
    tbArt.Seek "=", codArt
    If Not tbArt.NoMatch Then
        mFunObtengoCodIvaArticulo = tbArt("CodIvaArticulo")
    Else
        mFunObtengoCodIvaArticulo = 99
    End If
    Set tbArt = Nothing
End Function

Public Sub mSubMuestro_Totales(grilla As MSFlexGrid, _
                            TotalExento As Double, _
                            TotalImpMinimo As Double, _
                            TotalIvaMinimo As Double, _
                            TotalImpBasico As Double, _
                            TotalIvaBasico As Double, _
                            TotalGral As Double)

    If tbPARAMETROS("SisMostrarTotalesResumidos") = 0 Then
            grilla.FormatString = FunCabezalTotalesExtras
            grilla.col = 0
            grilla.Text = Format(TotalExento, "####0.00;;#")
            grilla.col = 1
            grilla.Text = Format(TotalImpMinimo, "####0.00;;#")
            grilla.col = 2
            grilla.Text = mFun_PorIva(2)
            grilla.col = 3
            grilla.Text = Format(TotalIvaMinimo, "####0.00;;#")
            grilla.col = 4
            grilla.Text = Format(TotalImpBasico, "####0.00;;#")
            grilla.col = 5
            grilla.Text = mFun_PorIva(1)
            grilla.col = 6
            grilla.Text = Format(TotalIvaBasico, "####0.00;;#")
            grilla.col = 7
            grilla.Text = Format(TotalGral, "####0.00")
            grilla.CellFontBold = True          'pongo total en negrita
            grilla.RowHeight(1) = 300

        
    Else    'Muestro totales resumidos
        'Achica el tamaño de la grilla de totales,
        'para mostrar el total en forma resumida.
        grilla.Width = 3650
        grilla.Left = 7800
        grilla.RowHeight(1) = 300
        
        grilla.FormatString = "Importe Total                                                          "
        grilla.col = 0
        grilla.Text = Format(TotalGral, "####0.00")
        grilla.CellFontWidth = 10
        grilla.CellFontBold = True          'pongo total en negrita
    End If
    
End Sub

Private Function FunCabezalTotalesExtras()
     FunCabezalTotalesExtras = "Exento       |" & _
                                "Imponible     |" & _
                                "% |" & _
                                "I.V.A.      |" & _
                                "Imponible     |" & _
                                "% |" & _
                                "I.V.A.      |" & _
                                "Total         "
End Function

Public Function mFunObtengoCotiDocu(tipoDoc As Byte, NroDoc As Long) As Single
    '---------------------------------------------------------------------------
    'Obtiene la cotización que se utilizó para realizar un documento.
    '---------------------------------------------------------------------------
    'Parámetros.
    '   Entrda [tipoDoc] tipo del documento que estoy consultando
    '          [NroDoc] número del documento que estoy consultando
    '
    '   Salida Valor de la cotización que se utilizó para crear el documento.
    '----------------------------------------------------------------------------
    'declaro variables para acceder al cabezal del documento
    Dim tbCab As Recordset
    Set tbCab = tbCABEZAL
    tbCab.Index = "i_cabezal"
    tbCab.Seek "=", tipoDoc, NroDoc
    If Not tbCab.NoMatch Then
            mFunObtengoCotiDocu = tbCab("cotizacion")
    Else
        mFunObtengoCotiDocu = 0
    End If
    Set tbCab = Nothing
    
End Function

Public Sub subCabezalGrilla(grilla As MSFlexGrid)
    '---------------------------------------------------------------------
    'Configuro cabezal de la grilla.
    'Utilizado para grilla de facturación y deolución
    '---------------------------------------------------------------------
    grilla.FormatString = _
    "| Fecha    " & _
    "| Hab. " & _
    "| Código     " & _
    "| Descripción                                                        " & _
    "| Tipo " & _
    "| Cant. " & _
    "| P. unitario " & _
    "| Total          " & _
    "|  Total conv." & _
    "|           Khab  " & _
    "| Kcorr    " & _
    "| Ktipo"
End Sub

Public Function mFunObtengoCantidadViasImp() As Byte
    '------------------------------------------------------------------------
    'Devuelve la cantidad de vías que se deben de imprimir de cada documento
    '------------------------------------------------------------------------
    'Parámetros.
    '   Salida: cantidad de vías a imprimir ya sea de facturas o devoluciones.
    '-------------------------------------------------------------------------
    mFunObtengoCantidadViasImp = tbPARAMETROS("factCantViasImpresas")
End Function

Public Function mFunObtengoImpExtranjero() As Byte
    '---------------------------------------------------------------------------------
    'Devuelve el tipo de impuesto que se le aplica a los gastos de tipo alojamiento
    'cuando el cliente es extranjero.
    'Según las reglamentaciones actuales (04/2003), los extranjeros tienen exonerado
    'el pago de impuestos en el alojamiento.
    'Esto no es así para todos los hoteles.
    '----------------------------------------------------------------------------------
    'Parámetros.
    '   Salida: tipo de impuesto amplicado
    '           1 = basico
    '           2 = mínimo
    '           3 = exento
    'Estos valor corresponden a los existentes en el archivo IVA
    '-----------------------------------------------------------------------------------
    mFunObtengoImpExtranjero = tbPARAMETROS("factTipoImpAlojaExtranjeros")
End Function

Public Function mFunObtengoCantViasImp() As Byte
    '-------------------------------------------------------------------------
    'Devuelve la cantidad de vías de los documentos que hay que imprimir
    '-------------------------------------------------------------------------
    mFunObtengoCantViasImp = tbPARAMETROS("factCantViasImpresas")
End Function

Public Sub mSubArmoReporteFactura(tipoDoc As Byte, NroDoc As Long)
    '-----------------------------------------------------------------------------------
    'Configuro el control Crystal genérico,obtengo datos y documento de tipo facturas
    '-----------------------------------------------------------------------------------
    'Parámetros.
    '   Entrada:    [tipoDoc] tipo del documento a imprimir
    '               [nroDoc]  número del documento a imprimir
    '-----------------------------------------------------------------------------------
    Dim totViasAImprimir As Byte
    'inicializo control data
     subInicializoControlData frmMAIN.Data1CrystalReport
    
    'NOTA: esta consulta debe de ser la misma que se utilizo para crear el reporte,
    'por ese motivo no debe de ser modificada en aspectos tales como campos a mostrar y
    'tablas utilizadas.
    frmMAIN.Data1CrystalReport.RecordSource = _
    "select * from fac_cabezal,fac_lineas where " & _
    "fac_cabezal.tipo_docu = fac_lineas.tipo_linea and " & _
    "fac_cabezal.nro_docu = fac_lineas.nro_factura and " & _
    "fac_cabezal.tipo_docu = " & tipoDoc & _
    "and fac_cabezal.nro_docu = " & NroDoc
    
    
    'ejecuto consulta control data
    frmMAIN.Data1CrystalReport.Refresh
    
    'como de aquí en más voy a trabajar directamente con el control data
    'me aseguro de que se allan encontrado reservas
    If frmMAIN.Data1CrystalReport.Recordset.RecordCount > 0 Then
        'determino el nombre del reporte a utilizar
        frmMAIN.CrystalReport1.ReportFileName = vardir2 + "rptfacturas.rpt"

        'inicializo fórmulas
        With frmMAIN.CrystalReport1
            .Formulas(0) = "cabRegistro = '" & mFunImpRegistro & "'"                'hotel propietario de la aplicación
            .Formulas(1) = "signoMn = '" & gblSignoMonedaNacional & "'"
            .Formulas(2) = "signoDol = '" & gblSignoDolares & "'"
        End With
        
        'imprimo la cantidad de vías establecidas en el cuadro de configuración
        totViasAImprimir = mFunObtengoCantViasImp
        'genero listado
        frmMAIN.CrystalReport1.Action = 1
        frmMAIN.CrystalReport1.CopiesToPrinter = totViasAImprimir
    
        'inicializo fórmulas
        mSubInicializoFormulas 2
    Else
        'aviso de que no hay datos para imprimir
        mSubMensaje 3, 9
    End If
End Sub

