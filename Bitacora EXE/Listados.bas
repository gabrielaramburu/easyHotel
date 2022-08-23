Attribute VB_Name = "Listados"
Option Explicit
Private DescFecha As String     'Se utiliza para almazenar un string
                                'que aparecerá impreso en el listado
                                'conteniendo información hacerca de las fechas
                                'utilizadas.

Public Sub mSubMuestroListados()
    'Recorro el archivo de listados y creo un nodo por cada
    'listado existente.
    
    Dim imagen As Byte
    frmMain.Frame1.Visible = True
    frmMain.lwOpr.Width = 8150
    
    'limpio arbol
    frmMain.twListados.Nodes.Clear
    
    'creo nodo principal
    frmMain.twListados.Nodes.Add , , "Lst", "Listados", 1

    tbSISTEMA_BITACORAlistados.Index = "pk_listado"
    tbSISTEMA_BITACORAlistados.Seek ">=", " "
    If Not tbSISTEMA_BITACORAlistados.NoMatch Then
        Do While Not tbSISTEMA_BITACORAlistados.EOF
            'creo nodo
            If funListadoPrede(tbSISTEMA_BITACORAlistados("NomLst")) Then
                imagen = 7 'imagen predeterminado
            Else
                imagen = 2  'imagen normal
            End If
            frmMain.twListados.Nodes.Add "Lst", 4, , tbSISTEMA_BITACORAlistados("NomLst"), imagen
            tbSISTEMA_BITACORAlistados.MoveNext
        Loop
        'Muestro el árbol expandido
        frmMain.twListados.Nodes.Item(1).Expanded = True
    Else
        'Si no hay listados ingresados muestro mensaje
        subMuestroLeyenda 2
    End If
    frmMain.Frame1.Refresh
End Sub

Public Sub mSubOcultoCuadroListados()
    'Oculta el cuadro de listados
    frmMain.Frame1.Visible = False
    frmMain.lwOpr.Width = 11900
End Sub

Public Sub mSubMuestroInfListado(lst As String)
    'Muestro la descripción del listado
    'en la venta de descripción
    If mfunBuscoListado(lst) Then    'si existe listado
        frmMain.txtInfLst.Text = tbSISTEMA_BITACORAlistados("InfLst")
    Else
        frmMain.txtInfLst.Text = ""
    End If
End Sub

Public Sub mSubEjecutoListado(listado As String)
    'Crea una sentencia SQL con los datos obtenidos del
    'archivo de listados
    
    Dim consulta As String
    Dim LineaFecha As String
    Dim qdf As QueryDef
    
    'busco listado
    If mfunBuscoListado(listado) Then  'si existe
        'obtengo fecha a filtrar
        LineaFecha = _
        funObtengoFecha(tbSISTEMA_BITACORAlistados("WhereTipoFechaLst")) & ")"
        If Not CanceloIngresoFechas Then
            consulta = _
            "Select " & tbSISTEMA_BITACORAlistados("columnlst") & _
            " From sistema_bitacora " & _
            "Where (" & tbSISTEMA_BITACORAlistados("WhereUsrLst") & ") and (" & _
            tbSISTEMA_BITACORAlistados("WhereOprLst") & _
            LineaFecha _
            & tbSISTEMA_BITACORAlistados("OrdenLst")
            
            'ejecuto consulta
            Set qdf = bdAplicacion.CreateQueryDef("")
            qdf.SQL = consulta
             
            Set rst_opr = qdf.OpenRecordset(dbOpenSnapshot)
                 
            'limpio listview
            mSubLimpioLista
             
            subArmoListado
            
            'muestro en barra de estado el nombre del listado
            mSubMuestroListadoActual listado
            
            'Muestro leyenda en barra de estado
            subMuestroLeyenda 1, rst_opr.RecordCount
        End If
    End If
End Sub

Private Sub subMuestroLeyenda(tipo As Byte, Optional inf As String)
    'Muestra leyenda en el primer panel de la barra de estado
    Dim texto As String
    Select Case tipo
        Case 1
            texto = inf & " operaciones seleccionadas."
        Case 2
            texto = "No hay listados para ejecutar."
    End Select
    frmMain.StatusBar1.Panels(1).Text = texto
End Sub

Private Sub subArmoListado()
    'Recorro el recordeset y lo muestro en pantalla
    If rst_opr.RecordCount > 0 Then
        'creo cabezal de la lista
        subCreoCabezalPantalla
    
        If rst_opr.RecordCount > 0 Then
            rst_opr.MoveFirst
            Do While Not rst_opr.EOF
                subCargoLista
                rst_opr.MoveNext
            Loop
        End If
        'muestro lista
        frmMain.lwOpr.View = lvwReport
    End If
End Sub

Private Sub subCargoLista()
    'Muestro los datos contenidos en el recordset en el control listview
    'y ademas muestro la descripción de la operación en lugar
    'del código.
    'IMPORTANTE: Se identifica al campo del recordset por el nombre
    'por este motivo no se puede cambiar el nombre de este campo en el archivo.
    
    Dim itmX As ListItem
    Dim i As Byte
    i = 0
    
    'Recorro los campos del recordset
    If rst_opr.Fields(i).Name = "CodOprBit" Then
        Set itmX = frmMain.lwOpr.ListItems.Add(, , mFunBuscoDescOpr(rst_opr(i)))
    Else
        Set itmX = frmMain.lwOpr.ListItems.Add(, , rst_opr(i))
    End If
    
    i = i + 1
    Do While i <= rst_opr.Fields.Count - 1
        If rst_opr.Fields(i).Name = "CodOprBit" Then
            'tengo que ir a buscar el nombre de la operación
            itmX.SubItems(i) = mFunBuscoDescOpr(rst_opr(i))
        Else
            itmX.SubItems(i) = rst_opr(i)
        End If
        i = i + 1
    Loop
End Sub

Private Function funObtengoFecha(tipo As Byte)
    'Dependiendo del tipo de listado obtengo fecha
    CanceloIngresoFechas = False
    Select Case tipo
        Case 1      'fecha sistema
            funObtengoFecha = " ) and (fechaBit = " & fechaSQL(Date)
            DescFecha = "Operaciones realizadas el día " & Date
            
        Case 2      'fecha aplicacion
            funObtengoFecha = " ) and (fechaBit = " & fechaSQL(tbSISTEMA_PARAMETROS("fecha_ultimo_cierre_realizado"))
            DescFecha = "Operaciones realizadas el día " & tbSISTEMA_PARAMETROS("fecha_ultimo_cierre_realizado")
            
        Case 3      'pido fecha
            tipo_accion_fechas = 1
            frmFechas.Show 1
            funObtengoFecha = " ) and (fechaBit = " & fechaSQL(frmFechas.fFechaIni.Text)
            DescFecha = "Operaciones realizadas el día " & frmFechas.fFechaIni.Text
            'descargo formulario de fechas
            Unload frmFechas
            
        Case 4      'pido rango
            tipo_accion_fechas = 2
            frmFechas.Show 1
            funObtengoFecha = " ) and (fechaBit >= " & _
                            fechaSQL(frmFechas.fFechaIni.Text) & _
                            " and FechaBit <= " & _
                            fechaSQL(frmFechas.fFechaFin.Text)
            DescFecha = "Operaciones realizadas desde el " & frmFechas.fFechaIni.Text _
            & " hasta el " & frmFechas.fFechaFin.Text
            'descargo formulario de fechas
            Unload frmFechas
            
        Case 5      'todas
            'no hago nada
            
    End Select
End Function

Private Sub subCreoCabezalPantalla()
    'creo el cabezal del control listview dependiendo de la
    'configuración del listado a ejecutar
    Dim i As Integer
    Dim columna As String
    Dim caracter As String
    Dim Columnas As String
    
    'La comilla al final es para evitar que la última palabra
    'también se procese ya que si no se llega al fin de la
    'cadena sin encontrar una comilla y esto origina que se no se
    'cree la última columna
    Columnas = tbSISTEMA_BITACORAlistados("ColumnDescLst") & ","
    
    i = 1
    'recorro el campo ColumnDescLst del archvio sistema_bitacoralistados
    'separando las diferentes palabras
    Do While i <= Len(Columnas)
        caracter = Mid(Columnas, i, 1)
        If caracter = "," Then
            'tengo una nueva palabra (columna)
            frmMain.lwOpr.ColumnHeaders.Add , columna, columna
            columna = ""
        Else
            columna = columna & caracter
        End If
        i = i + 1
    Loop
End Sub

Public Sub mSubEliminoListado(listado As String)
    'Elimina un listado creado
    
    'busco listado
    If mfunBuscoListado(listado) Then  'si existe
        If MsgBox("¿Esta seguro que desea eliminar el listado " & listado & " ?", vbOKCancel + vbExclamation) = vbOK Then
            tbSISTEMA_BITACORAlistados.Delete
        End If
    End If
End Sub

Public Sub mSubPredeterminarListado(listado As String)
    'Establece como predeterminado un listado
    
    'busco listado
    tbSISTEMA_BITACORAparametros.Edit
        tbSISTEMA_BITACORAparametros("ListadoPredeterminado") = listado
    tbSISTEMA_BITACORAparametros.Update
    If Len(listado) > 0 Then 'se paso listado como parametro
        MsgBox "El listado " & listado & " se establecio como predeterminado.", vbInformation
    Else
        'no se establec ningúm liastdo como predeterminado
        MsgBox "No se estableció ningún listado como predeterminado", vbInformation
    End If
End Sub

Public Function funListadoPrede(listado As String)
    'Determina si un listado esta establecido como predeterminado
    funListadoPrede = False
    If listado = tbSISTEMA_BITACORAparametros("ListadoPredeterminado") Then
        funListadoPrede = True
    End If
End Function

Public Sub mSubImprimoListado(listado As String)
    'Ejecuta consulta nuevamente pero esta vez el resultado saldrá
    'por la impresora.
    'Imprimo listado por impresora
        
    Dim LineaFecha As String
    Dim consulta As String
    
    'busco listado
    If mfunBuscoListado(listado) Then  'si existe
        'obtengo fecha a filtrar
        LineaFecha = _
        funObtengoFecha(tbSISTEMA_BITACORAlistados("WhereTipoFechaLst")) & ")"
        If Not CanceloIngresoFechas Then
            'A diferencia de la consulta por pantalla quí,
            'agrego el campo de descripción de operaciones.
            consulta = _
            "Select  DescOpr," & tbSISTEMA_BITACORAlistados("columnlst") & _
            " From sistema_bitacora,sistema_operaciones " & _
            " Where sistema_operaciones.codopr = sistema_bitacora.codoprbit and " & _
            "(" & tbSISTEMA_BITACORAlistados("WhereUsrLst") & ") and (" & _
            tbSISTEMA_BITACORAlistados("WhereOprLst") & _
            LineaFecha _
            & tbSISTEMA_BITACORAlistados("OrdenLst")
        
            frmMain.Data1.RecordSource = consulta
            frmMain.Data1.Refresh
            frmMain.CrystalReport1.ReportFileName = m_vardirRpt & "rptbitacorasimple.rpt"
            
            'Cargo formulas dependiendo del orden de los campos
            subCargoFormulas tbSISTEMA_BITACORAlistados("ColumnDescLst")
            
            'Configuro sección de corte de control
            'indicando si se realiza salto de página o no
            subTipoCorte
            
            frmMain.CrystalReport1.WindowState = crptMaximized
            
            frmMain.CrystalReport1.Action = 1
        End If
    End If
End Sub

Private Sub subCargoFormulas(Columnas As String)
    'El problema que hay que solucionar es que los campos del listado
    'tienen que salir en el mismo orden en el que se configuraron (paso 3)
    'manteniendo de esta manera una choerencia entre lo que se ve en la pantalla
    'y lo que sale impreso.
    
    'Para solucionarlo se crearon 6 fórmulas (es la cantidad máxima de columnas a
    'imprimir).
    'La primera fórmula contiene un string que determina el primer campo a mostrar
    'la segunda determina el segundo campo y así susecivamente.
    Dim caracter As String
    Dim columna As String
    Dim j As Byte
    Dim i As Integer
    
    j = 0
    i = 1
    'Esta linea es importante para que se procese la última palabra
    Columnas = Columnas & ","
    
    'recorro el campo ColumnDescLst del archvio sistema_bitacoralistados
    'separando las diferentes palabras
    Do While i <= Len(Columnas)
        caracter = Mid(Columnas, i, 1)
        If caracter = "," Then
            'tengo una nueva palabra (columna)
            frmMain.CrystalReport1.Formulas(j) = "TipoCampo" & j + 1 & "='" & _
            LTrim(columna) & "'"
            columna = ""
            j = j + 1
        Else
            columna = columna & caracter
        End If
        i = i + 1
    Loop
    'cargo fórmulas de títulos
    frmMain.CrystalReport1.Formulas(j) = "TituloListado='" & tbSISTEMA_BITACORAlistados("NomLst") & "'"
    frmMain.CrystalReport1.Formulas(j + 1) = "SubTituloListado='" & tbSISTEMA_BITACORAlistados("DescLst") & "'"

    frmMain.CrystalReport1.Formulas(j + 2) = "TipoCampoCorte='" & tbSISTEMA_BITACORAlistados("CampoCorteLst") & "'"
    frmMain.CrystalReport1.Formulas(j + 3) = "DescFiltroFechas='" & DescFecha & "'"
End Sub

Private Sub subTipoCorte()
    If tbSISTEMA_BITACORAlistados("RealizoCorte") = 1 Then
        'realizo corte de control sin salto de página
        frmMain.CrystalReport1.SectionFormat(0) = "GH1;X;X;X;X;X;X;X"
    End If
    If tbSISTEMA_BITACORAlistados("RealizoCorte") = 2 Then
        'realizo corte de control con salto de página
        frmMain.CrystalReport1.SectionFormat(0) = "GH1;X;T;X;X;X;X;X"
    End If
End Sub
