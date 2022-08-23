Attribute VB_Name = "MenuArbol"
'El 1 de abril del 2002 se quita este módulo ya que los procedimientos que
'contiene son usados por frmMAIN el cual será remplazado por frmMAINver2
Option Explicit

Public Sub mSubMenuMuestroReservas()
    'Muestra las reservas activas del hotel
    
    'muestro cabezal
    subCabezalReserva
    
    'Recorro el archivo de reservas y muestro las que la fecha de ingreso
    'sea mayor igual a la fecha de hoy
    tbRESERVAS.Index = "i_res_fing"
    tbRESERVAS.Seek ">=", m_FechaSis
    If Not tbRESERVAS.NoMatch Then  'si hay reservas para ingresar
        Do While Not tbRESERVAS.EOF
            'no muestro las reservas ya ingresadas
            If Not busco_reservaCheckinTF(tbRESERVAS("nroreserva")) Then
                subLineaReservas
            End If
            tbRESERVAS.MoveNext
        Loop
    End If
    'Muestro datos
    frmMAIN.lwDerecha.View = lvwReport
End Sub

Public Sub mSubMenuMuestroIngresos()
    'Muestro los ingresos previstos para hoy
    
    Dim rst_ing As Recordset    'almaceno resultado consulta
    Dim qdf_ing As QueryDef
    
    'ejecuto consulta
    Set qdf_ing = bdHOTEL.CreateQueryDef("")
    qdf_ing.SQL = SQLIngresosPrevistos(m_FechaSis)
    Set rst_ing = qdf_ing.OpenRecordset(dbOpenSnapshot)
    
    'muestro cabezal
    subCabezalIngresos
    'recorro recordset para mostrar en grilla
    If rst_ing.RecordCount > 0 Then
        rst_ing.MoveFirst
        Do While Not rst_ing.EOF
            subLineaIngresos rst_ing
            rst_ing.MoveNext
        Loop
    End If
    
    'Muestro datos
    frmMAIN.lwDerecha.View = lvwReport
End Sub

Public Sub mSubMenuMuestroEgresos(grilla As MSFlexGrid)
    'Muestro los pasajeros que engresan hoy y los que ya lo hicieron.
    Dim rst_egr As Recordset    'almaceno resultado consulta
    Dim qdf_egr As QueryDef
    
    Dim hab_ant As Long
    
    'borro datos grilla
    limpio_grilla grilla
    
    'ejecuto consulta
    Set qdf_egr = bdHOTEL.CreateQueryDef("")
    qdf_egr.SQL = SQLEgresosPorRealizar(m_FechaSis) & " UNION " & SQLEgresosYaRealizados(m_FechaSis) & _
    " order by habitacion"  'ordeno por número de habitación para realizar el corte de control.
    Set rst_egr = qdf_egr.OpenRecordset(dbOpenSnapshot)
    
    'muestro cabezal
    grilla.FormatString = funCabezalEgresos
    
    If rst_egr.RecordCount > 0 Then
        rst_egr.MoveFirst
        Do While Not rst_egr.EOF
        
            'muestro habitación
            grilla.AddItem FunLineaHabEgresos(rst_egr)
            'muestro lineas en negrita
            mSubLineaEnNegrita grilla, grilla.Rows - 1
            'muestro linea como titulo
            mSubLineaComoTitulo grilla, grilla.Rows - 1
            
            hab_ant = rst_egr("habitacion")
            Do While Not rst_egr.EOF
                If hab_ant = rst_egr("habitacion") Then
                    'muestro pasajeros
                    grilla.AddItem funLineaPasaEgresos(rst_egr)
                    'si ya se fue del hotel
                    If rst_egr(4) <> "" Then
                        'muestro icono de salida del hotel
                        grilla.col = 2
                        grilla.Row = grilla.Rows - 1
                        grilla.CellPictureAlignment = 4
                        Set grilla.CellPicture = frmMAIN.ImageList2.ListImages(3).Picture
                    End If
                    rst_egr.MoveNext
                Else
                    Exit Do
                End If
            Loop
        Loop
        'borro primera fila de la grilla porque queda vacia
        grilla.RemoveItem (1)

    End If
End Sub

Public Sub mSubMenuMuestroFacturas()
    'Muestro las facturas realizadas hoy
    Dim rst_fact As Recordset    'almaceno resultado consulta
    Dim qdf_fact As QueryDef
    
    'ejecuto consulta
    Set qdf_fact = bdHOTEL.CreateQueryDef("")
    qdf_fact.SQL = SQLFacturasRealizadas(m_FechaSis)
    Set rst_fact = qdf_fact.OpenRecordset(dbOpenSnapshot)
    
    'muestro cabezal
    subCabezalFacturas
    
    'recorro recordset para mostrar en grilla
    If rst_fact.RecordCount > 0 Then
        rst_fact.MoveFirst
        Do While Not rst_fact.EOF
            subLineaFacturas rst_fact
            rst_fact.MoveNext
        Loop
    End If
End Sub

Public Sub mSubMenuMuestroDevoluciones()
    'Muestro las devoluciones realizadas hoy
    Dim rst_dev As Recordset    'almaceno resultado consulta
    Dim qdf_dev As QueryDef
    
    'ejecuto consulta
    Set qdf_dev = bdHOTEL.CreateQueryDef("")
    qdf_dev.SQL = SQLDevolucionesRealizadas(m_FechaSis)
    Set rst_dev = qdf_dev.OpenRecordset(dbOpenSnapshot)
    
    'muestro cabezal
    subCabezalDevoluciones
    
    'recorro recordset para mostrar en grilla
    If rst_dev.RecordCount > 0 Then
        rst_dev.MoveFirst
        Do While Not rst_dev.EOF
            subLineaDevoluciones rst_dev
            rst_dev.MoveNext
        Loop
    End If
End Sub

Public Sub mSubMenuMuestroRecivos()
    'Muestro los recivos realizadas hoy
    Dim rst_res As Recordset    'almaceno resultado consulta
    Dim qdf_res As QueryDef
    
    'ejecuto consulta
    Set qdf_res = bdHOTEL.CreateQueryDef("")
    qdf_res.SQL = SQLRecivosRealizados(m_FechaSis)
    Set rst_res = qdf_res.OpenRecordset(dbOpenSnapshot)
    
    'muestro cabezal
    subCabezalRecivos
    
    'recorro recordset para mostrar en grilla
    If rst_res.RecordCount > 0 Then
        rst_res.MoveFirst
        Do While Not rst_res.EOF
            subLineaRecivos rst_res
            rst_res.MoveNext
        Loop
    End If
End Sub

Public Sub mSubMenuMuestroHabitaciones(grilla As MSFlexGrid)
    'Muestro todas la habitaciones del hotel, su estado y su situación.
    
    'borro datos grilla
    limpio_grilla grilla
    
    'muestro cabezal
    grilla.FormatString = funCabezalHabitaciones
    
    'recorro todas la habitaciones
    tbHABITACIONES.Index = "inrohab"
    tbHABITACIONES.Seek ">=", 0
    If Not tbHABITACIONES.NoMatch Then
        Do While Not tbHABITACIONES.EOF
            grilla.AddItem funLineaHabitaciones
            'muestro icono de situación
            grilla.col = 4
            grilla.Row = grilla.Rows - 1
            grilla.CellPictureAlignment = 1
            If tbHABITACIONES("situacionhab") = 1 Then  'limpia
                Set grilla.CellPicture = frmMAIN.ImageList2.ListImages(5).Picture
            End If
            If tbHABITACIONES("situacionhab") = 2 Then  'sucia
                Set grilla.CellPicture = frmMAIN.ImageList2.ListImages(4).Picture
            End If
            
            tbHABITACIONES.MoveNext
        Loop
        'borro primera fila de la grilla porque queda vacia
        grilla.RemoveItem (1)
    End If
End Sub

'*********************************************
'*
'*  Muestro cabezal y doy formato a las lineas
'*
'*
'*********************************************

Private Sub subCabezalReserva()
    frmMAIN.lwDerecha.ColumnHeaders.Add , "FechaIng", "Fecha ingreso", 1200
    frmMAIN.lwDerecha.ColumnHeaders.Add , "FechaEgr", "Fecha egreso", 1200
    frmMAIN.lwDerecha.ColumnHeaders.Add , "Reserva", "Reserva", 1200
    frmMAIN.lwDerecha.ColumnHeaders.Add , "Noches", "Noches", 500
    frmMAIN.lwDerecha.ColumnHeaders.Add , "titular", "Titular de la reserva", 4000
End Sub

Private Sub subCabezalIngresos()
    'Realizo cabezal de grilla de ingresos previstos
    frmMAIN.lwDerecha.ColumnHeaders.Add , "Reserva", "Reserva", 1200
    frmMAIN.lwDerecha.ColumnHeaders.Add , "Hab", "Habitación", 1200
    frmMAIN.lwDerecha.ColumnHeaders.Add , "Tipo", "Tipo", 1200
    frmMAIN.lwDerecha.ColumnHeaders.Add , "Pasajeros", "Pasajeros", 1200
    frmMAIN.lwDerecha.ColumnHeaders.Add , "titular", "Titular de la reserva", 4000
End Sub

Private Function funCabezalEgresos()
    'realizo cabeal de grilla de egresos previstos
    funCabezalEgresos = _
    "Habitación |" & _
    "Tipo               |" & _
    "       |" & _
    "Pasajeros                                                                           |" & _
    "Hora salida        "
End Function

Private Sub subCabezalFacturas()
    'Realizo cabezal de grilla de facturas
    frmMAIN.lwDerecha.ColumnHeaders.Add , "Numero", "Número", 1200
    frmMAIN.lwDerecha.ColumnHeaders.Add , "Tipo", "Tipo", 1200
    frmMAIN.lwDerecha.ColumnHeaders.Add , "Fecha", "Fecha", 1200
    frmMAIN.lwDerecha.ColumnHeaders.Add , "Total", "Total", 1200
    frmMAIN.lwDerecha.ColumnHeaders.Add , "Cliente", "Titular de la reserva", 4000
End Sub

Private Sub subCabezalDevoluciones()
    'Realizo cabezal de grilla de devoluciones
    frmMAIN.lwDerecha.ColumnHeaders.Add , "Numero", "Número", 1200
    frmMAIN.lwDerecha.ColumnHeaders.Add , "Tipo", "Tipo", 1200
    frmMAIN.lwDerecha.ColumnHeaders.Add , "Fecha", "Fecha", 1200
    frmMAIN.lwDerecha.ColumnHeaders.Add , "Total", "Total", 1200
    frmMAIN.lwDerecha.ColumnHeaders.Add , "Factura", "Factura", 1200
    frmMAIN.lwDerecha.ColumnHeaders.Add , "Cliente", "Titular de la reserva", 4000
End Sub

Private Sub subCabezalRecivos()
    'Realizo cabezal de grilla de recivos manuales
    frmMAIN.lwDerecha.ColumnHeaders.Add , "Numero", "Número", 1200
    frmMAIN.lwDerecha.ColumnHeaders.Add , "Moneda", "Moneda", 1200
    frmMAIN.lwDerecha.ColumnHeaders.Add , "Fecha", "Fecha", 1200
    frmMAIN.lwDerecha.ColumnHeaders.Add , "Importe", "Importe", 1200
    frmMAIN.lwDerecha.ColumnHeaders.Add , "Cliente", "Realizado a", 4000
End Sub

Private Function funCabezalHabitaciones()
    'Realizo cabezal de grilla de habitaciones
    funCabezalHabitaciones = _
    "Habitación      |" & _
    "Tipo                                |" & _
    "     |" & _
    "Estado                              |" & _
    "     |" & _
    "Situación                                  "
End Function

Private Sub subLineaReservas()
    'Muestro datos de la reservas
    Dim itmX As ListItem
    
    Set itmX = frmMAIN.lwDerecha.ListItems.Add(, , tbRESERVAS("fechaing"))
    itmX.SubItems(1) = tbRESERVAS("fechaegr")
    itmX.SubItems(2) = tbRESERVAS("nroreserva")
    itmX.SubItems(3) = tbRESERVAS("cantnoches")
    itmX.SubItems(4) = tbRESERVAS("primer_ape_titular") + " " + tbRESERVAS("segundo_ape_titular") + " " + _
    tbRESERVAS("primer_nom_titular") + " " + tbRESERVAS("segundo_nom_titular")
    'Si la reserva ingresa hoy muestro icono
    If tbRESERVAS("fechaing") = m_FechaSis Then
        itmX.SmallIcon = 1
    End If
End Sub

Private Sub subLineaIngresos(rst As Recordset)
    'Muestro datos de los ingresos previstos para hoy
    Dim itmX As ListItem
    
    Set itmX = frmMAIN.lwDerecha.ListItems.Add(, , rst("NroReserva"))
    itmX.SubItems(1) = rst("nrohabitacion")
    itmX.SubItems(2) = rst("descripcion")
    itmX.SubItems(3) = rst("pasajeros")
    itmX.SubItems(4) = rst("primer_ape_titular") + " " + rst("segundo_ape_titular") + " " + _
    rst("primer_nom_titular") + " " + rst("segundo_nom_titular")
    
    'si ya ingreso muestro hora e icono
    If busco_habita_checkin(rst("nrohabitacion")) Then
        itmX.SmallIcon = 2
    End If
End Sub

Private Function funLineaPasaEgresos(rst As Recordset)
    'Muestro datos de los pasajeros de cada habitación
    funLineaPasaEgresos = _
    Chr(9) & _
    Chr(9) & _
    Chr(9) & _
    rst("nombre_completo_titular") & _
    Chr(9) & _
    rst(4)  'hora egreso
End Function

Private Function FunLineaHabEgresos(rst As Recordset)
    'Muestro datos de las habitaciones que se fueron o estan por irse del hotel
    FunLineaHabEgresos = _
    rst("habitacion") & _
    Chr(9) & _
    mFun_BuscoDescriTipoHab(rst("tipohab"))
End Function

Private Sub subLineaFacturas(rst As Recordset)
    'Muestro datos de las facturas realizadas el día de hoy.
    
    Dim itmX As ListItem
    
    Set itmX = frmMAIN.lwDerecha.ListItems.Add(, , rst("nro_docu"))
    itmX.SubItems(1) = mFunDescripcionTipoDocu(rst("tipo_docu"))
    itmX.SubItems(2) = rst("fecha_docu")
    itmX.SubItems(3) = rst("tot_docu")
    itmX.SubItems(4) = rst("nom_docu")
End Sub

Private Sub subLineaDevoluciones(rst As Recordset)
    'Muestro datos de las devoluciones realizadas el día de hoy.
    Dim itmX As ListItem
    
    Set itmX = frmMAIN.lwDerecha.ListItems.Add(, , rst("nro_docu"))
    itmX.SubItems(1) = mFunDescripcionTipoDocu(rst("tipo_docu"))
    itmX.SubItems(2) = rst("fecha_docu")
    itmX.SubItems(3) = rst("nro_fact_docu")
    itmX.SubItems(4) = rst("nom_docu")
End Sub

Private Sub subLineaRecivos(rst As Recordset)
    'Muestro datos de los recivos manuales ingresados el día de hoy.
    Dim itmX As ListItem
    
    Set itmX = frmMAIN.lwDerecha.ListItems.Add(, , rst("nro_recivo"))
    itmX.SubItems(1) = paso_moneda_a_desc(rst("moneda_recivo"))
    itmX.SubItems(2) = rst("fecha_recivo")
    itmX.SubItems(3) = rst("importe_recivo")
    itmX.SubItems(4) = rst("nomcli_recivo")
End Sub

Private Function funLineaHabitaciones()
    'Muestro datos de todas la habitaciones del hotel.
    funLineaHabitaciones = _
    tbHABITACIONES("nrohab") & _
    Chr(9) & _
    mFun_BuscoDescriTipoHab(tbHABITACIONES("tipohab")) & _
    Chr(9) & _
    Chr(9) & _
    mFunObtengoEstadoHab(tbHABITACIONES("nrohab")) & _
    Chr(9) & _
    Chr(9) & _
    mFunObtengoSituacionHab(tbHABITACIONES("situacionhab"))
End Function
