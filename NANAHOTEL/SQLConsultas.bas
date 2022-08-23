Attribute VB_Name = "SQLConsultas"
Option Explicit

Public Function fechaSQL(f As Variant)
    'Devuelve una cadena de caracters de formato: #MM/DD/AA#
    'con el objetivo de poder usarla en una consulta SQL
    
    Dim aux As String
    If IsDate(f) Then
        aux = Format(f, "mm/dd/yy")
        fechaSQL = "#" & aux & "#"
    End If
End Function

Public Sub SQLpasajeros_habitacion(nrohab As Long, Control As Data)
    'Selecciono los pasajeros de una habitación alojada
    Dim consulta As String
    consulta = "Select clientes.nombre_completo_titular, " & _
                "       checkin.nrocorrcli " & _
                "From   checkin, clientes " & _
                "Where  checkin.nrocorrcli = clientes.nrocorr " & _
                "and checkin.nrohab = " & nrohab
    
    ejecuto_consulta Control, consulta
End Sub

Private Sub ejecuto_consulta(control_data As Data, consulta As String)
    'Ejecuto la consulta que paso por parámetros
    control_data.RecordSource = consulta
    control_data.Refresh
End Sub

Public Function SQLIngresosPrevistos(fecha As Date)
    'Selecciono las habitaciones reservadas en una fecha dada
    
    SQLIngresosPrevistos = _
    "SELECT reservas.nroreserva," & _
        "primer_ape_titular," & _
        "segundo_ape_titular," & _
        "primer_nom_titular," & _
        "segundo_nom_titular," & _
        "hab_reserva.nrohabitacion," & _
        "hab_reserva.tipohabitacion," & _
        "hab_reserva.pasajeros," & _
        "descripcion," & _
        "nroagenciaempresa " & _
        "FROM reservas, hab_reserva,tipo_habitaciones " & _
        "WHERE reservas.nroreserva = hab_reserva.nroreserva and " & _
        "hab_reserva.tipohabitacion = tipo_habitaciones.tipohab and " & _
        "reservas.fechaing = " & fechaSQL(fecha)
End Function

Public Function SQLFacturasRealizadas(fecha As Date)
    'Selecciona las facturas crédito y contado en ambas monedas
    'realizadas en la fecha
    SQLFacturasRealizadas = _
    "SELECT tipo_docu," & _
        "nro_docu," & _
        "fecha_docu," & _
        "nom_docu," & _
        "tot_docu " & _
        "FROM fac_cabezal " & _
        "WHERE fecha_docu = " & fechaSQL(fecha) & " and " & _
        "(tipo_docu = 1 or tipo_docu=2 or tipo_docu=3 or tipo_docu=4)"
End Function

Public Function SQLDevolucionesRealizadas(fecha As Date)
    'Selecciona las devoluciones crédito y contado en ambas monedas
    'realizadas en la fecha
    SQLDevolucionesRealizadas = _
    "SELECT tipo_docu," & _
        "nro_docu," & _
        "fecha_docu," & _
        "nom_docu," & _
        "tot_docu," & _
        "nro_fact_docu " & _
        "FROM fac_cabezal " & _
        "WHERE fecha_docu = " & fechaSQL(fecha) & " and " & _
        "(tipo_docu = 5 or tipo_docu=6 or tipo_docu=7 or tipo_docu=8)"
End Function

Public Function SQLRecivosRealizados(fecha As Date)
    'Selecciona los recivos MANUALES ingresados al sistema
    'de hambas monedas, realizados en la fecha.
    SQLRecivosRealizados = _
    "SELECT tipo_recivo," & _
        "nro_recivo," & _
        "fecha_recivo," & _
        "nomcli_recivo," & _
        "moneda_recivo," & _
        "importe_recivo " & _
        "FROM recivos " & _
        "WHERE fecha_recivo = " & fechaSQL(fecha) & " and " & _
        "tipo_recivo = 2"
End Function

Public Function SQLEgresosPorRealizar(fecha As Variant)
    'Selecciona los pasajeros hospedados que se tienen que ir en la fecha correspondiente
    'y todavía no lo han hecho
    'El campo de tipo texto ('HoraEgr') es utilizado pra realizar la UNION con la consulta
    'SQLEgresosYaRealizados ya que la cantidad de campos debe ser la misma en este tipo
    'de operacion. El mismo ocupa el lugar del campo HoreDeEgreso del archivo Checkout
    
    'El procedimiento es de tipo variant ya que puede recivir una string o un date
    
    SQLEgresosPorRealizar = _
    "SELECT checkin.nrohab," & _
    "nrocorrcli," & _
    "nombre_completo_titular," & _
    "habitaciones.tipohab," & _
    "checkin.nroReserva,'' as HoraEgr " & _
    "FROM checkin, clientes, habitaciones " & _
    "WHERE checkin.nrocorrcli = clientes.nrocorr and " & _
    "habitaciones.nrohab = checkin.nrohab and " & _
    "fcheckhas = " & fechaSQL(fecha)
End Function

Public Function SQLEgresosYaRealizados(fecha As Variant)
    'Selecciona los pasajeros hospedados que se tienen que ir hoy
    'y ya lo hicieron
    
    'El procedimiento es de tipo variant ya que puede recivir una string o un date
    
    SQLEgresosYaRealizados = _
    "SELECT checkout.nrohab," & _
    "nrocorrcli," & _
    "nombre_completo_titular," & _
    "habitaciones.tipohab," & _
    "checkout.nroreserva," & _
    "horaegrhab as HoraEgr " & _
    "FROM checkout, clientes, habitaciones " & _
    "WHERE checkout.nrocorrcli = clientes.nrocorr and " & _
    "habitaciones.nrohab = checkout.nrohab and " & _
    "fhas = " & fechaSQL(fecha)
End Function

