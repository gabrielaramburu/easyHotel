Attribute VB_Name = "Busquedas"
Option Explicit

Public Function mFunBusqueda(tipo As Byte, Optional tipodocu As Byte) As Variant
    'Llamo a el formulario de b�squeda.

    'La ayuda de los mantenimientos que se realizan en el formulario de mantenimiento
    'general, no se encuntran aqui.
    
    'Tipodocu es utilizado para determinar el tipo de documento que deseo mostrar en la
    'ayuda de decumentos
    
    Dim TituloDocumento As String
    Dim seleccion As String
    Dim consCompleAux As String 'utilizada para formar propiedad propSeleccionComplementaria
                                'en la ayuda de habitaciones ocupadas (9)
    
    'configuro tama�o del formulario
    frmBusq.propAncho = 6825
    frmBusq.propLargo = 11370
    'cargo las propiedades comunes a todas las ayudas
    frmBusq.propTeclaSeleccion = 13             'la tecla de selecci�n es el enter
    frmBusq.propTablasRelacionadas = ""         'por defecto se accede solamente a una tabla
    frmBusq.propSeleccionComplementaria = ""    'por defecto no se realiza selecci�n complementaria
    Select Case tipo
        Case 1  'todos los pasajeros pasajeros
            frmBusq.propNroCampoInicial = 1          'por defecto ordeno por nombre completo
            frmBusq.propIndiceCampoRetorno = 20      'retorno el c�digo del cliente
            frmBusq.propTabla = "CLIENTES"
            frmBusq.propCampos = "1;CLIENTES;NombreCompleto;4500@2;CLIENTES;PrimerNombre;1500@3;CLIENTES;SegundoNombre;1500@4;CLIENTES;PrimerApellido;1500@5;CLIENTES;SegundoApellido;1500@17;CLIENTES;Documento;1500@6;CLIENTES;Direcci�n;3000@7;CLIENTES;Localidad;1500@1;PAISES;Pa�sResidencia;1500@9;CLIENTES;C�digoPostal;1500@10;CLIENTES;Tel�fono;1500@11;CLIENTES;Fax;1500@20;CLIENTES;Email;1500@12;CLIENTES;OtrosTelFax;1500@1;SEXO;Sexo;1000@1;NACIONALIDADES;Nacionalidad;1500@15;CLIENTES;FechaNacimiento;1500@1;ESTADO_CIVIL;EstadoCivil;1000@18;CLIENTES;Ruc;1500@19;CLIENTES;Observaciones;5500@0;CLIENTES;C�digo;750@"
            frmBusq.propTablasRelacionadas = "8;PAISES;0@13;SEXO;0@14;NACIONALIDADES;0@16;ESTADO_CIVIL;0@"
            frmBusq.propTituloFormulario = "Clientes"
        
        Case 2  'solo pasajeros hospedados
            frmBusq.propNroCampoInicial = 2          'por defecto ordeno por nombre completo
            frmBusq.propIndiceCampoRetorno = 21      'retorno el c�digo del cliente
            frmBusq.propTabla = "CLIENTES"
            frmBusq.propCampos = "0;CHECKIN;Habitaci�n;1500@1;CLIENTES;NombreCompleto;4500@2;CLIENTES;PrimerNombre;1500@3;CLIENTES;SegundoNombre;1500@4;CLIENTES;PrimerApellido;1500@5;CLIENTES;SegundoApellido;1500@17;CLIENTES;Documento;1500@6;CLIENTES;Direcci�n;3000@7;CLIENTES;Localidad;1500@1;PAISES;Pa�sResidencia;1500@9;CLIENTES;C�digoPostal;1500@10;CLIENTES;Tel�fono;1500@11;CLIENTES;Fax;1500@20;CLIENTES;Email;1500@12;CLIENTES;OtrosTelFax;1500@1;SEXO;Sexo;1000@1;NACIONALIDADES;Nacionalidad;1500@15;CLIENTES;FechaNacimiento;1500@1;ESTADO_CIVIL;EstadoCivil;1000@18;CLIENTES;Ruc;1500@19;CLIENTES;Observaciones;5500@0;CLIENTES;C�digo;750@"
            frmBusq.propTablasRelacionadas = "8;PAISES;0@13;SEXO;0@14;NACIONALIDADES;0@16;ESTADO_CIVIL;0@" & "0;CHECKIN;1@"
            'Selecciono solo los pasajeros hospedados
            frmBusq.propSeleccionComplementaria = " CLIENTES.nrocorr = CHECKIN.nrocorrcli "
            frmBusq.propTituloFormulario = "Pasajeros alojados actualmente en el hotel."
            
        Case 3  'empresas
            frmBusq.propNroCampoInicial = 1         'por defecto ordeno por descripci�n
            frmBusq.propIndiceCampoRetorno = 8      'retorno el c�digo de empresa
            frmBusq.propTabla = "EMPRESAS"
            frmBusq.propCampos = "1;EMPRESAS;Nombre;2500@2;EMPRESAS;Raz�nSocial;2500@3;EMPRESAS;Ruc;1500@4;EMPRESAS;Direcci�n;2500@5;EMPRESAS;Tel�fono;1500@6;EMPRESAS;Fax;1500@7;EMPRESAS;Email;1500@8;EMPRESAS;Contacto;2000@0;EMPRESAS;C�digo;750@"
            frmBusq.propTituloFormulario = "Empresas"
            
        Case 4  'documentos (facturas y devoluciones)
            frmBusq.propNroCampoInicial = 0         'por defecto ordeno por n�mero de documento
            frmBusq.propIndiceCampoRetorno = 0      'retorno el n�mero del documento
            frmBusq.propTabla = "FAC_CABEZAL"
            frmBusq.propCampos = "1;FAC_CABEZAL;NroDocumento;1500@2;FAC_CABEZAL;Fecha;1500@3;FAC_CABEZAL;Nombre;4500@4;FAC_CABEZAL;Direcci�n;3000@5;FAC_CABEZAL;Localidad;1500@10;FAC_CABEZAL;Total;1500@"
            'Selecciono solo los pasajeros hospedados
            frmBusq.propSeleccionComplementaria = " tipo_docu = " & tipodocu
            frmBusq.propTituloFormulario = "Documentos de tipo: " & mFunDescripcionTipoDocu(tipodocu)
    
        Case 5  'recivos
            frmBusq.propNroCampoInicial = 1         'por defecto ordeno por n�mero
            frmBusq.propIndiceCampoRetorno = 1      'retorno el n�mero del recivo
            frmBusq.propTabla = "RECIVOS"
            frmBusq.propCampos = "0;RECIVOS;TipoRecivo;1500@1;RECIVOS;NroRecivo;1500@2;RECIVOS;Fecha;1500@3;RECIVOS;RealizadoA;4500@5;RECIVOS;Importe;1500@1;MONEDAS;Moneda;1500@"
            frmBusq.propTablasRelacionadas = "6;MONEDAS;0@"
            'Selecciono solo los recivos de un tipo determinado
            frmBusq.propSeleccionComplementaria = " tipo_recivo = " & tipodocu
            frmBusq.propTituloFormulario = "Recivo"
            
        Case 6  'art�culos
            frmBusq.propNroCampoInicial = 1         'por defecto ordeno por descripci�n
            frmBusq.propIndiceCampoRetorno = 0      'retorno el c�digo de art�culo
            frmBusq.propTabla = "ARTICULOS"
            frmBusq.propCampos = "0;ARTICULOS;C�digo;750@1;ARTICULOS;Descripci�n;3500@1;PUNTO_VENTA;PuntoDeVenta;3000@1;MONEDAS;Moneda;1500@1;IVA;TipoIVA;1000@5;ARTICULOS;PrecioSinIVA;1100@"
            frmBusq.propTablasRelacionadas = "4;PUNTO_VENTA;0@3;MONEDAS;0@2;IVA;0@"
            frmBusq.propTituloFormulario = "Art�culos"
            
        Case 7  'muestra solo los pasajeros no alojados
            frmBusq.propNroCampoInicial = 1          'por defecto ordeno por nombre completo
            frmBusq.propIndiceCampoRetorno = 20      'retorno el c�digo del cliente
            frmBusq.propTabla = "CLIENTES"
            frmBusq.propCampos = "1;CLIENTES;NombreCompleto;4500@2;CLIENTES;PrimerNombre;1500@3;CLIENTES;SegundoNombre;1500@4;CLIENTES;PrimerApellido;1500@5;CLIENTES;SegundoApellido;1500@17;CLIENTES;Documento;1500@6;CLIENTES;Direcci�n;3000@7;CLIENTES;Localidad;1500@1;PAISES;Pa�sResidencia;1500@9;CLIENTES;C�digoPostal;1500@10;CLIENTES;Tel�fono;1500@11;CLIENTES;Fax;1500@20;CLIENTES;Email;1500@12;CLIENTES;OtrosTelFax;1500@1;SEXO;Sexo;1000@1;NACIONALIDADES;Nacionalidad;1500@15;CLIENTES;FechaNacimiento;1500@1;ESTADO_CIVIL;EstadoCivil;1000@18;CLIENTES;Ruc;1500@19;CLIENTES;Observaciones;5500@0;CLIENTES;C�digo;750@"
            frmBusq.propTablasRelacionadas = "8;PAISES;0@13;SEXO;0@14;NACIONALIDADES;0@16;ESTADO_CIVIL;0@"
            frmBusq.propTituloFormulario = "Clientes no alojados actualmente en el hotel."
            'no muestro los pasajeros ya hospedados
            frmBusq.propSeleccionComplementaria = " CLIENTES.nrocorr NOT IN " & _
            "(Select CLIENTES.nrocorr " & _
            " from CLIENTES,CHECKIN " & _
            " where CLIENTES.nrocorr = CHECKIN.nrocorrcli) "
            
        Case 8  'todas las habitaciones del hotel
            'NOTA: si la habitaci�n no tiene situaci�n establecida no se muestra. Esto
            'no tendr�a por que pasar.
            
            'modifico el largo del formulario para la ayuda de habitaciones
            frmBusq.propLargo = 9000
            frmBusq.propNroCampoInicial = 0          'por defecto ordeno por n�mero de habitaci�n
            frmBusq.propIndiceCampoRetorno = 0       'retorno el n�mero de habitaci�n
            frmBusq.propTabla = "HABITACIONES"
            frmBusq.propCampos = "0;HABITACIONES;Habitaci�n;1500@1;TIPO_HABITACIONES;TipoHabitaci�n;2500@2;TIPO_ESTADO_HAB;Situaci�n;2500@9;HABITACIONES;Tarifa;1500@"
            frmBusq.propTablasRelacionadas = "1;TIPO_HABITACIONES;0@10;TIPO_ESTADO_HAB;1@"
            frmBusq.propTituloFormulario = "Habitaciones del hotel."
            'el join con el archivo TIPO_ESTADO_HAB debe der ser con los registros de tipo 2 (situaciones)
            frmBusq.propSeleccionComplementaria = " TIPO_ESTADO_HAB.tipo_cod = 2 "
            
        Case 9  'habitaciones del hotel ocupadas
            'Muestro todos los pasajeros hospedados del hotel
            frmBusq.propNroCampoInicial = 0          'por defecto ordeno por n�mero de habitaci�n
            frmBusq.propIndiceCampoRetorno = 0       'retorno el n�mero de habitaci�n
            frmBusq.propTabla = "CHECKIN"
            frmBusq.propCampos = "0;HABITACIONES;Habitaci�n;1500@1;CLIENTES;NombrePasajero;4500@"
            frmBusq.propTablasRelacionadas = "0;HABITACIONES;0@1;CLIENTES;0@"
            frmBusq.propTituloFormulario = "Habitaciones del hotel ocupadas."
            'Las limitaciones del control de selecci�n me impiden mostrar el tipo de habitaci�n
            'y la situaci�n de la misma, porque el join solo se realiza con la tabla CHECKIN
            'y no hay posibilidad de realizarlo con m�s de una tabla.
    End Select
    'muestro el formulario
    frmBusq.Show 1
    mFunBusqueda = frmBusq.propRetorno
    Unload frmBusq
End Function

Public Function mFunBuscarReserva(tipoAccion As Byte) As String
    'Llamo al formulario de b�squeda de reservas
    
    'cargo propiedad que indica el tipo de acci�n a realizar
    frmBusqReservas.propTipoAccion = tipoAccion
    
    frmBusqReservas.Show 1
    mFunBuscarReserva = frmBusqReservas.propRetorno
    Unload frmBusqReservas
End Function

