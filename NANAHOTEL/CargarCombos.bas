Attribute VB_Name = "CargarCombos"

'tipos de documentos a emitir
Public tipo_documento(3) As String

Public Sub mSubcargo_combos_vectores()
    'carga tipo de documento a emitir
    tipo_documento(0) = "Factura M/N"
    tipo_documento(1) = "Factura U$S"
    tipo_documento(2) = "Contado M / N"
    tipo_documento(3) = "Contado U$S"
    
    'cargo estados de habitaciones
    vec_estados(0) = "Libre"
    vec_estados(1) = "Ocupada"
    vec_estados(2) = "Reservada"
    vec_estados(3) = "Bloqueada"
    vec_estados(4) = "No asignada"

End Sub

Sub carga_tipo_hab(combo As ComboBox)
    On Error Resume Next
    tbTIPO_HABITACIONES.Index = "i_tipo_hab"
    tbTIPO_HABITACIONES.MoveFirst
    Do While Not tbTIPO_HABITACIONES.EOF
        combo.AddItem tbTIPO_HABITACIONES("descripcion")
        combo.ItemData(combo.NewIndex) = tbTIPO_HABITACIONES("tipohab")
        tbTIPO_HABITACIONES.MoveNext
    Loop
End Sub

Sub carga_tipo_nacionalidad(combo As ComboBox)
    On Error Resume Next
    tbNACIONALIDADES.Index = "descri"
    tbNACIONALIDADES.MoveFirst
    Do While Not tbNACIONALIDADES.EOF
        combo.AddItem tbNACIONALIDADES("descri_nacionalidad")
        combo.ItemData(combo.NewIndex) = tbNACIONALIDADES("cod_nacionalidad")
        tbNACIONALIDADES.MoveNext
    Loop
End Sub

Sub carga_tipo_pais(combo As ComboBox)
    On Error Resume Next
    tbPAISES.Index = "descri"
    tbPAISES.MoveFirst
    Do While Not tbPAISES.EOF
        combo.AddItem tbPAISES("descri_pais")
        combo.ItemData(combo.NewIndex) = tbPAISES("cod_pais")
        tbPAISES.MoveNext
    Loop
End Sub

Sub carga_punto_venta(combo As ComboBox)
    On Error Resume Next
    tbPUNTO_VENTA.Index = "i_punto_venta"
    tbPUNTO_VENTA.MoveFirst
    Do While Not tbPUNTO_VENTA.EOF
        combo.AddItem tbPUNTO_VENTA("DescripcionPV")
        combo.ItemData(combo.NewIndex) = tbPUNTO_VENTA("nropv")
        tbPUNTO_VENTA.MoveNext
    Loop
End Sub

Sub carga_tipo_estado_hab(combo As ComboBox, tipo As Byte)
    'tipo 1 = motivos bloqueo
    'tipo 2 = situacion habitación
    tbTIPO_ESTADO_HAB.Index = "i_estado"
    tbTIPO_ESTADO_HAB.Seek ">=", tipo, 1
    Do While Not tbTIPO_ESTADO_HAB.EOF
        If tbTIPO_ESTADO_HAB("tipo_cod") = tipo Then
            combo.AddItem tbTIPO_ESTADO_HAB("descri")
            combo.ItemData(combo.NewIndex) = tbTIPO_ESTADO_HAB("cod")
            tbTIPO_ESTADO_HAB.MoveNext
        Else
            Exit Do
        End If
    Loop
End Sub

Public Sub carga_tipoIVA(combo As ComboBox)
    On Error Resume Next
    tbIVA.Index = "pk_iva"
    tbIVA.MoveFirst
    Do While Not tbIVA.EOF
        combo.AddItem tbIVA("DescIva")
        combo.ItemData(combo.NewIndex) = tbIVA("CodIva")
        tbIVA.MoveNext
    Loop
End Sub

Sub carga_tipo_moneda(combo As ComboBox)
    'Recorro el archivo MONEDAS y cargo el combo con esta información
    On Error Resume Next
    tbMONEDAS.Index = "pk_moneda"
    tbMONEDAS.MoveFirst
    Do While Not tbMONEDAS.EOF
        combo.AddItem tbMONEDAS("descMoneda")
        combo.ItemData(combo.NewIndex) = tbMONEDAS("codMoneda")
        tbMONEDAS.MoveNext
    Loop
End Sub

Sub carga_tipo_sexo(combo As ComboBox)
    'Recorro el archivo SEXO y cargo el combo con esta información
    On Error Resume Next
    tbSEXO.Index = "pk_sexo"
    tbSEXO.MoveFirst
    Do While Not tbSEXO.EOF
        combo.AddItem tbSEXO("descSexo")
        combo.ItemData(combo.NewIndex) = tbSEXO("codSexo")
        tbSEXO.MoveNext
    Loop
End Sub

Sub carga_tipo_estadocivil(combo As ComboBox)
    'Recorro el archivo ESTADO_CIVIL y cargo el combo con esta información
    On Error Resume Next
    tbESTADO_CIVIL.Index = "pk_estadoCivil"
    tbESTADO_CIVIL.MoveFirst
    Do While Not tbESTADO_CIVIL.EOF
        combo.AddItem tbESTADO_CIVIL("descEstadoCivil")
        combo.ItemData(combo.NewIndex) = tbESTADO_CIVIL("codEstadoCivil")
        tbESTADO_CIVIL.MoveNext
    Loop
End Sub

Sub mSubCargoComboTarjetasCredito(combo As ComboBox)
    'Cargo el combo que paso como parámetro con todas las tarjetas de crédito utilizadas
    'por el hotel.
    On Error Resume Next
    tbTARJETAS.Index = "pkTarjetas"
    tbTARJETAS.MoveFirst
    Do While Not tbTARJETAS.EOF
        combo.AddItem tbTARJETAS("descTarjeta")
        combo.ItemData(combo.NewIndex) = tbTARJETAS("codTarjeta")
        tbTARJETAS.MoveNext
    Loop
End Sub

Sub carga_tipo_docu(combo As ComboBox)
    Dim i As Byte
    i = 0
    Do While i < 3
       combo.AddItem tipo_documento(i)
       combo.ItemData(combo.NewIndex) = i
       i = i + 1
    Loop
End Sub

Sub limpio_combo(combo As ComboBox)
    combo.Clear
End Sub

Sub posiciono_combo(combo As ComboBox, indice As Long)
    'recorre todo el combo y muestra la posición
    'que corresponda.
    Dim i As Integer
    i = 0
    For i = 0 To combo.ListCount - 1
        If combo.ItemData(i) = indice Then
            combo.Text = combo.List(i)
            Exit For
        End If
    Next i
End Sub

Public Sub mSubCargoCombosFuentes(combo As ComboBox)
    'Recorro los fuentes instalados en el sistema
    Dim i As Integer
    For i = 0 To Screen.FontCount - 1
        combo.AddItem Screen.Fonts(i)
    Next i
End Sub

Public Sub mSubCargoComboMotivoAlojaManual(combo As ComboBox)
    '-------------------------------------------------------------------------------
    'Recorro todas las constantes de tipos de alojamiento y las cargo
    'en el combo de motivos de alojamiento manual.
    'Parámetros.
    '   Entrada:    [combo] control de tipo comboBox donde se cargan las constantes
    '
    'Nota: no incluyo el tipo 0, ya que el mismo solo se asigan automáticamente a los
    'alojamientos cargados automáticamente.(esta restricción me impide utilizar el
    'procedimiento genérico que carga a un combo las constantes de un tipo determinado.)
    '----------------------------------------------------------------------------------
    'declaración de variables de archivo
    Dim tbConst As Recordset
    Set tbConst = tbSISTEMA_CONSTANTES
    
    'recorro constantes de tipo 1 (tipos de alojamiento)
    tbConst.Index = "pkConst"
    tbConst.Seek ">=", 1, 0
    If Not tbConst.NoMatch Then
        Do While Not tbConst.EOF
            'cargo solo las contanstes de tipo 1
            If tbConst("tipoConst") = 1 Then 'tipos de alojamiento
                'no cargo la constante 1 ya que se usa solo al realizarse un
                'alojamiento automático en el proceso de cierre diario.
                If tbConst("codConst") > 1 Then
                    combo.AddItem tbConst("descConst")
                    combo.ItemData(combo.NewIndex) = tbConst("codConst")
                End If
            Else
                Exit Do
            End If
            tbConst.MoveNext
        Loop
    End If
    Set tbConst = Nothing
End Sub

Public Sub mSubCargoComboConstantes(tipoConst As Integer, combo As ComboBox)
    '----------------------------------------------------------------------------------
    'Carga a un combo box la descripción de las constantes establecidas de
    'un tipo determinado.
    '----------------------------------------------------------------------------------
    'En el archvio SISTEMA_CONSTANTES, se almacenan las constantes utilizadas por el
    'sistema, estas constantes bien podrían estar declaradas e inicializadas mediante
    'código, el principal motivo por el cual se implementa esta tabla es que muchas
    'consultas SQL, requieren conocer el valor de dichas contantes para poder ejecutar
    'consultas determinadas, facilitando de esta manera su implementación al existir
    'este archivo con estos datos, el cual se inicializa en la etapa de diseño.
    '-----------------------------------------------------------------------------------
    'Parámetros.
    '   Entrada [tipoConst] 1= almacena tipos de carga de alojamiento manual
    '                       2= tipos de documentos de clientes
    '
    '           [combo] nombre del combo en el cual se cargan las const.
    '------------------------------------------------------------------------------------
    'declaración de variables de archivo
    Dim tbConst As Recordset
    Set tbConst = tbSISTEMA_CONSTANTES
    
    tbConst.Index = "pkConst"
    tbConst.Seek ">=", tipoConst, 0
    If Not tbConst.NoMatch Then
        Do While Not tbConst.EOF
            'cargo solo las contanstes que paso como parámetros
            If tbConst("tipoConst") = tipoConst Then
                combo.AddItem tbConst("descConst")
                combo.ItemData(combo.NewIndex) = tbConst("codConst")
            Else
                Exit Do
            End If
            tbConst.MoveNext
        Loop
    End If
    Set tbConst = Nothing
End Sub
