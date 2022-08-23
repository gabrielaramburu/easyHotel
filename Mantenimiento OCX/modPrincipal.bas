Attribute VB_Name = "modPrincipal"
Option Explicit
'NOTA: cualquier error inseperado qu surja en alguno de estos procedimientos
'será interceptado por el on error del procedimiento que los llama.

Public Sub mSubValidoNum(KeyAscii As Integer, Menos As Boolean, Punto As Boolean)
    Select Case KeyAscii
        Case 48 To 57       ' Permite los dígitos
        Case 8              ' Permite el carácter de retroceso
        Case 46 And Punto   ' Permite el punto
        Case 45 And Menos
        Case Else
            KeyAscii = 0
            Beep
    End Select
End Sub

Public Sub mSubLimpioCombo(combo As ComboBox)
    'borro todos los elementos de un combo
    On Error GoTo error
    Dim i As Integer
    i = combo.ListCount - 1
    Do While i >= 0
        combo.RemoveItem i
        i = i - 1
    Loop
Exit Sub
error:
    mSubControloErroresPropiedades "mSubLimpioCombo"
End Sub

Public Sub mSubRangoCeldas(grilla As MSFlexGrid, coli As Integer, _
                                                    colf As Integer, _
                                                    rowi As Integer, _
                                                    rowf As Integer)
    'Establece un rango de celdas seleccionadas en una grilla.
    On Error GoTo error
    grilla.FillStyle = flexFillRepeat
    grilla.Col = coli
    grilla.Row = rowi
    grilla.ColSel = colf
    grilla.RowSel = rowf
Exit Sub
error:
    mSubControloErroresPropiedades "mSubRangoCeldas"
End Sub

Public Sub mSubBajarLinea(g As MSFlexGrid)
    'Bajo una fila la fila seleccionada
    'La idea es que despues de crear un campo en la grilla se pueda cambiar la posicion
    'en la que aparece en el control
    On Error GoTo error
    Dim nuevaFila As Integer
    nuevaFila = g.Row + 1
    'valido que no se valla del rango de filas
    If nuevaFila < g.Rows Then
        g.RowPosition(g.Row) = nuevaFila
        'selecciono la fila en la nueva posición
        mSubRangoCeldas g, 0, g.Cols - 1, nuevaFila, nuevaFila
    End If
Exit Sub
error:
    mSubControloErroresPropiedades "mSubBajarLinea"

End Sub

Public Sub mSubSubirLinea(g As MSFlexGrid)
    'Subo una fila la fila seleccionada
    'La idea es que despues de crear un campo en la grilla se pueda cambiar la posicion
    'en la que aparece en el control
    On Error GoTo error
    Dim nuevaFila As Integer
    nuevaFila = g.Row - 1
    'valido que no se valla del rango de filas
    If nuevaFila >= 2 Then
        g.RowPosition(g.Row) = nuevaFila
        'selecciono la fila en la nueva posición
        mSubRangoCeldas g, 0, g.Cols - 1, nuevaFila, nuevaFila
    End If
Exit Sub
error:
    mSubControloErroresPropiedades "mSubSubirLinea"

End Sub

Public Sub mSubLimpioGrilla(grilla As MSFlexGrid, eliminoCol As Boolean)
    'Inicializo grilla
    'Elimino filas
    On Error GoTo error
    grilla.Row = grilla.Rows - 1
    Do While grilla.Rows > 2
        grilla.RemoveItem (grilla.Row)
    Loop
    If eliminoCol Then
        'elimino las columnas
        grilla.Cols = 1
    End If
Exit Sub
error:
    mSubControloErroresPropiedades "mSubLimpioGrilla"
End Sub

Public Function mfunObtengoValorDesdeStr(cadena As String, PosValor As Byte, caracterDiv As String) As String
    'Dado un string, devuelve los caracters que se encuentran entre los dos
    'signos divisorios inmediatos, a partir del número de indice que se  especifíque en posvalor.
    'Ej:    cadena ="aa;bb;cc;dd" ---> se transforma en ";aa;bb;cc;dd;"
    '       funObtengoValorDesdeStr(cadena,3,";") ---> cc
    '       funObtengoValorDesdeStr(cadena,4,";") ---> dd
    '       funObtengoValorDesdeStr(cadena,1,";") ---> aa
    
    'para simplificar el procedimeinto agrego a al cadena punto y coma al principio
    'y final de la misma
    On Error GoTo error
    Dim caracter As String
    Dim Valor As String
    Dim i As Integer
    Dim largo As Integer
    Dim contCampos As Integer
    Dim FinRecorrido As Boolean
    Dim cadenaAux As String
    
    cadenaAux = ";" & cadena & ";"
    largo = Len(cadenaAux)
    contCampos = 0
    FinRecorrido = False
    i = 1
    Do While (i <= largo) And Not FinRecorrido
        'proceso cada caracter de la cadena por vez
        caracter = Mid(cadenaAux, i, 1)
        If caracter = ";" Then
            contCampos = contCampos + 1
        End If
        'verifico si comienza el campo que me interesa
        If contCampos = PosValor Then
            'comienzo a obtener el valor
            i = i + 1                           'dejo de lado el caracter de ;
            Do While (i <= largo) And Not FinRecorrido
                caracter = Mid(cadenaAux, i, 1)    'obtengo sigiente caracter despues
                                                'del punto y coma
                If caracter = ";" Then          'termina el campo que me interesa
                    FinRecorrido = True
                Else
                    Valor = Valor & caracter
                End If
                i = i + 1
            Loop
        End If
        i = i + 1
    Loop
    mfunObtengoValorDesdeStr = Valor
Exit Function
error:
    mSubControloErroresPropiedades "mfunObtengoValorDesdeStr"

End Function

Public Sub mSubPosicionoCombo(combo As ListBox, indice As Long)
    'recorre todo el combo y muestra la posición
    'que corresponda.
    On Error GoTo error
    Dim i As Integer
    i = 0
    For i = 0 To combo.ListCount - 1
        If combo.ItemData(i) = indice Then
            combo.Text = combo.List(i)
            combo.Selected(i) = True
            Exit For
        End If
    Next i
Exit Sub
error:
    mSubControloErroresPropiedades "mSubPosicionoCombo"

End Sub

Public Sub mSubControloErroresPropiedades(Propdesde As String)
    'Muestro mensaje indicando que ha ocurrido un error
    'Este mensaje le indicará al programador que esta utilizando el control,
    'que dicho control ha realizado una operación no válida, debido
    'a problemas de hardware, o problmas generados por mál funcionamiento del ocx
    'en sí, que se originan al realizar el usuario operaciones inseperadas que no se tuvieron
    'en cuenta por parte del diseñador del ocx. (o sea yo)
    Dim mensaje As String
    mensaje = Err.Number & " " & Err.Description & Chr(10) & _
            "Error producido en " & Propdesde & Chr(10) & _
              "Consulte al proveedor del ocx"
    MsgBox mensaje, vbCritical
End Sub



