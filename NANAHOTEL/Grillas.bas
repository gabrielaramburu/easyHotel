Attribute VB_Name = "Grillas"
Option Explicit

Public Sub limpio_grilla(grilla As MSFlexGrid)
    'Inicializo grilla
    'Elimino filas
    grilla.row = grilla.Rows - 1
    Do While grilla.Rows > 2
        grilla.RemoveItem (grilla.row)
    Loop
    'elimino las columnas
    grilla.Cols = 1
End Sub

Public Sub marco_celdas_grilla(grilla As MSFlexGrid, coli As Integer, colf As Integer, rowi As Integer, rowf As Integer)
    'Establece un rango de celdas seleccionadas en una grilla.
    grilla.FillStyle = flexFillRepeat
    grilla.col = coli
    grilla.row = rowi
    grilla.ColSel = colf
    grilla.RowSel = rowf
End Sub

Public Sub marco_grilla(grilla As MSFlexGrid, ini As Byte, fin As Byte)
    'Al hacer doble clik sobre la grilla se marca toda la fila
    ' y si ya esta marcada se desmarca.
    
    'ini, fin son la columna inicial y final
    Dim colorback As String
    Dim colorfore As String
    Dim col As Integer
    If grilla.CellBackColor = mSisColor_15FilaSeleccionada Then    'si esta marcado
        colorback = &H80000005                  'desmarco
        colorfore = &H80000008
    Else                                        'marco
        colorback = mSisColor_15FilaSeleccionada
        colorfore = mSisColor_19FilaSeleccionadaTexto
    End If
    
    For col = ini To fin
        grilla.col = col
        grilla.CellBackColor = colorback
        grilla.CellForeColor = colorfore
    Next
End Sub

Public Sub mSubLineaEnNegrita(grilla As MSFlexGrid, linea As Integer)
    'Marca la linea de la grilla en negrita
    marco_celdas_grilla grilla, 0, grilla.Cols - 1, linea, linea
    grilla.CellFontBold = True
End Sub

Public Sub mSubLineaComoTitulo(grilla As MSFlexGrid, linea As Integer)
    'Marca la linea de la grilla con el fondo más oscuro
    marco_celdas_grilla grilla, 0, grilla.Cols - 1, linea, linea
    grilla.CellBackColor = mSisColor_18ControlesNoHabilitados
End Sub

Public Sub mSubAparienciaGrilla(grilla As MSFlexGrid, diferencia As Integer)
    'Reparto el espacio de la grilla, uniformemente entre cada columna
    'el parámetro diferencia establece el espacio que hay que dejar para la barra de desplazamiento
    'vertical y el espacio de la columna 0.
    Dim anchoCol As Long
    Dim i As Long
    anchoCol = (grilla.Width - diferencia) / (grilla.Cols - 1)
    For i = 1 To grilla.Cols - 1
        grilla.ColWidth(i) = anchoCol
    Next
End Sub

Public Sub mSubMuestroIcono(grilla As MSFlexGrid, col As Long)
    'Cada vez que ordeno una grilla por uno de sus campos
    'dibujo un ícono en el campo (columna) correspondiente.
    
    Dim i As Long
    
    grilla.row = 0    'el icono se dibuja en el cabezal de cada campo
  
    
    'primero borro el icono de la columna anterior
    'para eso recorro todas las columnas y borro el icono
    For i = 1 To grilla.Cols - 1
        grilla.col = i
        grilla.CellPictureAlignment = 7    'muestro el ícono sobre la derecha
        Set grilla.CellPicture = frmMAIN.ImageList1.ListImages(4).Picture
    Next
    'dibujo icono
    grilla.col = col
    grilla.CellPictureAlignment = 7    'muestro el ícono sobre la derecha
    Set grilla.CellPicture = frmMAIN.ImageList1.ListImages(3).Picture
End Sub

Public Sub mSubMuestroIconoGrilla(grilla As MSFlexGrid, muestro As Boolean, centrado As Byte)
    'Muestro ícono solo si tiene más de una línea.
    grilla.row = 0
    grilla.col = 0
    'centro imagen en la celda
    grilla.CellAlignment = centrado
    If muestro Then
        'verifico cantidad de líneas antes de mostrar ícono
        If grilla.Rows > 1 Then
            Set grilla.CellPicture = frmMAIN.ImageList1.ListImages(5).Picture
        End If
    Else
        'muestro ícono en blanco
        Set grilla.CellPicture = frmMAIN.ImageList1.ListImages(6).Picture
    End If
End Sub

Public Function mFunctionValorCeldaMSFGrid(grilla As MSFlexGrid, _
                                            col As Integer, _
                                            row As Integer) As Variant

    '------------------------------------------------------------------------
    'Obtiene el valor de una celda determina de un control grilla de tipo
    'MsFlexGrid.
    '------------------------------------------------------------------------
    'Parámetros.
    '   Entrada.
    '       [grilla]    control grilla de tipo MsFlexGrid
    '       [col]       columna donde esta ubicada la celda
    '       [row]       fila donde esta ubicada la celda.
    '
    '   Saldia.         valor almacenado en la celda
    '
    'Nota: esta función no desencadena los eventos de cambio de celda, ya que
    'utiliza la sentencia TextMatrix
    '-------------------------------------------------------------------------
    On Error Resume Next
    mFunctionValorCeldaMSFGrid = _
    grilla.TextMatrix(row, col)
End Function

Public Sub subOrdenoGrillaMSFlex(grilla As MSFlexGrid, col As Integer)
    '--------------------------------------------------------------------------------
    'Ordeno grilla por una columna determinada.
    '--------------------------------------------------------------------------------
    'Parámetros.
    '   Entrada.    [grilla] puede ser la primer grilla (cuenta única)
    '                        la segunda grilla (cuentas separadas)
    '               [col]   columna por la cual quiero ordenar la grilla
    '--------------------------------------------------------------------------------
    On Error Resume Next
    grilla.col = col       'columna por la cual ordeno
    grilla.Sort = 5     'propiedades de la ordenación
End Sub

