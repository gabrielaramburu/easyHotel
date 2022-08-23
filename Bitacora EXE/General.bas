Attribute VB_Name = "General"
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

Public Sub mSubSeleccionoTodos(lst As ListBox)
    'recorro el listbox y selecciono todos los
    'elementos
    Dim i As Integer
    i = 0
    Do While i < lst.ListCount
        lst.Selected(i) = True
        i = i + 1
    Loop
End Sub

Public Sub mSubLimpioLista()
    'Inicializo la lista para mostrar un nuevo listado
    frmMain.lwOpr.ColumnHeaders.Clear
    frmMain.lwOpr.ListItems.Clear
End Sub

Public Sub subGraboConfActual()
    'Grabo la configuración actual del sistema para
    'que se permanesca la próxima vez que ingrese
    tbSISTEMA_BITACORAparametros.Edit
    If frmMain.mnuVerListados.Checked Then
        'muestro cuadro listado al iniciar aplicación
        tbSISTEMA_BITACORAparametros("cuadroListados") = 1
    Else
        'no lo muestro
        tbSISTEMA_BITACORAparametros("cuadroListados") = 0
    End If
    tbSISTEMA_BITACORAparametros.Update
End Sub

Public Sub mSubMuestroListadoActual(listado As String)
    'Muestra en la barra de estado el último
    'listado ejecutado (actual)
    Dim imagen As Byte
    If funListadoPrede(listado) Then
        imagen = 9
    Else
        imagen = 10
    End If
    frmMain.StatusBar1.Panels(2).Picture = _
        frmMain.ImageList1.ListImages(imagen).Picture
    frmMain.StatusBar1.Panels(2).Text = listado
End Sub

Public Function funExisteOtraInstancia() As Boolean
    'Determino si ya hay una instancia de la aplicación ejecutándose.
    Dim msg As String
    If App.PrevInstance Then
        msg = App.EXEName & ".EXE" & " ya está en ejecución"
        MsgBox msg, 16, "Aplicación."
        funExisteOtraInstancia = True
    Else
        'no existe ninguna instancia
        funExisteOtraInstancia = False
    End If
End Function

