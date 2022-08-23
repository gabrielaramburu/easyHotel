Attribute VB_Name = "MuestroDerecha"
Option Explicit
'Contiene procedimientos que trabajan con el listview

Public Sub mSubMuestroOperaciones()
    'Muestro en la lista todas las operaciones
    'disponibles en el sistema
    
    'realizo cabezal de operaciones
    frmMain.lwDerecha.ColumnHeaders.Add , "DesOpr", "Descripción", 6000
    
    'recorro operaciones
    tbSISTEMA_OPERACIONES.MoveFirst
    tbSISTEMA_OPERACIONES.Index = "pk_operaciones"
    Do While Not tbSISTEMA_OPERACIONES.EOF
        If Not tbSISTEMA_OPERACIONES.NoMatch Then
            'creo reglon de la lista
            subCreoLineaOpr
        End If
        tbSISTEMA_OPERACIONES.MoveNext
    Loop
    frmMain.lwDerecha.View = m_TipoLista
End Sub

Public Sub mSubMuestroUsuarios()
    'Muestro en la lista todos los usuarios
    'disponibles en la lista de usuarios
    Dim existenUsr As Boolean
    Dim itmX As ListItem
    
    'por defecto asumo que no existen usuarios
    existenUsr = False
    
    'realizo cabezal de usuario
    frmMain.lwDerecha.ColumnHeaders.Add , "NomUsr", "Nombre ", 6000
    'recorro archivo de usuarios
    tbSISTEMA_USUARIOS.Index = "iclaves"
    tbSISTEMA_USUARIOS.Seek ">=", ""
    Do While Not tbSISTEMA_USUARIOS.EOF
        If Not tbSISTEMA_USUARIOS.NoMatch Then
            'creo reglon de la lista
            subCreoLineaUsr
            'tengo 1 o más usuarios
            existenUsr = True
        End If
        tbSISTEMA_USUARIOS.MoveNext
    Loop
    'verifico si tengo usuarios
    If Not existenUsr Then
        'muestro línea de no usuarios
        Set itmX = frmMain.lwDerecha.ListItems.Add(, , "No existen usuarios defeinidos")
    End If
    frmMain.lwDerecha.View = m_TipoLista
End Sub

Private Sub subCreoLineaUsr()
    'Crea una nueva linea en la lista de usuarios
    Dim itmX As ListItem
    Set itmX = frmMain.lwDerecha.ListItems.Add(, , tbSISTEMA_USUARIOS("NomUsr"), , 5)
End Sub

Private Sub subCreoLineaOpr()
    'Crea una nueva linea en la lista de operaciones.
    
    Dim itmX As ListItem
    Dim imagen As Byte
    
    If tbSISTEMA_OPERACIONES("tipoOpr") = 1 Then
        imagen = 2
    Else
        imagen = 3
    End If
    Set itmX = frmMain.lwDerecha.ListItems.Add(, , tbSISTEMA_OPERACIONES("DescOpr"), , imagen)
End Sub

Public Sub mSubMuestroUsuariosPermitidos()
    'Dada una operación, muestro los usuarios que están habilitados
    'para trabajar con ella.
     
    Dim Opr As Integer
     
    'creo cabezal de la lista
    frmMain.lwDerecha.ColumnHeaders.Add , "NomUsr", "Usuarios", 5000
      
    Opr = Val(Mid(frmMain.twUsuarios.SelectedItem.Key, 4))
    'recorro el archivo de perfiles y
    'obtengo los usuarios habilitados para la operación
    tbSISTEMA_PERFILES.Index = "pk_perfiles"
    tbSISTEMA_PERFILES.Seek ">=", Opr, ""
    If Not tbSISTEMA_PERFILES.NoMatch Then
        Do While Not tbSISTEMA_PERFILES.EOF
            If tbSISTEMA_PERFILES("CodOpr") = Opr Then
                'Cargo usuarios a la lista
                subCreoLineaUsrAuto
            Else
                Exit Do
            End If
            tbSISTEMA_PERFILES.MoveNext
        Loop
    End If
    frmMain.lwDerecha.View = m_TipoLista
End Sub

Public Sub mSubMuestroOperacionesPermitidas()
    'Dada un usuario, muestro las operaciones que están habilitados
    'para trabajar con él.
    
    Dim Usr As String
     
    'creo cabezal de la lista
    frmMain.lwDerecha.ColumnHeaders.Add , "DescOpr", "Operaciones", 6000
    
     
    Usr = frmMain.twUsuarios.SelectedItem.Key
    'recorro el archivo de perfiles y
    'obtengo las operaciones permitidas para ese usuario
    tbSISTEMA_PERFILES.Index = "i_NomUsr"
    tbSISTEMA_PERFILES.Seek ">=", Usr
    If Not tbSISTEMA_PERFILES.NoMatch Then
        Do While Not tbSISTEMA_PERFILES.EOF
            If tbSISTEMA_PERFILES("NomUsr") = Usr Then
                'Cargo usuarios a la lista
                subCreoLineaOpePermitida
            Else
                Exit Do
            End If
            tbSISTEMA_PERFILES.MoveNext
        Loop
    End If
    frmMain.lwDerecha.View = m_TipoLista
End Sub

Private Sub subCreoLineaUsrAuto()
    'Creo linea de usuarios autorizados
    'para una determinada operación
    
    Dim itmX As ListItem
    Set itmX = frmMain.lwDerecha.ListItems.Add(, , tbSISTEMA_PERFILES("NomUsr"), , 7)
End Sub

Private Sub subCreoLineaOpePermitida()
    'Creo linea de operaciones permitidas
    'para un determinado usuario
    
    Dim imagen As Byte
    'busco tipo de operación para poder determinar el tipo
    
    If funBuscoOperacionTF(tbSISTEMA_PERFILES("CodOpr")) Then
        If tbSISTEMA_OPERACIONES("tipoOpr") = 1 Then
            imagen = 2
        Else
            imagen = 3
        End If
        
        Dim itmX As ListItem
        Set itmX = frmMain.lwDerecha.ListItems.Add(, , tbSISTEMA_OPERACIONES("DescOpr"), , imagen)
    End If
End Sub
