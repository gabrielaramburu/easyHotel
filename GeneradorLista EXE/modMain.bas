Attribute VB_Name = "modMain"
'Este m�dulo se encuntra el procedimeinto que se ejecuta primero, al
'iniciar un programa.

Public Sub Main()
    On Error Resume Next
    'abro base de datos
    mSubAbroBaseDeDatos
    'muestro formulario
    frmGeneral.Show 1
End Sub
