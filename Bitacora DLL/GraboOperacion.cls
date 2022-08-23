VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GraboOperacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private tablaBit As Recordset
Public Sub GraboOperacionEnBaseDeDatos(Fecha As Date, _
                                        Usuario As String, _
                                        Operacion As Integer, _
                                        HoraIni As String, _
                                        HoraFin As String, _
                                        Observaciones As String, _
                                        ArchivoBit As Recordset)
                                        
    'Graba la operación realizada en el archivo sistema_bitacora
    
    Dim NroCorrBit As Single
    Dim nomUsuarioAux As String
    
    'Para poder trabajar con el archivo desde todos los procedimientos
    'lo asigno a una variable general.
    Set tablaBit = ArchivoBit
    
    NroCorrBit = funObtengoNroCorrBit(Fecha)
    
    'Verifico el nombre del usuario.
    'Si se ejecuta la aplicación principal sin tener definidos los perfiles de usuario,
    'el parámetro usuario viene en blanco.
    'Por ese motivo se inicializa aquí con el valor "Usuario por defecto".
    If Trim(Usuario) = Empty Then
        'para no asignarle un valor directamente al parámetro
        'trabajo con una variable auxiliar
        nomUsuarioAux = "Usuario por defecto."
    Else
        nomUsuarioAux = Usuario
    End If
    
    'agrego nuevo registro a la tabla sistema_bitacora
    tablaBit.AddNew
        tablaBit("FechaBit") = Fecha
        tablaBit("NroCorrBit") = NroCorrBit
        tablaBit("NomUsrBit") = nomUsuarioAux
        tablaBit("CodOprBit") = Operacion
        tablaBit("HoraIniBit") = HoraIni
        tablaBit("HoraFinBit") = HoraFin
        tablaBit("ObservacionBit") = Observaciones
    tablaBit.Update
End Sub
                                        
Private Function funObtengoNroCorrBit(Fechas As Date)
    'Obtengo el siguiente número correlativo para una
    'fecha determinada
    
    funObtengoNroCorrBit = 1
    tablaBit.Index = "pk_bitacora"  'fecha,nrocorr
    tablaBit.Seek ">=", Fechas, 0
    If Not tablaBit.NoMatch Then    'si existe
        Do While Not tablaBit.EOF
            If tablaBit("FechaBit") = Fechas Then
                funObtengoNroCorrBit = tablaBit("NroCorrBit") + 1
            Else
                Exit Do
            End If
            tablaBit.MoveNext
        Loop
    End If
End Function

                                        
                                        
                

