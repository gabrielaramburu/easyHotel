Attribute VB_Name = "LogicaDeDisponibilidad"
Option Explicit

Public Function habitacion_reservada(hab As Long, fd_par As Date, fh_par As Date)
    'Determino si la habitación esta reservada en un perído de fechas.
    
    'Devuelve TRUE si esta reservada.
    
    Dim fd As Date
    Dim fh As Date

    habitacion_reservada = False
    tbHAB_RESERVAS.Index = "ihab_reserva_fecha"
    tbHAB_RESERVAS.Seek ">", hab, fd_par
    'Estudiando este procedimiento el 28/04/02 llegué a la conclusión que la condición ">="
    'puede cambiarse por ">", para dejar fuera de la seleción a las reservas que se van
    'en la fecha desde. fd_par. Esto implicaría un aumento en la eficiencia del procedimiento.
    If Not tbHAB_RESERVAS.NoMatch Then
        Do While Not tbHAB_RESERVAS.EOF
            If tbHAB_RESERVAS("nrohabitacion") = hab Then
                fd = tbHAB_RESERVAS("fechaing")
                fh = tbHAB_RESERVAS("fechaegr")
                
                'ACLARACIÓN:    10/06/01
                'Puede conciderarse hacer estos dos controles redundantes,
                'ya que una reserva No show vigente (las No show no vigente no se recorren)
                ',no se tendrá en cuenta de todas maneras,
                'ya que entra dentro del tipo de Vigente ocupadas
                
                'Si la reserva cayó noshow no la tomo en cuenta
                If tbHAB_RESERVAS("noshow") = 0 Then    'no noshow
                    'Si la reserva es del tipo Vigente Ocupada tampoco la tomo en cuanta
                    'ya que en ese caso se toma en cuenta la ocupación
                    If fd < m_FechaSis Then
                    Else
                        'Caso 1
                        If fd_par < fd Then
                            If fh_par > fd Then
                                habitacion_reservada = True
                                Exit Function
                            End If
                        End If
                        
                        'Caso 2
                        If fd_par >= fd Then
                            habitacion_reservada = True
                            Exit Function
                        End If
                    End If
                End If
            Else
                Exit Do
            End If
            tbHAB_RESERVAS.MoveNext
        Loop
    End If
End Function

Public Function habitacion_bloqueada(habi As Long, fd_par As Date, fh_par As Date)
    'Determino si la habitación esta bloqueada en un perído de fechas.
    'Devuelve TRUE si esta bloqueada.

    Dim fd As Date
    Dim fh As Date
    
    habitacion_bloqueada = False
    tbBLOQUEO_HAB.Index = "i_bloq_fh"
    tbBLOQUEO_HAB.Seek ">", habi, fd_par
    Do While Not tbBLOQUEO_HAB.EOF
        If Not tbBLOQUEO_HAB.NoMatch Then
            If tbBLOQUEO_HAB("hab_bloq") = habi Then
                fd = tbBLOQUEO_HAB("FDesdeBloq")
                fh = tbBLOQUEO_HAB("FHastaBloq")
                'Caso 1
                If fd_par < fd Then
                    If fh_par > fd Then
                        habitacion_bloqueada = True
                        Exit Function
                    End If
                End If
                'Caso 2
                If fd_par >= fd Then
                    habitacion_bloqueada = True
                    Exit Function
                End If
            Else
                Exit Do
            End If
        End If
        tbBLOQUEO_HAB.MoveNext
    Loop
End Function

Public Function habitacion_ocupada(hab As Long, fd_par As Date)
    'Determino si la habitación esta ocupada en el período comprendido entre
    'fecha de hoy y fecha des.
    'Esta función se diferencia de busco_habita_checkin, en que determina la ocupación
    'dependiendo de la fecha de inicio del nuevo período, mientras que busco_habita_checkin
    'determina si la habitación esta ocupada actualmente.
    '---------------------------------------------------------------------------------
    'Parámetros.
    '   Entrada:
    '       [hab]       habitación con la cual estoy trabajando
    '       [fd_par]    fecha desde, del nuevo período de trabajo
    '                   (reserva, ocupación, bloqueado)
    '   Salida:
    '                   True, si esta ocupada.
    '                   False, si no esta ocupada
    '
    ' NOTA: 1) En el caso de un walkin y de un checkin "No Asignada"
    '       la fechas des = a fecha de hoy por lo tanto se cumple
    '       que la habitación esta ocupada actualmente
    
    '       2) Es muy importante determinar si la habitación esta o no dentro del período de alojamiento.
    '----------------------------------------------------------------------------------
    
    Dim fh As Date
    
    habitacion_ocupada = False
    tbCHECKIN.Index = "i_habitacion"
    tbCHECKIN.Seek "=", hab
    'me posiciono en el primer pasajero de la habitación
    If Not tbCHECKIN.NoMatch Then
        fh = tbCHECKIN("fcheckhas")
        
        'Determino si la habitación está dentro del período de alojamiento.
        'La fh siempre será mayor igual a la fecha de hoy
        'eseptuando las habitaciones en la que no se ha realizado el chekout
        '(fuera del período de alojamiento). En dicho caso fh será menor a la fecha
        'de hoy.
        If fh < m_FechaSis Then
            'la habitación está fuera del período de alojamiento
            'verifico si la fecha desde del nuevo período es igual al día de hoy
            If fd_par = m_FechaSis Then
                'determino que la habitación esta actualmente ocupada.
                habitacion_ocupada = True
            Else
                'para los futuros días (fecha desde > fecha de hoy) determino que
                'la habitación esta disponible, ya que hay tiempo de que se le ralize
                'un checkout a la habitación fuera del período.
            End If
        Else
            'la habitación está dentro del período de alojamiento
            If fd_par < fh Then 'caso 1
                habitacion_ocupada = True
            End If
        End If
    End If
End Function


