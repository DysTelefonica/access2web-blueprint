Attribute VB_Name = "Módulo1"
Option Compare Database
Option Explicit

Public Function RegularizarCalibracionesEqMedida(Optional ByRef p_Error As String) As String
    
    
    Dim m_EventoEquipoMedidaAntes As EventoEquipoMedidaAntes
    Dim m_Col As Scripting.Dictionary
    Dim m_ID As Variant
    Dim m_Calibracion As EquipoMedidaCalibracion
    
    On Error GoTo errores
    
    Set m_Col = getEquiposMedidaEnEventosAntes1(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        Exit Function
    End If
    For Each m_ID In m_Col
        Set m_EventoEquipoMedidaAntes = m_Col(m_ID)
        Set m_Calibracion = m_EventoEquipoMedidaAntes.CalibracionEnAltaEvento
        p_Error = m_EventoEquipoMedidaAntes.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If Not m_Calibracion Is Nothing Then
            RegistrarCalibracionEnNuevaTbEVEQ m_ID, m_Calibracion.IDCalibracion, p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        End If
        Set m_EventoEquipoMedidaAntes = Nothing
    Next
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RegularizarCalibracionesEqMedida ha devuelto el error: " & Err.Description
    End If
    
End Function
Private Function RegistrarCalibracionEnNuevaTbEVEQ( _
                                                    m_ID As Variant, _
                                                    m_IDCalibracion As String, _
                                                    Optional ByRef p_Error As String _
                                                    ) As String
    
    
    Dim m_SQL As String
    
    
    On Error GoTo errores
    If CStr(m_ID) = "" Or m_IDCalibracion = "" Then
        Exit Function
    End If
    m_SQL = "UPDATE TbEventosEquipoMedida SET " & _
            "IDCalibracion =" & m_IDCalibracion & " " & _
            "WHERE IDEventoEquipoMedida=" & m_ID & ";"
    getdb().Execute m_SQL
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RegistrarCalibracionEnNuevaTbEVEQ ha devuelto el error: " & Err.Description
    End If
    
End Function
Private Function getEquiposMedidaEnEventosAntes( _
                                            Optional ByRef p_Error As String _
                                            ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    
    Dim m_EventoEquipoMedidaAntes As EventoEquipoMedidaAntes
    On Error GoTo errores
    
    m_SQL = "TbEventosEquipoMedida_antes"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_EventoEquipoMedidaAntes = New EventoEquipoMedidaAntes
            For Each m_Campo In m_EventoEquipoMedidaAntes.ColCampos
                m_EventoEquipoMedidaAntes.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            
            If getEquiposMedidaEnEventosAntes Is Nothing Then
                Set getEquiposMedidaEnEventosAntes = New Scripting.Dictionary
                getEquiposMedidaEnEventosAntes.CompareMode = TextCompare
            End If
            If Not getEquiposMedidaEnEventosAntes.Exists(m_EventoEquipoMedidaAntes.IDEquipoMedida) Then
                getEquiposMedidaEnEventosAntes.Add m_EventoEquipoMedidaAntes.IDEquipoMedida, m_EventoEquipoMedidaAntes
            End If
            Set m_EventoEquipoMedidaAntes = Nothing
            
            .MoveNext
        Loop
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getEquiposMedidaEnEventosAntes ha devuelto el error: " & Err.Description
    End If
End Function

Private Function getEquiposMedidaEnEventosAntes1( _
                                            Optional ByRef p_Error As String _
                                            ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    
    Dim m_EventoEquipoMedidaAntes As EventoEquipoMedidaAntes
    On Error GoTo errores
    
    m_SQL = "SELECT TbEventosEquipoMedida_antes.* " & _
            "FROM TbEventosEquipoMedida INNER JOIN TbEventosEquipoMedida_antes " & _
            "ON TbEventosEquipoMedida.IDEventoEquipoMedida = TbEventosEquipoMedida_antes.IDEventoEquipoMedida " & _
            "WHERE (((TbEventosEquipoMedida.IDCalibracion) Is Null));"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_EventoEquipoMedidaAntes = New EventoEquipoMedidaAntes
            For Each m_Campo In m_EventoEquipoMedidaAntes.ColCampos
                m_EventoEquipoMedidaAntes.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            
            If getEquiposMedidaEnEventosAntes1 Is Nothing Then
                Set getEquiposMedidaEnEventosAntes1 = New Scripting.Dictionary
                getEquiposMedidaEnEventosAntes1.CompareMode = TextCompare
            End If
            If Not getEquiposMedidaEnEventosAntes1.Exists(m_EventoEquipoMedidaAntes.IDEventoEquipoMedida) Then
                getEquiposMedidaEnEventosAntes1.Add m_EventoEquipoMedidaAntes.IDEventoEquipoMedida, m_EventoEquipoMedidaAntes
            End If
            Set m_EventoEquipoMedidaAntes = Nothing
            
            .MoveNext
        Loop
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getEquiposMedidaEnEventosAntes1 ha devuelto el error: " & Err.Description
    End If
End Function

