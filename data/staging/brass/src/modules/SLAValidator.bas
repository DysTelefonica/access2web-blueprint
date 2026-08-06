Attribute VB_Name = "SLAValidator"
Option Compare Database
Option Explicit

Private Function NormalizarSiNo(ByVal p_Valor As String) As String
    Dim m_Valor As String

    m_Valor = Trim$(Nz(p_Valor, ""))
    m_Valor = Replace(m_Valor, "í", "i")
    m_Valor = Replace(m_Valor, "Í", "I")
    m_Valor = UCase$(m_Valor)

    If m_Valor = "SI" Then
        NormalizarSiNo = "SI"
    ElseIf m_Valor = "NO" Then
        NormalizarSiNo = "NO"
    Else
        NormalizarSiNo = ""
    End If
End Function

Public Function ValidarCamposSLA( _
    ByVal p_EnFranqueo As Boolean, _
    ByVal p_FechaRecepcion As String, _
    ByVal p_FechaInicioContacto As String, _
    ByVal p_IncidenciaAveria As String, _
    ByVal p_FechaInicioAdquisicion As String, _
    ByVal p_FechaFinAdquisicion As String, _
    ByVal p_TipoReparacion As String, _
    ByVal p_Urgente As Variant, _
    ByVal p_EventoConServicioAfectado As String, _
    ByVal p_FechaRestablecimiento As String, _
    Optional ByRef p_Mensaje As String, _
    Optional ByRef p_Error As String) As String
    
    Dim m_CodError As String
    Dim m_Mensaje As String
    Dim m_Acumulador As String
    
    m_CodError = ""
    m_Mensaje = ""
    m_Acumulador = ""
    p_Error = ""
    
    On Error GoTo errores
    
    Dim m_FechaRecepcion As String
    Dim m_FechaInicioContacto As String
    Dim m_IncidenciaAveria As String
    Dim m_FechaInicioAdquisicion As String
    Dim m_FechaFinAdquisicion As String
    Dim m_TipoReparacion As String
    Dim m_Urgente As Variant
    Dim m_EventoConServicio As String
    Dim m_FechaRestablecimiento As String
    
    m_FechaRecepcion = p_FechaRecepcion
    m_FechaInicioContacto = p_FechaInicioContacto
    m_IncidenciaAveria = p_IncidenciaAveria
    m_FechaInicioAdquisicion = p_FechaInicioAdquisicion
    m_FechaFinAdquisicion = p_FechaFinAdquisicion
    m_TipoReparacion = p_TipoReparacion
    m_Urgente = p_Urgente
    m_EventoConServicio = NormalizarSiNo(p_EventoConServicioAfectado)
    m_FechaRestablecimiento = p_FechaRestablecimiento
    
    ' --- Validaciones generales (siempre aplican) ---
    
    If IsDate(m_FechaInicioContacto) And Not IsDate(m_FechaRecepcion) Then
        If m_CodError = "" Then m_CodError = "SLA-001"
        m_Mensaje = "No puede haber fecha de fin (Inicio contacto) sin fecha de inicio (Recepcion notificacion)"
        m_Acumulador = m_Acumulador & m_Mensaje & vbCrLf
    End If
    
    If IsDate(m_FechaFinAdquisicion) And Not IsDate(m_FechaInicioAdquisicion) Then
        If m_CodError = "" Then m_CodError = "SLA-001"
        m_Mensaje = "No puede haber fecha de fin (Fin adquisicion) sin fecha de inicio (Inicio adquisicion)"
        m_Acumulador = m_Acumulador & m_Mensaje & vbCrLf
    End If
    
    If IsDate(m_FechaRestablecimiento) And m_EventoConServicio <> "SI" Then
        If m_CodError = "" Then m_CodError = "SLA-001"
        m_Mensaje = "No puede haber fecha de restablecimiento si el evento no afecta al servicio"
        m_Acumulador = m_Acumulador & m_Mensaje & vbCrLf
    End If
    
    If IsDate(m_FechaRecepcion) And IsDate(m_FechaInicioContacto) Then
        If CDate(m_FechaInicioContacto) < CDate(m_FechaRecepcion) Then
            If m_CodError = "" Then m_CodError = "SLA-007"
            m_Mensaje = "La fecha de inicio de contacto no puede ser anterior a la fecha de recepcion de notificacion"
            m_Acumulador = m_Acumulador & m_Mensaje & vbCrLf
        End If
    End If
    
    If IsDate(m_FechaInicioAdquisicion) And IsDate(m_FechaFinAdquisicion) Then
        If CDate(m_FechaFinAdquisicion) < CDate(m_FechaInicioAdquisicion) Then
            If m_CodError = "" Then m_CodError = "SLA-007"
            m_Mensaje = "La fecha de fin de adquisicion no puede ser anterior a la fecha de inicio de adquisicion"
            m_Acumulador = m_Acumulador & m_Mensaje & vbCrLf
        End If
    End If
    
    If IsDate(m_FechaRecepcion) And IsDate(m_FechaRestablecimiento) Then
        If CDate(m_FechaRestablecimiento) < CDate(m_FechaRecepcion) Then
            If m_CodError = "" Then m_CodError = "SLA-007"
            m_Mensaje = "La fecha de restablecimiento no puede ser anterior a la fecha de recepcion de notificacion"
            m_Acumulador = m_Acumulador & m_Mensaje & vbCrLf
        End If
    End If
    
    ' --- Validaciones de franqueo (solo cuando p_EnFranqueo=True) ---
    
    If p_EnFranqueo Then
        
        If NormalizarSiNo(m_IncidenciaAveria) = "SI" Then
            If Not IsDate(m_FechaInicioAdquisicion) Then
                If m_CodError = "" Then m_CodError = "SLA-002"
                m_Mensaje = "Para franquear: la fecha de inicio de adquisicion es obligatoria para reparaciones"
                m_Acumulador = m_Acumulador & m_Mensaje & vbCrLf
            End If
            
            If Not IsDate(m_FechaFinAdquisicion) Then
                If m_CodError = "" Then m_CodError = "SLA-003"
                m_Mensaje = "Para franquear: la fecha de fin de adquisicion es obligatoria para reparaciones"
                m_Acumulador = m_Acumulador & m_Mensaje & vbCrLf
            End If
            
            If m_TipoReparacion = "" Then
                If m_CodError = "" Then m_CodError = "SLA-004"
                m_Mensaje = "Para franquear: el tipo de reparacion es obligatorio cuando es una incidencia/reparacion"
                m_Acumulador = m_Acumulador & m_Mensaje & vbCrLf
            End If
        End If
        
        If m_EventoConServicio = "SI" And Not IsDate(m_FechaRestablecimiento) Then
            If m_CodError = "" Then m_CodError = "SLA-005"
            m_Mensaje = "Para franquear: la fecha de restablecimiento de servicio es obligatoria cuando el evento afecta al servicio"
            m_Acumulador = m_Acumulador & m_Mensaje & vbCrLf
        End If
        
    End If
    
    ' Si hubo errores, quitar el vbCrLf final y devolver el acumulador como mensaje
    If m_Acumulador <> "" Then
        m_Acumulador = Left$(m_Acumulador, Len(m_Acumulador) - Len(vbCrLf))
    End If
    
    ValidarCamposSLA = m_CodError
    p_Mensaje = m_Acumulador
    p_Error = m_Acumulador
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        m_CodError = "SLA-999"
        m_Acumulador = "Error inesperado en validacion: " & Err.Description
    End If
    p_Error = m_Acumulador
    ValidarCamposSLA = m_CodError
    p_Mensaje = m_Acumulador
End Function





