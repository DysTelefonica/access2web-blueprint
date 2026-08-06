Attribute VB_Name = "AutomatizacionRepositorio"

Option Compare Database
Option Explicit

' Borra correos programados de una solicitud específica que aún no han salido.
' Se llama cuando el usuario avanza de fase manualmente o cancela la solicitud.
Public Function AnularProgramadosPorFase(p_IDSolicitud As String, p_Fase As String, ByRef p_Error As String)
    Dim m_SQL As String
    Dim filtroTipos As String
    
    On Error GoTo errores
    
    ' Definimos qué tipos de correo borrar según el hito que acaba de ocurrir
    Select Case p_Fase
        Case "RecepcionExcel"
            ' El usuario ha mandado el Excel -> Borramos recordatorios de envío de Excel y su cancelación
            filtroTipos = "(" & EnumTipoEnvioCorreo.RecordatorioExcel1 & "," & _
                              EnumTipoEnvioCorreo.RecordatorioExcel2 & "," & _
                              EnumTipoEnvioCorreo.CancelacionPreMarga & ")"
                              
        Case "AltaMarga"
            ' Por seguridad, al dar de alta en Marga, nos aseguramos que no quede nada de la fase anterior
            filtroTipos = "(" & EnumTipoEnvioCorreo.RecordatorioExcel1 & "," & _
                              EnumTipoEnvioCorreo.RecordatorioExcel2 & "," & _
                              EnumTipoEnvioCorreo.CancelacionPreMarga & ")"
        
        Case "RecepcionDPS"
            ' El usuario ha mandado el DPS (o lo hemos recibido) -> Borramos recordatorios de Marga
            filtroTipos = "(" & EnumTipoEnvioCorreo.RecordatorioMarga1 & "," & _
                              EnumTipoEnvioCorreo.RecordatorioMarga2 & "," & _
                              EnumTipoEnvioCorreo.CancelacionMarga & ")"
                              
        Case "TODO"
            ' Cancelación o Desestimación de la solicitud -> Borrar TODO lo futuro pendiente
            filtroTipos = "" ' Se gestiona diferente en el SQL
            
    End Select
    
    If p_Fase = "TODO" Then
        m_SQL = "DELETE * FROM TbCorreosEnviados " & _
                "WHERE IDSolicitud = " & p_IDSolicitud & " " & _
                "AND FechaEnvio Is Null " & _
                "AND Programado = 'Sí';"
    ElseIf filtroTipos <> "" Then
        m_SQL = "DELETE * FROM TbCorreosEnviados " & _
                "WHERE IDSolicitud = " & p_IDSolicitud & " " & _
                "AND FechaEnvio Is Null " & _
                "AND Programado = 'Sí' " & _
                "AND TipoCorreo IN " & filtroTipos & ";"
    End If
    
    If m_SQL <> "" Then
        getdb().Execute m_SQL
    End If
    
    Exit Function
errores:
    p_Error = "Error anulando programados en AutomatizacionRepositorio: " & Err.Description
End Function

