Attribute VB_Name = "modPublicacionCalidadExecutionHelper"
' =============================================================================
' modPublicacionCalidadExecutionHelper.bas
'
' Helper module para rechazo de propuesta de publicación.
' Project: gestion_riesgos
' Branch: feat/metodologia-e2e-riesgos-2026-06-19
' Refactor: 2026-06-19
'
' Helper público (1):
'   RechazarPropuestaPublicacion — rechaza propuesta con motivo y envía correo
'
' CORRECCIÓN 2026-06-19: métodos asumidos en el audit original NO existen.
' Realidad: Edicion.RechazoPropuestaParaPublicacion(p_Error) As CORREO es el
' método canónico. Internamente setea PropuestaRechazadaPorCalidadMotivo,
' actualiza TbProyectosEdiciones, llama a EnviarCorreoRechazoPropuestaPublicacion
' y registra el correo vía EdicionCorreoRevision.RegistrarCorreoEnviado.
' =============================================================================
Option Compare Database
Option Explicit

Public Function RechazarPropuestaPublicacion( _
    ByVal p_IDEdicion As String, _
    ByVal p_Motivo As String, _
    Optional ByRef p_CorreoRechazo As Object, _
    Optional ByRef db As DAO.Database = Nothing, _
    Optional ByRef p_PromptResult As Long) As String

    ' -----------------------------------------------------------------------
    ' Orchestrate: resolver edición > asignar motivo > rechazar > retornar.
    '
    ' El form llama a este helper en lugar de hacerlo inline.
    ' El form recibe el CORREO devuelto para verificar si fue enviado.
    ' -----------------------------------------------------------------------

    Dim m_Edicion As Edicion
    Dim m_CorreoRechazo As Object
    Dim m_Error As String
    Dim logs(0 To 3) As String

    logs(0) = "1. Resolver edición por IDEdicion=" & CStr(p_IDEdicion)
    logs(1) = "2. Asignar PropuestaRechazadaPorCalidadMotivo"
    logs(2) = "3. Llamar Edicion.RechazoPropuestaParaPublicacion"
    logs(3) = "4. Retornar JSON con resultado"

    On Error GoTo errores

    ' --- 1. Resolver edición ---
    Set m_Edicion = Constructor.getEdicion(p_IDEdicion:=p_IDEdicion, p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If

    ' --- 2. Asignar el motivo antes de rechazar ---
    m_Edicion.PropuestaRechazadaPorCalidadMotivo = p_Motivo

    ' --- 3. Rechazar: el método internamente envía el correo y lo registra ---
    Set m_CorreoRechazo = m_Edicion.RechazoPropuestaParaPublicacion(m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If

    ' --- 4. Devolver el correo al caller (para que el test átomo lo verifique) ---
    Set p_CorreoRechazo = m_CorreoRechazo

    logs(1) = "2. Asignado motivo: " & IIf(Len(p_Motivo) > 30, Left(p_Motivo, 30) & "...", p_Motivo)
    logs(3) = "4. Rechazo completado; correo=" & _
              IIf(m_CorreoRechazo Is Nothing, "Nothing", m_CorreoRechazo.IDCorreo)

    RechazarPropuestaPublicacion = BuildJsonOk("rechazado", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        m_Error = "RechazarPropuestaPublicacion: error " & Err.Number & ": " & Err.description
    End If
    ' logs(3) ya describe el paso donde falló
    RechazarPropuestaPublicacion = BuildJsonFail(m_Error, logs)
End Function

