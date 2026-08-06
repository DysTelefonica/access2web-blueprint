Attribute VB_Name = "modAccionFechasHelper"
Option Compare Database
Option Explicit

' =============================================================================
' modAccionFechasHelper.bas
'
' Helper module para validación cronológica de fechas de acciones de planes
' de mitigación (PMAccion) y contingencia (PCAccion).
' Project: gestion_riesgos
' SDD:     e2e-form-by-form-2026-06-22 - Bloque 4 - REQ-CAL-04 (parte 2)
' Issue:   #90 — Punto 04 parte 2 — Constraints cronológicas de fecha
'                       prevista e inicio.
'
' Helper público (1):
'   ValidarFechasAccion - valida las constraints cronológicas (a) y (b)
'                         que aplican a PMAccion y PCAccion (no al Riesgo).
'
' DECISIÓN 2026-07-09 (sdd-apply issue #90):
'   * Helper PURO (no DAO): la regla es función pura de
'     (p_FechaInicio, p_FechaFinPrevista, p_FechaBaseProyecto).
'     El cache-first se hace en el form layer (m_FechaBaseProyecto
'     resuelto UNA sola vez en Form_Load vía modEdicionFechasPresenter.
'     ObtenerFechaBaseProyecto).
'   * Devuelve Boolean (no String JSON). Mismo estilo que
'     modRiesgoFechasHelper.bas:ValidarFechaCampo y
'     modPlanAccionAuthorizationHelper.bas.
'   * p_Error ByRef es OBLIGATORIO (convención Telefónica D&S); el caller
'     lo inspecciona y muestra el mensaje al usuario.
'   * Sin DAO: el helper NO consulta la BD. Toda dependencia de datos
'     se inyecta desde el caller (form). Esto preserva la regla de
'     oro "cache-first" — la query SQL para resolver
'     MIN(FechaEdicion) por IDProyecto se ejecuta UNA vez por form load,
'     no una vez por acción ni una vez por BeforeUpdate.
'
' Reglas de negocio (issue #90):
'   Constraint (a): FechaFinPrevista >= FechaInicio.
'                   Si FechaFinPrevista < FechaInicio → rechazar
'                   (mensaje: "La fecha prevista de la acción no puede
'                    ser anterior a la fecha de inicio").
'                   Si igual o posterior → aceptar.
'                   Si FechaInicio está vacío → constraint (a) NO aplica.
'
'   Constraint (b): FechaInicio >= FechaBaseProyecto.
'                   Si FechaInicio < FechaBaseProyecto → rechazar
'                   (mensaje: "La fecha de inicio de la acción no puede
'                    ser anterior a la fecha de creación de la primera
'                    edición del proyecto").
'                   Si igual o posterior → aceptar.
'                   Si FechaBaseProyecto está vacío → constraint (b)
'                   NO aplica (skip defensivo, no es nuestro problema
'                   que el form layer no haya cacheado el dato).
'
'   Vacío / no-fecha: NO es nuestro problema (la regla de "campo
'                     requerido" la aplica el formulario, igual que
'                     en modRiesgoFechasHelper.bas).
' =============================================================================

' -----------------------------------------------------------------------------
' ValidarFechasAccion
'
' Valida que las fechas de inicio y fin prevista de una acción
' (PMAccion / PCAccion) cumplan las constraints cronológicas (a) y (b).
'
' Parámetros:
'   p_FechaInicio (String)            - fecha de inicio de la acción
'                                       (formato VB: dd/mm/yyyy o vacío).
'   p_FechaFinPrevista (String)       - fecha prevista de fin de la acción
'                                       (formato VB: dd/mm/yyyy o vacío).
'   p_FechaBaseProyecto (String)      - MIN(FechaEdicion) de TbProyectosEdiciones
'                                       filtrado por IDProyecto (formato VB:
'                                       dd/mm/yyyy o vacío si el form no
'                                       cacheó el dato).
'   p_Error (out String, opcional)    - mensaje legible si retorna False;
'                                       vacío si True.
'
' Retorna:
'   True  si las fechas cumplen las constraints cronológicas (o si la
'         constraint no aplica por falta del dato correspondiente).
'   False si alguna constraint se viola, con p_Error poblado en español.
' -----------------------------------------------------------------------------
Public Function ValidarFechasAccion( _
    ByVal p_FechaInicio As String, _
    ByVal p_FechaFinPrevista As String, _
    ByVal p_FechaBaseProyecto As String, _
    Optional ByRef p_Error As String) As Boolean

    On Error GoTo errores

    p_Error = ""
    ValidarFechasAccion = False

    ' --- Constraint (a): FechaFinPrevista >= FechaInicio ---
    ' Solo aplica si FechaInicio está presente (FechaInicio vacío = no aplica
    ' per spec del issue #90 escenario 6). Si FechaFinPrevista no es fecha,
    ' es problema del form layer, no nuestro.
    If Len(Trim$(Nz(p_FechaInicio, ""))) > 0 Then
        If Len(Trim$(Nz(p_FechaFinPrevista, ""))) > 0 Then
            If IsDate(p_FechaInicio) And IsDate(p_FechaFinPrevista) Then
                If CDate(p_FechaFinPrevista) < CDate(p_FechaInicio) Then
                    p_Error = "La fecha prevista de la acción no puede ser anterior a la fecha de inicio"
                    ValidarFechasAccion = False
                    Exit Function
                End If
            End If
            ' Si alguna no es fecha: skip defensivo (no es nuestro problema).
        End If
    End If

    ' --- Constraint (b): FechaInicio >= FechaBaseProyecto ---
    ' Solo aplica si ambos datos están presentes. Si FechaBaseProyecto está
    ' vacío, el form layer no cacheó el dato (skip defensivo).
    If Len(Trim$(Nz(p_FechaInicio, ""))) > 0 Then
        If Len(Trim$(Nz(p_FechaBaseProyecto, ""))) > 0 Then
            If IsDate(p_FechaInicio) And IsDate(p_FechaBaseProyecto) Then
                If CDate(p_FechaInicio) < CDate(p_FechaBaseProyecto) Then
                    p_Error = "La fecha de inicio de la acción no puede ser anterior a la fecha de creación de la primera edición del proyecto"
                    ValidarFechasAccion = False
                    Exit Function
                End If
            End If
        End If
    End If

    ' Todas las constraints aplicables pasaron.
    ValidarFechasAccion = True
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "ValidarFechasAccion: " & Err.Number & " - " & Err.description
    End If
    ValidarFechasAccion = False
End Function