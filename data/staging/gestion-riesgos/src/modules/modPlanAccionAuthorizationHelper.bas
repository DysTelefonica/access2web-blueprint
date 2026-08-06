Attribute VB_Name = "modPlanAccionAuthorizationHelper"
Option Compare Database
Option Explicit

' =============================================================================
' modPlanAccionAuthorizationHelper.bas
'
' Helper module para validación de eliminación de acciones de plan.
' Gemelo PM ? PC (workflow types) — ambos planes usan el mismo helper.
' Project: gestion_riesgos
' Branch: feat/metodologia-e2e-riesgos-2026-06-19
' Refactor original: 2026-06-19 (extracción desde Form_FormRiesgosGestionPlanAcciones.cls)
' Realineación:    2026-06-24 (overrides 2026-06-19 deviation)
'
' Helper público (1):
'   ValidarEliminacionAccion — valida eliminación de acción de plan (PM o PC)
'
' ALINEADO AL PATRÓN CANÓNICO desde 2026-06-24 (overrides 2026-06-19 deviation).
'   Usa Optional ByRef p_PromptResult As Long para bypass de MsgBox en tests.
'   Documentado en átomos (Test_PlanAccionAuthorizationHelper.bas).
'   Patrón alineado a modNCHelper.bas, modRiesgoEstadoGateHelper.bas,
'   modNCExclusionHelper.bas, modAvisoPublicacionHelper.bas,
'   modInformePublicacionHelper.bas.
' =============================================================================

' --- Bypass de MsgBox vía parámetro p_PromptResult ---
'   0       = fallback a MsgBox real (modo producción)
'   vbYes   = simula "Sí" automáticamente (modo test)
'   vbNo    = simula "No" (modo test)

' =============================================================================
' ValidarEliminacionAccion
'
' Valida si la acción de plan (PM o PC) puede eliminarse:
'   1) Verifica que la acción existe en la base de datos
'   2) Verifica que el usuario conectado tiene autorización para eliminar
'   3) Solicita confirmación via MsgBox (o bypass automático en tests)
'
' Parámetros:
'   p_EsMitigacion   - True = acción de Plan de Mitigación (PM)
'                       False = acción de Plan de Contingencia (PC)
'   p_IDAccion       - ID de la acción a eliminar
'   p_PromptResult   - Bypass de MsgBox para tests TDD:
'                       0      = fallback a MsgBox real (modo producción)
'                       vbYes  = simula "Sí" (modo test, átomo Happy/Edge/Adv)
'                       vbNo   = simula "No" (modo test, cancel path)
'   p_Error          - Motivo del rechazo (solo errores reales).
'                       Convención Telefónica D&S: "no encontrado",
'                       "no autorizado" y "cancelado por usuario" se
'                       retornan con p_Error="" y result=False; la
'                       distinción la hace el caller por el contexto.
'                       (Regla también aplicada en Constructor.getPMAccion
'                       / getPCAccion — EOF branch.)
'
' Retorna:
'   True  = la acción puede eliminarse (el usuario confirmó con "Sí")
'   False = la acción NO puede eliminarse (no existe, no autorizada,
'           o el usuario canceló la confirmación)
'
' Notas:
'   Esta función es el gemelo PM ? PC: la misma lógica aplica a ambos
'   tipos de plan, variando solo en qué clase de acción se consulta.
' =============================================================================
Public Function ValidarEliminacionAccion( _
    ByVal p_EsMitigacion As Boolean, _
    ByVal p_IDAccion As String, _
    Optional ByRef p_PromptResult As Long, _
    Optional ByRef p_Error As String) As Boolean

    Dim m_ObjAccion As Object
    Dim m_Pregunta As Long

    On Error GoTo errores

    p_Error = ""
    ValidarEliminacionAccion = False

    ' --- Paso 1: Obtener la acción desde la base de datos ---
    If p_EsMitigacion Then
        Set m_ObjAccion = Constructor.getPMAccion(p_IDAccion, p_Error)
        If p_Error <> "" Then Err.Raise 1000
        If m_ObjAccion Is Nothing Then
            ValidarEliminacionAccion = False
            Exit Function
        End If
    Else
        Set m_ObjAccion = Constructor.getPCAccion(p_IDAccion, p_Error)
        If p_Error <> "" Then Err.Raise 1000
        If m_ObjAccion Is Nothing Then
            ValidarEliminacionAccion = False
            Exit Function
        End If
    End If

    ' --- Paso 2: Verificar autorización del usuario conectado ---
    '   El authorization se resuelve via la acción, que delega al plan
    '   (PMAccion.UsuarioConectadoAutorizado o PCAccion.UsuarioConectadoAutorizado)
    '   y este a su vez llama a UsuarioAutorizado() con el objeto del plan.
    '   Convención Telefónica D&S: "no autorizado" ? p_Error="" (caller path).
    If m_ObjAccion.UsuarioConectadoAutorizado <> EnumSiNo.Sí Then
        ValidarEliminacionAccion = False
        Exit Function
    End If

    ' --- Paso 3: Confirmación del usuario (bypass vía p_PromptResult) ---
    '   Patrón canónico alineado a modNCHelper.bas: si p_PromptResult <> 0
    '   se usa ese valor (inyectado por átomos TDD); si es 0, cae al MsgBox
    '   real en modo producción.
    If p_PromptResult <> 0 Then
        m_Pregunta = p_PromptResult
    Else
        m_Pregunta = MsgBox( _
            "¿Desea realmente borrar la acción seleccionada?", _
            vbExclamation + vbYesNo + vbDefaultButton2, _
            "Borrado de una acción")
    End If

    If m_Pregunta <> vbYes Then
        ' Convención Telefónica D&S: "cancelado por usuario" ? p_Error="".
        ValidarEliminacionAccion = False
        Exit Function
    End If

    ' --- Autorizado y confirmado ---
    ValidarEliminacionAccion = True
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "ValidarEliminacionAccion: " & Err.description
    End If
    ValidarEliminacionAccion = False
End Function

