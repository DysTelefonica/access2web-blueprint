Attribute VB_Name = "modArbolRefrescoHelper"
' =============================================================================
' modArbolRefrescoHelper.bas
'
' Helper para refresco del arbol de riesgos en el form FormRiesgosGestion.
' Project: gestion_riesgos
' Slice: forms-thin-phase0-testeable-2026-06-25 / Tier-1 form refactor
'
' Reglas:
'   - NO UI, NO MsgBox. Helpers reciben Optional ByRef p_Error (Telefonica D&S).
'   - Helpers devuelven vacio en exito; p_Error <> "" en fallo (sin raise
'     visible para atomos).
'   - Naming: Refresco_<Verbo> por convencion del proyecto.
'   - Re-exporta la logica que antes vivia como Private Function
'     RefrescarArbolRiesgosCalidadBridge dentro de cada form (Calidad,
'     Mitigacion, Materializado, Retirado). Calidad migro en el slice
'     545b152; Mitigacion, Materializado, Retirado migraron en el slice
'     de desduplicacion gemelos Tier-3→Tier-1.
'
' Por que se extrae:
'   - Hard rule 1 (no business logic in event handlers): los 4 handlers
'     m_FormVisado_* de Form_FormCalidadTareaRiesgosAceptadosRetirados
'     llamaban al bridge privado. Ahora el bridge es publico y testeable.
'   - Hard rule 11 (gemelo invariant): el bridge estaba duplicado en 4 forms
'     (Calidad, Mitigacion, Materializado, Retirado). Todos migrados.
' =============================================================================
Option Compare Database
Option Explicit

' -----------------------------------------------------------------------------
' Refresco_ArbolRiesgosCalidadBridge — equivalente al antiguo Private
'   Function RefrescarArbolRiesgosCalidadBridge del form Calidad.
'   Recibe p_Mutacion (string libre, mapeada via ResolverScopeRefrescoCalidadRiesgo),
'   resuelve el scope y dispara RefrescarArbolRiesgosScope en FormRiesgosGestion
'   si esta abierto. Si el form de gestion no esta abierto, sale silenciosamente
'   sin error (caso normal en tests headless).
'
' Pre-condicion: FormRiesgosGestion debe estar abierto para que el refresco
'   tenga efecto visual. Si no esta abierto, no-op silencioso (p_Error = "").
'   Esto preserva la semantica original del bridge privado.
'
' Hard rule 7: NO MsgBox aqui. Hard rule 1: NO DoEvents / DoCmd.OpenForm aqui.
'   El caller (helper FormOutcomeHandler_Run) decide que hacer con p_Error.
' -----------------------------------------------------------------------------
Public Sub Refresco_ArbolRiesgosCalidadBridge( _
                                            ByVal p_Mutacion As String, _
                                            Optional ByRef p_Error As String _
                                            )
    Dim m_Scope As String
    Dim m_ScopeError As String

    On Error GoTo errores

    p_Error = ""

    If Not FormularioAbierto("FormRiesgosGestion") Then
        ' No-op silencioso: en tests headless este form no esta abierto.
        Exit Sub
    End If

    m_ScopeError = ""
    m_Scope = ResolverScopeRefrescoCalidadRiesgo(p_Mutacion, m_ScopeError)
    If m_ScopeError <> "" Then
        p_Error = m_ScopeError
        Exit Sub
    End If

    Form_FormRiesgosGestion.RefrescarArbolRiesgosScope p_Scope:=m_Scope, p_Error:=p_Error
    Exit Sub

errores:
    If p_Error = "" Then
        p_Error = "El metodo Refresco_ArbolRiesgosCalidadBridge ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Sub