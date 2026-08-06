Attribute VB_Name = "modFormOutcomeHandler"
' =============================================================================
' modFormOutcomeHandler.bas
'
' Helper GENERAL para outcomes de formularios de gestion de riesgos.
' Project: gestion_riesgos
' Slice: forms-thin-phase0-testeable-2026-06-25 / Tier-1 form refactor
'        + generalize-visado-outcome-handler-2026-07-01
'
' Historia:
'   - Nacio como modVisadoOutcomeHandler (slice 545b152) para los 4 outcomes
'     del form Form_FormCalidadTareaRiesgosAceptadosRetirados.
'   - Renombrado a modFormOutcomeHandler (este slice, 2026-07-01) porque el
'     patron es GENERAL: cualquier form que necesite ejecutar un outcome
'     discreto contra el mismo flujo EstablecerTareasCalidad + refresco
'     del arbol de riesgos reutiliza este helper.
'
' Reglas:
'   - NO UI, NO MsgBox. El form renderiza el error con MostrarMsgBoxSiNoEsTest.
'   - p_Error ByRef (Telefonica D&S); vacio en exito.
'   - Naming: FormOutcomeHandler_<Verbo> por convencion del proyecto
'     (ModuleName_ prefix por Hard rule 5).
'   - Es un Sub (no Function): no devuelve JSON; los atomos Test_*_bas
'     envuelven la llamada y devuelven JSON via TestCore_BuildOk/BuildFail.
'
' Por que existe (general):
'   - Hard rule 1: forms thin; los event handlers deben ser una sola linea
'     de llamada al helper + el wiring de error del form (On Error / MostrarMsgBox).
'   - Hard rule 11 (gemelo invariant): cualquier form que comparta el mismo
'     patron algoritmico (outcome discreto -> EstablecerTareasCalidad(Si, No)
'     -> refresco del arbol via bridge) reusa este helper sin duplicacion.
'
' Outcomes soportados (parametro p_Outcome):
'   - "AceptacionAprobada"        → EstablecerTareasCalidad(Si, No) + refresco calidad
'   - "AceptacionAprobadaQuitado" → idem
'   - "AceptacionRechazada"       → idem
'   - "AceptacionRechazadaQuitada"→ idem
'   - "" o cualquier otro         → p_Error = outcome no soportado (sad path)
'
' Decisiones de diseno:
'   - p_Outcome se acepta y se valida, pero NO discrimina el flujo interno:
'     los 4 outcomes llaman EstablecerTareasCalidad(EnumSiNo.Sí, EnumSiNo.No)
'     identicamente. El parametro se conserva para trazabilidad / extension
'     futura (e.g. si algun outcome requiere una variante del bridge).
'   - Sin parametro db (Hard rule 2 honest signature): el helper no lee DAO
'     directamente. Toda persistencia ocurre dentro de las funciones
'     EstablecerTareasCalidad / Refresco_ArbolRiesgosCalidadBridge, que ya
'     usan getdb() internamente.
'   - Scope "calidad" hard-coded en el bridge: este helper sirve los forms
'     del gemelo Calidad. Si en el futuro un gemelo Tecnico necesita el
'     mismo patron, se introducira un parametro p_Scope en este helper
'     (parametro opcional con default "calidad") para preservar la
'     backward-compat de los callers actuales.
' =============================================================================
Option Compare Database
Option Explicit

' -----------------------------------------------------------------------------
' FormOutcomeHandler_Run — ejecuta el flujo de un outcome discreto de un form
'   de gestion de riesgos (los 4 outcomes del gemelo Calidad actualmente).
'
' Parametros:
'   p_Outcome : uno de los 4 outcomes soportados (ver bloque de arriba).
'   p_Error   : ByRef, vacio en exito. NO se hace Err.Raise al caller;
'               el caller inspecciona p_Error y eleva Err.Raise 1000 si
'               quiere que el handler del form muestre MostrarMsgBox.
'
' Comportamiento:
'   1. Si p_Outcome esta vacio o no es uno de los 4 outcomes soportados,
'      setea p_Error y sale SIN tocar EstablecerTareasCalidad ni el bridge.
'   2. Llama a EstablecerTareasCalidad(EnumSiNo.Sí, EnumSiNo.No, p_Error).
'      Si p_Error <> "", sale (no llama al bridge).
'   3. Llama a modArbolRefrescoHelper.Refresco_ArbolRiesgosCalidadBridge
'      con scope "calidad" (unico scope del gemelo Calidad actualmente).
'
' Garantias (Hard rule 7):
'   - CERO MsgBox / InputBox / DoCmd.Hourglass / DoCmd.OpenForm / DoEvents.
'   - CERO referencias cruzadas a Forms("...") — eso vive en el bridge
'     extraido, que es el unico punto de contacto con FormRiesgosGestion.
' -----------------------------------------------------------------------------
Public Sub FormOutcomeHandler_Run( _
                                    ByVal p_Outcome As String, _
                                    Optional ByRef p_Error As String _
                                    )

    On Error GoTo errores

    p_Error = ""

    ' 1. Validacion del outcome (sad path determinista)
    If Not FormOutcomeHandler_OutcomeEsValido(p_Outcome) Then
        p_Error = "FormOutcomeHandler_Run: outcome no soportado '" & p_Outcome & "'"
        Exit Sub
    End If

    ' 2. EstablecerTareasCalidad(Si, No) — re-uso del helper existente
    '    (Funciones Generales.bas:2204). Si falla, p_Error queda rellenado.
    EstablecerTareasCalidad EnumSiNo.Sí, EnumSiNo.No, p_Error
    If p_Error <> "" Then
        Exit Sub
    End If

    ' 3. Refresco del arbol via bridge extraido (no-op si FormRiesgosGestion
    '    no esta abierto — caso normal en tests headless).
    modArbolRefrescoHelper.Refresco_ArbolRiesgosCalidadBridge "calidad", p_Error
    Exit Sub

errores:
    If p_Error = "" Then
        p_Error = "FormOutcomeHandler_Run: " & Err.Number & " - " & Err.Description
    End If
End Sub

' -----------------------------------------------------------------------------
' FormOutcomeHandler_OutcomeEsValido — predicado puro, sin UI, sin DAO.
'   Vive en este modulo (no en modTestingCoreHelper) porque es la regla
'   de aceptacion del helper, no de testing.
' -----------------------------------------------------------------------------
Private Function FormOutcomeHandler_OutcomeEsValido(ByVal p_Outcome As String) As Boolean
    Select Case p_Outcome
        Case "AceptacionAprobada", "AceptacionAprobadaQuitado", _
             "AceptacionRechazada", "AceptacionRechazadaQuitada"
            FormOutcomeHandler_OutcomeEsValido = True
        Case Else
            FormOutcomeHandler_OutcomeEsValido = False
    End Select
End Function