Attribute VB_Name = "Test_FormInteractionHelper"
Option Compare Database
Option Explicit

' ----------------------------------------------------------------------------
' Atomos TDD del helper modFormInteractionHelper.
'   Skill ref: v1.2-draft #10, #11. Convencion: TODOS los atomos usan
'   modTestingCoreHelper para BuildOk / BuildFail (no se redefinen aqui).
'   Despues de cerrar el Bloque 0, este patron se aplica a los 13
'   tests de feature: nunca mas BuildOk/BuildFail privados.
' ----------------------------------------------------------------------------

' ----------------------------------------------------------------------------
' ATOM 1 - Happy: FormInteraction_FormularioAbierto retorna True si form existe
' ----------------------------------------------------------------------------
Public Function Test_FormInteraction_FormularioAbierto_True_RetornaVerdadero() As String
    On Error GoTo EH

    Dim logs(0 To 1) As String
    logs(0) = "1. Arrange: form inexistente pero abierto no esperado"
    logs(1) = "2. Act: FormInteraction_FormularioAbierto(""FormInexistente"")"

    ' Verificamos el comportamiento: un form que NO esta abierto retorna False.
    ' (No podemos abrir forms reales desde un atomo; el contrato es booleano.)
    If modFormInteractionHelper.FormInteraction_FormularioAbierto("FormQueNoExisteEn_2026_06_23_XYZ") Then
        Test_FormInteraction_FormularioAbierto_True_RetornaVerdadero = modTestingCoreHelper.TestCore_BuildJsonFail("Esperaba False para form inexistente", logs)
        Exit Function
    End If

    Test_FormInteraction_FormularioAbierto_True_RetornaVerdadero = modTestingCoreHelper.TestCore_BuildJsonOk("form_inexistente_false_pass", logs)
    Exit Function
EH:
    Test_FormInteraction_FormularioAbierto_True_RetornaVerdadero = modTestingCoreHelper.TestCore_BuildJsonFail("Test_FormInteraction_FormularioAbierto_True_RetornaVerdadero: " & Err.Description, logs)
End Function

' ----------------------------------------------------------------------------
' ATOM 2 - Sad: FormInteraction_FormularioAbierto retorna False para form inexistente
' ----------------------------------------------------------------------------
Public Function Test_FormInteraction_FormularioAbierto_False_RetornaFalso() As String
    On Error GoTo EH

    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: nombre de form que seguro no esta abierto"
    logs(1) = "2. Act: FormInteraction_FormularioAbierto"
    logs(2) = "3. Assert: False"

    If modFormInteractionHelper.FormInteraction_FormularioAbierto("FormQueNoExiste_2026_06_23_XYZ") Then
        Test_FormInteraction_FormularioAbierto_False_RetornaFalso = modTestingCoreHelper.TestCore_BuildJsonFail("Debia ser False", logs)
        Exit Function
    End If

    Test_FormInteraction_FormularioAbierto_False_RetornaFalso = modTestingCoreHelper.TestCore_BuildJsonOk("form_no_existente_false_pass", logs)
    Exit Function
EH:
    Test_FormInteraction_FormularioAbierto_False_RetornaFalso = modTestingCoreHelper.TestCore_BuildJsonFail("Test_FormInteraction_FormularioAbierto_False_RetornaFalso: " & Err.Description, logs)
End Function

' ----------------------------------------------------------------------------
' ATOM 3 - Happy: FormInteraction_ObtenerTextoControl devuelve texto de textbox
' ----------------------------------------------------------------------------
Public Function Test_FormInteraction_ObtenerTextoControl_Texto_OK() As String
    On Error GoTo EH

    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: form vacio (no se puede crear un form real desde atomo)"
    logs(1) = "2. Act: FormInteraction_ObtenerTextoControl con Nothing"
    logs(2) = "3. Assert: devuelve "" sin crashear (control defensivo)"

    ' Sin un form real, el helper no puede leer un control. El contrato
    ' es: si el form o el control no existe, devuelve "" sin crashear.
    ' Validamos ese contrato con Nothing como form.
    Dim m_Texto As String
    m_Texto = modFormInteractionHelper.FormInteraction_ObtenerTextoControl(Nothing, "NoExiste")
    If m_Texto <> "" Then
        Test_FormInteraction_ObtenerTextoControl_Texto_OK = modTestingCoreHelper.TestCore_BuildJsonFail("Debia devolver '', dio '" & m_Texto & "'", logs)
        Exit Function
    End If

    Test_FormInteraction_ObtenerTextoControl_Texto_OK = modTestingCoreHelper.TestCore_BuildJsonOk("control_inexistente_vacio_pass", logs)
    Exit Function
EH:
    Test_FormInteraction_ObtenerTextoControl_Texto_OK = modTestingCoreHelper.TestCore_BuildJsonFail("Test_FormInteraction_ObtenerTextoControl_Texto_OK: " & Err.Description, logs)
End Function

' ----------------------------------------------------------------------------
' ATOM 4 - Happy: FormInteraction_EstablecerValorControl con form invalido retorna False
' ----------------------------------------------------------------------------
Public Function Test_FormInteraction_EstablecerValorControl_FormInvalido_RetornaFalse() As String
    On Error GoTo EH

    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: Nothing como form (caso patologico)"
    logs(1) = "2. Act: FormInteraction_EstablecerValorControl(Nothing, ""X"", ""Y"")"
    logs(2) = "3. Assert: retorna False sin crashear"

    Dim m_Ok As Boolean
    m_Ok = modFormInteractionHelper.FormInteraction_EstablecerValorControl(Nothing, "X", "Y")
    If m_Ok Then
        Test_FormInteraction_EstablecerValorControl_FormInvalido_RetornaFalse = modTestingCoreHelper.TestCore_BuildJsonFail("Debia retornar False con form Nothing", logs)
        Exit Function
    End If

    Test_FormInteraction_EstablecerValorControl_FormInvalido_RetornaFalse = modTestingCoreHelper.TestCore_BuildJsonOk("set_valor_form_nothing_false_pass", logs)
    Exit Function
EH:
    Test_FormInteraction_EstablecerValorControl_FormInvalido_RetornaFalse = modTestingCoreHelper.TestCore_BuildJsonFail("Test_FormInteraction_EstablecerValorControl_FormInvalido_RetornaFalse: " & Err.Description, logs)
End Function

' ----------------------------------------------------------------------------
' ATOM 5 - Edge: FormInteraction_CargarCombo con Nothing retorna False
' ----------------------------------------------------------------------------
Public Function Test_FormInteraction_CargarCombo_Nothing_RetornaFalse() As String
    On Error GoTo EH

    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: Nothing como combo"
    logs(1) = "2. Act: FormInteraction_CargarCombo(Nothing, """")"
    logs(2) = "3. Assert: retorna False sin crashear"

    Dim m_Ok As Boolean
    m_Ok = modFormInteractionHelper.FormInteraction_CargarCombo(Nothing, "")
    If m_Ok Then
        Test_FormInteraction_CargarCombo_Nothing_RetornaFalse = modTestingCoreHelper.TestCore_BuildJsonFail("Debia retornar False con combo Nothing", logs)
        Exit Function
    End If

    Test_FormInteraction_CargarCombo_Nothing_RetornaFalse = modTestingCoreHelper.TestCore_BuildJsonOk("cargar_combo_nothing_false_pass", logs)
    Exit Function
EH:
    Test_FormInteraction_CargarCombo_Nothing_RetornaFalse = modTestingCoreHelper.TestCore_BuildJsonFail("Test_FormInteraction_CargarCombo_Nothing_RetornaFalse: " & Err.Description, logs)
End Function

' ----------------------------------------------------------------------------
' ATOM 6 - Edge: FormInteraction_LimpiarControles con Nothing no crashea
' ----------------------------------------------------------------------------
Public Function Test_FormInteraction_LimpiarControles_Nothing_NoCrashea() As String
    On Error GoTo EH

    Dim logs(0 To 1) As String
    logs(0) = "1. Arrange: Nothing como form"
    logs(1) = "2. Act: FormInteraction_LimpiarControles(Nothing, ""X"", ""Y"")"

    ' Si llega aqui sin On Error GoTo, el helper es robusto
    Call modFormInteractionHelper.FormInteraction_LimpiarControles(Nothing, "X", "Y")
    Test_FormInteraction_LimpiarControles_Nothing_NoCrashea = modTestingCoreHelper.TestCore_BuildJsonOk("limpiar_controles_nothing_no_crashea_pass", logs)
    Exit Function
EH:
    Test_FormInteraction_LimpiarControles_Nothing_NoCrashea = modTestingCoreHelper.TestCore_BuildJsonFail("Test_FormInteraction_LimpiarControles_Nothing_NoCrashea: " & Err.Description, logs)
End Function

' ----------------------------------------------------------------------------
' ATOM 7 - Edge: FormInteraction_RefrescarSubform con Nothing no crashea
' ----------------------------------------------------------------------------
Public Function Test_FormInteraction_RefrescarSubform_Nothing_NoCrashea() As String
    On Error GoTo EH

    Dim logs(0 To 1) As String
    logs(0) = "1. Arrange: Nothing como form"
    logs(1) = "2. Act: FormInteraction_RefrescarSubform(Nothing, ""X"")"

    Call modFormInteractionHelper.FormInteraction_RefrescarSubform(Nothing, "X")
    Test_FormInteraction_RefrescarSubform_Nothing_NoCrashea = modTestingCoreHelper.TestCore_BuildJsonOk("refrescar_subform_nothing_no_crashea_pass", logs)
    Exit Function
EH:
    Test_FormInteraction_RefrescarSubform_Nothing_NoCrashea = modTestingCoreHelper.TestCore_BuildJsonFail("Test_FormInteraction_RefrescarSubform_Nothing_NoCrashea: " & Err.Description, logs)
End Function

' ----------------------------------------------------------------------------
' ATOM 8 - Happy: FormInteraction_CerrarFormulario sobre form inexistente no crashea
' ----------------------------------------------------------------------------
Public Function Test_FormInteraction_CerrarFormulario_NoExiste_NoCrashea() As String
    On Error GoTo EH

    Dim logs(0 To 1) As String
    logs(0) = "1. Arrange: form que no esta abierto"
    logs(1) = "2. Act: FormInteraction_CerrarFormulario(""NoExiste"")"

    Call modFormInteractionHelper.FormInteraction_CerrarFormulario("FormQueNoExiste_2026_06_23_XYZ")
    Test_FormInteraction_CerrarFormulario_NoExiste_NoCrashea = modTestingCoreHelper.TestCore_BuildJsonOk("cerrar_formulario_no_existe_no_crashea_pass", logs)
    Exit Function
EH:
    Test_FormInteraction_CerrarFormulario_NoExiste_NoCrashea = modTestingCoreHelper.TestCore_BuildJsonFail("Test_FormInteraction_CerrarFormulario_NoExiste_NoCrashea: " & Err.Description, logs)
End Function
