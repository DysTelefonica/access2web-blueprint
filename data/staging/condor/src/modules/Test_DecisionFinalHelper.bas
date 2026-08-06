Attribute VB_Name = "Test_DecisionFinalHelper"
Option Compare Database
Option Explicit

' ==========================================================================
' Test_DecisionFinalHelper — Slice 3.1 DecisionFinal shared pure contract
' ==========================================================================

Public Function Test_DecisionFinalHelper_SharedContract_AppliesToAllGemelos() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(5)

    logs(0) = "1. Matriz gemelos: PC, PCSUB, CDCA y CDCASUB comparten decisionFinal + NombreFirmanteFinal"

    If Not DecisionFinalHelper_EsCompleta("APROBADO", "Firmante PC") Then Err.Raise 513, , "PC debería estar completo"
    logs(1) = "2. PC completo"

    If Not DecisionFinalHelper_EsCompleta("APROBADO", "Firmante PCSUB") Then Err.Raise 513, , "PCSUB debería estar completo"
    logs(2) = "3. PCSUB completo"

    If Not DecisionFinalHelper_EsCompleta("APROBADO", "Firmante CDCA") Then Err.Raise 513, , "CDCA debería estar completo"
    logs(3) = "4. CDCA completo"

    If Not DecisionFinalHelper_EsCompleta("APROBADO", "Firmante CDCASUB") Then Err.Raise 513, , "CDCASUB debería estar completo"
    logs(4) = "5. CDCASUB completo"

    Test_DecisionFinalHelper_SharedContract_AppliesToAllGemelos = TestHelper.BuildJsonOk("true", logs)
    Exit Function

EH:
    logs(UBound(logs)) = "ERR " & Err.Number & " - " & Err.Description
    Test_DecisionFinalHelper_SharedContract_AppliesToAllGemelos = TestHelper.BuildJsonFail(Err.Description, logs)
End Function

Public Function Test_DecisionFinalHelper_MissingDecision_ReturnsFalse() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(3)

    logs(0) = "1. Arrange: decisión vacía y firmante informado"
    If DecisionFinalHelper_EsCompleta(vbNullString, "Firmante") Then Err.Raise 513, , "Una decisión vacía no debe estar completa"
    logs(1) = "2. PASS: decisión vacía devuelve False"

    If DecisionFinalHelper_EsCompleta(Null, "Firmante") Then Err.Raise 513, , "Una decisión Null no debe estar completa"
    logs(2) = "3. PASS: decisión Null devuelve False"

    Test_DecisionFinalHelper_MissingDecision_ReturnsFalse = TestHelper.BuildJsonOk("true", logs)
    Exit Function

EH:
    logs(UBound(logs)) = "ERR " & Err.Number & " - " & Err.Description
    Test_DecisionFinalHelper_MissingDecision_ReturnsFalse = TestHelper.BuildJsonFail(Err.Description, logs)
End Function

Public Function Test_DecisionFinalHelper_MissingFirmante_ReturnsFalse() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(3)

    logs(0) = "1. Arrange: decisión informada y firmante vacío"
    If DecisionFinalHelper_EsCompleta("APROBADO", vbNullString) Then Err.Raise 513, , "Un firmante vacío no debe estar completo"
    logs(1) = "2. PASS: firmante vacío devuelve False"

    If DecisionFinalHelper_EsCompleta("APROBADO", Null) Then Err.Raise 513, , "Un firmante Null no debe estar completo"
    logs(2) = "3. PASS: firmante Null devuelve False"

    Test_DecisionFinalHelper_MissingFirmante_ReturnsFalse = TestHelper.BuildJsonOk("true", logs)
    Exit Function

EH:
    logs(UBound(logs)) = "ERR " & Err.Number & " - " & Err.Description
    Test_DecisionFinalHelper_MissingFirmante_ReturnsFalse = TestHelper.BuildJsonFail(Err.Description, logs)
End Function

Public Function Test_DecisionFinalHelper_TrimsWhitespaceBeforeValidation() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(4)

    logs(0) = "1. Arrange: espacios alrededor de valores válidos"
    If Not DecisionFinalHelper_EsCompleta("  APROBADO  ", "  Firmante  ") Then Err.Raise 513, , "Debe aceptar valores con espacios exteriores"
    logs(1) = "2. PASS: trim de valores válidos"

    If DecisionFinalHelper_EsCompleta("   ", "  Firmante  ") Then Err.Raise 513, , "Solo espacios en decisión debe devolver False"
    logs(2) = "3. PASS: solo espacios en decisión devuelve False"

    If DecisionFinalHelper_EsCompleta("APROBADO", "   ") Then Err.Raise 513, , "Solo espacios en firmante debe devolver False"
    logs(3) = "4. PASS: solo espacios en firmante devuelve False"

    Test_DecisionFinalHelper_TrimsWhitespaceBeforeValidation = TestHelper.BuildJsonOk("true", logs)
    Exit Function

EH:
    logs(UBound(logs)) = "ERR " & Err.Number & " - " & Err.Description
    Test_DecisionFinalHelper_TrimsWhitespaceBeforeValidation = TestHelper.BuildJsonFail(Err.Description, logs)
End Function
