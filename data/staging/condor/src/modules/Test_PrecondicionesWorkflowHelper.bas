Attribute VB_Name = "Test_PrecondicionesWorkflowHelper"
Option Compare Database
Option Explicit

' ==========================================================================
' Test_PrecondicionesWorkflowHelper - Slice 3.3 BR-002..BR-006 atoms
'
' Each atom is a Public Function (global, unique) that returns the
' canonical JSON via TestHelper.BuildJsonOk / BuildJsonFail. Pure
' helper tests: NO DAO, NO controls, NO MsgBox, no fixtures required.
' ==========================================================================

' --- ParteTecnicaCompleta ----------------------------------------------------

Public Function Test_PrecondicionesWorkflowHelper_ParteTecnicaCompleta_HappyPathAllGemelos() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(5)
    Dim entidad As Object
    Dim msg As String
    Dim ok As Boolean

    logs(0) = "1. Arrange: Dictionary con ParteTecnicaCompleta=True para los 4 gemelos"

    Set entidad = CreateObject("Scripting.Dictionary")
    entidad.Add "ParteTecnicaCompleta", True

    ok = PrecondicionesWorkflowHelper_ParteTecnicaCompleta("PC", "estadoModificacion", entidad, msg)
    logs(1) = "2. PC: ok=" & CStr(ok) & " msg='" & msg & "'"
    If Not ok Then Err.Raise 513, , "PC con ParteTecnicaCompleta=True debe devolver True"

    ok = PrecondicionesWorkflowHelper_ParteTecnicaCompleta("PCSUB", "estadoModificacion", True, msg)
    logs(2) = "3. PCSUB (Boolean): ok=" & CStr(ok) & " msg='" & msg & "'"
    If Not ok Then Err.Raise 513, , "PCSUB con True debe devolver True"

    ok = PrecondicionesWorkflowHelper_ParteTecnicaCompleta("CD_CA", "estadoValidacion", True, msg)
    logs(3) = "4. CD_CA: ok=" & CStr(ok)
    If Not ok Then Err.Raise 513, , "CD_CA con True debe devolver True"

    ok = PrecondicionesWorkflowHelper_ParteTecnicaCompleta("CD_CA_SUB", "estadoValidacion", True, msg)
    logs(4) = "5. CDCASUB: ok=" & CStr(ok)
    If Not ok Then Err.Raise 513, , "CD_CA_SUB con True debe devolver True"

    Test_PrecondicionesWorkflowHelper_ParteTecnicaCompleta_HappyPathAllGemelos = TestHelper.BuildJsonOk("true", logs)
    Exit Function

EH:
    logs(UBound(logs)) = "ERR " & Err.Number & " - " & Err.Description
    Test_PrecondicionesWorkflowHelper_ParteTecnicaCompleta_HappyPathAllGemelos = TestHelper.BuildJsonFail(Err.Description, logs)
End Function

Public Function Test_PrecondicionesWorkflowHelper_ParteTecnicaCompleta_FaltanDatos_False() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(4)
    Dim entidad As Object
    Dim msg As String
    Dim ok As Boolean

    logs(0) = "1. Arrange: Dictionary con ParteTecnicaCompleta=False"

    Set entidad = CreateObject("Scripting.Dictionary")
    entidad.Add "ParteTecnicaCompleta", False

    ok = PrecondicionesWorkflowHelper_ParteTecnicaCompleta("PC", "estadoModificacion", entidad, msg)
    logs(1) = "2. Act: ParteTecnicaCompleta=False"

    If ok Then Err.Raise 513, , "ParteTecnicaCompleta=False debe devolver False"
    If Len(Trim$(msg)) = 0 Then Err.Raise 513, , "RazonFallo no debe estar vacia cuando faltan datos"
    If InStr(msg, "Propuesta") = 0 Or InStr(msg, "Impacto") = 0 Then _
        Err.Raise 513, , "RazonFallo debe mencionar Propuesta/Impacto (CAP-007 §2 BR-002). msg='" & msg & "'"
    logs(2) = "3. PASS: ok=False, msg contiene Propuesta+Impacto"
    logs(3) = "4. PASS: wording BR-002 mantenido"

    Test_PrecondicionesWorkflowHelper_ParteTecnicaCompleta_FaltanDatos_False = TestHelper.BuildJsonOk("false", logs)
    Exit Function

EH:
    logs(UBound(logs)) = "ERR " & Err.Number & " - " & Err.Description
    Test_PrecondicionesWorkflowHelper_ParteTecnicaCompleta_FaltanDatos_False = TestHelper.BuildJsonFail(Err.Description, logs)
End Function

' --- RACCompleto -------------------------------------------------------------

Public Function Test_PrecondicionesWorkflowHelper_RACCompleto_AllGreen() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(3)

    logs(0) = "1. Arrange: codigo=X, nombre=Firmante, decision=APROBADO, aprobacion=True"

    If Not PrecondicionesWorkflowHelper_RACCompleto("X", "Firmante", "APROBADO", True) Then _
        Err.Raise 513, , "RAC con todos los campos llenos + aprobacion=True debe devolver True"
    logs(1) = "2. PASS: RACCompleto=True"

    ' Sad complementario: misma entrada pero aprobacion=False -> False.
    If PrecondicionesWorkflowHelper_RACCompleto("X", "Firmante", "APROBADO", False) Then _
        Err.Raise 513, , "AprobacionSuministrador=False debe rechazar"
    logs(2) = "3. PASS: aprobacion=False -> RACCompleto=False"

    Test_PrecondicionesWorkflowHelper_RACCompleto_AllGreen = TestHelper.BuildJsonOk("true", logs)
    Exit Function

EH:
    logs(UBound(logs)) = "ERR " & Err.Number & " - " & Err.Description
    Test_PrecondicionesWorkflowHelper_RACCompleto_AllGreen = TestHelper.BuildJsonFail(Err.Description, logs)
End Function

Public Function Test_PrecondicionesWorkflowHelper_RACCompleto_MissingAprobacionSuministrador_False() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(3)

    logs(0) = "1. Arrange: codigo/nombre/decision llenos pero aprobacion=False"

    If PrecondicionesWorkflowHelper_RACCompleto("X", "Firmante", "APROBADO", False) Then _
        Err.Raise 513, , "Falta Aprobacion Suministrador -> RACCompleto=False"
    logs(1) = "2. PASS: aprobacion=False -> False"

    ' Variante sin aprobacion pero con RECHAZADO (la exencion de BR-008
    ' no exime la Aprobacion Suministrador).
    If PrecondicionesWorkflowHelper_RACCompleto("X", "Firmante", "RECHAZADO", False) Then _
        Err.Raise 513, , "RECHAZADO sin aprobacion -> RACCompleto=False"
    logs(2) = "3. PASS: RECHAZADO sin aprobacion -> False"

    Test_PrecondicionesWorkflowHelper_RACCompleto_MissingAprobacionSuministrador_False = TestHelper.BuildJsonOk("false", logs)
    Exit Function

EH:
    logs(UBound(logs)) = "ERR " & Err.Number & " - " & Err.Description
    Test_PrecondicionesWorkflowHelper_RACCompleto_MissingAprobacionSuministrador_False = TestHelper.BuildJsonFail(Err.Description, logs)
End Function

Public Function Test_PrecondicionesWorkflowHelper_RACCompleto_NullDecision_False() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(3)

    logs(0) = "1. Arrange: codigo=X, nombre=Firmante, decision=Null (DAO Null, no string)"

    ' Variant absorbe Null sin error 94.
    If PrecondicionesWorkflowHelper_RACCompleto("X", "Firmante", Null, True) Then _
        Err.Raise 513, , "decision=Null debe devolver False (BR-008 exige decision)"
    logs(1) = "2. PASS: decision=Null -> False sin error 94"

    ' Variante: nombre=Null.
    If PrecondicionesWorkflowHelper_RACCompleto("X", Null, "APROBADO", True) Then _
        Err.Raise 513, , "nombre=Null debe devolver False"
    logs(2) = "3. PASS: nombre=Null -> False sin error 94"

    Test_PrecondicionesWorkflowHelper_RACCompleto_NullDecision_False = TestHelper.BuildJsonOk("false", logs)
    Exit Function

EH:
    logs(UBound(logs)) = "ERR " & Err.Number & " - " & Err.Description
    Test_PrecondicionesWorkflowHelper_RACCompleto_NullDecision_False = TestHelper.BuildJsonFail(Err.Description, logs)
End Function

Public Function Test_PrecondicionesWorkflowHelper_RACCompleto_Rechazado_AllowsEmptyCodigo() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(4)

    logs(0) = "1. Arrange: decision=RECHAZADO, codigo vacio, nombre=Firmante, aprobacion=True (BR-008)"

    If Not PrecondicionesWorkflowHelper_RACCompleto("", "Firmante", "RECHAZADO", True) Then _
        Err.Raise 513, , "RECHAZADO debe eximir codigo cuando nombre y aprobacion estan"
    logs(1) = "2. PASS: RECHAZADO + codigo vacio + nombre + aprobacion=True -> True"

    ' Case-insensitive: rechazado en minusculas.
    If Not PrecondicionesWorkflowHelper_RACCompleto("", "Firmante", "rechazado", True) Then _
        Err.Raise 513, , "BR-008 debe ser case-insensitive (rechazado)"
    logs(2) = "3. PASS: case-insensitive"

    ' Sin embargo, sin aprobacion -> False.
    If PrecondicionesWorkflowHelper_RACCompleto("", "Firmante", "RECHAZADO", False) Then _
        Err.Raise 513, , "AprobacionSuministrador sigue siendo obligatoria bajo RECHAZADO"
    logs(3) = "4. PASS: RECHAZADO sin aprobacion -> False"

    Test_PrecondicionesWorkflowHelper_RACCompleto_Rechazado_AllowsEmptyCodigo = TestHelper.BuildJsonOk("true", logs)
    Exit Function

EH:
    logs(UBound(logs)) = "ERR " & Err.Number & " - " & Err.Description
    Test_PrecondicionesWorkflowHelper_RACCompleto_Rechazado_AllowsEmptyCodigo = TestHelper.BuildJsonFail(Err.Description, logs)
End Function

' --- AdjuntoEnEtapa ----------------------------------------------------------

Public Function Test_PrecondicionesWorkflowHelper_AdjuntoEnEtapa_Validacion_True() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(3)

    logs(0) = "1. Arrange: etapa=Validacion, adjunto=True (BR-004)"

    If Not PrecondicionesWorkflowHelper_AdjuntoEnEtapa("Validacion", True) Then _
        Err.Raise 513, , "Validacion con adjunto=True debe devolver True"
    logs(1) = "2. PASS: etapa Valida + adjunto=True -> True"

    ' Case-insensitive en etapa.
    If Not PrecondicionesWorkflowHelper_AdjuntoEnEtapa("VALIDACION", True) Then _
        Err.Raise 513, , "etapa case-insensitive"
    logs(2) = "3. PASS: case-insensitive"

    Test_PrecondicionesWorkflowHelper_AdjuntoEnEtapa_Validacion_True = TestHelper.BuildJsonOk("true", logs)
    Exit Function

EH:
    logs(UBound(logs)) = "ERR " & Err.Number & " - " & Err.Description
    Test_PrecondicionesWorkflowHelper_AdjuntoEnEtapa_Validacion_True = TestHelper.BuildJsonFail(Err.Description, logs)
End Function

Public Function Test_PrecondicionesWorkflowHelper_AdjuntoEnEtapa_Cierre_FalseWithoutAdjunto() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(3)

    logs(0) = "1. Arrange: etapa=Cierre, adjunto=False (BR-006)"

    If PrecondicionesWorkflowHelper_AdjuntoEnEtapa("Cierre", False) Then _
        Err.Raise 513, , "Cierre sin adjunto debe devolver False"
    logs(1) = "2. PASS: etapa Cierre + adjunto=False -> False"

    ' Adjunto Long 0 (DAO Long) tambien cuenta como False.
    If PrecondicionesWorkflowHelper_AdjuntoEnEtapa("Cierre", 0) Then _
        Err.Raise 513, , "Long 0 debe tratarse como False"
    logs(2) = "3. PASS: adjunto Long 0 -> False"

    Test_PrecondicionesWorkflowHelper_AdjuntoEnEtapa_Cierre_FalseWithoutAdjunto = TestHelper.BuildJsonOk("false", logs)
    Exit Function

EH:
    logs(UBound(logs)) = "ERR " & Err.Number & " - " & Err.Description
    Test_PrecondicionesWorkflowHelper_AdjuntoEnEtapa_Cierre_FalseWithoutAdjunto = TestHelper.BuildJsonFail(Err.Description, logs)
End Function

Public Function Test_PrecondicionesWorkflowHelper_AdjuntoEnEtapa_UnknownEtapa_False() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(4)

    logs(0) = "1. Arrange: etapa=Patata (no canonica), adjunto=True"

    If PrecondicionesWorkflowHelper_AdjuntoEnEtapa("Patata", True) Then _
        Err.Raise 513, , "Etapa desconocida debe devolver False aunque adjunto=True"
    logs(1) = "2. PASS: etapa Patata -> False"

    ' Etapa vacia -> False.
    If PrecondicionesWorkflowHelper_AdjuntoEnEtapa("", True) Then _
        Err.Raise 513, , "Etapa vacia debe devolver False"
    logs(2) = "3. PASS: etapa vacia -> False"

    ' Etapa Null -> False (Variant absorbe Null).
    If PrecondicionesWorkflowHelper_AdjuntoEnEtapa(Null, True) Then _
        Err.Raise 513, , "Etapa Null debe devolver False"
    logs(3) = "4. PASS: etapa Null -> False"

    Test_PrecondicionesWorkflowHelper_AdjuntoEnEtapa_UnknownEtapa_False = TestHelper.BuildJsonOk("false", logs)
    Exit Function

EH:
    logs(UBound(logs)) = "ERR " & Err.Number & " - " & Err.Description
    Test_PrecondicionesWorkflowHelper_AdjuntoEnEtapa_UnknownEtapa_False = TestHelper.BuildJsonFail(Err.Description, logs)
End Function

' --- ResultadoValidacionAprobado ---------------------------------------------

Public Function Test_PrecondicionesWorkflowHelper_ResultadoValidacionAprobado_APROBADO_True() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(3)

    logs(0) = "1. Arrange: p_UltimoResultado='APROBADO' (BR-005)"

    If Not PrecondicionesWorkflowHelper_ResultadoValidacionAprobado("APROBADO") Then _
        Err.Raise 513, , "APROBADO debe devolver True"
    logs(1) = "2. PASS: APROBADO -> True"

    ' Con espacios exteriores.
    If Not PrecondicionesWorkflowHelper_ResultadoValidacionAprobado("  APROBADO  ") Then _
        Err.Raise 513, , "trim antes de comparar"
    logs(2) = "3. PASS: trim de APROBADO con espacios"

    Test_PrecondicionesWorkflowHelper_ResultadoValidacionAprobado_APROBADO_True = TestHelper.BuildJsonOk("true", logs)
    Exit Function

EH:
    logs(UBound(logs)) = "ERR " & Err.Number & " - " & Err.Description
    Test_PrecondicionesWorkflowHelper_ResultadoValidacionAprobado_APROBADO_True = TestHelper.BuildJsonFail(Err.Description, logs)
End Function

Public Function Test_PrecondicionesWorkflowHelper_ResultadoValidacionAprobado_PENDIENTE_False() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(4)

    logs(0) = "1. Arrange: resultados != APROBADO"

    If PrecondicionesWorkflowHelper_ResultadoValidacionAprobado("PENDIENTE") Then _
        Err.Raise 513, , "PENDIENTE debe devolver False"
    logs(1) = "2. PASS: PENDIENTE -> False"

    If PrecondicionesWorkflowHelper_ResultadoValidacionAprobado("RECHAZADO") Then _
        Err.Raise 513, , "RECHAZADO debe devolver False"
    logs(2) = "3. PASS: RECHAZADO -> False"

    If PrecondicionesWorkflowHelper_ResultadoValidacionAprobado("CON_COMENTARIOS") Then _
        Err.Raise 513, , "CON_COMENTARIOS debe devolver False"
    logs(3) = "4. PASS: CON_COMENTARIOS -> False"

    Test_PrecondicionesWorkflowHelper_ResultadoValidacionAprobado_PENDIENTE_False = TestHelper.BuildJsonOk("false", logs)
    Exit Function

EH:
    logs(UBound(logs)) = "ERR " & Err.Number & " - " & Err.Description
    Test_PrecondicionesWorkflowHelper_ResultadoValidacionAprobado_PENDIENTE_False = TestHelper.BuildJsonFail(Err.Description, logs)
End Function

Public Function Test_PrecondicionesWorkflowHelper_ResultadoValidacionAprobado_CaseInsensitive() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(4)

    logs(0) = "1. Arrange: variantes de case para APROBADO"

    If Not PrecondicionesWorkflowHelper_ResultadoValidacionAprobado("aprobado") Then _
        Err.Raise 513, , "minusculas deben contar como APROBADO"
    logs(1) = "2. PASS: 'aprobado' -> True"

    If Not PrecondicionesWorkflowHelper_ResultadoValidacionAprobado("Aprobado") Then _
        Err.Raise 513, , "capitalizado debe contar como APROBADO"
    logs(2) = "3. PASS: 'Aprobado' -> True"

    If Not PrecondicionesWorkflowHelper_ResultadoValidacionAprobado("ApRoBaDo") Then _
        Err.Raise 513, , "mixto debe contar como APROBADO"
    logs(3) = "4. PASS: 'ApRoBaDo' -> True"

    Test_PrecondicionesWorkflowHelper_ResultadoValidacionAprobado_CaseInsensitive = TestHelper.BuildJsonOk("true", logs)
    Exit Function

EH:
    logs(UBound(logs)) = "ERR " & Err.Number & " - " & Err.Description
    Test_PrecondicionesWorkflowHelper_ResultadoValidacionAprobado_CaseInsensitive = TestHelper.BuildJsonFail(Err.Description, logs)
End Function
