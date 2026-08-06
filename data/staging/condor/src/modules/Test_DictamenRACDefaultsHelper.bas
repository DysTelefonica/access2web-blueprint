Attribute VB_Name = "Test_DictamenRACDefaultsHelper"
Option Compare Database
Option Explicit

' ==========================================================================
' Test_DictamenRACDefaultsHelper - Slice 3.2 DictamenRAC pure helper atoms
'
' Each atom is a Public Function (global, unique) that returns the
' canonical JSON via TestHelper.BuildJsonOk / BuildJsonFail. Pure
' helper tests: NO DAO, NO controls, NO MsgBox, no fixtures required.
' ==========================================================================

Public Function Test_DictamenRACDefaultsHelper_PC_ApplyWhenRacAvailable() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(4)
    Dim prompt As String
    Dim msg As String
    Dim applied As Boolean

    logs(0) = "1. Arrange: PC, bloque editable, expediente con 'RAC-1234'"

    applied = DictamenRACDefaultsHelper_ConstruirPlan( _
        "PC", "Principal", "RAC-1234", True, "", prompt, msg)
    logs(1) = "2. Act: ConstruirPlan con tipo=PC, RAC presente"

    If Not applied Then Err.Raise 513, , "PC debe aplicar defaults cuando hay RAC"
    If prompt <> "OK" Then Err.Raise 513, , "PromptResult esperado 'OK', obtenido '" & prompt & "'"
    If InStr(msg, "racNombre") = 0 Then _
        Err.Raise 513, , "MessagePlan debe codificar target 'racNombre', obtenido '" & msg & "'"
    If InStr(msg, "RAC cargado") = 0 Then _
        Err.Raise 513, , "MessagePlan debe incluir 'RAC cargado', obtenido '" & msg & "'"
    logs(2) = "3. PASS: apply=True, prompt='OK', msg target=racNombre"
    logs(3) = "4. PASS: PC happy path con seam rendereable"

    Test_DictamenRACDefaultsHelper_PC_ApplyWhenRacAvailable = TestHelper.BuildJsonOk("true", logs)
    Exit Function

EH:
    logs(UBound(logs)) = "ERR " & Err.Number & " - " & Err.Description
    Test_DictamenRACDefaultsHelper_PC_ApplyWhenRacAvailable = TestHelper.BuildJsonFail(Err.Description, logs)
End Function

Public Function Test_DictamenRACDefaultsHelper_PCSUB_RoutesToDelegadoNombre() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(4)
    Dim prompt As String
    Dim msg As String
    Dim applied As Boolean

    logs(0) = "1. Arrange: PCSUB, bloque editable, expediente con 'RAC-1234' (asimetria)"

    applied = DictamenRACDefaultsHelper_ConstruirPlan( _
        "PCSUB", "Principal", "RAC-1234", True, "", prompt, msg)
    logs(1) = "2. Act: ConstruirPlan con tipo=PCSUB"

    If Not applied Then Err.Raise 513, , "PCSUB debe aplicar defaults cuando hay RAC"
    If prompt <> "OK" Then Err.Raise 513, , "PromptResult esperado 'OK', obtenido '" & prompt & "'"
    If InStr(msg, "racDelegadoNombre") = 0 Then _
        Err.Raise 513, , "MessagePlan debe codificar target 'racDelegadoNombre' (asimetria PCSUB), obtenido '" & msg & "'"
    If InStr(msg, "RAC cargado") = 0 Then _
        Err.Raise 513, , "MessagePlan debe incluir 'RAC cargado', obtenido '" & msg & "'"
    logs(2) = "3. PASS: target field es racDelegadoNombre (no racNombre)"
    logs(3) = "4. PASS: asimetria PCSUB resuelta en helper, no en Form"

    Test_DictamenRACDefaultsHelper_PCSUB_RoutesToDelegadoNombre = TestHelper.BuildJsonOk("true", logs)
    Exit Function

EH:
    logs(UBound(logs)) = "ERR " & Err.Number & " - " & Err.Description
    Test_DictamenRACDefaultsHelper_PCSUB_RoutesToDelegadoNombre = TestHelper.BuildJsonFail(Err.Description, logs)
End Function

Public Function Test_DictamenRACDefaultsHelper_CDCA_BlocksWhenNotEditable() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(4)
    Dim prompt As String
    Dim msg As String
    Dim applied As Boolean

    logs(0) = "1. Arrange: CDCA, bloque NO editable (workflow), expediente con RAC"

    applied = DictamenRACDefaultsHelper_ConstruirPlan( _
        "CDCA", "Principal", "RAC-1234", False, "", prompt, msg)
    logs(1) = "2. Act: ConstruirPlan con p_BloqueEditable=False"

    If applied Then Err.Raise 513, , "CDCA NO debe aplicar cuando bloque no editable"
    If prompt <> "BloqueNoEditable" Then _
        Err.Raise 513, , "PromptResult esperado 'BloqueNoEditable', obtenido '" & prompt & "'"
    If Len(Trim$(msg)) = 0 Then Err.Raise 513, , "MessagePlan debe describir motivo del bloqueo, vacio"
    logs(2) = "3. PASS: apply=False, prompt='BloqueNoEditable'"
    logs(3) = "4. PASS: CDCA bloqueado por workflow, sin MsgBox"

    Test_DictamenRACDefaultsHelper_CDCA_BlocksWhenNotEditable = TestHelper.BuildJsonOk("true", logs)
    Exit Function

EH:
    logs(UBound(logs)) = "ERR " & Err.Number & " - " & Err.Description
    Test_DictamenRACDefaultsHelper_CDCA_BlocksWhenNotEditable = TestHelper.BuildJsonFail(Err.Description, logs)
End Function

Public Function Test_DictamenRACDefaultsHelper_CDCASUB_DeniesWhenMissingRac() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(4)
    Dim prompt As String
    Dim msg As String
    Dim applied As Boolean

    logs(0) = "1. Arrange: CDCASUB, bloque editable, expediente SIN RAC"

    applied = DictamenRACDefaultsHelper_ConstruirPlan( _
        "CDCASUB", "Principal", "", True, "", prompt, msg)
    logs(1) = "2. Act: ConstruirPlan con racAsignado vacio"

    If applied Then Err.Raise 513, , "CDCASUB NO debe aplicar cuando expediente sin RAC"
    If prompt <> "SinRACAsignado" Then _
        Err.Raise 513, , "PromptResult esperado 'SinRACAsignado', obtenido '" & prompt & "'"
    If Len(Trim$(msg)) = 0 Then Err.Raise 513, , "MessagePlan debe describir motivo (sin RAC), vacio"
    logs(2) = "3. PASS: apply=False, prompt='SinRACAsignado'"
    logs(3) = "4. PASS: CDCASUB niega con prompt seam rendereable"

    Test_DictamenRACDefaultsHelper_CDCASUB_DeniesWhenMissingRac = TestHelper.BuildJsonOk("true", logs)
    Exit Function

EH:
    logs(UBound(logs)) = "ERR " & Err.Number & " - " & Err.Description
    Test_DictamenRACDefaultsHelper_CDCASUB_DeniesWhenMissingRac = TestHelper.BuildJsonFail(Err.Description, logs)
End Function

Public Function Test_DictamenRACDefaultsHelper_Regla_Rechazado_AllowsEmptyCodigo() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(4)

    logs(0) = "1. Arrange: decision=RECHAZADO, codigo vacio, nombre=Firmante X (BR-008)"

    If Not DictamenRACDefaultsHelper_ReglaEsCompleta("", "Firmante X", "RECHAZADO") Then _
        Err.Raise 513, , "RECHAZADO debe eximir codigo si nombre esta presente"
    logs(1) = "2. PASS: RECHAZADO + nombre no-vacio -> completo=True"

    If Not DictamenRACDefaultsHelper_ReglaEsCompleta("", "Firmante X", "rechazado") Then _
        Err.Raise 513, , "La regla debe ser case-insensitive (rechazado en minusculas)"
    logs(2) = "3. PASS: case-insensitive"

    ' Sad complementario: RECHAZADO exige nombre (no solo codigo).
    If DictamenRACDefaultsHelper_ReglaEsCompleta("X", "", "RECHAZADO") Then _
        Err.Raise 513, , "RECHAZADO sin nombre debe rechazar aunque codigo este presente"
    logs(3) = "4. PASS: RECHAZADO sin nombre -> completo=False"

    Test_DictamenRACDefaultsHelper_Regla_Rechazado_AllowsEmptyCodigo = TestHelper.BuildJsonOk("true", logs)
    Exit Function

EH:
    logs(UBound(logs)) = "ERR " & Err.Number & " - " & Err.Description
    Test_DictamenRACDefaultsHelper_Regla_Rechazado_AllowsEmptyCodigo = TestHelper.BuildJsonFail(Err.Description, logs)
End Function

Public Function Test_DictamenRACDefaultsHelper_Regla_RejectsBlankNombre() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(4)

    logs(0) = "1. Arrange: codigo=X, decision=APROBADO, nombre vacio"

    If DictamenRACDefaultsHelper_ReglaEsCompleta("X", "", "APROBADO") Then _
        Err.Raise 513, , "Nombre vacio debe rechazar incluso con codigo y decision"
    logs(1) = "2. PASS: codigo+decision sin nombre -> completo=False"

    If DictamenRACDefaultsHelper_ReglaEsCompleta("X", "   ", "APROBADO") Then _
        Err.Raise 513, , "Nombre solo espacios debe rechazarse"
    logs(2) = "3. PASS: nombre con solo espacios -> completo=False"

    ' Sanity check: tres campos no vacios en APROBADO -> completo=True
    If Not DictamenRACDefaultsHelper_ReglaEsCompleta("X", "Firmante", "APROBADO") Then _
        Err.Raise 513, , "Sanity: tres campos no vacios en APROBADO debe ser completo"
    logs(3) = "4. PASS: sanity check (codigo+nombre+decision no-vacios -> completo=True)"

    Test_DictamenRACDefaultsHelper_Regla_RejectsBlankNombre = TestHelper.BuildJsonOk("true", logs)
    Exit Function

EH:
    logs(UBound(logs)) = "ERR " & Err.Number & " - " & Err.Description
    Test_DictamenRACDefaultsHelper_Regla_RejectsBlankNombre = TestHelper.BuildJsonFail(Err.Description, logs)
End Function

Public Function Test_DictamenRACDefaultsHelper_ConstruirPlan_HandlesNullRacAsignado() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(4)
    Dim prompt As String
    Dim msg As String
    Dim applied As Boolean

    logs(0) = "1. Arrange: PC, expediente devuelve Null (DAO Null) en lugar de string vacio"

    applied = DictamenRACDefaultsHelper_ConstruirPlan( _
        "PC", "Principal", Null, True, "", prompt, msg)
    logs(1) = "2. Act: ConstruirPlan con p_RacAsignadoByExpediente=Null"

    If applied Then Err.Raise 513, , "Null en racAsignado debe tratarse como sin RAC"
    If prompt <> "SinRACAsignado" Then _
        Err.Raise 513, , "PromptResult esperado 'SinRACAsignado', obtenido '" & prompt & "'"
    logs(2) = "3. PASS: Null -> SinRACAsignado (robusto)"
    logs(3) = "4. PASS: helper tolera Null en el input Variant"

    Test_DictamenRACDefaultsHelper_ConstruirPlan_HandlesNullRacAsignado = TestHelper.BuildJsonOk("true", logs)
    Exit Function

EH:
    logs(UBound(logs)) = "ERR " & Err.Number & " - " & Err.Description
    Test_DictamenRACDefaultsHelper_ConstruirPlan_HandlesNullRacAsignado = TestHelper.BuildJsonFail(Err.Description, logs)
End Function

Public Function Test_DictamenRACDefaultsHelper_Regla_TrimsWhitespaceBeforeValidation() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(4)

    logs(0) = "1. Arrange: valores rodeados de espacios"

    ' RECHAZADO con codigo vacio rodeado de espacios y nombre valido.
    If Not DictamenRACDefaultsHelper_ReglaEsCompleta("   ", "  Firmante  ", "  RECHAZADO  ") Then _
        Err.Raise 513, , "Debe aceptar codigo solo espacios bajo RECHAZADO (exento)"
    logs(1) = "2. PASS: RECHAZADO codigo solo espacios -> completo=True"

    ' APROBADO con codigo valido y nombre solo espacios -> rechaza.
    If DictamenRACDefaultsHelper_ReglaEsCompleta("X", "   ", "APROBADO") Then _
        Err.Raise 513, , "Solo espacios en nombre debe rechazarse"
    logs(2) = "3. PASS: nombre solo espacios -> completo=False"

    ' APROBADO con todo valido -> acepta.
    If Not DictamenRACDefaultsHelper_ReglaEsCompleta("  X  ", "  Firmante  ", "  APROBADO  ") Then _
        Err.Raise 513, , "Debe aceptar valores con espacios exteriores"
    logs(3) = "4. PASS: trim acepta valores validos con espacios exteriores"

    Test_DictamenRACDefaultsHelper_Regla_TrimsWhitespaceBeforeValidation = TestHelper.BuildJsonOk("true", logs)
    Exit Function

EH:
    logs(UBound(logs)) = "ERR " & Err.Number & " - " & Err.Description
    Test_DictamenRACDefaultsHelper_Regla_TrimsWhitespaceBeforeValidation = TestHelper.BuildJsonFail(Err.Description, logs)
End Function