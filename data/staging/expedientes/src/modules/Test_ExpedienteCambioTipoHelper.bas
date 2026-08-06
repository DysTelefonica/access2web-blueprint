Attribute VB_Name = "Test_ExpedienteCambioTipoHelper"
Option Compare Database
Option Explicit

' Test_ExpedienteCambioTipoHelper — REWORK (2026-06-26, Phase 3.2b / PR-8b)
' Pure-data TDD atoms for modExpedienteCambioTipoHelper.bas
' (Form_FormExpedienteCambioTipo, see docs/audit/form-expediente-cambio-tipo-pure-data.md).
'
' Anti-pattern removed:
'   - No DoCmd.OpenForm / Forms(...) / Screen.ActiveForm references.
'   - Stubs are Scripting.Dictionary instances.
'
' Implements access-vba-tdd skill:
'   - §1.8 declaration ordering: all Private Const/Function/Sub at top, Public atoms after.
'   - §1.10 atom signature MUST match helper signature EXACTLY.
'   - §4.2 no-humo: atoms assert concrete values.
'   - §4.4 strong assertions.

' === Module-level constants (all at top per vba-access §10.1) ====================

Private Const TEST_BASE_ID As Long = 900820


' === Local helpers (all at top per vba-access §10.1) ============================

Private Function BuildOk(ByVal p_Value As Variant, ByRef p_Logs() As String) As String
    BuildOk = TestingCore_BuildOk(p_Value, p_Logs)
End Function

Private Function BuildFail(ByVal p_Error As String, ByRef p_Logs() As String) As String
    BuildFail = TestingCore_BuildFail(p_Error, p_Logs)
End Function

' --- BuildStubExpediente ---------------------------------------------------------
' Builds a Scripting.Dictionary mimicking the Expediente class shape with the
' EsAM/EsLote/EsExpediente/EsBasado flags. Only one is "Sí" at a time.
Private Function BuildStubExpediente( _
    ByVal p_EsAM As String, _
    ByVal p_EsLote As String, _
    ByVal p_EsExpediente As String, _
    ByVal p_EsBasado As String _
) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("EsAM") = p_EsAM
    d("EsLote") = p_EsLote
    d("EsExpediente") = p_EsExpediente
    d("EsBasado") = p_EsBasado
    Set BuildStubExpediente = d
End Function


' === Public atoms ===============================================================

' ---------------------------------------------------------------------------
' 1. EstablecerDatos — happy AM (3 opciones: Lote, Basado, ExpIndep)
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteCambioTipoHelper_EstablecerDatos_HappyAM() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim stub As Object
    Set stub = BuildStubExpediente("Sí", "No", "No", "No")

    Dim p_Error As String
    Dim json As String
    json = modExpedienteCambioTipoHelper.ExpedienteCambioTipo_EstablecerDatos(stub, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If CStr(parsed("payload")("tipoActual")) <> "AM" Then _
        Err.Raise 1002, , "expected tipoActual=AM"

    Dim opts As Object
    Set opts = parsed("payload")("opciones")
    If CLng(opts.Count) <> 3 Then _
        Err.Raise 1002, , "expected 3 opciones for AM"

    Test_ExpedienteCambioTipoHelper_EstablecerDatos_HappyAM = BuildOk("AM", logs)
    Exit Function
EH:
    Test_ExpedienteCambioTipoHelper_EstablecerDatos_HappyAM = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 2. EstablecerDatos — happy LOTE (4 opciones)
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteCambioTipoHelper_EstablecerDatos_HappyLote() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim stub As Object
    Set stub = BuildStubExpediente("No", "Sí", "No", "No")

    Dim p_Error As String
    Dim json As String
    json = modExpedienteCambioTipoHelper.ExpedienteCambioTipo_EstablecerDatos(stub, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If CStr(parsed("payload")("tipoActual")) <> "LOTE" Then _
        Err.Raise 1002, , "expected tipoActual=LOTE"
    If CLng(parsed("payload")("opciones").Count) <> 4 Then _
        Err.Raise 1002, , "expected 4 opciones for LOTE"

    Test_ExpedienteCambioTipoHelper_EstablecerDatos_HappyLote = BuildOk("LOTE", logs)
    Exit Function
EH:
    Test_ExpedienteCambioTipoHelper_EstablecerDatos_HappyLote = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 3. EstablecerDatos — sad (Expediente Nothing)
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteCambioTipoHelper_EstablecerDatos_SadNothing() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modExpedienteCambioTipoHelper.ExpedienteCambioTipo_EstablecerDatos(Nothing, p_Error)
    If InStr(json, """ok"":false") = 0 Then Err.Raise 1001, , "expected ok=false"
    If p_Error = "" Then Err.Raise 1002, , "expected p_Error set"

    Test_ExpedienteCambioTipoHelper_EstablecerDatos_SadNothing = BuildOk("rejected", logs)
    Exit Function
EH:
    Test_ExpedienteCambioTipoHelper_EstablecerDatos_SadNothing = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 4. EstablecerDatos — edge (no flag set -> empty opciones)
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteCambioTipoHelper_EstablecerDatos_EdgeNoFlag() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim stub As Object
    Set stub = BuildStubExpediente("No", "No", "No", "No")

    Dim p_Error As String
    Dim json As String
    json = modExpedienteCambioTipoHelper.ExpedienteCambioTipo_EstablecerDatos(stub, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If CStr(parsed("payload")("tipoActual")) <> "" Then _
        Err.Raise 1002, , "expected tipoActual=''"

    Test_ExpedienteCambioTipoHelper_EstablecerDatos_EdgeNoFlag = BuildOk("empty", logs)
    Exit Function
EH:
    Test_ExpedienteCambioTipoHelper_EstablecerDatos_EdgeNoFlag = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 5. Tipo_AfterUpdate — Convertir a Acuerdo Marco -> locked
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteCambioTipoHelper_Tipo_AfterUpdate_AM_Locked() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modExpedienteCambioTipoHelper.ExpedienteCambioTipo_Tipo_AfterUpdate( _
        "Convertir a Acuerdo Marco", p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If CBool(parsed("payload")("idPadreLocked")) <> True Then _
        Err.Raise 1002, , "expected idPadreLocked=true"
    If CBool(parsed("payload")("idPadreEnabled")) <> False Then _
        Err.Raise 1002, , "expected idPadreEnabled=false"

    Test_ExpedienteCambioTipoHelper_Tipo_AfterUpdate_AM_Locked = BuildOk("locked", logs)
    Exit Function
EH:
    Test_ExpedienteCambioTipoHelper_Tipo_AfterUpdate_AM_Locked = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 6. Tipo_AfterUpdate — Convertir a Lote -> enabled
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteCambioTipoHelper_Tipo_AfterUpdate_Lote_Enabled() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modExpedienteCambioTipoHelper.ExpedienteCambioTipo_Tipo_AfterUpdate( _
        "Convertir a Lote", p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If CBool(parsed("payload")("idPadreEnabled")) <> True Then _
        Err.Raise 1002, , "expected idPadreEnabled=true"

    Test_ExpedienteCambioTipoHelper_Tipo_AfterUpdate_Lote_Enabled = BuildOk("enabled", logs)
    Exit Function
EH:
    Test_ExpedienteCambioTipoHelper_Tipo_AfterUpdate_Lote_Enabled = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 7. Tipo_AfterUpdate — empty tipo -> defaults
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteCambioTipoHelper_Tipo_AfterUpdate_Empty_Defaults() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modExpedienteCambioTipoHelper.ExpedienteCambioTipo_Tipo_AfterUpdate("", p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If CBool(parsed("payload")("idPadreLocked")) <> True Then _
        Err.Raise 1002, , "expected idPadreLocked=true (default)"

    Test_ExpedienteCambioTipoHelper_Tipo_AfterUpdate_Empty_Defaults = BuildOk("defaults", logs)
    Exit Function
EH:
    Test_ExpedienteCambioTipoHelper_Tipo_AfterUpdate_Empty_Defaults = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 8. ValidarPadre — Convertir a Lote + padre AM -> valid
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteCambioTipoHelper_ValidarPadre_LoteConAM_Valid() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modExpedienteCambioTipoHelper.ExpedienteCambioTipo_ValidarPadre( _
        "Convertir a Lote", "PADRE-AM-1", True, False, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If CBool(parsed("payload")("valido")) <> True Then _
        Err.Raise 1002, , "expected valido=true"

    Test_ExpedienteCambioTipoHelper_ValidarPadre_LoteConAM_Valid = BuildOk("valid", logs)
    Exit Function
EH:
    Test_ExpedienteCambioTipoHelper_ValidarPadre_LoteConAM_Valid = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 9. ValidarPadre — Convertir a Lote + padre Lote -> invalid (must be AM)
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteCambioTipoHelper_ValidarPadre_LoteConLote_Invalid() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modExpedienteCambioTipoHelper.ExpedienteCambioTipo_ValidarPadre( _
        "Convertir a Lote", "PADRE-LOTE-1", False, True, p_Error)
    If InStr(json, """ok"":false") = 0 Then _
        Err.Raise 1001, , "expected ok=false"
    If InStr(CStr(JsonConverter.ParseJson(json)("payload")("error")), "Acuerdo Marco") = 0 Then _
        Err.Raise 1002, , "expected error mentioning Acuerdo Marco"

    Test_ExpedienteCambioTipoHelper_ValidarPadre_LoteConLote_Invalid = BuildOk("invalid", logs)
    Exit Function
EH:
    Test_ExpedienteCambioTipoHelper_ValidarPadre_LoteConLote_Invalid = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 10. ValidarPadre — Convertir a Lote sin padre -> invalid
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteCambioTipoHelper_ValidarPadre_LoteSinPadre_Invalid() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modExpedienteCambioTipoHelper.ExpedienteCambioTipo_ValidarPadre( _
        "Convertir a Lote", "", False, False, p_Error)
    If InStr(json, """ok"":false") = 0 Then _
        Err.Raise 1001, , "expected ok=false"

    Test_ExpedienteCambioTipoHelper_ValidarPadre_LoteSinPadre_Invalid = BuildOk("invalid", logs)
    Exit Function
EH:
    Test_ExpedienteCambioTipoHelper_ValidarPadre_LoteSinPadre_Invalid = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 11. ValidarPadre — Convertir a Basado + padre AM -> valid
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteCambioTipoHelper_ValidarPadre_BasadoConAM_Valid() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modExpedienteCambioTipoHelper.ExpedienteCambioTipo_ValidarPadre( _
        "Convertir a Basado", "PADRE-AM-1", True, False, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If CBool(parsed("payload")("valido")) <> True Then _
        Err.Raise 1002, , "expected valido=true"

    Test_ExpedienteCambioTipoHelper_ValidarPadre_BasadoConAM_Valid = BuildOk("valid", logs)
    Exit Function
EH:
    Test_ExpedienteCambioTipoHelper_ValidarPadre_BasadoConAM_Valid = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 12. ValidarPadre — Convertir a Basado + padre Lote -> valid
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteCambioTipoHelper_ValidarPadre_BasadoConLote_Valid() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modExpedienteCambioTipoHelper.ExpedienteCambioTipo_ValidarPadre( _
        "Convertir a Basado", "PADRE-LOTE-1", False, True, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If CBool(parsed("payload")("valido")) <> True Then _
        Err.Raise 1002, , "expected valido=true"

    Test_ExpedienteCambioTipoHelper_ValidarPadre_BasadoConLote_Valid = BuildOk("valid", logs)
    Exit Function
EH:
    Test_ExpedienteCambioTipoHelper_ValidarPadre_BasadoConLote_Valid = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 13. ValidarPadre — Convertir a Basado + padre Indep -> invalid
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteCambioTipoHelper_ValidarPadre_BasadoConIndep_Invalid() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modExpedienteCambioTipoHelper.ExpedienteCambioTipo_ValidarPadre( _
        "Convertir a Basado", "PADRE-INDEP-1", False, False, p_Error)
    If InStr(json, """ok"":false") = 0 Then _
        Err.Raise 1001, , "expected ok=false"

    Test_ExpedienteCambioTipoHelper_ValidarPadre_BasadoConIndep_Invalid = BuildOk("invalid", logs)
    Exit Function
EH:
    Test_ExpedienteCambioTipoHelper_ValidarPadre_BasadoConIndep_Invalid = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 14. ValidarPadre — Convertir a Expediente Independiente con padre -> invalid
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteCambioTipoHelper_ValidarPadre_IndepConPadre_Invalid() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modExpedienteCambioTipoHelper.ExpedienteCambioTipo_ValidarPadre( _
        "Convertir a Expediente Independiente", "PADRE-X", False, False, p_Error)
    If InStr(json, """ok"":false") = 0 Then _
        Err.Raise 1001, , "expected ok=false"

    Test_ExpedienteCambioTipoHelper_ValidarPadre_IndepConPadre_Invalid = BuildOk("invalid", logs)
    Exit Function
EH:
    Test_ExpedienteCambioTipoHelper_ValidarPadre_IndepConPadre_Invalid = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 15. ValidarPadre — Convertir a Expediente Independiente sin padre -> valid
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteCambioTipoHelper_ValidarPadre_IndepSinPadre_Valid() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modExpedienteCambioTipoHelper.ExpedienteCambioTipo_ValidarPadre( _
        "Convertir a Expediente Independiente", "", False, False, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If CBool(parsed("payload")("valido")) <> True Then _
        Err.Raise 1002, , "expected valido=true"

    Test_ExpedienteCambioTipoHelper_ValidarPadre_IndepSinPadre_Valid = BuildOk("valid", logs)
    Exit Function
EH:
    Test_ExpedienteCambioTipoHelper_ValidarPadre_IndepSinPadre_Valid = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 16. ComandoRegistrar — Convertir a Acuerdo Marco -> flags correct
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteCambioTipoHelper_ComandoRegistrar_AM() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim stub As Object
    Set stub = BuildStubExpediente("No", "No", "Sí", "No")

    Dim p_Error As String
    Dim json As String
    json = modExpedienteCambioTipoHelper.ExpedienteCambioTipo_ComandoRegistrar( _
        stub, "Convertir a Acuerdo Marco", "", p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    Dim flags As Object
    Set flags = parsed("payload")("flagsApplied")
    If CStr(flags("EsAM")) <> "Sí" Then Err.Raise 1002, , "expected EsAM=Sí"
    If CStr(flags("EsLote")) <> "No" Then Err.Raise 1002, , "expected EsLote=No"
    If CStr(flags("EsExpediente")) <> "No" Then Err.Raise 1002, , "expected EsExpediente=No"
    If CStr(flags("EsBasado")) <> "No" Then Err.Raise 1002, , "expected EsBasado=No"
    If CStr(flags("IDExpedientePadre")) <> "" Then Err.Raise 1002, , "expected IDExpedientePadre=''"

    Test_ExpedienteCambioTipoHelper_ComandoRegistrar_AM = BuildOk("ok", logs)
    Exit Function
EH:
    Test_ExpedienteCambioTipoHelper_ComandoRegistrar_AM = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 17. ComandoRegistrar — Convertir a Basado -> flags correct
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteCambioTipoHelper_ComandoRegistrar_Basado() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim stub As Object
    Set stub = BuildStubExpediente("No", "No", "No", "No")

    Dim p_Error As String
    Dim json As String
    json = modExpedienteCambioTipoHelper.ExpedienteCambioTipo_ComandoRegistrar( _
        stub, "Convertir a Basado", "PADRE-AM-1", p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    Dim flags As Object
    Set flags = parsed("payload")("flagsApplied")
    If CStr(flags("EsBasado")) <> "Sí" Then Err.Raise 1002, , "expected EsBasado=Sí"
    If CStr(flags("IDExpedientePadre")) <> "PADRE-AM-1" Then _
        Err.Raise 1002, , "expected IDExpedientePadre='PADRE-AM-1'"

    Test_ExpedienteCambioTipoHelper_ComandoRegistrar_Basado = BuildOk("ok", logs)
    Exit Function
EH:
    Test_ExpedienteCambioTipoHelper_ComandoRegistrar_Basado = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 18. ComandoRegistrar — Convertir a Lote -> flags correct
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteCambioTipoHelper_ComandoRegistrar_Lote() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim stub As Object
    Set stub = BuildStubExpediente("No", "No", "No", "No")

    Dim p_Error As String
    Dim json As String
    json = modExpedienteCambioTipoHelper.ExpedienteCambioTipo_ComandoRegistrar( _
        stub, "Convertir a Lote", "PADRE-AM-1", p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    Dim flags As Object
    Set flags = parsed("payload")("flagsApplied")
    If CStr(flags("EsLote")) <> "Sí" Then Err.Raise 1002, , "expected EsLote=Sí"
    If CStr(flags("IDExpedientePadre")) <> "PADRE-AM-1" Then _
        Err.Raise 1002, , "expected IDExpedientePadre='PADRE-AM-1'"

    Test_ExpedienteCambioTipoHelper_ComandoRegistrar_Lote = BuildOk("ok", logs)
    Exit Function
EH:
    Test_ExpedienteCambioTipoHelper_ComandoRegistrar_Lote = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 19. ComandoRegistrar — Convertir a Expediente Independiente -> flags correct
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteCambioTipoHelper_ComandoRegistrar_Indep() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim stub As Object
    Set stub = BuildStubExpediente("No", "No", "No", "No")

    Dim p_Error As String
    Dim json As String
    json = modExpedienteCambioTipoHelper.ExpedienteCambioTipo_ComandoRegistrar( _
        stub, "Convertir a Expediente Independiente", "", p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    Dim flags As Object
    Set flags = parsed("payload")("flagsApplied")
    If CStr(flags("EsExpediente")) <> "Sí" Then Err.Raise 1002, , "expected EsExpediente=Sí"
    If CStr(flags("IDExpedientePadre")) <> "" Then _
        Err.Raise 1002, , "expected IDExpedientePadre=''"

    Test_ExpedienteCambioTipoHelper_ComandoRegistrar_Indep = BuildOk("ok", logs)
    Exit Function
EH:
    Test_ExpedienteCambioTipoHelper_ComandoRegistrar_Indep = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 20. ComandoRegistrar — Convertir a Lote de Acuerdo Marco -> flags correct
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteCambioTipoHelper_ComandoRegistrar_LoteAM() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim stub As Object
    Set stub = BuildStubExpediente("No", "No", "No", "No")

    Dim p_Error As String
    Dim json As String
    json = modExpedienteCambioTipoHelper.ExpedienteCambioTipo_ComandoRegistrar( _
        stub, "Convertir a Lote de Acuerdo Marco", "PADRE-AM-1", p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    Dim flags As Object
    Set flags = parsed("payload")("flagsApplied")
    If CStr(flags("EsLote")) <> "Sí" Then Err.Raise 1002, , "expected EsLote=Sí"
    If CStr(flags("IDExpedientePadre")) <> "PADRE-AM-1" Then _
        Err.Raise 1002, , "expected IDExpedientePadre='PADRE-AM-1'"

    Test_ExpedienteCambioTipoHelper_ComandoRegistrar_LoteAM = BuildOk("ok", logs)
    Exit Function
EH:
    Test_ExpedienteCambioTipoHelper_ComandoRegistrar_LoteAM = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 21. ComandoRegistrar — sad (empty tipo)
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteCambioTipoHelper_ComandoRegistrar_SadEmptyTipo() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim stub As Object
    Set stub = BuildStubExpediente("No", "No", "No", "No")

    Dim p_Error As String
    Dim json As String
    json = modExpedienteCambioTipoHelper.ExpedienteCambioTipo_ComandoRegistrar( _
        stub, "", "", p_Error)
    If InStr(json, """ok"":false") = 0 Then _
        Err.Raise 1001, , "expected ok=false"

    Test_ExpedienteCambioTipoHelper_ComandoRegistrar_SadEmptyTipo = BuildOk("rejected", logs)
    Exit Function
EH:
    Test_ExpedienteCambioTipoHelper_ComandoRegistrar_SadEmptyTipo = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 22. ComandoRegistrar — sad (Expediente Nothing)
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteCambioTipoHelper_ComandoRegistrar_SadNothing() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modExpedienteCambioTipoHelper.ExpedienteCambioTipo_ComandoRegistrar( _
        Nothing, "Convertir a Acuerdo Marco", "", p_Error)
    If InStr(json, """ok"":false") = 0 Then _
        Err.Raise 1001, , "expected ok=false"

    Test_ExpedienteCambioTipoHelper_ComandoRegistrar_SadNothing = BuildOk("rejected", logs)
    Exit Function
EH:
    Test_ExpedienteCambioTipoHelper_ComandoRegistrar_SadNothing = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 23. ComandoRegistrar — sad (unknown tipo)
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteCambioTipoHelper_ComandoRegistrar_SadUnknownTipo() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim stub As Object
    Set stub = BuildStubExpediente("No", "No", "No", "No")

    Dim p_Error As String
    Dim json As String
    json = modExpedienteCambioTipoHelper.ExpedienteCambioTipo_ComandoRegistrar( _
        stub, "Convertir a Wat", "", p_Error)
    If InStr(json, """ok"":false") = 0 Then _
        Err.Raise 1001, , "expected ok=false"

    Test_ExpedienteCambioTipoHelper_ComandoRegistrar_SadUnknownTipo = BuildOk("rejected", logs)
    Exit Function
EH:
    Test_ExpedienteCambioTipoHelper_ComandoRegistrar_SadUnknownTipo = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 24. RunAll — wrapper for Dysflow manifest discovery
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteCambioTipoHelper_RunAll() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim atoms As Variant
    atoms = Array( _
        "Test_ExpedienteCambioTipoHelper_EstablecerDatos_HappyAM", _
        "Test_ExpedienteCambioTipoHelper_EstablecerDatos_HappyLote", _
        "Test_ExpedienteCambioTipoHelper_EstablecerDatos_SadNothing", _
        "Test_ExpedienteCambioTipoHelper_EstablecerDatos_EdgeNoFlag", _
        "Test_ExpedienteCambioTipoHelper_Tipo_AfterUpdate_AM_Locked", _
        "Test_ExpedienteCambioTipoHelper_Tipo_AfterUpdate_Lote_Enabled", _
        "Test_ExpedienteCambioTipoHelper_Tipo_AfterUpdate_Empty_Defaults", _
        "Test_ExpedienteCambioTipoHelper_ValidarPadre_LoteConAM_Valid", _
        "Test_ExpedienteCambioTipoHelper_ValidarPadre_LoteConLote_Invalid", _
        "Test_ExpedienteCambioTipoHelper_ValidarPadre_LoteSinPadre_Invalid", _
        "Test_ExpedienteCambioTipoHelper_ValidarPadre_BasadoConAM_Valid", _
        "Test_ExpedienteCambioTipoHelper_ValidarPadre_BasadoConLote_Valid", _
        "Test_ExpedienteCambioTipoHelper_ValidarPadre_BasadoConIndep_Invalid", _
        "Test_ExpedienteCambioTipoHelper_ValidarPadre_IndepConPadre_Invalid", _
        "Test_ExpedienteCambioTipoHelper_ValidarPadre_IndepSinPadre_Valid", _
        "Test_ExpedienteCambioTipoHelper_ComandoRegistrar_AM", _
        "Test_ExpedienteCambioTipoHelper_ComandoRegistrar_Basado", _
        "Test_ExpedienteCambioTipoHelper_ComandoRegistrar_Lote", _
        "Test_ExpedienteCambioTipoHelper_ComandoRegistrar_Indep", _
        "Test_ExpedienteCambioTipoHelper_ComandoRegistrar_LoteAM", _
        "Test_ExpedienteCambioTipoHelper_ComandoRegistrar_SadEmptyTipo", _
        "Test_ExpedienteCambioTipoHelper_ComandoRegistrar_SadNothing", _
        "Test_ExpedienteCambioTipoHelper_ComandoRegistrar_SadUnknownTipo" _
    )

    Dim passed As Long
    Dim failed As Long
    Dim firstFailure As String
    Dim i As Long
    For i = LBound(atoms) To UBound(atoms)
        Dim result As String
        result = Application.Run(CStr(atoms(i)))
        If InStr(result, """ok"":true") > 0 Then
            passed = passed + 1
        Else
            failed = failed + 1
            If Len(firstFailure) = 0 Then firstFailure = CStr(atoms(i)) & " -> " & result
        End If
    Next i

    logs(0) = "RunAll: passed=" & passed & " failed=" & failed
    If failed > 0 Then
        logs(1) = "RunAll: firstFailure=" & firstFailure
        Test_ExpedienteCambioTipoHelper_RunAll = BuildFail("firstFailure=" & firstFailure, logs)
    Else
        Test_ExpedienteCambioTipoHelper_RunAll = BuildOk("all-passed", logs)
    End If
    Exit Function
EH:
    Dim p_Error As String
    p_Error = "Test_ExpedienteCambioTipoHelper_RunAll EH: " & Err.Description
    Test_ExpedienteCambioTipoHelper_RunAll = BuildFail(p_Error, logs)
End Function