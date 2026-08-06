Attribute VB_Name = "Test_ExpedienteAltaHelper"
Option Compare Database
Option Explicit

' Test_ExpedienteAltaHelper — REWORK (2026-06-26, Phase 3.2b / PR-8b)
' Pure-data TDD atoms for modExpedienteAltaHelper.bas
' (Form_FormExpedienteAlta, see docs/audit/form-expediente-alta-pure-data.md).
'
' Implements access-vba-tdd skill:
'   - §1.8 declaration ordering: all Private Const/Function/Sub at top, Public atoms after.
'   - §1.10 atom signature MUST match helper signature EXACTLY.

' === Module-level constants (all at top per vba-access §10.1) ====================

Private Const TEST_BASE_ID As Long = 900830


' === Local helpers (all at top per vba-access §10.1) ============================

Private Function BuildOk(ByVal p_Value As Variant, ByRef p_Logs() As String) As String
    BuildOk = TestingCore_BuildOk(p_Value, p_Logs)
End Function

Private Function BuildFail(ByVal p_Error As String, ByRef p_Logs() As String) As String
    BuildFail = TestingCore_BuildFail(p_Error, p_Logs)
End Function

' --- BuildStubExpediente ---------------------------------------------------------
Private Function BuildStubExpediente( _
    ByVal p_IDExpediente As String, _
    ByVal p_Ordinal As String, _
    ByVal p_OrdinalCalculado As String _
) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("IDExpediente") = p_IDExpediente
    d("Ordinal") = p_Ordinal
    d("OrdinalCalculado") = p_OrdinalCalculado
    Set BuildStubExpediente = d
End Function

' --- BuildStubDTO ----------------------------------------------------------------
Private Function BuildStubDTO(ByVal p_Expediente As Object) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Set d("Expediente") = p_Expediente
    Set BuildStubDTO = d
End Function


' === Public atoms ===============================================================

' ---------------------------------------------------------------------------
' 1. EstablecerDatos — happy AM (ordinal hidden)
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteAltaHelper_EstablecerDatos_HappyAMOrdinalHidden() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim stub As Object
    Set stub = BuildStubDTO(BuildStubExpediente("", "", ""))

    Dim p_Error As String
    Dim json As String
    json = modExpedienteAltaHelper.ExpedienteAlta_EstablecerDatos( _
        stub, 1, "", p_Error)  ' AM=1
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If CBool(parsed("payload")("ordinalVisible")) <> False Then _
        Err.Raise 1002, , "expected ordinalVisible=false for AM"

    Test_ExpedienteAltaHelper_EstablecerDatos_HappyAMOrdinalHidden = BuildOk("ok", logs)
    Exit Function
EH:
    Test_ExpedienteAltaHelper_EstablecerDatos_HappyAMOrdinalHidden = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 2. EstablecerDatos — happy EXPIndep (ordinal hidden)
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteAltaHelper_EstablecerDatos_HappyIndepOrdinalHidden() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim stub As Object
    Set stub = BuildStubDTO(BuildStubExpediente("", "", ""))

    Dim p_Error As String
    Dim json As String
    json = modExpedienteAltaHelper.ExpedienteAlta_EstablecerDatos( _
        stub, 5, "", p_Error)  ' EXP_INDEP=5
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If CBool(parsed("payload")("ordinalVisible")) <> False Then _
        Err.Raise 1002, , "expected ordinalVisible=false for Indep"

    Test_ExpedienteAltaHelper_EstablecerDatos_HappyIndepOrdinalHidden = BuildOk("ok", logs)
    Exit Function
EH:
    Test_ExpedienteAltaHelper_EstablecerDatos_HappyIndepOrdinalHidden = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 3. EstablecerDatos — happy Lote sin padre (ordinal hidden)
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteAltaHelper_EstablecerDatos_HappyLoteSinPadre() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim stub As Object
    Set stub = BuildStubDTO(BuildStubExpediente("", "", ""))

    Dim p_Error As String
    Dim json As String
    json = modExpedienteAltaHelper.ExpedienteAlta_EstablecerDatos( _
        stub, 2, "", p_Error)  ' LOTE=2, no padre
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If CBool(parsed("payload")("ordinalVisible")) <> False Then _
        Err.Raise 1002, , "expected ordinalVisible=false for Lote sin padre"

    Test_ExpedienteAltaHelper_EstablecerDatos_HappyLoteSinPadre = BuildOk("ok", logs)
    Exit Function
EH:
    Test_ExpedienteAltaHelper_EstablecerDatos_HappyLoteSinPadre = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 4. EstablecerDatos — happy Lote con padre (ordinal visible, caption "Nº DE LOTE")
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteAltaHelper_EstablecerDatos_HappyLoteConPadre() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim stub As Object
    Set stub = BuildStubDTO(BuildStubExpediente("", "3", "1"))

    Dim p_Error As String
    Dim json As String
    json = modExpedienteAltaHelper.ExpedienteAlta_EstablecerDatos( _
        stub, 2, "PADRE-AM-1", p_Error)  ' LOTE=2, con padre
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If CBool(parsed("payload")("ordinalVisible")) <> True Then _
        Err.Raise 1002, , "expected ordinalVisible=true"
    If CStr(parsed("payload")("ordinalCaption")) <> "Nº DE LOTE" Then _
        Err.Raise 1002, , "expected ordinalCaption='Nº DE LOTE'"
    If CStr(parsed("payload")("ordinalValue")) <> "3" Then _
        Err.Raise 1002, , "expected ordinalValue='3' (stored non-calc)"

    Test_ExpedienteAltaHelper_EstablecerDatos_HappyLoteConPadre = BuildOk("ok", logs)
    Exit Function
EH:
    Test_ExpedienteAltaHelper_EstablecerDatos_HappyLoteConPadre = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 5. EstablecerDatos — happy Basado con padre (caption "Nº BASADO")
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteAltaHelper_EstablecerDatos_HappyBasadoConPadre() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim stub As Object
    Set stub = BuildStubDTO(BuildStubExpediente("", "", ""))

    Dim p_Error As String
    Dim json As String
    json = modExpedienteAltaHelper.ExpedienteAlta_EstablecerDatos( _
        stub, 3, "PADRE-AM-1", p_Error)  ' BASADO_DE_AM=3, con padre
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If CStr(parsed("payload")("ordinalCaption")) <> "Nº BASADO" Then _
        Err.Raise 1002, , "expected ordinalCaption='Nº BASADO'"

    Test_ExpedienteAltaHelper_EstablecerDatos_HappyBasadoConPadre = BuildOk("ok", logs)
    Exit Function
EH:
    Test_ExpedienteAltaHelper_EstablecerDatos_HappyBasadoConPadre = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 6. EstablecerDatos — sad (DTO Nothing)
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteAltaHelper_EstablecerDatos_SadNothing() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modExpedienteAltaHelper.ExpedienteAlta_EstablecerDatos( _
        Nothing, 1, "", p_Error)
    If InStr(json, """ok"":false") = 0 Then Err.Raise 1001, , "expected ok=false"

    Test_ExpedienteAltaHelper_EstablecerDatos_SadNothing = BuildOk("rejected", logs)
    Exit Function
EH:
    Test_ExpedienteAltaHelper_EstablecerDatos_SadNothing = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 7. EstablecerDatosConTipo — happy AM
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteAltaHelper_EstablecerDatosConTipo_HappyAM() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim stub As Object
    Set stub = BuildStubDTO(BuildStubExpediente("", "", ""))

    Dim p_Error As String
    Dim json As String
    json = modExpedienteAltaHelper.ExpedienteAlta_EstablecerDatosConTipo( _
        stub, "Acuerdo Marco", p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If CStr(parsed("payload")("esAM")) <> "Sí" Then _
        Err.Raise 1002, , "expected esAM=Sí"
    If CStr(parsed("payload")("esLote")) <> "No" Then _
        Err.Raise 1002, , "expected esLote=No"
    If CStr(parsed("payload")("esExpediente")) <> "No" Then _
        Err.Raise 1002, , "expected esExpediente=No"

    Test_ExpedienteAltaHelper_EstablecerDatosConTipo_HappyAM = BuildOk("ok", logs)
    Exit Function
EH:
    Test_ExpedienteAltaHelper_EstablecerDatosConTipo_HappyAM = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 8. EstablecerDatosConTipo — happy Lote sin Acuerdo Marco
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteAltaHelper_EstablecerDatosConTipo_HappyLoteSAM() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim stub As Object
    Set stub = BuildStubDTO(BuildStubExpediente("", "", ""))

    Dim p_Error As String
    Dim json As String
    json = modExpedienteAltaHelper.ExpedienteAlta_EstablecerDatosConTipo( _
        stub, "Lote sin Acuero Marco", p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If CStr(parsed("payload")("esLote")) <> "Sí" Then _
        Err.Raise 1002, , "expected esLote=Sí"

    Test_ExpedienteAltaHelper_EstablecerDatosConTipo_HappyLoteSAM = BuildOk("ok", logs)
    Exit Function
EH:
    Test_ExpedienteAltaHelper_EstablecerDatosConTipo_HappyLoteSAM = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 9. EstablecerDatosConTipo — happy Expediente individual
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteAltaHelper_EstablecerDatosConTipo_HappyIndep() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim stub As Object
    Set stub = BuildStubDTO(BuildStubExpediente("", "", ""))

    Dim p_Error As String
    Dim json As String
    json = modExpedienteAltaHelper.ExpedienteAlta_EstablecerDatosConTipo( _
        stub, "Expediente individual", p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If CStr(parsed("payload")("esExpediente")) <> "Sí" Then _
        Err.Raise 1002, , "expected esExpediente=Sí"

    Test_ExpedienteAltaHelper_EstablecerDatosConTipo_HappyIndep = BuildOk("ok", logs)
    Exit Function
EH:
    Test_ExpedienteAltaHelper_EstablecerDatosConTipo_HappyIndep = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 10. EstablecerDatosConTipo — edge (empty tipo -> empty flags)
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteAltaHelper_EstablecerDatosConTipo_EdgeEmpty() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim stub As Object
    Set stub = BuildStubDTO(BuildStubExpediente("", "", ""))

    Dim p_Error As String
    Dim json As String
    json = modExpedienteAltaHelper.ExpedienteAlta_EstablecerDatosConTipo( _
        stub, "", p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If CStr(parsed("payload")("esAM")) <> "" Then _
        Err.Raise 1002, , "expected esAM=''"

    Test_ExpedienteAltaHelper_EstablecerDatosConTipo_EdgeEmpty = BuildOk("ok", logs)
    Exit Function
EH:
    Test_ExpedienteAltaHelper_EstablecerDatosConTipo_EdgeEmpty = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 11. EstablecerDatosConTipo — sad (unknown tipo)
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteAltaHelper_EstablecerDatosConTipo_SadUnknown() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim stub As Object
    Set stub = BuildStubDTO(BuildStubExpediente("", "", ""))

    Dim p_Error As String
    Dim json As String
    json = modExpedienteAltaHelper.ExpedienteAlta_EstablecerDatosConTipo( _
        stub, "Tipo Inexistente", p_Error)
    If InStr(json, """ok"":false") = 0 Then Err.Raise 1001, , "expected ok=false"

    Test_ExpedienteAltaHelper_EstablecerDatosConTipo_SadUnknown = BuildOk("rejected", logs)
    Exit Function
EH:
    Test_ExpedienteAltaHelper_EstablecerDatosConTipo_SadUnknown = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 12. Ambito_AfterUpdate — HPS -> Sí
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteAltaHelper_Ambito_AfterUpdate_HPS() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modExpedienteAltaHelper.ExpedienteAlta_Ambito_AfterUpdate("HPS", p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If CStr(parsed("payload")("hpsAplica")) <> "Sí" Then _
        Err.Raise 1002, , "expected hpsAplica=Sí"

    Test_ExpedienteAltaHelper_Ambito_AfterUpdate_HPS = BuildOk("ok", logs)
    Exit Function
EH:
    Test_ExpedienteAltaHelper_Ambito_AfterUpdate_HPS = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 13. Ambito_AfterUpdate — non-HPS -> empty
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteAltaHelper_Ambito_AfterUpdate_NonHPS() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modExpedienteAltaHelper.ExpedienteAlta_Ambito_AfterUpdate("Otro", p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If CStr(parsed("payload")("hpsAplica")) <> "" Then _
        Err.Raise 1002, , "expected hpsAplica=''"

    Test_ExpedienteAltaHelper_Ambito_AfterUpdate_NonHPS = BuildOk("ok", logs)
    Exit Function
EH:
    Test_ExpedienteAltaHelper_Ambito_AfterUpdate_NonHPS = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 14. Ordinal_AfterUpdate — happy no conflict
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteAltaHelper_Ordinal_AfterUpdate_NoConflict() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modExpedienteAltaHelper.ExpedienteAlta_Ordinal_AfterUpdate( _
        "5", "PADRE-AM-1", Nothing, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If CBool(parsed("payload")("showWarning")) <> False Then _
        Err.Raise 1002, , "expected showWarning=false"

    Test_ExpedienteAltaHelper_Ordinal_AfterUpdate_NoConflict = BuildOk("ok", logs)
    Exit Function
EH:
    Test_ExpedienteAltaHelper_Ordinal_AfterUpdate_NoConflict = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 15. Ordinal_AfterUpdate — happy conflict (CodExp available)
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteAltaHelper_Ordinal_AfterUpdate_ConflictCodExp() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim stubConf As Object
    Set stubConf = CreateObject("Scripting.Dictionary")
    stubConf("CodExp") = "EXP-2024-001"
    stubConf("Nemotecnico") = "EXP-NEMO"
    stubConf("Titulo") = "Expediente Test"

    Dim p_Error As String
    Dim json As String
    json = modExpedienteAltaHelper.ExpedienteAlta_Ordinal_AfterUpdate( _
        "3", "PADRE-AM-1", stubConf, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If CBool(parsed("payload")("showWarning")) <> True Then _
        Err.Raise 1002, , "expected showWarning=true"
    If InStr(CStr(parsed("payload")("conflictoTexto")), "EXP-2024-001") = 0 Then _
        Err.Raise 1002, , "expected conflictoTexto to contain CodExp"

    Test_ExpedienteAltaHelper_Ordinal_AfterUpdate_ConflictCodExp = BuildOk("ok", logs)
    Exit Function
EH:
    Test_ExpedienteAltaHelper_Ordinal_AfterUpdate_ConflictCodExp = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 16. Ordinal_AfterUpdate — conflict falls back to Nemotecnico
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteAltaHelper_Ordinal_AfterUpdate_ConflictNemotecnico() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim stubConf As Object
    Set stubConf = CreateObject("Scripting.Dictionary")
    stubConf("CodExp") = ""
    stubConf("Nemotecnico") = "EXP-NEMO-2"
    stubConf("Titulo") = "Expediente Test"

    Dim p_Error As String
    Dim json As String
    json = modExpedienteAltaHelper.ExpedienteAlta_Ordinal_AfterUpdate( _
        "3", "PADRE-AM-1", stubConf, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If InStr(CStr(parsed("payload")("conflictoTexto")), "EXP-NEMO-2") = 0 Then _
        Err.Raise 1002, , "expected fallback to Nemotecnico"

    Test_ExpedienteAltaHelper_Ordinal_AfterUpdate_ConflictNemotecnico = BuildOk("ok", logs)
    Exit Function
EH:
    Test_ExpedienteAltaHelper_Ordinal_AfterUpdate_ConflictNemotecnico = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 17. Ordinal_AfterUpdate — conflict falls back to Titulo
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteAltaHelper_Ordinal_AfterUpdate_ConflictTitulo() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim stubConf As Object
    Set stubConf = CreateObject("Scripting.Dictionary")
    stubConf("CodExp") = ""
    stubConf("Nemotecnico") = ""
    stubConf("Titulo") = "Mi Expediente de Prueba"

    Dim p_Error As String
    Dim json As String
    json = modExpedienteAltaHelper.ExpedienteAlta_Ordinal_AfterUpdate( _
        "3", "PADRE-AM-1", stubConf, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If InStr(CStr(parsed("payload")("conflictoTexto")), "Mi Expediente de Prueba") = 0 Then _
        Err.Raise 1002, , "expected fallback to Titulo"

    Test_ExpedienteAltaHelper_Ordinal_AfterUpdate_ConflictTitulo = BuildOk("ok", logs)
    Exit Function
EH:
    Test_ExpedienteAltaHelper_Ordinal_AfterUpdate_ConflictTitulo = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 18. Ordinal_AfterUpdate — edge (empty ordinal -> no validation)
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteAltaHelper_Ordinal_AfterUpdate_EdgeEmpty() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modExpedienteAltaHelper.ExpedienteAlta_Ordinal_AfterUpdate( _
        "", "PADRE-AM-1", Nothing, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If CBool(parsed("payload")("showWarning")) <> False Then _
        Err.Raise 1002, , "expected showWarning=false for empty"

    Test_ExpedienteAltaHelper_Ordinal_AfterUpdate_EdgeEmpty = BuildOk("ok", logs)
    Exit Function
EH:
    Test_ExpedienteAltaHelper_Ordinal_AfterUpdate_EdgeEmpty = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 19. Ordinal_AfterUpdate — edge (no padre -> no validation)
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteAltaHelper_Ordinal_AfterUpdate_EdgeNoPadre() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modExpedienteAltaHelper.ExpedienteAlta_Ordinal_AfterUpdate( _
        "3", "", Nothing, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If CBool(parsed("payload")("showWarning")) <> False Then _
        Err.Raise 1002, , "expected showWarning=false for no padre"

    Test_ExpedienteAltaHelper_Ordinal_AfterUpdate_EdgeNoPadre = BuildOk("ok", logs)
    Exit Function
EH:
    Test_ExpedienteAltaHelper_Ordinal_AfterUpdate_EdgeNoPadre = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 20. NotificarAltaTipo — happy (open)
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteAltaHelper_NotificarAltaTipo_HappyOpen() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modExpedienteAltaHelper.ExpedienteAlta_NotificarAltaTipo(True, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If CBool(parsed("payload")("notified")) <> True Then _
        Err.Raise 1002, , "expected notified=true"

    Test_ExpedienteAltaHelper_NotificarAltaTipo_HappyOpen = BuildOk("ok", logs)
    Exit Function
EH:
    Test_ExpedienteAltaHelper_NotificarAltaTipo_HappyOpen = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 21. NotificarAltaTipo — edge (closed)
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteAltaHelper_NotificarAltaTipo_EdgeClosed() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modExpedienteAltaHelper.ExpedienteAlta_NotificarAltaTipo(False, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If CBool(parsed("payload")("notified")) <> False Then _
        Err.Raise 1002, , "expected notified=false"

    Test_ExpedienteAltaHelper_NotificarAltaTipo_EdgeClosed = BuildOk("ok", logs)
    Exit Function
EH:
    Test_ExpedienteAltaHelper_NotificarAltaTipo_EdgeClosed = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 22. RunAll — wrapper for Dysflow manifest discovery
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteAltaHelper_RunAll() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim atoms As Variant
    atoms = Array( _
        "Test_ExpedienteAltaHelper_EstablecerDatos_HappyAMOrdinalHidden", _
        "Test_ExpedienteAltaHelper_EstablecerDatos_HappyIndepOrdinalHidden", _
        "Test_ExpedienteAltaHelper_EstablecerDatos_HappyLoteSinPadre", _
        "Test_ExpedienteAltaHelper_EstablecerDatos_HappyLoteConPadre", _
        "Test_ExpedienteAltaHelper_EstablecerDatos_HappyBasadoConPadre", _
        "Test_ExpedienteAltaHelper_EstablecerDatos_SadNothing", _
        "Test_ExpedienteAltaHelper_EstablecerDatosConTipo_HappyAM", _
        "Test_ExpedienteAltaHelper_EstablecerDatosConTipo_HappyLoteSAM", _
        "Test_ExpedienteAltaHelper_EstablecerDatosConTipo_HappyIndep", _
        "Test_ExpedienteAltaHelper_EstablecerDatosConTipo_EdgeEmpty", _
        "Test_ExpedienteAltaHelper_EstablecerDatosConTipo_SadUnknown", _
        "Test_ExpedienteAltaHelper_Ambito_AfterUpdate_HPS", _
        "Test_ExpedienteAltaHelper_Ambito_AfterUpdate_NonHPS", _
        "Test_ExpedienteAltaHelper_Ordinal_AfterUpdate_NoConflict", _
        "Test_ExpedienteAltaHelper_Ordinal_AfterUpdate_ConflictCodExp", _
        "Test_ExpedienteAltaHelper_Ordinal_AfterUpdate_ConflictNemotecnico", _
        "Test_ExpedienteAltaHelper_Ordinal_AfterUpdate_ConflictTitulo", _
        "Test_ExpedienteAltaHelper_Ordinal_AfterUpdate_EdgeEmpty", _
        "Test_ExpedienteAltaHelper_Ordinal_AfterUpdate_EdgeNoPadre", _
        "Test_ExpedienteAltaHelper_NotificarAltaTipo_HappyOpen", _
        "Test_ExpedienteAltaHelper_NotificarAltaTipo_EdgeClosed" _
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
        Test_ExpedienteAltaHelper_RunAll = BuildFail("firstFailure=" & firstFailure, logs)
    Else
        Test_ExpedienteAltaHelper_RunAll = BuildOk("all-passed", logs)
    End If
    Exit Function
EH:
    Dim p_Error As String
    p_Error = "Test_ExpedienteAltaHelper_RunAll EH: " & Err.Description
    Test_ExpedienteAltaHelper_RunAll = BuildFail(p_Error, logs)
End Function