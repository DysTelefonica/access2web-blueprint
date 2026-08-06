Attribute VB_Name = "Test_ExpedienteAltaParaHPSHelper"
Option Compare Database
Option Explicit

' Test_ExpedienteAltaParaHPSHelper — REWORK (2026-06-26, Phase 3.2b / PR-8b)
' Pure-data TDD atoms for modExpedienteAltaParaHPSHelper.bas
' (Form_FormExpedienteAltaParaHPS, see docs/audit/form-expediente-alta-para-hps-pure-data.md).
'
' Anti-pattern removed:
'   - tests assume UI-driven Rellenar/Establecer — they don't.
'   - No DoCmd.OpenForm / Forms(...) / Screen.ActiveForm references.
'   - Stubs are Scripting.Dictionary instances.
'
' Implements access-vba-tdd skill:
'   - §1.8 declaration ordering: all Private Const/Function/Sub at top, Public atoms after.
'   - §1.10 atom signature MUST match helper signature EXACTLY.
'   - §4.2 no-humo: atoms assert concrete values (counts, fields, JSON), not "did not crash".
'   - §4.4 strong assertions: count=0 vs >0, payload keys present, etc.
'   - §5.1 fixture IDs in test range (>= 900000).
'
' Convention:
'   - Atoms use ONLY Global public names.
'   - Each atom returns JSON: {"ok":true|false,"value":...,"payload":null,"error":...,"logs":[...]}.
'   - Atoms never call MsgBox or pop up UI.

' === Module-level constants (all at top per vba-access §10.1) ====================

' Fixture ID base for test entities.
Private Const TEST_BASE_ID As Long = 900810

' Field separator used in rowSource assertions. Mirrors the helper's constant.
Private Const TEST_FIELD_SEP As String = ";"


' === Local helpers (all at top per vba-access §10.1) ============================

Private Function BuildOk(ByVal p_Value As Variant, ByRef p_Logs() As String) As String
    BuildOk = TestingCore_BuildOk(p_Value, p_Logs)
End Function

Private Function BuildFail(ByVal p_Error As String, ByRef p_Logs() As String) As String
    BuildFail = TestingCore_BuildFail(p_Error, p_Logs)
End Function

' --- BuildStubExpediente ---------------------------------------------------------
' Builds a Scripting.Dictionary mimicking the Expediente class shape.
' Used as p_DTO.Expediente and for direct property reads.
Private Function BuildStubExpediente() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("Ambito") = ""
    d("HPSAplica") = ""
    d("POSTAGEDO") = ""
    d("TIpo") = ""
    d("IDExpediente") = ""
    Set BuildStubExpediente = d
End Function

' --- BuildStubDTO ----------------------------------------------------------------
' Builds a Scripting.Dictionary wrapping Expediente + ColArbolSuministradores
' + ColLugaresEjecucion. Mirrors the form's Dictionary wrapper.
Private Function BuildStubDTO( _
    ByVal p_Expediente As Object, _
    ByVal p_ColArbol As Object, _
    ByVal p_ColLugares As Object _
) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Set d("Expediente") = p_Expediente
    Set d("ColArbolSuministradores") = p_ColArbol
    Set d("ColLugaresEjecucion") = p_ColLugares
    Set BuildStubDTO = d
End Function

' --- BuildStubColArbol -----------------------------------------------------------
' Builds a ColArbolSuministradores Dictionary. Each value is a Dictionary
' with IDSuministrador, ContratistaPrincipal, SubContratista, IdPadre.
Private Function BuildStubColArbol() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Set BuildStubColArbol = d
End Function

' --- BuildStubColLugares ---------------------------------------------------------
Private Function BuildStubColLugares() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Set BuildStubColLugares = d
End Function

' --- BuildStubItemArbol ----------------------------------------------------------
' Builds one ExpedienteSuministrador-like Dictionary with the canonical shape.
Private Function BuildStubItemArbol( _
    ByVal p_IDExpSum As String, _
    ByVal p_IDSuministrador As String, _
    ByVal p_ContratistaPrincipal As String, _
    ByVal p_SubContratista As String, _
    ByVal p_IdPadre As String _
) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("IDExpedienteSuministrador") = p_IDExpSum
    d("IDSuministrador") = p_IDSuministrador
    d("IdPadre") = p_IdPadre
    d("Tag") = p_IDExpSum
    d("ContratistaPrincipal") = p_ContratistaPrincipal
    d("SubContratista") = p_SubContratista
    d("Descripcon") = ""
    Set BuildStubItemArbol = d
End Function

' --- BuildStubLugar --------------------------------------------------------------
Private Function BuildStubLugar( _
    ByVal p_IDLugar As String, _
    ByVal p_Lugar As String _
) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("IDLugarEjecucion") = p_IDLugar
    d("LugarEjecucion") = p_Lugar
    Set BuildStubLugar = d
End Function


' === Public atoms ===============================================================

' ---------------------------------------------------------------------------
' 1. Form_Load — happy (DTO present, ColArbol absent -> initialized)
'    Expected: colArbolInitialized=true, allowEdits=true
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteAltaParaHPSHelper_Form_Load_HappyInicializaColArbol() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim stubExp As Object
    Set stubExp = BuildStubExpediente()
    Dim stubDTO As Object
    Set stubDTO = BuildStubDTO(stubExp, Nothing, Nothing)

    Dim p_Error As String
    Dim json As String
    json = modExpedienteAltaParaHPSHelper.ExpedienteHPS_Form_Load(stubDTO, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error
    If InStr(json, """ok"":true") = 0 Then Err.Raise 1001, , "expected ok=true, got " & json

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If CBool(parsed("payload")("colArbolInitialized")) <> True Then _
        Err.Raise 1002, , "expected colArbolInitialized=true"
    If CBool(parsed("payload")("allowEdits")) <> True Then _
        Err.Raise 1002, , "expected allowEdits=true"

    Test_ExpedienteAltaParaHPSHelper_Form_Load_HappyInicializaColArbol = BuildOk("ok", logs)
    Exit Function
EH:
    Test_ExpedienteAltaParaHPSHelper_Form_Load_HappyInicializaColArbol = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 2. Form_Load — sad (DTO Nothing)
'    Expected: ok=false, p_Error set
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteAltaParaHPSHelper_Form_Load_SadDTONothing() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modExpedienteAltaParaHPSHelper.ExpedienteHPS_Form_Load(Nothing, p_Error)
    If InStr(json, """ok"":false") = 0 Then _
        Err.Raise 1001, , "expected ok=false"
    If p_Error = "" Then Err.Raise 1002, , "expected p_Error set"

    Test_ExpedienteAltaParaHPSHelper_Form_Load_SadDTONothing = BuildOk("dto-rejected", logs)
    Exit Function
EH:
    Test_ExpedienteAltaParaHPSHelper_Form_Load_SadDTONothing = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 3. Form_Load — edge (DTO has ColArbol already set -> not reinitialized)
'    Expected: colArbolInitialized=false
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteAltaParaHPSHelper_Form_Load_EdgeColArbolYaExiste() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim stubExp As Object
    Set stubExp = BuildStubExpediente()
    Dim stubCol As Object
    Set stubCol = BuildStubColArbol()
    ' Add a sentinel entry so the helper sees it as existing.
    Dim sentinel As Object
    Set sentinel = BuildStubItemArbol("TMP_SENTINEL", "S1", "Sí", "No", "")
    stubCol.Add "TMP_SENTINEL", sentinel

    Dim stubDTO As Object
    Set stubDTO = BuildStubDTO(stubExp, stubCol, Nothing)

    Dim p_Error As String
    Dim json As String
    json = modExpedienteAltaParaHPSHelper.ExpedienteHPS_Form_Load(stubDTO, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If CBool(parsed("payload")("colArbolInitialized")) <> False Then _
        Err.Raise 1002, , "expected colArbolInitialized=false (already set)"

    Test_ExpedienteAltaParaHPSHelper_Form_Load_EdgeColArbolYaExiste = BuildOk("ok", logs)
    Exit Function
EH:
    Test_ExpedienteAltaParaHPSHelper_Form_Load_EdgeColArbolYaExiste = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 4. EstablecerCombos_Construir — happy (with Entorno Dictionary)
'    Expected: 4 keys present, all string rowSources (maybe empty)
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteAltaParaHPSHelper_EstablecerCombos_HappyConEntorno() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim stubExp As Object
    Set stubExp = BuildStubExpediente()
    Dim stubDTO As Object
    Set stubDTO = BuildStubDTO(stubExp, Nothing, Nothing)

    ' Build a stub Entorno with all 3 collections.
    Dim stubEntorno As Object
    Set stubEntorno = CreateObject("Scripting.Dictionary")
    Dim clasif As Object
    Set clasif = CreateObject("Scripting.Dictionary")
    Dim itemC As Object
    Set itemC = CreateObject("Scripting.Dictionary")
    itemC("IdGradoClasificacion") = "C1"
    itemC("GradoClasificacion") = "Reservado"
    clasif.Add "C1", itemC
    stubEntorno.Add "GradosClasificaciones", clasif

    Dim sums As Object
    Set sums = CreateObject("Scripting.Dictionary")
    Dim itemS As Object
    Set itemS = CreateObject("Scripting.Dictionary")
    itemS("IDSuministrador") = "S1"
    itemS("Nombre") = "Acme"
    sums.Add "S1", itemS
    stubEntorno.Add "Suministradores", sums

    Dim lugares As Object
    Set lugares = CreateObject("Scripting.Dictionary")
    Dim itemL As Object
    Set itemL = CreateObject("Scripting.Dictionary")
    itemL("IDLugarEjecucion") = "L1"
    itemL("LugarEjecucion") = "Madrid"
    lugares.Add "L1", itemL
    stubEntorno.Add "LugaresEjecucion", lugares

    stubDTO.Add "Entorno", stubEntorno

    Dim p_Error As String
    Dim json As String
    json = modExpedienteAltaParaHPSHelper.ExpedienteHPS_EstablecerCombos_Construir(stubDTO, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    Dim rs As String
    rs = CStr(parsed("payload")("clasificaciones"))
    If InStr(rs, "C1") = 0 Then Err.Raise 1002, , "clasificaciones rowSource missing C1"
    rs = CStr(parsed("payload")("contratistas"))
    If InStr(rs, "S1") = 0 Then Err.Raise 1002, , "contratistas rowSource missing S1"
    rs = CStr(parsed("payload")("subcontratistas"))
    If InStr(rs, "S1") = 0 Then Err.Raise 1002, , "subcontratistas rowSource missing S1"
    rs = CStr(parsed("payload")("lugares"))
    If InStr(rs, "L1") = 0 Then Err.Raise 1002, , "lugares rowSource missing L1"

    Test_ExpedienteAltaParaHPSHelper_EstablecerCombos_HappyConEntorno = BuildOk("ok", logs)
    Exit Function
EH:
    Test_ExpedienteAltaParaHPSHelper_EstablecerCombos_HappyConEntorno = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 5. RellenarLista_Construir — happy CONTRATISTAS (1 item marked "Sí")
'    Expected: count=1, rowSource contains the IDSuministrador
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteAltaParaHPSHelper_RellenarLista_Contratistas_Happy() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim stubCol As Object
    Set stubCol = BuildStubColArbol()
    stubCol.Add "TMP_C1", BuildStubItemArbol("TMP_C1", "S1", "Sí", "No", "")
    stubCol.Add "TMP_S1", BuildStubItemArbol("TMP_S1", "S2", "No", "Sí", "")

    Dim stubDTO As Object
    Set stubDTO = BuildStubDTO(BuildStubExpediente(), stubCol, Nothing)

    Dim p_Error As String
    Dim json As String
    json = modExpedienteAltaParaHPSHelper.ExpedienteHPS_RellenarLista_Construir( _
        stubDTO, "CONTRATISTAS", p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If CLng(parsed("payload")("count")) <> 1 Then _
        Err.Raise 1002, , "expected count=1 (only the Contratista Principal)"
    If InStr(CStr(parsed("payload")("rowSource")), "S1") = 0 Then _
        Err.Raise 1002, , "rowSource should contain S1"

    Test_ExpedienteAltaParaHPSHelper_RellenarLista_Contratistas_Happy = BuildOk(1, logs)
    Exit Function
EH:
    Test_ExpedienteAltaParaHPSHelper_RellenarLista_Contratistas_Happy = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 6. RellenarLista_Construir — happy SUBCONTRATISTAS (1 root, 1 nested)
'    Expected: count=1 (only the root Sub)
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteAltaParaHPSHelper_RellenarLista_Subcontratistas_Happy() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim stubCol As Object
    Set stubCol = BuildStubColArbol()
    stubCol.Add "TMP_ROOT_SUB", BuildStubItemArbol("TMP_ROOT_SUB", "S2", "No", "Sí", "")
    stubCol.Add "TMP_NESTED", BuildStubItemArbol("TMP_NESTED", "S3", "No", "Sí", "S2")

    Dim stubDTO As Object
    Set stubDTO = BuildStubDTO(BuildStubExpediente(), stubCol, Nothing)

    Dim p_Error As String
    Dim json As String
    json = modExpedienteAltaParaHPSHelper.ExpedienteHPS_RellenarLista_Construir( _
        stubDTO, "SUBCONTRATISTAS", p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If CLng(parsed("payload")("count")) <> 1 Then _
        Err.Raise 1002, , "expected count=1 (only the root Sub)"

    Test_ExpedienteAltaParaHPSHelper_RellenarLista_Subcontratistas_Happy = BuildOk(1, logs)
    Exit Function
EH:
    Test_ExpedienteAltaParaHPSHelper_RellenarLista_Subcontratistas_Happy = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 7. RellenarLista_Construir — happy LUGARES (2 items)
'    Expected: count=2
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteAltaParaHPSHelper_RellenarLista_Lugares_Happy() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim stubLug As Object
    Set stubLug = BuildStubColLugares()
    stubLug.Add "L1", BuildStubLugar("L1", "Madrid")
    stubLug.Add "L2", BuildStubLugar("L2", "Barcelona")

    Dim stubDTO As Object
    Set stubDTO = BuildStubDTO(BuildStubExpediente(), Nothing, stubLug)

    Dim p_Error As String
    Dim json As String
    json = modExpedienteAltaParaHPSHelper.ExpedienteHPS_RellenarLista_Construir( _
        stubDTO, "LUGARES", p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If CLng(parsed("payload")("count")) <> 2 Then _
        Err.Raise 1002, , "expected count=2"
    If InStr(CStr(parsed("payload")("rowSource")), "Madrid") = 0 Then _
        Err.Raise 1002, , "rowSource missing Madrid"

    Test_ExpedienteAltaParaHPSHelper_RellenarLista_Lugares_Happy = BuildOk(2, logs)
    Exit Function
EH:
    Test_ExpedienteAltaParaHPSHelper_RellenarLista_Lugares_Happy = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 8. RellenarLista_Construir — sad (invalid tipo)
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteAltaParaHPSHelper_RellenarLista_SadTipoInvalido() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim stubDTO As Object
    Set stubDTO = BuildStubDTO(BuildStubExpediente(), Nothing, Nothing)

    Dim p_Error As String
    Dim json As String
    json = modExpedienteAltaParaHPSHelper.ExpedienteHPS_RellenarLista_Construir( _
        stubDTO, "INVALID", p_Error)
    If InStr(json, """ok"":false") = 0 Then _
        Err.Raise 1001, , "expected ok=false"
    If p_Error = "" Then Err.Raise 1002, , "expected p_Error set"

    Test_ExpedienteAltaParaHPSHelper_RellenarLista_SadTipoInvalido = BuildOk("rejected", logs)
    Exit Function
EH:
    Test_ExpedienteAltaParaHPSHelper_RellenarLista_SadTipoInvalido = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 9. AltaSuministrador — happy CONTRATISTA
'    Expected: added=true, count=1, idExpedienteSuministrador non-empty
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteAltaParaHPSHelper_AltaSuministrador_HappyContratista() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim stubDTO As Object
    Set stubDTO = BuildStubDTO(BuildStubExpediente(), Nothing, Nothing)

    Dim p_Error As String
    Dim json As String
    json = modExpedienteAltaParaHPSHelper.ExpedienteHPS_AltaSuministrador( _
        stubDTO, "S1", "ROOT_CONTR", p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If CBool(parsed("payload")("added")) <> True Then _
        Err.Raise 1002, , "expected added=true"
    If CBool(parsed("payload")("duplicate")) <> False Then _
        Err.Raise 1002, , "expected duplicate=false"
    If Len(CStr(parsed("payload")("idExpedienteSuministrador"))) = 0 Then _
        Err.Raise 1002, , "expected non-empty idExpedienteSuministrador"

    Test_ExpedienteAltaParaHPSHelper_AltaSuministrador_HappyContratista = BuildOk("ok", logs)
    Exit Function
EH:
    Test_ExpedienteAltaParaHPSHelper_AltaSuministrador_HappyContratista = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 10. AltaSuministrador — sad (duplicate detection)
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteAltaParaHPSHelper_AltaSuministrador_SadDuplicado() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim stubCol As Object
    Set stubCol = BuildStubColArbol()
    stubCol.Add "TMP_X", BuildStubItemArbol("TMP_X", "S1", "Sí", "No", "")
    Dim stubDTO As Object
    Set stubDTO = BuildStubDTO(BuildStubExpediente(), stubCol, Nothing)

    Dim p_Error As String
    Dim json As String
    json = modExpedienteAltaParaHPSHelper.ExpedienteHPS_AltaSuministrador( _
        stubDTO, "S1", "ROOT_CONTR", p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If CBool(parsed("payload")("added")) <> False Then _
        Err.Raise 1002, , "expected added=false"
    If CBool(parsed("payload")("duplicate")) <> True Then _
        Err.Raise 1002, , "expected duplicate=true"

    Test_ExpedienteAltaParaHPSHelper_AltaSuministrador_SadDuplicado = BuildOk("dup", logs)
    Exit Function
EH:
    Test_ExpedienteAltaParaHPSHelper_AltaSuministrador_SadDuplicado = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 11. AltaSuministrador — sad (invalid tipo)
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteAltaParaHPSHelper_AltaSuministrador_SadTipoInvalido() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim stubDTO As Object
    Set stubDTO = BuildStubDTO(BuildStubExpediente(), Nothing, Nothing)

    Dim p_Error As String
    Dim json As String
    json = modExpedienteAltaParaHPSHelper.ExpedienteHPS_AltaSuministrador( _
        stubDTO, "S1", "ROOT_X", p_Error)
    If InStr(json, """ok"":false") = 0 Then _
        Err.Raise 1001, , "expected ok=false"
    If p_Error = "" Then Err.Raise 1002, , "expected p_Error set"

    Test_ExpedienteAltaParaHPSHelper_AltaSuministrador_SadTipoInvalido = BuildOk("rejected", logs)
    Exit Function
EH:
    Test_ExpedienteAltaParaHPSHelper_AltaSuministrador_SadTipoInvalido = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 12. AltaSuministrador — sad (empty id)
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteAltaParaHPSHelper_AltaSuministrador_SadEmptyId() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim stubDTO As Object
    Set stubDTO = BuildStubDTO(BuildStubExpediente(), Nothing, Nothing)

    Dim p_Error As String
    Dim json As String
    json = modExpedienteAltaParaHPSHelper.ExpedienteHPS_AltaSuministrador( _
        stubDTO, "", "ROOT_CONTR", p_Error)
    If InStr(json, """ok"":false") = 0 Then _
        Err.Raise 1001, , "expected ok=false"

    Test_ExpedienteAltaParaHPSHelper_AltaSuministrador_SadEmptyId = BuildOk("rejected", logs)
    Exit Function
EH:
    Test_ExpedienteAltaParaHPSHelper_AltaSuministrador_SadEmptyId = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 13. AltaSuministrador — sad (DTO Nothing)
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteAltaParaHPSHelper_AltaSuministrador_SadDTONothing() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modExpedienteAltaParaHPSHelper.ExpedienteHPS_AltaSuministrador( _
        Nothing, "S1", "ROOT_CONTR", p_Error)
    If InStr(json, """ok"":false") = 0 Then _
        Err.Raise 1001, , "expected ok=false"

    Test_ExpedienteAltaParaHPSHelper_AltaSuministrador_SadDTONothing = BuildOk("rejected", logs)
    Exit Function
EH:
    Test_ExpedienteAltaParaHPSHelper_AltaSuministrador_SadDTONothing = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 14. EliminarSuministrador — happy (1 item removed, countAfter=0)
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteAltaParaHPSHelper_EliminarSuministrador_Happy() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim stubCol As Object
    Set stubCol = BuildStubColArbol()
    stubCol.Add "TMP_X", BuildStubItemArbol("TMP_X", "S1", "Sí", "No", "")
    Dim stubDTO As Object
    Set stubDTO = BuildStubDTO(BuildStubExpediente(), stubCol, Nothing)

    Dim p_Error As String
    Dim json As String
    json = modExpedienteAltaParaHPSHelper.ExpedienteHPS_EliminarSuministrador( _
        stubDTO, "S1", "ROOT_CONTR", p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If CBool(parsed("payload")("removed")) <> True Then _
        Err.Raise 1002, , "expected removed=true"
    If CLng(parsed("payload")("countAfter")) <> 0 Then _
        Err.Raise 1002, , "expected countAfter=0"

    Test_ExpedienteAltaParaHPSHelper_EliminarSuministrador_Happy = BuildOk("ok", logs)
    Exit Function
EH:
    Test_ExpedienteAltaParaHPSHelper_EliminarSuministrador_Happy = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 15. EliminarSuministrador — edge (not present -> removed=false, countAfter=1)
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteAltaParaHPSHelper_EliminarSuministrador_EdgeNoExiste() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim stubCol As Object
    Set stubCol = BuildStubColArbol()
    stubCol.Add "TMP_X", BuildStubItemArbol("TMP_X", "S1", "Sí", "No", "")
    Dim stubDTO As Object
    Set stubDTO = BuildStubDTO(BuildStubExpediente(), stubCol, Nothing)

    Dim p_Error As String
    Dim json As String
    json = modExpedienteAltaParaHPSHelper.ExpedienteHPS_EliminarSuministrador( _
        stubDTO, "S-DOES-NOT-EXIST", "ROOT_CONTR", p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If CBool(parsed("payload")("removed")) <> False Then _
        Err.Raise 1002, , "expected removed=false"
    If CLng(parsed("payload")("countAfter")) <> 1 Then _
        Err.Raise 1002, , "expected countAfter=1"

    Test_ExpedienteAltaParaHPSHelper_EliminarSuministrador_EdgeNoExiste = BuildOk("noop", logs)
    Exit Function
EH:
    Test_ExpedienteAltaParaHPSHelper_EliminarSuministrador_EdgeNoExiste = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 16. AltaLugar — happy
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteAltaParaHPSHelper_AltaLugar_Happy() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim stubDTO As Object
    Set stubDTO = BuildStubDTO(BuildStubExpediente(), Nothing, Nothing)

    Dim p_Error As String
    Dim json As String
    json = modExpedienteAltaParaHPSHelper.ExpedienteHPS_AltaLugar(stubDTO, "L1", p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If CBool(parsed("payload")("added")) <> True Then _
        Err.Raise 1002, , "expected added=true"

    Test_ExpedienteAltaParaHPSHelper_AltaLugar_Happy = BuildOk("ok", logs)
    Exit Function
EH:
    Test_ExpedienteAltaParaHPSHelper_AltaLugar_Happy = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 17. AltaLugar — sad (duplicate)
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteAltaParaHPSHelper_AltaLugar_SadDuplicado() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim stubLug As Object
    Set stubLug = BuildStubColLugares()
    stubLug.Add "L1", BuildStubLugar("L1", "Madrid")
    Dim stubDTO As Object
    Set stubDTO = BuildStubDTO(BuildStubExpediente(), Nothing, stubLug)

    Dim p_Error As String
    Dim json As String
    json = modExpedienteAltaParaHPSHelper.ExpedienteHPS_AltaLugar(stubDTO, "L1", p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If CBool(parsed("payload")("added")) <> False Then _
        Err.Raise 1002, , "expected added=false"
    If CBool(parsed("payload")("duplicate")) <> True Then _
        Err.Raise 1002, , "expected duplicate=true"

    Test_ExpedienteAltaParaHPSHelper_AltaLugar_SadDuplicado = BuildOk("dup", logs)
    Exit Function
EH:
    Test_ExpedienteAltaParaHPSHelper_AltaLugar_SadDuplicado = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 18. AltaLugar — sad (empty id)
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteAltaParaHPSHelper_AltaLugar_SadEmptyId() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim stubDTO As Object
    Set stubDTO = BuildStubDTO(BuildStubExpediente(), Nothing, Nothing)

    Dim p_Error As String
    Dim json As String
    json = modExpedienteAltaParaHPSHelper.ExpedienteHPS_AltaLugar(stubDTO, "", p_Error)
    If InStr(json, """ok"":false") = 0 Then _
        Err.Raise 1001, , "expected ok=false"

    Test_ExpedienteAltaParaHPSHelper_AltaLugar_SadEmptyId = BuildOk("rejected", logs)
    Exit Function
EH:
    Test_ExpedienteAltaParaHPSHelper_AltaLugar_SadEmptyId = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 19. EliminarLugar — happy
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteAltaParaHPSHelper_EliminarLugar_Happy() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim stubLug As Object
    Set stubLug = BuildStubColLugares()
    stubLug.Add "L1", BuildStubLugar("L1", "Madrid")
    Dim stubDTO As Object
    Set stubDTO = BuildStubDTO(BuildStubExpediente(), Nothing, stubLug)

    Dim p_Error As String
    Dim json As String
    json = modExpedienteAltaParaHPSHelper.ExpedienteHPS_EliminarLugar(stubDTO, "L1", p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If CBool(parsed("payload")("removed")) <> True Then _
        Err.Raise 1002, , "expected removed=true"

    Test_ExpedienteAltaParaHPSHelper_EliminarLugar_Happy = BuildOk("ok", logs)
    Exit Function
EH:
    Test_ExpedienteAltaParaHPSHelper_EliminarLugar_Happy = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 20. EliminarLugar — edge (not present)
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteAltaParaHPSHelper_EliminarLugar_EdgeNoExiste() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim stubDTO As Object
    Set stubDTO = BuildStubDTO(BuildStubExpediente(), Nothing, Nothing)

    Dim p_Error As String
    Dim json As String
    json = modExpedienteAltaParaHPSHelper.ExpedienteHPS_EliminarLugar(stubDTO, "L1", p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If CBool(parsed("payload")("removed")) <> False Then _
        Err.Raise 1002, , "expected removed=false"

    Test_ExpedienteAltaParaHPSHelper_EliminarLugar_EdgeNoExiste = BuildOk("noop", logs)
    Exit Function
EH:
    Test_ExpedienteAltaParaHPSHelper_EliminarLugar_EdgeNoExiste = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 21. ComandoRegistrar_Click — happy (DTO + Expediente present)
'    Expected: ok=true, payload has ambito="HPS" and ready=true
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteAltaParaHPSHelper_ComandoRegistrar_Happy() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim stubExp As Object
    Set stubExp = BuildStubExpediente()
    Dim stubDTO As Object
    Set stubDTO = BuildStubDTO(stubExp, Nothing, Nothing)

    Dim p_Error As String
    Dim json As String
    json = modExpedienteAltaParaHPSHelper.ExpedienteHPS_ComandoRegistrar_Click(stubDTO, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If CStr(parsed("payload")("ambito")) <> "HPS" Then _
        Err.Raise 1002, , "expected ambito=HPS"
    If CStr(parsed("payload")("hpsAplica")) <> "Sí" Then _
        Err.Raise 1002, , "expected hpsAplica=Sí"
    If CStr(parsed("payload")("postagedo")) <> "Sí" Then _
        Err.Raise 1002, , "expected postagedo=Sí"
    If InStr(CStr(parsed("payload")("tipo")), "HPS") = 0 Then _
        Err.Raise 1002, , "expected tipo contains HPS"
    If CBool(parsed("payload")("ready")) <> True Then _
        Err.Raise 1002, , "expected ready=true"

    Test_ExpedienteAltaParaHPSHelper_ComandoRegistrar_Happy = BuildOk("ok", logs)
    Exit Function
EH:
    Test_ExpedienteAltaParaHPSHelper_ComandoRegistrar_Happy = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 22. ComandoRegistrar_Click — sad (DTO Nothing)
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteAltaParaHPSHelper_ComandoRegistrar_SadDTONothing() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modExpedienteAltaParaHPSHelper.ExpedienteHPS_ComandoRegistrar_Click(Nothing, p_Error)
    If InStr(json, """ok"":false") = 0 Then _
        Err.Raise 1001, , "expected ok=false"

    Test_ExpedienteAltaParaHPSHelper_ComandoRegistrar_SadDTONothing = BuildOk("rejected", logs)
    Exit Function
EH:
    Test_ExpedienteAltaParaHPSHelper_ComandoRegistrar_SadDTONothing = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 23. ComandoRegistrar_Click — sad (DTO without Expediente)
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteAltaParaHPSHelper_ComandoRegistrar_SadExpedienteNothing() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim stubDTO As Object
    Set stubDTO = BuildStubDTO(Nothing, Nothing, Nothing)

    Dim p_Error As String
    Dim json As String
    json = modExpedienteAltaParaHPSHelper.ExpedienteHPS_ComandoRegistrar_Click(stubDTO, p_Error)
    If InStr(json, """ok"":false") = 0 Then _
        Err.Raise 1001, , "expected ok=false"

    Test_ExpedienteAltaParaHPSHelper_ComandoRegistrar_SadExpedienteNothing = BuildOk("rejected", logs)
    Exit Function
EH:
    Test_ExpedienteAltaParaHPSHelper_ComandoRegistrar_SadExpedienteNothing = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 24. RunAll — wrapper for Dysflow manifest discovery
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteAltaParaHPSHelper_RunAll() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim atoms As Variant
    atoms = Array( _
        "Test_ExpedienteAltaParaHPSHelper_Form_Load_HappyInicializaColArbol", _
        "Test_ExpedienteAltaParaHPSHelper_Form_Load_SadDTONothing", _
        "Test_ExpedienteAltaParaHPSHelper_Form_Load_EdgeColArbolYaExiste", _
        "Test_ExpedienteAltaParaHPSHelper_EstablecerCombos_HappyConEntorno", _
        "Test_ExpedienteAltaParaHPSHelper_RellenarLista_Contratistas_Happy", _
        "Test_ExpedienteAltaParaHPSHelper_RellenarLista_Subcontratistas_Happy", _
        "Test_ExpedienteAltaParaHPSHelper_RellenarLista_Lugares_Happy", _
        "Test_ExpedienteAltaParaHPSHelper_RellenarLista_SadTipoInvalido", _
        "Test_ExpedienteAltaParaHPSHelper_AltaSuministrador_HappyContratista", _
        "Test_ExpedienteAltaParaHPSHelper_AltaSuministrador_SadDuplicado", _
        "Test_ExpedienteAltaParaHPSHelper_AltaSuministrador_SadTipoInvalido", _
        "Test_ExpedienteAltaParaHPSHelper_AltaSuministrador_SadEmptyId", _
        "Test_ExpedienteAltaParaHPSHelper_AltaSuministrador_SadDTONothing", _
        "Test_ExpedienteAltaParaHPSHelper_EliminarSuministrador_Happy", _
        "Test_ExpedienteAltaParaHPSHelper_EliminarSuministrador_EdgeNoExiste", _
        "Test_ExpedienteAltaParaHPSHelper_AltaLugar_Happy", _
        "Test_ExpedienteAltaParaHPSHelper_AltaLugar_SadDuplicado", _
        "Test_ExpedienteAltaParaHPSHelper_AltaLugar_SadEmptyId", _
        "Test_ExpedienteAltaParaHPSHelper_EliminarLugar_Happy", _
        "Test_ExpedienteAltaParaHPSHelper_EliminarLugar_EdgeNoExiste", _
        "Test_ExpedienteAltaParaHPSHelper_ComandoRegistrar_Happy", _
        "Test_ExpedienteAltaParaHPSHelper_ComandoRegistrar_SadDTONothing", _
        "Test_ExpedienteAltaParaHPSHelper_ComandoRegistrar_SadExpedienteNothing" _
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
        Test_ExpedienteAltaParaHPSHelper_RunAll = BuildFail("firstFailure=" & firstFailure, logs)
    Else
        Test_ExpedienteAltaParaHPSHelper_RunAll = BuildOk("all-passed", logs)
    End If
    Exit Function
EH:
    Dim p_Error As String
    p_Error = "Test_ExpedienteAltaParaHPSHelper_RunAll EH: " & Err.Description
    Test_ExpedienteAltaParaHPSHelper_RunAll = BuildFail(p_Error, logs)
End Function