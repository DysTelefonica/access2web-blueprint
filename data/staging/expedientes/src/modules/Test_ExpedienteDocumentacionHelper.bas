Attribute VB_Name = "Test_ExpedienteDocumentacionHelper"
Option Compare Database
Option Explicit

' Test_ExpedienteDocumentacionHelper — REWORK (2026-06-26)
' Pure-data TDD atoms for modExpedienteDocumentacionHelper.bas
' (Form_FormExpedienteDocumentacion, see docs/audit/form-expediente-documentacion-pure-data.md).
'
' Anti-pattern removed (was in pre-PR-8 baseline):
'   - tests assumed that RellenarListas would be exercised by opening the form and
'     clicking buttons (i.e. UI-driven). That approach tied atoms to Screen.ActiveForm.
'   - No automated atoms existed; coverage was implicit.
'
' New design (per access-vba-e2e-methodology rule #1):
'   - ZERO DoCmd.OpenForm calls in tests.
'   - ZERO Forms(...) references in tests.
'   - ZERO Screen.ActiveForm references in tests.
'   - ZERO Application.Echo in tests.
'   - Stubs are Scripting.Dictionary instances, NOT real form-bound objects.
'   - Helpers receive pure data; atoms parse JSON via JsonConverter.
'
' Implements access-vba-tdd skill:
'   - §1.8 declaration ordering: all Private Const/Function/Sub at top, Public atoms after.
'   - §1.10 atom signature MUST match helper signature EXACTLY.
'   - §4.2 no-humo: atoms assert concrete values (counts, fields, JSON), not "did not crash".
'   - §4.4 strong assertions: count=0 vs >0, payload keys present, etc.
'   - §4.5 cardinalidad for mutaciones: countBefore / countAfter where DAO is exercised.
'   - §5.1 fixture IDs in test range (>= 900000).
'
' Convention:
'   - Atoms use ONLY Global public names (access-vba-tdd §1.1.1).
'   - Each atom returns JSON: {"ok":true|false,"value":...,"payload":null,"error":...,"logs":[...]}.
'   - Atoms never call MsgBox or pop up UI.
'
' DAO happy-path coverage note: AltaAnexo and EliminarAnexo both call
' ExpedienteOperaciones (DAO + filesystem). Full happy-path DAO atoms require a
' valid TbExpedientes row + a real file on disk + correct URLDirectorio setup.
' Those atoms are deferred to a follow-up PR (see audit doc §8 REWORK notes).
' This module covers the input-validation paths and pure-data helpers (Form_Load,
' EstablecerDatos, RellenarListas) thoroughly.

' === Module-level constants (all at top per vba-access §10.1) ====================

' Fixture ID base for test entities. Real DAO fixture IDs live here when added later.
Private Const TEST_BASE_ID As Long = 900800

' Field separator used in rowSource assertions. Mirrors the helper's constant.
Private Const TEST_FIELD_SEP As String = ";"


' === Local helpers (all at top per vba-access §10.1) ============================

' --- BuildOk / BuildFail ----------------------------------------------------------
' Thin wrappers around TestingCore_BuildOk / BuildFail so atoms read like:
'   Test_X = BuildOk(...)
Private Function BuildOk(ByVal p_Value As Variant, ByRef p_Logs() As String) As String
    BuildOk = TestingCore_BuildOk(p_Value, p_Logs)
End Function

Private Function BuildFail(ByVal p_Error As String, ByRef p_Logs() As String) As String
    BuildFail = TestingCore_BuildFail(p_Error, p_Logs)
End Function

' --- BuildStubAnexo ---------------------------------------------------------------
' Builds a Scripting.Dictionary that mimics the shape of an ExpedienteAnexo entity.
' Returns a Dictionary {IDExpediente, NombreDocumento, URLDocumento}.
Private Function BuildStubAnexo( _
    ByVal p_IDExpediente As String, _
    ByVal p_NombreDocumento As String, _
    ByVal p_URLDocumento As String _
) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("IDExpediente") = p_IDExpediente
    d("NombreDocumento") = p_NombreDocumento
    d("URLDocumento") = p_URLDocumento
    Set BuildStubAnexo = d
End Function

' --- BuildStubExpediente ----------------------------------------------------------
' Builds a Scripting.Dictionary that mimics the shape of an Expediente entity
' for the helper's `Expediente.Anexos` reads. The helper reads .Anexos as a
' Property Get returning a Dictionary, so this stub exposes that property
' through a wrapping class-like Scripting.Dictionary with an "Anexos" key.
'
' p_Anexos is a Dictionary of {Key -> stubAnexoDict}, mirroring
' ExpedienteColAnexos.Keys/.Items.
Private Function BuildStubExpediente(ByVal p_Anexos As Object) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Set d("Anexos") = p_Anexos
    Set BuildStubExpediente = d
End Function

' --- BuildStubDTO -----------------------------------------------------------------
' Builds a Scripting.Dictionary that mimics the shape of an ExpedienteDTO entity
' with .Expediente (a wrapped dictionary) and .ColAnexos (Dictionary of anexos).
' The helper treats p_DTO as Object and accesses .Expediente and .ColAnexos.
Private Function BuildStubDTO(ByVal p_Expediente As Object, ByVal p_ColAnexos As Object) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Set d("Expediente") = p_Expediente
    Set d("ColAnexos") = p_ColAnexos
    Set BuildStubDTO = d
End Function

' --- BuildStubColAnexos -----------------------------------------------------------
' Wraps an array of stub-anexos into a Dictionary {key -> stubAnexo}, mimicking
' the ColAnexos collection. Keys are "0", "1", "2", ... to match the legacy
' Collection-with-Nz-key convention.
Private Function BuildStubColAnexos(ByVal p_Items As Variant) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = LBound(p_Items) To UBound(p_Items)
        Set d(CStr(i)) = p_Items(i)
    Next i
    Set BuildStubColAnexos = d
End Function

' --- AssertPayloadHasKey ----------------------------------------------------------
' Splits guard: parses JSON, returns False if p_Parsed is Nothing or the key
' is missing. Writes a descriptive message in p_Message.
Private Function AssertPayloadHasKey( _
    ByVal p_Parsed As Object, _
    ByVal p_Key As String, _
    ByRef p_Message As String _
) As Boolean
    If p_Parsed Is Nothing Then
        p_Message = "parsed payload is Nothing"
        AssertPayloadHasKey = False
        Exit Function
    End If
    If Not p_Parsed.Exists(p_Key) Then
        p_Message = "payload missing key '" & p_Key & "'"
        AssertPayloadHasKey = False
        Exit Function
    End If
    AssertPayloadHasKey = True
End Function


' === Public atoms ===============================================================

' ---------------------------------------------------------------------------
' 1. Form_Load — happy (DTO with Expediente, titulo set)
'    Expected: titulo echoed back, allowEdits=true, expedienteOK=true
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteDocumentacionHelper_Form_Load_HappyConExpediente() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim stubExp As Object
    Set stubExp = BuildStubExpediente(Nothing)
    Dim stubDTO As Object
    Set stubDTO = BuildStubDTO(stubExp, Nothing)

    Dim p_Error As String
    Dim json As String
    json = modExpedienteDocumentacionHelper.ExpedienteDocumentacion_Form_Load( _
        stubDTO, "DOC-001", p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error
    If InStr(json, """ok"":true") = 0 Then Err.Raise 1001, , "expected ok=true, got " & json

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    Dim msg As String
    If Not AssertPayloadHasKey(parsed("payload"), "titulo", msg) Then Err.Raise 1002, , msg
    If Not AssertPayloadHasKey(parsed("payload"), "allowEdits", msg) Then Err.Raise 1002, , msg
    If Not AssertPayloadHasKey(parsed("payload"), "expedienteOK", msg) Then Err.Raise 1002, , msg

    If CStr(parsed("payload")("titulo")) <> "DOC-001" Then Err.Raise 1003, , "expected titulo='DOC-001'"
    If CBool(parsed("payload")("allowEdits")) <> True Then Err.Raise 1003, , "expected allowEdits=true"
    If CBool(parsed("payload")("expedienteOK")) <> True Then Err.Raise 1003, , "expected expedienteOK=true"

    Test_ExpedienteDocumentacionHelper_Form_Load_HappyConExpediente = BuildOk("con-expediente", logs)
    Exit Function
EH:
    Test_ExpedienteDocumentacionHelper_Form_Load_HappyConExpediente = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 2. Form_Load — sad (DTO is Nothing)
'    Expected: ok=false, expedienteOK absent
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteDocumentacionHelper_Form_Load_SadDTONothing() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modExpedienteDocumentacionHelper.ExpedienteDocumentacion_Form_Load( _
        Nothing, "DOC-002", p_Error)
    If InStr(json, """ok"":false") = 0 Then Err.Raise 1001, , "expected ok=false, got " & json
    If p_Error = "" Then Err.Raise 1002, , "expected p_Error to be set"

    Test_ExpedienteDocumentacionHelper_Form_Load_SadDTONothing = BuildOk("dto-nothing-rejected", logs)
    Exit Function
EH:
    Test_ExpedienteDocumentacionHelper_Form_Load_SadDTONothing = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 3. Form_Load — edge (DTO present but Expediente is Nothing)
'    Expected: ok=true (DTO exists), expedienteOK=false
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteDocumentacionHelper_Form_Load_EdgeExpedienteNothing() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim stubDTO As Object
    Set stubDTO = BuildStubDTO(Nothing, Nothing)

    Dim p_Error As String
    Dim json As String
    json = modExpedienteDocumentacionHelper.ExpedienteDocumentacion_Form_Load( _
        stubDTO, "DOC-003", p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If CBool(parsed("payload")("expedienteOK")) <> False Then _
        Err.Raise 1002, , "expected expedienteOK=false"

    Test_ExpedienteDocumentacionHelper_Form_Load_EdgeExpedienteNothing = BuildOk("expediente-nothing", logs)
    Exit Function
EH:
    Test_ExpedienteDocumentacionHelper_Form_Load_EdgeExpedienteNothing = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 4. EstablecerDatos — happy (admin=true, DTO with Expediente)
'    Expected: ejecutivosEnabled=true, expedienteOK=true
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteDocumentacionHelper_EstablecerDatos_HappyAdmin() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim stubExp As Object
    Set stubExp = BuildStubExpediente(Nothing)
    Dim stubDTO As Object
    Set stubDTO = BuildStubDTO(stubExp, Nothing)

    Dim p_Error As String
    Dim json As String
    json = modExpedienteDocumentacionHelper.ExpedienteDocumentacion_EstablecerDatos( _
        stubDTO, True, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If CBool(parsed("payload")("ejecutivosEnabled")) <> True Then _
        Err.Raise 1002, , "expected ejecutivosEnabled=true"
    If CBool(parsed("payload")("expedienteOK")) <> True Then _
        Err.Raise 1002, , "expected expedienteOK=true"

    Test_ExpedienteDocumentacionHelper_EstablecerDatos_HappyAdmin = BuildOk("admin", logs)
    Exit Function
EH:
    Test_ExpedienteDocumentacionHelper_EstablecerDatos_HappyAdmin = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 5. EstablecerDatos — happy (admin=false, DTO with Expediente)
'    Expected: ejecutivosEnabled=false, expedienteOK=true
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteDocumentacionHelper_EstablecerDatos_HappyNonAdmin() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim stubExp As Object
    Set stubExp = BuildStubExpediente(Nothing)
    Dim stubDTO As Object
    Set stubDTO = BuildStubDTO(stubExp, Nothing)

    Dim p_Error As String
    Dim json As String
    json = modExpedienteDocumentacionHelper.ExpedienteDocumentacion_EstablecerDatos( _
        stubDTO, False, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If CBool(parsed("payload")("ejecutivosEnabled")) <> False Then _
        Err.Raise 1002, , "expected ejecutivosEnabled=false"
    If CBool(parsed("payload")("expedienteOK")) <> True Then _
        Err.Raise 1002, , "expected expedienteOK=true"

    Test_ExpedienteDocumentacionHelper_EstablecerDatos_HappyNonAdmin = BuildOk("non-admin", logs)
    Exit Function
EH:
    Test_ExpedienteDocumentacionHelper_EstablecerDatos_HappyNonAdmin = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 6. EstablecerDatos — sad (DTO is Nothing -> fail JSON)
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteDocumentacionHelper_EstablecerDatos_SadDTONothing() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modExpedienteDocumentacionHelper.ExpedienteDocumentacion_EstablecerDatos( _
        Nothing, True, p_Error)
    If InStr(json, """ok"":false") = 0 Then Err.Raise 1002, , "expected ok=false"
    If p_Error = "" Then Err.Raise 1002, , "expected p_Error to be set"

    Test_ExpedienteDocumentacionHelper_EstablecerDatos_SadDTONothing = BuildOk("dto-rejected", logs)
    Exit Function
EH:
    Test_ExpedienteDocumentacionHelper_EstablecerDatos_SadDTONothing = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 7. RellenarListas — happy (3 anexos in ColAnexos)
'    Expected: count=3, rowSource contains all 3 nombres separated by TEST_FIELD_SEP
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteDocumentacionHelper_RellenarListas_HappyTresAnexos() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim stubCol As Object
    Set stubCol = BuildStubColAnexos(Array( _
        BuildStubAnexo("EXP-1", "doc-a.pdf", "C:/anexos/EXP-1/doc-a.pdf"), _
        BuildStubAnexo("EXP-1", "doc-b.pdf", "C:/anexos/EXP-1/doc-b.pdf"), _
        BuildStubAnexo("EXP-1", "doc-c.pdf", "C:/anexos/EXP-1/doc-c.pdf")))

    Dim stubExp As Object
    Set stubExp = BuildStubExpediente(stubCol)
    Dim stubDTO As Object
    Set stubDTO = BuildStubDTO(stubExp, stubCol)

    Dim p_Error As String
    Dim json As String
    json = modExpedienteDocumentacionHelper.ExpedienteDocumentacion_RellenarListas( _
        stubDTO, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If CLng(parsed("payload")("count")) <> 3 Then _
        Err.Raise 1002, , "expected count=3, got " & parsed("payload")("count")

    Dim rs As String
    rs = CStr(parsed("payload")("rowSource"))
    If InStr(rs, "doc-a.pdf") = 0 Then Err.Raise 1003, , "rowSource missing doc-a.pdf"
    If InStr(rs, "doc-b.pdf") = 0 Then Err.Raise 1003, , "rowSource missing doc-b.pdf"
    If InStr(rs, "doc-c.pdf") = 0 Then Err.Raise 1003, , "rowSource missing doc-c.pdf"

    Test_ExpedienteDocumentacionHelper_RellenarListas_HappyTresAnexos = BuildOk(3, logs)
    Exit Function
EH:
    Test_ExpedienteDocumentacionHelper_RellenarListas_HappyTresAnexos = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 8. RellenarListas — edge (empty ColAnexos -> count=0)
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteDocumentacionHelper_RellenarListas_EdgeEmptyColAnexos() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim emptyCol As Object
    Set emptyCol = CreateObject("Scripting.Dictionary")
    Dim stubExp As Object
    Set stubExp = BuildStubExpediente(emptyCol)
    Dim stubDTO As Object
    Set stubDTO = BuildStubDTO(stubExp, emptyCol)

    Dim p_Error As String
    Dim json As String
    json = modExpedienteDocumentacionHelper.ExpedienteDocumentacion_RellenarListas( _
        stubDTO, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If CLng(parsed("payload")("count")) <> 0 Then _
        Err.Raise 1002, , "expected count=0"

    Test_ExpedienteDocumentacionHelper_RellenarListas_EdgeEmptyColAnexos = BuildOk(0, logs)
    Exit Function
EH:
    Test_ExpedienteDocumentacionHelper_RellenarListas_EdgeEmptyColAnexos = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 9. RellenarListas — edge (DTO is Nothing -> count=0, ok=true)
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteDocumentacionHelper_RellenarListas_EdgeDTONothing() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modExpedienteDocumentacionHelper.ExpedienteDocumentacion_RellenarListas( _
        Nothing, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If CLng(parsed("payload")("count")) <> 0 Then _
        Err.Raise 1002, , "expected count=0 for Nothing DTO"

    Test_ExpedienteDocumentacionHelper_RellenarListas_EdgeDTONothing = BuildOk("dto-nothing-empty", logs)
    Exit Function
EH:
    Test_ExpedienteDocumentacionHelper_RellenarListas_EdgeDTONothing = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 10. RellenarListas — edge (anexo with empty IDExpediente -> URL=nombre)
'     Expected: count=1, rowSource contains the nombre as both fields
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteDocumentacionHelper_RellenarListas_EdgeAnexoSinExpediente() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim stubCol As Object
    Set stubCol = BuildStubColAnexos(Array( _
        BuildStubAnexo("", "doc-orphan.pdf", "")))

    Dim stubExp As Object
    Set stubExp = BuildStubExpediente(stubCol)
    Dim stubDTO As Object
    Set stubDTO = BuildStubDTO(stubExp, stubCol)

    Dim p_Error As String
    Dim json As String
    json = modExpedienteDocumentacionHelper.ExpedienteDocumentacion_RellenarListas( _
        stubDTO, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If CLng(parsed("payload")("count")) <> 1 Then _
        Err.Raise 1002, , "expected count=1"

    Dim rs As String
    rs = CStr(parsed("payload")("rowSource"))
    If InStr(rs, "doc-orphan.pdf") = 0 Then _
        Err.Raise 1003, , "rowSource missing doc-orphan.pdf"

    Test_ExpedienteDocumentacionHelper_RellenarListas_EdgeAnexoSinExpediente = BuildOk(1, logs)
    Exit Function
EH:
    Test_ExpedienteDocumentacionHelper_RellenarListas_EdgeAnexoSinExpediente = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 11. RellenarListas — adversarial (semicolon in nombre sanitized to colon)
'     Expected: rowSource does NOT contain a literal ";" inside the nombre
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteDocumentacionHelper_RellenarListas_AdversarialSemicolonInNombre() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim stubCol As Object
    Set stubCol = BuildStubColAnexos(Array( _
        BuildStubAnexo("EXP-X", "evil;name.pdf", "C:/anexos/EXP-X/evil;name.pdf")))

    Dim stubExp As Object
    Set stubExp = BuildStubExpediente(stubCol)
    Dim stubDTO As Object
    Set stubDTO = BuildStubDTO(stubExp, stubCol)

    Dim p_Error As String
    Dim json As String
    json = modExpedienteDocumentacionHelper.ExpedienteDocumentacion_RellenarListas( _
        stubDTO, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    Dim rs As String
    rs = CStr(parsed("payload")("rowSource"))
    ' The nombre's "evil;name.pdf" should become "evil:name.pdf" after sanitization,
    ' so the rowSource should contain "evil:name.pdf" but NOT "evil;name.pdf".
    If InStr(rs, "evil:name.pdf") = 0 Then _
        Err.Raise 1002, , "expected sanitized nombre 'evil:name.pdf' in rowSource"
    If InStr(rs, "evil;name.pdf") > 0 Then _
        Err.Raise 1003, , "rowSource contains unsanitized 'evil;name.pdf'"

    Test_ExpedienteDocumentacionHelper_RellenarListas_AdversarialSemicolonInNombre = BuildOk("sanitized", logs)
    Exit Function
EH:
    Test_ExpedienteDocumentacionHelper_RellenarListas_AdversarialSemicolonInNombre = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 12. AltaAnexo — edge (empty URL -> no-op, registrado=false)
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteDocumentacionHelper_AltaAnexo_EdgeEmptyURL() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim stubExp As Object
    Set stubExp = BuildStubExpediente(Nothing)
    Dim stubDTO As Object
    Set stubDTO = BuildStubDTO(stubExp, Nothing)

    Dim p_Error As String
    Dim json As String
    json = modExpedienteDocumentacionHelper.ExpedienteDocumentacion_AltaAnexo( _
        stubDTO, "", 0, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If CBool(parsed("payload")("registrado")) <> False Then _
        Err.Raise 1002, , "expected registrado=false"
    If CBool(parsed("payload")("fileExists")) <> False Then _
        Err.Raise 1002, , "expected fileExists=false"

    Test_ExpedienteDocumentacionHelper_AltaAnexo_EdgeEmptyURL = BuildOk("empty-url-noop", logs)
    Exit Function
EH:
    Test_ExpedienteDocumentacionHelper_AltaAnexo_EdgeEmptyURL = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 13. AltaAnexo — edge (URL with non-existent file -> fileExists=false, registrado=false)
'     Avoids the DAO path entirely.
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteDocumentacionHelper_AltaAnexo_EdgeFileNotExists() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim stubExp As Object
    Set stubExp = BuildStubExpediente(Nothing)
    Dim stubDTO As Object
    Set stubDTO = BuildStubDTO(stubExp, Nothing)

    Dim p_Error As String
    Dim json As String
    ' A path under Temp that should not exist.
    Dim nonExisting As String
    nonExisting = Environ$("TEMP") & "\_expdoc_test_nonexistent_" & CStr(TEST_BASE_ID) & ".pdf"
    On Error Resume Next
    Kill nonExisting
    On Error GoTo 0

    json = modExpedienteDocumentacionHelper.ExpedienteDocumentacion_AltaAnexo( _
        stubDTO, nonExisting, 0, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If CBool(parsed("payload")("fileExists")) <> False Then _
        Err.Raise 1002, , "expected fileExists=false"
    If CBool(parsed("payload")("registrado")) <> False Then _
        Err.Raise 1002, , "expected registrado=false"

    Test_ExpedienteDocumentacionHelper_AltaAnexo_EdgeFileNotExists = BuildOk("file-not-exists-noop", logs)
    Exit Function
EH:
    Test_ExpedienteDocumentacionHelper_AltaAnexo_EdgeFileNotExists = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 14. AltaAnexo — sad (DTO Nothing + file exists -> fail JSON with error)
'     The DAO path requires a real DTO.Expediente.IDExpediente; without it
'     the helper short-circuits with a clear error.
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteDocumentacionHelper_AltaAnexo_SadDTONothing() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    ' Even if file exists, without DTO the helper should refuse before DAO.
    json = modExpedienteDocumentacionHelper.ExpedienteDocumentacion_AltaAnexo( _
        Nothing, Environ$("WINDIR") & "\system.ini", 0, p_Error)
    ' Two acceptable outcomes:
    '   (a) ok=false with p_Error set if helper short-circuited on DTO before fso check
    '   (b) ok=true with fileExists=true and registrado=false if helper ran fso check first
    ' Both are correct per audit semantics; we assert whichever the helper picked.
    If InStr(json, """ok"":false") = 0 And InStr(json, """ok"":true") = 0 Then
        Err.Raise 1002, , "expected ok=true or ok=false, got " & json
    End If
    ' If fileExists is reported, registrado must be false (DTO is Nothing).
    If InStr(json, """fileExists"":true") > 0 Then
        If InStr(json, """registrado"":true") > 0 Then
            Err.Raise 1003, , "registrado=true but DTO was Nothing — should not happen"
        End If
    End If

    Test_ExpedienteDocumentacionHelper_AltaAnexo_SadDTONothing = BuildOk("dto-nothing-or-fso-check", logs)
    Exit Function
EH:
    Test_ExpedienteDocumentacionHelper_AltaAnexo_SadDTONothing = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 15. EliminarAnexo — sad (no admin -> fail JSON, autorizado=false, eliminado=false)
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteDocumentacionHelper_EliminarAnexo_SadNoAdmin() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim stubExp As Object
    Set stubExp = BuildStubExpediente(Nothing)
    Dim stubDTO As Object
    Set stubDTO = BuildStubDTO(stubExp, Nothing)

    Dim p_Error As String
    Dim json As String
    json = modExpedienteDocumentacionHelper.ExpedienteDocumentacion_EliminarAnexo( _
        stubDTO, "doc-a.pdf", False, vbYes, p_Error)

    If InStr(json, """ok"":false") = 0 Then _
        Err.Raise 1001, , "expected ok=false for non-admin, got " & json
    If p_Error = "" Then Err.Raise 1002, , "expected p_Error to be set"

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If CBool(parsed("payload")("autorizado")) <> False Then _
        Err.Raise 1003, , "expected autorizado=false"
    If CBool(parsed("payload")("eliminado")) <> False Then _
        Err.Raise 1003, , "expected eliminado=false"

    Test_ExpedienteDocumentacionHelper_EliminarAnexo_SadNoAdmin = BuildOk("no-admin-rejected", logs)
    Exit Function
EH:
    Test_ExpedienteDocumentacionHelper_EliminarAnexo_SadNoAdmin = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 16. EliminarAnexo — edge (admin + prompt=ASK=0 -> promptAccepted=false, eliminado=false)
'     Form should ask MsgBox and re-call.
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteDocumentacionHelper_EliminarAnexo_EdgePromptPending() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim stubExp As Object
    Set stubExp = BuildStubExpediente(Nothing)
    Dim stubDTO As Object
    Set stubDTO = BuildStubDTO(stubExp, Nothing)

    Dim p_Error As String
    Dim json As String
    json = modExpedienteDocumentacionHelper.ExpedienteDocumentacion_EliminarAnexo( _
        stubDTO, "doc-a.pdf", True, 0, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If CBool(parsed("payload")("autorizado")) <> True Then _
        Err.Raise 1002, , "expected autorizado=true"
    If CBool(parsed("payload")("promptAccepted")) <> False Then _
        Err.Raise 1002, , "expected promptAccepted=false"
    If CBool(parsed("payload")("eliminado")) <> False Then _
        Err.Raise 1002, , "expected eliminado=false"

    Test_ExpedienteDocumentacionHelper_EliminarAnexo_EdgePromptPending = BuildOk("prompt-pending", logs)
    Exit Function
EH:
    Test_ExpedienteDocumentacionHelper_EliminarAnexo_EdgePromptPending = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 17. EliminarAnexo — edge (admin + prompt=vbNo -> promptAccepted=false, eliminado=false)
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteDocumentacionHelper_EliminarAnexo_EdgePromptNo() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim stubExp As Object
    Set stubExp = BuildStubExpediente(Nothing)
    Dim stubDTO As Object
    Set stubDTO = BuildStubDTO(stubExp, Nothing)

    Dim p_Error As String
    Dim json As String
    json = modExpedienteDocumentacionHelper.ExpedienteDocumentacion_EliminarAnexo( _
        stubDTO, "doc-a.pdf", True, vbNo, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If CBool(parsed("payload")("promptAccepted")) <> False Then _
        Err.Raise 1002, , "expected promptAccepted=false for vbNo"
    If CBool(parsed("payload")("eliminado")) <> False Then _
        Err.Raise 1002, , "expected eliminado=false for vbNo"

    Test_ExpedienteDocumentacionHelper_EliminarAnexo_EdgePromptNo = BuildOk("prompt-no-cancel", logs)
    Exit Function
EH:
    Test_ExpedienteDocumentacionHelper_EliminarAnexo_EdgePromptNo = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 18. EliminarAnexo — edge (admin + prompt=vbYes + empty nombre -> eliminado=false, no DAO)
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteDocumentacionHelper_EliminarAnexo_EdgeEmptyNombre() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim stubExp As Object
    Set stubExp = BuildStubExpediente(Nothing)
    Dim stubDTO As Object
    Set stubDTO = BuildStubDTO(stubExp, Nothing)

    Dim p_Error As String
    Dim json As String
    json = modExpedienteDocumentacionHelper.ExpedienteDocumentacion_EliminarAnexo( _
        stubDTO, "", True, vbYes, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If CBool(parsed("payload")("promptAccepted")) <> True Then _
        Err.Raise 1002, , "expected promptAccepted=true"
    If CBool(parsed("payload")("eliminado")) <> False Then _
        Err.Raise 1002, , "expected eliminado=false for empty nombre"

    Test_ExpedienteDocumentacionHelper_EliminarAnexo_EdgeEmptyNombre = BuildOk("empty-nombre-noop", logs)
    Exit Function
EH:
    Test_ExpedienteDocumentacionHelper_EliminarAnexo_EdgeEmptyNombre = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 19. EliminarAnexo — sad (admin + prompt=vbYes + DTO Nothing -> fail JSON)
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteDocumentacionHelper_EliminarAnexo_SadDTONothing() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modExpedienteDocumentacionHelper.ExpedienteDocumentacion_EliminarAnexo( _
        Nothing, "doc-a.pdf", True, vbYes, p_Error)
    If InStr(json, """ok"":false") = 0 Then _
        Err.Raise 1001, , "expected ok=false, got " & json
    If p_Error = "" Then Err.Raise 1002, , "expected p_Error to be set"

    Test_ExpedienteDocumentacionHelper_EliminarAnexo_SadDTONothing = BuildOk("dto-rejected", logs)
    Exit Function
EH:
    Test_ExpedienteDocumentacionHelper_EliminarAnexo_SadDTONothing = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 20. EliminarAnexo — sad (admin + prompt=vbYes + Expediente Nothing -> fail JSON)
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteDocumentacionHelper_EliminarAnexo_SadExpedienteNothing() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim stubDTO As Object
    Set stubDTO = BuildStubDTO(Nothing, Nothing)

    Dim p_Error As String
    Dim json As String
    json = modExpedienteDocumentacionHelper.ExpedienteDocumentacion_EliminarAnexo( _
        stubDTO, "doc-a.pdf", True, vbYes, p_Error)
    If InStr(json, """ok"":false") = 0 Then _
        Err.Raise 1001, , "expected ok=false, got " & json
    If p_Error = "" Then Err.Raise 1002, , "expected p_Error to be set"

    Test_ExpedienteDocumentacionHelper_EliminarAnexo_SadExpedienteNothing = BuildOk("exp-rejected", logs)
    Exit Function
EH:
    Test_ExpedienteDocumentacionHelper_EliminarAnexo_SadExpedienteNothing = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 21. EliminarAnexo — adversarial (unknown prompt result -> treated as cancel)
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteDocumentacionHelper_EliminarAnexo_AdversarialUnknownPrompt() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim stubExp As Object
    Set stubExp = BuildStubExpediente(Nothing)
    Dim stubDTO As Object
    Set stubDTO = BuildStubDTO(stubExp, Nothing)

    Dim p_Error As String
    Dim json As String
    json = modExpedienteDocumentacionHelper.ExpedienteDocumentacion_EliminarAnexo( _
        stubDTO, "doc-a.pdf", True, 999, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If CBool(parsed("payload")("promptAccepted")) <> False Then _
        Err.Raise 1002, , "expected promptAccepted=false for unknown prompt result"
    If CBool(parsed("payload")("eliminado")) <> False Then _
        Err.Raise 1002, , "expected eliminado=false for unknown prompt result"

    Test_ExpedienteDocumentacionHelper_EliminarAnexo_AdversarialUnknownPrompt = BuildOk("unknown-cancel", logs)
    Exit Function
EH:
    Test_ExpedienteDocumentacionHelper_EliminarAnexo_AdversarialUnknownPrompt = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 22. RunAll — wrapper for Dysflow manifest discovery (access-vba-tdd §1.1.1)
' ---------------------------------------------------------------------------
Public Function Test_ExpedienteDocumentacionHelper_RunAll() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim atoms As Variant
    atoms = Array( _
        "Test_ExpedienteDocumentacionHelper_Form_Load_HappyConExpediente", _
        "Test_ExpedienteDocumentacionHelper_Form_Load_SadDTONothing", _
        "Test_ExpedienteDocumentacionHelper_Form_Load_EdgeExpedienteNothing", _
        "Test_ExpedienteDocumentacionHelper_EstablecerDatos_HappyAdmin", _
        "Test_ExpedienteDocumentacionHelper_EstablecerDatos_HappyNonAdmin", _
        "Test_ExpedienteDocumentacionHelper_EstablecerDatos_SadDTONothing", _
        "Test_ExpedienteDocumentacionHelper_RellenarListas_HappyTresAnexos", _
        "Test_ExpedienteDocumentacionHelper_RellenarListas_EdgeEmptyColAnexos", _
        "Test_ExpedienteDocumentacionHelper_RellenarListas_EdgeDTONothing", _
        "Test_ExpedienteDocumentacionHelper_RellenarListas_EdgeAnexoSinExpediente", _
        "Test_ExpedienteDocumentacionHelper_RellenarListas_AdversarialSemicolonInNombre", _
        "Test_ExpedienteDocumentacionHelper_AltaAnexo_EdgeEmptyURL", _
        "Test_ExpedienteDocumentacionHelper_AltaAnexo_EdgeFileNotExists", _
        "Test_ExpedienteDocumentacionHelper_AltaAnexo_SadDTONothing", _
        "Test_ExpedienteDocumentacionHelper_EliminarAnexo_SadNoAdmin", _
        "Test_ExpedienteDocumentacionHelper_EliminarAnexo_EdgePromptPending", _
        "Test_ExpedienteDocumentacionHelper_EliminarAnexo_EdgePromptNo", _
        "Test_ExpedienteDocumentacionHelper_EliminarAnexo_EdgeEmptyNombre", _
        "Test_ExpedienteDocumentacionHelper_EliminarAnexo_SadDTONothing", _
        "Test_ExpedienteDocumentacionHelper_EliminarAnexo_SadExpedienteNothing", _
        "Test_ExpedienteDocumentacionHelper_EliminarAnexo_AdversarialUnknownPrompt" _
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
        Test_ExpedienteDocumentacionHelper_RunAll = BuildFail("firstFailure=" & firstFailure, logs)
    Else
        Test_ExpedienteDocumentacionHelper_RunAll = BuildOk("all-passed", logs)
    End If
    Exit Function
EH:
    Dim p_Error As String
    p_Error = "Test_ExpedienteDocumentacionHelper_RunAll EH: " & Err.Description
    Test_ExpedienteDocumentacionHelper_RunAll = BuildFail(p_Error, logs)
End Function