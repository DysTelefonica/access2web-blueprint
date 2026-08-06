Attribute VB_Name = "Test_Helper_ExpedienteSuministradoresArbol"
Option Compare Database
Option Explicit

' Test module for src/classes/Helper_ExpedienteSuministradoresArbol.cls
'
' 8 strict-TDD atoms covering Happy / Sad / Edge / Adversarial cases for the
' pure in-memory tree model. No DAO, no UI, no MSComctlLib — only
' Scripting.Dictionary + JsonConverter.
'
' Each atom:
'   - seeds its own fixture (no shared global state, no env vars)
'   - asserts via canonical JSON envelope ({ok, value, payload, error, logs})
'   - avoids Debug.Print / MsgBox / InputBox
'   - uses TestingCore_* wrappers (canonical harness, vba-access §BEFORE WRITING)

Private Const K_ROOT_ORGANO As String = "ROOT_ORGANO"

Private Function BuildOk(ByVal p_Value As Variant, ByRef p_Logs() As String) As String
    BuildOk = TestingCore_BuildOk(p_Value, p_Logs)
End Function

Private Function BuildFail(ByVal p_Error As String, ByRef p_Logs() As String) As String
    BuildFail = TestingCore_BuildFail(p_Error, p_Logs)
End Function

Private Function ParsedJson(ByVal p_Json As String) As Object
    Set ParsedJson = JsonConverter.ParseJson(p_Json)
End Function

Private Function BuildItem( _
    ByVal p_IDExpedienteSuministrador As Long, _
    ByVal p_IdPadre As Long, _
    Optional ByVal p_Text As String = "", _
    Optional ByVal p_Tag As String = "", _
    Optional ByVal p_EsEmpresaPropia As Boolean = False _
) As Object
    Dim item As Object
    Set item = CreateObject("Scripting.Dictionary")
    item("IDExpedienteSuministrador") = p_IDExpedienteSuministrador
    item("IdPadre") = p_IdPadre
    item("Text") = p_Text
    item("Tag") = p_Tag
    item("EsEmpresaPropia") = p_EsEmpresaPropia
    Set BuildItem = item
End Function

Private Function BuildItemsDictionary(ByRef p_Items() As Variant) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = TextCompare

    Dim i As Long
    For i = LBound(p_Items) To UBound(p_Items)
        Dim k As String
        k = "k" & CStr(i)
        Set dict(k) = p_Items(i)
    Next i

    Set BuildItemsDictionary = dict
End Function

Private Function SafeCLng(ByVal p_Value As Variant) As Long
    If IsNull(p_Value) Or IsEmpty(p_Value) Then
        SafeCLng = 0
    Else
        SafeCLng = CLng(p_Value)
    End If
End Function

Private Function SafeCStr(ByVal p_Value As Variant) As String
    If IsNull(p_Value) Or IsEmpty(p_Value) Then
        SafeCStr = ""
    Else
        SafeCStr = CStr(p_Value)
    End If
End Function

Private Function SerializedKeySet(ByVal p_Helper As Helper_ExpedienteSuministradoresArbol, ByRef p_Error As String) As Object
    Dim parsed As Object
    Set parsed = ParsedJson(p_Helper.Serializar(p_Error))
    If p_Error <> "" Then Err.Raise 1000, , p_Error
    If CBool(parsed("ok")) <> True Then Err.Raise 1001, , "expected ok=true"

    Dim keys As Object
    Set keys = CreateObject("Scripting.Dictionary")
    keys.CompareMode = TextCompare

    Dim nodes As Object
    Set nodes = parsed("payload")("nodes")

    Dim i As Long
    For i = 1 To nodes.Count
        keys(SafeCStr(nodes(i)("key"))) = True
    Next i

    Set SerializedKeySet = keys
End Function

' --- ATOM 1: Reset clears state (Happy) -------------------------------------

Public Function Test_Helper_ExpedienteSuministradoresArbol_Reset_LimpiaEstado() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(4)
    On Error GoTo EH

    Dim p_Error As String
    Dim helper As New Helper_ExpedienteSuministradoresArbol

    helper.Init "ORGANO-X", p_Error
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim items(0 To 0) As Variant
    Set items(0) = BuildItem(900001, 0, "Sum-001")
    If helper.Cargar(BuildItemsDictionary(items), p_Error) <> 1 Then Err.Raise 1001, , "expected loaded count=1"
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    logs(0) = "loaded-one-item"

    helper.Reset p_Error
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    logs(1) = "after-reset"

    Dim parsed As Object
    Set parsed = ParsedJson(helper.Serializar(p_Error))
    If p_Error <> "" Then Err.Raise 1000, , p_Error
    If CBool(parsed("ok")) <> True Then Err.Raise 1001, , "expected ok=true after reset"

    If CLng(parsed("payload")("Count")) <> 0 Then Err.Raise 1002, , "expected Count=0 after reset"
    If SafeCStr(parsed("payload")("OrganoNombre")) <> "" Then Err.Raise 1003, , "expected empty organo name after reset"

    logs(2) = "count-zero"
    logs(3) = "organo-cleared"

    Test_Helper_ExpedienteSuministradoresArbol_Reset_LimpiaEstado = BuildOk("reset-clears-state", logs)
    Exit Function
EH:
    Test_Helper_ExpedienteSuministradoresArbol_Reset_LimpiaEstado = BuildFail(Err.Description, logs)
End Function

' --- ATOM 2: Cargar with valid rows returns count (Happy) --------------------

Public Function Test_Helper_ExpedienteSuministradoresArbol_Cargar_Valido_DevuelveConteo() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(5)
    On Error GoTo EH

    Dim p_Error As String
    Dim helper As New Helper_ExpedienteSuministradoresArbol
    helper.Init "DIRECCION GENERAL", p_Error
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim items(0 To 2) As Variant
    Set items(0) = BuildItem(900101, 0, "Sum-A", "tag-a", False)
    Set items(1) = BuildItem(900102, 900101, "Sum-B", "tag-b", True)
    Set items(2) = BuildItem(900103, 900101, "Sum-C", "tag-c", False)

    Dim count As Long
    count = helper.Cargar(BuildItemsDictionary(items), p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error
    If count <> 3 Then Err.Raise 1001, , "expected count=3, got " & count

    logs(0) = "loaded-3-items"

    Dim parsed As Object
    Set parsed = ParsedJson(helper.Serializar(p_Error))
    If p_Error <> "" Then Err.Raise 1000, , p_Error
    If CLng(parsed("payload")("Count")) <> 3 Then Err.Raise 1002, , "expected serialized count=3"

    logs(1) = "serialized-count=3"

    If SafeCStr(parsed("payload")("OrganoNombre")) <> "DIRECCION GENERAL" Then Err.Raise 1003, , "expected organo name preserved"

    logs(2) = "organo-name-preserved"

    Dim nodeB As Object
    Set nodeB = helper.GetNodo("N-900102", p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error
    If nodeB Is Nothing Then Err.Raise 1004, , "expected GetNodo to return N-900102"
    If SafeCLng(nodeB("IdPadre")) <> 900101 Then Err.Raise 1005, , "expected IdPadre=900101"
    If SafeCStr(nodeB("Text")) <> "Sum-B" Then Err.Raise 1006, , "expected text=Sum-B"
    If CBool(nodeB("EsEmpresaPropia")) <> True Then Err.Raise 1007, , "expected esEmpresaPropia=true"

    logs(3) = "getnodo-happy"

    Test_Helper_ExpedienteSuministradoresArbol_Cargar_Valido_DevuelveConteo = BuildOk("cargar-valid-count-3", logs)
    Exit Function
EH:
    Test_Helper_ExpedienteSuministradoresArbol_Cargar_Valido_DevuelveConteo = BuildFail(Err.Description, logs)
End Function

' --- ATOM 3: Cargar rejects duplicates atomically (Sad) ----------------------

Public Function Test_Helper_ExpedienteSuministradoresArbol_Cargar_Duplicados_RechazaAtomico() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(5)
    On Error GoTo EH

    Dim p_Error As String
    Dim helper As New Helper_ExpedienteSuministradoresArbol
    helper.Init "ORGANO-Y", p_Error
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim seed(0 To 0) As Variant
    Set seed(0) = BuildItem(900201, 0, "Original")
    If helper.Cargar(BuildItemsDictionary(seed), p_Error) <> 1 Then Err.Raise 1001, , "expected seed load count=1"
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    logs(0) = "seed-loaded"

    Dim items(0 To 1) As Variant
    Set items(0) = BuildItem(900202, 0, "First")
    Set items(1) = BuildItem(900202, 0, "Duplicate")

    Dim count As Long

    On Error GoTo ExpectedError
    count = helper.Cargar(BuildItemsDictionary(items), p_Error)
    Err.Raise 1001, , "expected duplicate to raise"

ExpectedError:
    On Error GoTo EH
    If Err.Number = 1001 Then Err.Raise 1001, , Err.Description
    If count <> 0 Then Err.Raise 1002, , "expected rejected count=0"
    If p_Error = "" Then Err.Raise 1003, , "expected p_Error set for duplicate"
    If InStr(1, p_Error, "duplicate", vbTextCompare) = 0 And _
       InStr(1, p_Error, "duplicado", vbTextCompare) = 0 Then Err.Raise 1004, , "expected duplicate keyword in p_Error"

    logs(1) = "duplicate-rejected"

    Dim parsed As Object
    Set parsed = ParsedJson(helper.Serializar(p_Error))
    If p_Error <> "" Then Err.Raise 1000, , p_Error
    If CLng(parsed("payload")("Count")) <> 1 Then Err.Raise 1005, , "expected state unchanged: count=1"
    If SafeCStr(parsed("payload")("OrganoNombre")) <> "ORGANO-Y" Then Err.Raise 1006, , "expected organo unchanged"

    logs(2) = "state-unchanged"

    Test_Helper_ExpedienteSuministradoresArbol_Cargar_Duplicados_RechazaAtomico = BuildOk("duplicates-rejected", logs)
    Exit Function
EH:
    Test_Helper_ExpedienteSuministradoresArbol_Cargar_Duplicados_RechazaAtomico = BuildFail(Err.Description, logs)
End Function

' --- ATOM 4: Cargar rejects orphans atomically (Adversarial) ----------------

Public Function Test_Helper_ExpedienteSuministradoresArbol_Cargar_Huerfanos_RechazaAtomico() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(5)
    On Error GoTo EH

    Dim p_Error As String
    Dim helper As New Helper_ExpedienteSuministradoresArbol
    helper.Init "ORGANO-Z", p_Error
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim items(0 To 1) As Variant
    Set items(0) = BuildItem(900301, 0, "Sum-301")
    Set items(1) = BuildItem(900302, 999999, "Sum-302-Orphan")

    Dim count As Long

    On Error GoTo ExpectedError
    count = helper.Cargar(BuildItemsDictionary(items), p_Error)
    Err.Raise 1001, , "expected orphan to raise"

ExpectedError:
    On Error GoTo EH
    If Err.Number = 1001 Then Err.Raise 1001, , Err.Description
    If count <> 0 Then Err.Raise 1002, , "expected orphan rejected count=0"
    If p_Error = "" Then Err.Raise 1003, , "expected p_Error set for orphan"
    If InStr(1, p_Error, "orphan", vbTextCompare) = 0 And _
       InStr(1, p_Error, "huerfano", vbTextCompare) = 0 And _
       InStr(1, p_Error, "huérfano", vbTextCompare) = 0 Then Err.Raise 1004, , "expected orphan keyword in p_Error"

    logs(0) = "orphan-rejected"

    Dim parsed As Object
    Set parsed = ParsedJson(helper.Serializar(p_Error))
    If p_Error <> "" Then Err.Raise 1000, , p_Error
    If CLng(parsed("payload")("Count")) <> 0 Then Err.Raise 1005, , "expected state unchanged: count=0"

    logs(1) = "state-stays-empty"

    Test_Helper_ExpedienteSuministradoresArbol_Cargar_Huerfanos_RechazaAtomico = BuildOk("orphans-rejected", logs)
    Exit Function
EH:
    Test_Helper_ExpedienteSuministradoresArbol_Cargar_Huerfanos_RechazaAtomico = BuildFail(Err.Description, logs)
End Function

' --- ATOM 5: Cargar rejects cycles atomically (Adversarial) -----------------

Public Function Test_Helper_ExpedienteSuministradoresArbol_Cargar_Ciclos_RechazaAtomico() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(5)
    On Error GoTo EH

    Dim p_Error As String
    Dim helper As New Helper_ExpedienteSuministradoresArbol
    helper.Init "ORGANO-W", p_Error
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim items(0 To 2) As Variant
    Set items(0) = BuildItem(900401, 900402, "A-cycles-to-C")
    Set items(1) = BuildItem(900402, 900403, "B-cycles-to-A-via-itself")
    Set items(2) = BuildItem(900403, 900401, "C-cycles-to-B")

    Dim count As Long

    On Error GoTo ExpectedError
    count = helper.Cargar(BuildItemsDictionary(items), p_Error)
    Err.Raise 1001, , "expected cycle to raise"

ExpectedError:
    On Error GoTo EH
    If Err.Number = 1001 Then Err.Raise 1001, , Err.Description
    If count <> 0 Then Err.Raise 1002, , "expected cycle rejected count=0"
    If p_Error = "" Then Err.Raise 1003, , "expected p_Error set for cycle"
    If InStr(1, p_Error, "cycle", vbTextCompare) = 0 And _
       InStr(1, p_Error, "ciclo", vbTextCompare) = 0 Then Err.Raise 1004, , "expected cycle keyword in p_Error"

    logs(0) = "cycle-rejected"

    Dim parsed As Object
    Set parsed = ParsedJson(helper.Serializar(p_Error))
    If p_Error <> "" Then Err.Raise 1000, , p_Error
    If CLng(parsed("payload")("Count")) <> 0 Then Err.Raise 1005, , "expected state unchanged: count=0"

    logs(1) = "state-stays-empty"

    Test_Helper_ExpedienteSuministradoresArbol_Cargar_Ciclos_RechazaAtomico = BuildOk("cycles-rejected", logs)
    Exit Function
EH:
    Test_Helper_ExpedienteSuministradoresArbol_Cargar_Ciclos_RechazaAtomico = BuildFail(Err.Description, logs)
End Function

' --- ATOM 6: Serializar produces deterministic order independent of input (Edge) ---

Public Function Test_Helper_ExpedienteSuministradoresArbol_Serializar_OrdenDeterminista() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(6)
    On Error GoTo EH

    Dim p_Error As String
    Dim helperA As New Helper_ExpedienteSuministradoresArbol
    helperA.Init "ORDEN-TEST", p_Error
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim itemsForward(0 To 4) As Variant
    Set itemsForward(0) = BuildItem(900501, 0, "A")
    Set itemsForward(1) = BuildItem(900502, 0, "B")
    Set itemsForward(2) = BuildItem(900503, 900501, "C-child-of-A")
    Set itemsForward(3) = BuildItem(900504, 900501, "D-child-of-A")
    Set itemsForward(4) = BuildItem(900505, 900502, "E-child-of-B")

    If helperA.Cargar(BuildItemsDictionary(itemsForward), p_Error) <> 5 Then Err.Raise 1001, , "expected forward load count=5"
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    logs(0) = "forward-loaded"

    Dim helperB As New Helper_ExpedienteSuministradoresArbol
    helperB.Init "ORDEN-TEST", p_Error
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim itemsReverse(0 To 4) As Variant
    Set itemsReverse(0) = BuildItem(900505, 900502, "E-child-of-B")
    Set itemsReverse(1) = BuildItem(900504, 900501, "D-child-of-A")
    Set itemsReverse(2) = BuildItem(900503, 900501, "C-child-of-A")
    Set itemsReverse(3) = BuildItem(900502, 0, "B")
    Set itemsReverse(4) = BuildItem(900501, 0, "A")

    If helperB.Cargar(BuildItemsDictionary(itemsReverse), p_Error) <> 5 Then Err.Raise 1002, , "expected reverse load count=5"
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    logs(1) = "reverse-loaded"

    Dim jsonA As String
    jsonA = helperA.Serializar(p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim jsonB As String
    jsonB = helperB.Serializar(p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    logs(2) = "both-serialized"

    If StrComp(jsonA, jsonB, vbBinaryCompare) <> 0 Then Err.Raise 1003, , "expected deterministic serial output"

    logs(3) = "outputs-identical"

    Dim parsed As Object
    Set parsed = ParsedJson(jsonA)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim nodes As Object
    Set nodes = parsed("payload")("nodes")

    If SafeCStr(nodes(1)("key")) <> "ROOT_ORGANO" Then Err.Raise 1004, , "expected ROOT_ORGANO first"

    logs(4) = "root-first"

    Dim i As Long
    Dim prevDepth As Long
    prevDepth = -1
    For i = 1 To nodes.Count
        Dim d As Long
        d = SafeCLng(nodes(i)("depth"))
        If d < prevDepth Then Err.Raise 1005, , "expected non-decreasing depth order"
        prevDepth = d
    Next i

    logs(5) = "depth-monotonic"

    Test_Helper_ExpedienteSuministradoresArbol_Serializar_OrdenDeterminista = BuildOk("serializar-deterministic", logs)
    Exit Function
EH:
    Test_Helper_ExpedienteSuministradoresArbol_Serializar_OrdenDeterminista = BuildFail(Err.Description, logs)
End Function

' --- ATOM 7: GetNodo for missing key returns Nothing (Sad) -------------------

Public Function Test_Helper_ExpedienteSuministradoresArbol_GetNodo_Inexistente_DevuelveNothing() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(4)
    On Error GoTo EH

    Dim p_Error As String
    Dim helper As New Helper_ExpedienteSuministradoresArbol
    helper.Init "ORGANO-NX", p_Error
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim items(0 To 0) As Variant
    Set items(0) = BuildItem(900601, 0, "Only")
    If helper.Cargar(BuildItemsDictionary(items), p_Error) <> 1 Then Err.Raise 1001, , "expected load count=1"
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    logs(0) = "loaded"

    Dim result As Object
    Set result = helper.GetNodo("N-999999", p_Error)
    If p_Error <> "" Then Err.Raise 1002, , "expected no p_Error on missing key"
    If Not result Is Nothing Then Err.Raise 1003, , "expected Nothing for missing key"

    logs(1) = "missing-returns-nothing"

    Set result = helper.GetNodo("DOES-NOT-EXIST", p_Error)
    If Not result Is Nothing Then Err.Raise 1004, , "expected Nothing for non-numeric key"

    logs(2) = "non-numeric-also-nothing"

    Test_Helper_ExpedienteSuministradoresArbol_GetNodo_Inexistente_DevuelveNothing = BuildOk("getnodo-missing-nothing", logs)
    Exit Function
EH:
    Test_Helper_ExpedienteSuministradoresArbol_GetNodo_Inexistente_DevuelveNothing = BuildFail(Err.Description, logs)
End Function

' --- ATOM 8: SeleccionarNodo unknown returns False without changing previous (Sad) --

Public Function Test_Helper_ExpedienteSuministradoresArbol_SeleccionarNodo_Inexistente_NoCambiaPrevio() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(5)
    On Error GoTo EH

    Dim p_Error As String
    Dim helper As New Helper_ExpedienteSuministradoresArbol
    helper.Init "ORGANO-SEL", p_Error
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim items(0 To 1) As Variant
    Set items(0) = BuildItem(900701, 0, "Sum-A")
    Set items(1) = BuildItem(900702, 900701, "Sum-B")
    If helper.Cargar(BuildItemsDictionary(items), p_Error) <> 2 Then Err.Raise 1001, , "expected load count=2"
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    logs(0) = "loaded"

    If helper.SeleccionarNodo("N-900701", p_Error) <> True Then Err.Raise 1002, , "expected first selection to succeed"
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    logs(1) = "first-selection-ok"

    Dim parsed As Object
    Set parsed = ParsedJson(helper.Serializar(p_Error))
    If p_Error <> "" Then Err.Raise 1000, , p_Error
    If SafeCStr(parsed("payload")("NodoSeleccionado")) <> "N-900701" Then Err.Raise 1003, , "expected NodoSeleccionado=N-900701"

    logs(2) = "selection-recorded"

    If helper.SeleccionarNodo("N-999999", p_Error) <> False Then Err.Raise 1004, , "expected unknown selection to return False"
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    logs(3) = "unknown-returns-false"

    Set parsed = ParsedJson(helper.Serializar(p_Error))
    If p_Error <> "" Then Err.Raise 1000, , p_Error
    If SafeCStr(parsed("payload")("NodoSeleccionado")) <> "N-900701" Then Err.Raise 1005, , "expected previous selection preserved"

    logs(4) = "previous-preserved"

    Test_Helper_ExpedienteSuministradoresArbol_SeleccionarNodo_Inexistente_NoCambiaPrevio = BuildOk("seleccionar-non-destructive", logs)
    Exit Function
EH:
    Test_Helper_ExpedienteSuministradoresArbol_SeleccionarNodo_Inexistente_NoCambiaPrevio = BuildFail(Err.Description, logs)
End Function

' ==============================================================================
' ATOMS 9 & 10 — wiring (PRUEBA-003/REFAC-4b2b, closes #60)
'
' These atoms exercise the form -> DAO -> helper chain. They follow the same
' canonical harness contract as atoms 1-8 (BuildOk/BuildFail envelopes, JSON
' payload via JsonConverter, schema-first, no shared state).
'
' Atom 9 exercises the form path: RefrescarAArbol delegates to the helper,
' observed through a Public Property Get that returns the serialized JSON.
'
' Atom 10 exercises the DAO injection path: GetDatosArbolTestable + Cargar
' with an injected sandboxDb must not modify the sandbox cardinality
' (no-op semantics — the helper is read-only over the rows).
' ==============================================================================

' --- ATOM 9: Form.RefrescarAArbol delegates to m_Helper (wiring) --------------

Public Function Test_Form_ExpedienteSuministradores_RefrescarAArbol_DelegateAlHelper() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(5)
    On Error GoTo EH

    ' Strategy (c) per issue #60: Public Property Get NodosSerializados_ in the
    ' form returns m_Helper.Serializar(m_Error). This keeps m_Helper Private
    ' (encapsulation) while exposing the verifiable JSON envelope. The test
    ' calls RefrescarAArbol (now Public per slice #60) and reads the property.
    '
    ' Test approach: instantiate the form class directly (no DoCmd.OpenForm),
    ' inject m_IDExpediente via the public Property Let setter, then call
    ' RefrescarAArbol. m_ObjExpedienteDTOActivo is left at its default (Nothing),
    ' so ResolveOrganoNombre falls back to "DESCONOCIDO" — we only assert that
    ' the helper received the rows from the seeded DAO via the wiring chain.
    ' The painting step (Me.tvCadena.Nodes.*) fails silently because the form
    ' is never OpenForm'd; the data wire-up completes successfully.
    Dim p_Error As String
    Dim frm As Form_FormExpedienteSuministradores
    Dim sandboxDb As DAO.Database
    Dim items As New Collection
    Dim seedCount As Long
    Dim json As String
    Dim parsed As Object

    Set sandboxDb = GetTestDb(p_Error)
    If p_Error <> "" Then Err.Raise 1000, , "GetTestDb: " & p_Error
    If sandboxDb Is Nothing Then Err.Raise 1000, , "GetTestDb returned Nothing"

    logs(0) = "sandbox-ready"

    ' Pre-isolation: clear fixture rows in the sandbox
    If Not TestHelper.EnsureWorkingTableClean(sandboxDb, "TbExpedientesSuministradores", _
            "IDExpedienteSuministrador >= 900000 AND IDExpedienteSuministrador <= 900099", p_Error) Then
        Err.Raise 1000, , "EnsureWorkingTableClean: " & p_Error
    End If

    ' Build 2-row fixture: 1 root + 1 child
    Dim item1 As New ExpedienteSuministrador
    item1.SetPropiedad "IDExpedienteSuministrador", 900010
    item1.SetPropiedad "IDExpediente", 900000
    item1.SetPropiedad "IDSuministrador", 900010
    item1.SetPropiedad "IdPadre", Null
    item1.SetPropiedad "ContratistaPrincipal", "Sí"
    item1.SetPropiedad "SubContratista", "No"
    item1.SetPropiedad "Descripcon", "WireAtom9-Root"
    items.Add item1

    Dim item2 As New ExpedienteSuministrador
    item2.SetPropiedad "IDExpedienteSuministrador", 900011
    item2.SetPropiedad "IDExpediente", 900000
    item2.SetPropiedad "IDSuministrador", 900011
    item2.SetPropiedad "IdPadre", 900010
    item2.SetPropiedad "ContratistaPrincipal", "No"
    item2.SetPropiedad "SubContratista", "Sí"
    item2.SetPropiedad "Descripcon", "WireAtom9-Child"
    items.Add item2

    seedCount = TestHelper.SeedArbolSuministradores(items, sandboxDb, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , "SeedArbolSuministradores: " & p_Error
    If seedCount <> 2 Then Err.Raise 1001, , "expected seed count=2, got " & seedCount

    logs(1) = "seeded-2-rows"

    ' Instantiate the form class directly (no OpenForm; the wiring test
    ' exercises DAO + Helper without needing the live UI).
    Set frm = New Form_FormExpedienteSuministradores

    logs(2) = "form-instantiated"

    ' Inject IDExpediente via the public Property Let and trigger the wire-up
    frm.IDExpediente_ForTest = "900000"
    frm.RefrescarAArbol

    logs(3) = "refrescar-called"

    ' Observe via the public property (strategy c)
    json = frm.NodosSerializados_
    Set parsed = ParsedJson(json)
    If CBool(parsed("ok")) <> True Then Err.Raise 1002, , "expected ok=true from helper envelope"

    Dim nodesCount As Long
    nodesCount = CLng(parsed("payload")("Count"))
    If nodesCount < 1 Then Err.Raise 1003, , "expected nodes.Count >= 1, got " & nodesCount

    logs(4) = "helper-has-nodes=" & nodesCount

    Test_Form_ExpedienteSuministradores_RefrescarAArbol_DelegateAlHelper = BuildOk("wiring-delegate", logs)
    Exit Function
EH:
    On Error Resume Next
    Test_Form_ExpedienteSuministradores_RefrescarAArbol_DelegateAlHelper = BuildFail(Err.Description, logs)
End Function

' --- ATOM 10: Cargar + DAO adapter over injected db = no-op cardinality ------

Public Function Test_Helper_ExpedienteSuministradoresArbol_CargarDesdeDaoInyectado_NoTocaSandbox() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(6)
    On Error GoTo EH

    Dim p_Error As String
    Dim sandboxDb As DAO.Database
    Dim items As New Collection
    Dim seedCount As Long
    Dim helper As New Helper_ExpedienteSuministradoresArbol
    Dim colDatos As Collection
    Dim itemsDict As Object
    Dim beforeCount As Long
    Dim afterCount As Long
    Dim rs As DAO.Recordset
    Dim sqlCount As String

    Set sandboxDb = GetTestDb(p_Error)
    If p_Error <> "" Then Err.Raise 1000, , "GetTestDb: " & p_Error
    If sandboxDb Is Nothing Then Err.Raise 1000, , "GetTestDb returned Nothing"

    logs(0) = "sandbox-ready"

    ' Pre-isolation: clear fixture rows in the sandbox
    If Not TestHelper.EnsureWorkingTableClean(sandboxDb, "TbExpedientesSuministradores", _
            "IDExpedienteSuministrador >= 900100 AND IDExpedienteSuministrador <= 900199", p_Error) Then
        Err.Raise 1000, , "EnsureWorkingTableClean: " & p_Error
    End If

    logs(1) = "table-clean"

    ' Seed 3 rows in the sandbox via the fixture helper (must pass sandboxDb)
    Dim a As New ExpedienteSuministrador
    a.SetPropiedad "IDExpedienteSuministrador", 900110
    a.SetPropiedad "IDExpediente", 900100
    a.SetPropiedad "IDSuministrador", 900110
    a.SetPropiedad "IdPadre", Null
    a.SetPropiedad "ContratistaPrincipal", "Sí"
    a.SetPropiedad "SubContratista", "No"
    a.SetPropiedad "Descripcon", "NoOpAtom10-Root"
    items.Add a

    Dim b As New ExpedienteSuministrador
    b.SetPropiedad "IDExpedienteSuministrador", 900111
    b.SetPropiedad "IDExpediente", 900100
    b.SetPropiedad "IDSuministrador", 900111
    b.SetPropiedad "IdPadre", 900110
    b.SetPropiedad "ContratistaPrincipal", "No"
    b.SetPropiedad "SubContratista", "Sí"
    b.SetPropiedad "Descripcon", "NoOpAtom10-Child"
    items.Add b

    Dim c As New ExpedienteSuministrador
    c.SetPropiedad "IDExpedienteSuministrador", 900112
    c.SetPropiedad "IDExpediente", 900100
    c.SetPropiedad "IDSuministrador", 900112
    c.SetPropiedad "IdPadre", 900111
    c.SetPropiedad "ContratistaPrincipal", "No"
    c.SetPropiedad "SubContratista", "Sí"
    c.SetPropiedad "Descripcon", "NoOpAtom10-Grand"
    items.Add c

    seedCount = TestHelper.SeedArbolSuministradores(items, sandboxDb, p_Error)
    If p_Error <> "" Then Err.Raise 1001, , "SeedArbolSuministradores: " & p_Error
    If seedCount <> 3 Then Err.Raise 1002, , "expected seed count=3, got " & seedCount

    logs(2) = "seeded-3-rows"

    ' Snapshot BEFORE
    sqlCount = "SELECT COUNT(*) AS n FROM TbExpedientesSuministradores " & _
               "WHERE IDExpedienteSuministrador >= 900100 AND IDExpedienteSuministrador <= 900199"
    Set rs = sandboxDb.OpenRecordset(sqlCount, dbOpenSnapshot)
    beforeCount = CLng(rs!n)
    rs.Close
    Set rs = Nothing

    logs(3) = "before-count=" & beforeCount

    ' Exercise the DAO adapter + helper with injected sandboxDb
    Set colDatos = ExpedienteSuministradorRepositorio.GetDatosArbolTestable(900100, sandboxDb, p_Error)
    If p_Error <> "" Then Err.Raise 1003, , "GetDatosArbolTestable: " & p_Error
    If colDatos Is Nothing Then Err.Raise 1004, , "GetDatosArbolTestable returned Nothing"
    If colDatos.Count <> 3 Then Err.Raise 1005, , "expected col.Count=3, got " & colDatos.Count

    logs(4) = "dao-loaded-3-from-sandbox"

    ' Build helper items dictionary
    Set itemsDict = CreateObject("Scripting.Dictionary")
    itemsDict.CompareMode = TextCompare
    Dim i As Long
    Dim k As String
    For i = 1 To colDatos.Count
        Dim itm As ExpedienteSuministrador
        Set itm = colDatos(i)
        k = "k" & CStr(i)
        Dim dictItem As Object
        Set dictItem = CreateObject("Scripting.Dictionary")
        dictItem("IDExpedienteSuministrador") = CLng(itm.getPropiedad("IDExpedienteSuministrador", p_Error))
        If p_Error <> "" Then Err.Raise 1006, , p_Error
        dictItem("IdPadre") = itm.getPropiedad("IdPadre", p_Error)
        If p_Error <> "" Then Err.Raise 1006, , p_Error
        dictItem("Text") = CStr(itm.getPropiedad("Descripcon", p_Error))
        If p_Error <> "" Then Err.Raise 1006, , p_Error
        dictItem("Tag") = "RELID=" & CStr(dictItem("IDExpedienteSuministrador")) & ";IDS=" & CStr(itm.getPropiedad("IDSuministrador", p_Error))
        dictItem("EsEmpresaPropia") = False
        itemsDict.Add k, dictItem
    Next i

    helper.Init "ORGANO-NOOP", p_Error
    If p_Error <> "" Then Err.Raise 1007, , p_Error
    If helper.Cargar(itemsDict, p_Error) <> 3 Then Err.Raise 1008, , "expected helper.Cargar count=3"
    If p_Error <> "" Then Err.Raise 1007, , p_Error

    logs(5) = "helper-loaded-3-no-tx"

    ' Snapshot AFTER
    Set rs = sandboxDb.OpenRecordset(sqlCount, dbOpenSnapshot)
    afterCount = CLng(rs!n)
    rs.Close
    Set rs = Nothing

    If afterCount <> beforeCount Then
        Err.Raise 1009, , "expected cardinality unchanged (before=" & beforeCount & ", after=" & afterCount & ")"
    End If

    logs(5) = logs(5) & ";cardinality-stable=" & afterCount

    Test_Helper_ExpedienteSuministradoresArbol_CargarDesdeDaoInyectado_NoTocaSandbox = BuildOk("no-op-cardinality", logs)
    Exit Function
EH:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Test_Helper_ExpedienteSuministradoresArbol_CargarDesdeDaoInyectado_NoTocaSandbox = BuildFail(Err.Description, logs)
End Function