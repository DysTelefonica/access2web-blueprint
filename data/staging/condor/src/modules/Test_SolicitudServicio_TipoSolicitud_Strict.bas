Attribute VB_Name = "Test_SolicitudServicio_TipoSolicitud_Strict"
Option Compare Database
Option Explicit

' ============================================================================
' Test_SolicitudServicio_TipoSolicitud_Strict — v2.4.2 strict TDD atoms
'
' Skill:    access-vba-tdd v2.4.2
' SDD/CAP:  nueva capacidad solicitud-validacion (Spec-001)
' Branch:   staging
' Project:  condor (Dysflow projectId)
'
' Helper under test:
'   SolicitudServicio.Validar(ByRef entidad As Solicitud, Optional ByVal esEdicion As Boolean, Optional ByRef db)
'
' Contract (from spec solicitud-validacion):
'   - Canonical tipoSolicitud set: PC, CD_CA, CD_CA_SUB, PC_SUB
'   - Legacy aliases PCSUB and CDCASUB MUST remain accepted (evidence-gated
'     by task 1.0 audit — used by SnapshotHelper and getNombreAmigableTipoSolicitud)
'   - Unknown values MUST raise Err 513 with descriptive message
'   - Empty tipoSolicitud MUST keep existing required-field behavior (also Err 513)
'
' Atom strategy
'   - Domain-reject atoms (unknown, empty): domain check fires BEFORE the
'     duplicate check, so DB is irrelevant for the rejection path. They run
'     with db = Nothing and a Solicitud that satisfies prior basic validations.
'   - Domain-accept atoms (canonical, alias): the duplicate check fires AFTER
'     domain accept. To keep atoms deterministic we set up a sandbox db and
'     use timestamped codigoSolicitud so the duplicate check returns False.
'     If BeginTestSession cannot run, accept atoms return TESTS BLOCKED.
' ============================================================================

' --- Canonical/alias tokens (used for assertions on the Solicitud entity) ---
Private Const TYPE_PC As String = "PC"
Private Const TYPE_CD_CA As String = "CD_CA"
Private Const TYPE_CD_CA_SUB As String = "CD_CA_SUB"
Private Const TYPE_PC_SUB As String = "PC_SUB"
Private Const TYPE_PCSUB_ALIAS As String = "PCSUB"
Private Const TYPE_CDCASUB_ALIAS As String = "CDCASUB"
Private Const TYPE_UNKNOWN As String = "XYZ_BOGUS"
Private Const ERR_DESC_UNKNOWN_FRAGMENT As String = "Tipo de solicitud no válido"
Private Const ERR_DESC_EMPTY_FRAGMENT As String = "Debe seleccionar un tipo de solicitud"

' --- Time-prefixed codigoSolicitud helper to avoid duplicate collisions ---
Private Const CODIGO_PREFIX As String = "DC-TESTSOL-DTS-"

Private m_SSTTempRoot As String

' ============================================================================
' JSON wrappers (per skill §1.8)
' ============================================================================
Private Function BuildOk(ByVal value As Variant, ByRef logs() As String) As String
    BuildOk = TestHelper.BuildJsonOk(CStr(value), logs)
End Function

Private Function BuildFail(ByVal msg As String, ByRef logs() As String) As String
    BuildFail = TestHelper.BuildJsonFail(msg, logs)
End Function

' ============================================================================
' Helpers
' ============================================================================
Private Function CodigoUnico(ByVal sufijo As String) As String
    CodigoUnico = CODIGO_PREFIX & sufijo & "-" & Format$(Now(), "yyyymmddhhnnss") & "-" & CStr(Int(Rnd * 100000))
End Function

Private Function ErrorDescriptionContains(ByVal errDesc As String, ByVal expectedFragment As String) As Boolean
    ErrorDescriptionContains = (InStr(1, errDesc, expectedFragment, vbTextCompare) > 0)
End Function

Private Function BeginSSTSession(ByRef logs() As String, ByRef p_Error As String) As Boolean
    p_Error = ""
    m_SSTTempRoot = ""
    If Not TestHelper.BeginTestSession(logs, p_Error) Then Exit Function

    Call TestHelper.SetupProdGlobalsForTest(p_Error, m_SSTTempRoot, "sst_tipo_solicitud", "SST Test User")
    If Len(p_Error) > 0 Then
        Call TestHelper.ResetTestSession
        Exit Function
    End If

    BeginSSTSession = True
End Function

Private Sub EndSSTSession(ByRef logs() As String)
    Call TestHelper.ResetTestSession
    Call TestHelper.CleanupProdTempRoot(m_SSTTempRoot)
    m_SSTTempRoot = ""
End Sub

Private Sub SetBasicosValidos(ByRef sol As Solicitud, ByVal tipoSolicitud As String)
    ' Establishes the minimum required-field state so the Validar path reaches
    ' the new domain check and (for accept atoms) the duplicate check passes.
    sol.idExpediente = 900200    ' test-only id; outside canonical ranges
    sol.tipoSolicitud = tipoSolicitud
    sol.codigoSolicitud = CodigoUnico("v")
    sol.idNCAsociada = 0
    sol.revisionCalidadEstado = ""    ' empty skips the revisionCalidadEstado check
    sol.idEstadoInterno = 1           ' any non-zero; Validar does not check this
End Sub

' ============================================================================
' ATOMS — Domain reject (no DB required; rejection fires before duplicate)
' ============================================================================
Public Function Test_SST_Validar_UnknownType_Raises513() As String
    Dim logs(0 To 3) As String
    Dim sol As New Solicitud
    Dim srv As New SolicitudServicio
    Dim errNum As Long
    Dim errDesc As String
    Dim sessionError As String
    Dim db As DAO.Database

    logs(0) = "1. Arrange: Solicitud with tipoSolicitud = '" & TYPE_UNKNOWN & "' (unknown)"
    If Not BeginSSTSession(logs, sessionError) Then
        Test_SST_Validar_UnknownType_Raises513 = BuildFail("TESTS BLOCKED: " & sessionError, logs)
        Exit Function
    End If

    Call SetBasicosValidos(sol, TYPE_UNKNOWN)
    Set db = TestHelper.GetTestDb()

    logs(1) = "2. Act: SolicitudServicio.Validar(sol, False, db) — expect Err 513"
    On Error Resume Next
    Call srv.Validar(sol, False, db)
    errNum = Err.Number
    errDesc = Err.description
    On Error GoTo 0
    Call EndSSTSession(logs)

    logs(2) = "3. Assert: errNum = " & errNum & "; desc = " & errDesc

    If errNum = 0 Then
        Test_SST_Validar_UnknownType_Raises513 = BuildFail("Validar did NOT raise for unknown tipoSolicitud '" & TYPE_UNKNOWN & "'", logs)
        Exit Function
    End If

    If errNum <> 513 Then
        Test_SST_Validar_UnknownType_Raises513 = BuildFail("Validar raised " & errNum & " (expected 513) for unknown tipoSolicitud; desc: " & errDesc, logs)
        Exit Function
    End If

    If Not ErrorDescriptionContains(errDesc, ERR_DESC_UNKNOWN_FRAGMENT) Then
        Test_SST_Validar_UnknownType_Raises513 = BuildFail("Validar raised Err 513 but description is not descriptive for invalid tipoSolicitud; expected fragment: " & ERR_DESC_UNKNOWN_FRAGMENT & "; desc: " & errDesc, logs)
        Exit Function
    End If

    If Not ErrorDescriptionContains(errDesc, TYPE_UNKNOWN) Then
        Test_SST_Validar_UnknownType_Raises513 = BuildFail("Validar raised Err 513 but description does not identify the invalid value; expected value: " & TYPE_UNKNOWN & "; desc: " & errDesc, logs)
        Exit Function
    End If

    Test_SST_Validar_UnknownType_Raises513 = BuildOk("unknown tipoSolicitud raises Err 513 with descriptive message", logs)
End Function

Public Function Test_SST_Validar_EmptyType_Raises513_PreservesExistingBehavior() As String
    Dim logs(0 To 3) As String
    Dim sol As New Solicitud
    Dim srv As New SolicitudServicio
    Dim errNum As Long
    Dim errDesc As String
    Dim sessionError As String
    Dim db As DAO.Database

    logs(0) = "1. Arrange: Solicitud with empty tipoSolicitud (existing required-field check)"
    If Not BeginSSTSession(logs, sessionError) Then
        Test_SST_Validar_EmptyType_Raises513_PreservesExistingBehavior = BuildFail("TESTS BLOCKED: " & sessionError, logs)
        Exit Function
    End If

    Call SetBasicosValidos(sol, "")
    sol.codigoSolicitud = CodigoUnico("e")
    Set db = TestHelper.GetTestDb()

    logs(1) = "2. Act: SolicitudServicio.Validar(sol, False, db) — expect Err 513 from existing empty check"
    On Error Resume Next
    Call srv.Validar(sol, False, db)
    errNum = Err.Number
    errDesc = Err.description
    On Error GoTo 0
    Call EndSSTSession(logs)

    logs(2) = "3. Assert: errNum = " & errNum & "; desc = " & errDesc

    If errNum = 0 Then
        Test_SST_Validar_EmptyType_Raises513_PreservesExistingBehavior = BuildFail("Validar did NOT raise for empty tipoSolicitud", logs)
        Exit Function
    End If

    If errNum <> 513 Then
        Test_SST_Validar_EmptyType_Raises513_PreservesExistingBehavior = BuildFail("Validar raised " & errNum & " (expected 513) for empty tipoSolicitud; desc: " & errDesc, logs)
        Exit Function
    End If

    If Not ErrorDescriptionContains(errDesc, ERR_DESC_EMPTY_FRAGMENT) Then
        Test_SST_Validar_EmptyType_Raises513_PreservesExistingBehavior = BuildFail("Validar raised Err 513 but description is not descriptive for empty tipoSolicitud; expected fragment: " & ERR_DESC_EMPTY_FRAGMENT & "; desc: " & errDesc, logs)
        Exit Function
    End If

    Test_SST_Validar_EmptyType_Raises513_PreservesExistingBehavior = BuildOk("empty tipoSolicitud keeps Err 513 with descriptive message", logs)
End Function

' ============================================================================
' ATOMS — Domain accept (require a working sandbox db; if BeginTestSession
' cannot run, the atom returns TESTS BLOCKED so the runner can distinguish
' "test infra blocked" from "test failed")
' ============================================================================

Private Function TryAcceptDomain(ByVal tipoSolicitud As String, ByRef logs() As String) As Boolean
    Dim sol As New Solicitud
    Dim srv As New SolicitudServicio
    Dim db As DAO.Database
    Dim errNum As Long
    Dim errDesc As String
    Dim sessionError As String

    On Error GoTo AcceptFail

    If Not BeginSSTSession(logs, sessionError) Then
        logs(1) = logs(1) & " | " & sessionError
        GoTo AcceptClean
    End If

    ' Require sandbox session so duplicate checks and error logging route safely.
    Set db = TestHelper.GetTestDb()
    If db Is Nothing Then GoTo AcceptFail

    Call SetBasicosValidos(sol, tipoSolicitud)
    ' Inject db so duplicate check routes to sandbox (codigoSolicitud is timestamped)
    Call srv.Validar(sol, False, db)

    ' If we got here without Err.Raise, Validar accepted the canonical/alias.
    TryAcceptDomain = True
    GoTo AcceptClean

AcceptFail:
    errNum = Err.Number
    errDesc = Err.description
    logs(1) = logs(1) & " | errNum=" & errNum & "; desc=" & errDesc
    TryAcceptDomain = False

AcceptClean:
    Call EndSSTSession(logs)
End Function

Public Function Test_SST_Validar_Canonical_PC_Accepted() As String
    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: Solicitud with canonical PC"
    logs(1) = "2. Act: Validar through sandbox"
    logs(2) = "3. Assert: no error raised"

    If Not TryAcceptDomain(TYPE_PC, logs) Then
        Test_SST_Validar_Canonical_PC_Accepted = TestHelper.BuildJsonFail("TESTS BLOCKED or Validar rejected canonical PC (sandbox missing or new domain check missing); see logs", logs)
        Exit Function
    End If

    Test_SST_Validar_Canonical_PC_Accepted = BuildOk("canonical PC accepted", logs)
End Function

Public Function Test_SST_Validar_Canonical_CD_CA_Accepted() As String
    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: Solicitud with canonical CD_CA"
    logs(1) = "2. Act: Validar through sandbox"
    logs(2) = "3. Assert: no error raised"

    If Not TryAcceptDomain(TYPE_CD_CA, logs) Then
        Test_SST_Validar_Canonical_CD_CA_Accepted = TestHelper.BuildJsonFail("TESTS BLOCKED or Validar rejected canonical CD_CA; see logs", logs)
        Exit Function
    End If

    Test_SST_Validar_Canonical_CD_CA_Accepted = BuildOk("canonical CD_CA accepted", logs)
End Function

Public Function Test_SST_Validar_Canonical_CD_CA_SUB_Accepted() As String
    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: Solicitud with canonical CD_CA_SUB"
    logs(1) = "2. Act: Validar through sandbox"
    logs(2) = "3. Assert: no error raised"

    If Not TryAcceptDomain(TYPE_CD_CA_SUB, logs) Then
        Test_SST_Validar_Canonical_CD_CA_SUB_Accepted = TestHelper.BuildJsonFail("TESTS BLOCKED or Validar rejected canonical CD_CA_SUB; see logs", logs)
        Exit Function
    End If

    Test_SST_Validar_Canonical_CD_CA_SUB_Accepted = BuildOk("canonical CD_CA_SUB accepted", logs)
End Function

Public Function Test_SST_Validar_Canonical_PC_SUB_Accepted() As String
    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: Solicitud with canonical PC_SUB"
    logs(1) = "2. Act: Validar through sandbox"
    logs(2) = "3. Assert: no error raised"

    If Not TryAcceptDomain(TYPE_PC_SUB, logs) Then
        Test_SST_Validar_Canonical_PC_SUB_Accepted = TestHelper.BuildJsonFail("TESTS BLOCKED or Validar rejected canonical PC_SUB; see logs", logs)
        Exit Function
    End If

    Test_SST_Validar_Canonical_PC_SUB_Accepted = BuildOk("canonical PC_SUB accepted", logs)
End Function

Public Function Test_SST_Validar_Alias_PCSUB_Accepted() As String
    ' Evidence (task 1.0 audit): PCSUB is already accepted by
    '   SolicitudServicio.getNombreAmigableTipoSolicitud
    '   SnapshotHelper (3 sites) as a value for snapshot key mapping
    ' Rejecting it would create a contract gap between data layer and service.
    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: Solicitud with legacy alias PCSUB"
    logs(1) = "2. Act: Validar through sandbox"
    logs(2) = "3. Assert: no error raised (alias stays accepted)"

    If Not TryAcceptDomain(TYPE_PCSUB_ALIAS, logs) Then
        Test_SST_Validar_Alias_PCSUB_Accepted = TestHelper.BuildJsonFail("TESTS BLOCKED or Validar rejected alias PCSUB; see logs", logs)
        Exit Function
    End If

    Test_SST_Validar_Alias_PCSUB_Accepted = BuildOk("alias PCSUB accepted (evidence-gated)", logs)
End Function

Public Function Test_SST_Validar_Alias_CDCASUB_Accepted() As String
    ' Evidence (task 1.0 audit): CDCASUB is already accepted by
    '   SolicitudServicio.getNombreAmigableTipoSolicitud
    '   SnapshotHelper (3 sites) as a value for snapshot key mapping
    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: Solicitud with legacy alias CDCASUB"
    logs(1) = "2. Act: Validar through sandbox"
    logs(2) = "3. Assert: no error raised (alias stays accepted)"

    If Not TryAcceptDomain(TYPE_CDCASUB_ALIAS, logs) Then
        Test_SST_Validar_Alias_CDCASUB_Accepted = TestHelper.BuildJsonFail("TESTS BLOCKED or Validar rejected alias CDCASUB; see logs", logs)
        Exit Function
    End If

    Test_SST_Validar_Alias_CDCASUB_Accepted = BuildOk("alias CDCASUB accepted (evidence-gated)", logs)
End Function
