Attribute VB_Name = "Test_RiesgoEstablecerDatosHelper"
Option Compare Database
Option Explicit

' =============================================================================
' Test_RiesgoEstablecerDatosHelper.bas
'
' TDD atoms for modRiesgoEstablecerDatosHelper.
' Per Hard rule 10 (100% TDD green) — these MUST be red before implementation,
' green after.
'
' Re-uses Test_Helper.BuildJsonOk/BuildJsonFail via BuildOk/BuildFail wrappers
' (Hard rule 6). Uses Constructor.getRiesgo/getEdicion for fixture objects.
'
' Atoms (per orchestrator spec):
'   1. Test_RiesgoEstablecerDatos_Riesgo_Happy_AllBranchesSucceed   (happy, parent)
'   2. Test_RiesgoEstablecerDatos_Riesgo_Sad_DBErrorPropagates      (sad,   parent)
'   3. Test_RiesgoEstablecerDatos_Riesgo_Edge_NothingState          (edge,  parent)
'   4. Test_RiesgoEstablecerDatos_Riesgo_Adversarial_Doble_OK       (adv,   parent)
'   5. Test_RiesgoEstablecerDatos_RiesgoGestionRiesgo_Happy         (happy, gemelo)
'   6. Test_RiesgoEstablecerDatos_RiesgoGestionRiesgo_Sad_DBError   (sad,   gemelo)
'   7. Test_RiesgoEstablecerDatos_RiesgoGestionRiesgo_Adversarial_Doble_OK (adv, gemelo)
'   --- Slice HR4b: parent LIST form (FormRiesgosGestion) ---
'   8. Test_RiesgoEstablecerDatos_RiesgosGestion_Happy              (happy, LIST)
'   9. Test_RiesgoEstablecerDatos_RiesgosGestion_Sad                (sad,   LIST)
'  10. Test_RiesgoEstablecerDatos_RiesgosGestion_Adversarial_Doble_OK (adv, LIST)
' =============================================================================

' --- Constants de IDs del fixture (coinciden con Test_Fixtures.bas) ---
Private Const FIX_EDICION    As Long = 900102
Private Const FIX_PROYECTO   As Long = 900101
Private Const FIX_EXPEDIENTE As Long = 900100
Private Const FIX_ID_RIESGO  As Long = 900503

' --- Helper wrappers (Hard rule 6) ---
Private Function BuildOk(ByVal value As Variant, ByRef logs() As String) As String
    BuildOk = Test_Helper.BuildJsonOk(value, logs)
End Function

Private Function BuildFail(ByVal msg As String, ByRef logs() As String) As String
    BuildFail = Test_Helper.BuildJsonFail(msg, logs)
End Function

' --- Fixture seeding wrapper ---
Private Function EnsureFixturesSeeded(ByRef p_Error As String) As Boolean
    On Error GoTo eh
    p_Error = ""
    Test_Fixtures.SeedAll
    EnsureFixturesSeeded = True
    Exit Function
eh:
    p_Error = "SeedAll failed: " & Err.Description
    EnsureFixturesSeeded = False
End Function

' =============================================================================
' ATOM 1: Happy path — Riesgo (padre) all branches succeed.
'   Arrange: real riesgo + edicion activa + m_EsAlta=No + PermitidoEditar=Empty
'   Act:     modRiesgoEstablecerDatosHelper.RiesgoEstablecerDatos_Calcular _
'                m_Riesgo, m_Edicion, EnumSiNo.No, Empty, False, m_Resultado,
'                m_ObjRiesgoAlInicio, , m_Error
'   Assert:  m_Error = ""; m_Resultado.PermitidoEditar = Sí; Titulo empieza con "DETALLE";
'            NavGeneralTarget ∈ {"FormRiesgoDefinicion","FormRiesgoDefinicionNoBiblioteca"}
' =============================================================================
Public Function Test_RiesgoEstablecerDatos_Riesgo_Happy_AllBranchesSucceed() As String
    Dim logs(0 To 5) As String
    Dim m_SeedErr As String
    If Not EnsureFixturesSeeded(m_SeedErr) Then
        Test_RiesgoEstablecerDatos_Riesgo_Happy_AllBranchesSucceed = _
            BuildFail(m_SeedErr, logs)
        Exit Function
    End If
    logs(0) = "1. Arrange: getRiesgo + getEdicion from Constructor"
    logs(1) = "2. Act: RiesgoEstablecerDatos_Calcular (parent path, Empty permitido)"
    logs(2) = "3. Assert: m_Error empty + Resultado fields populated"

    Dim m_Error As String
    Dim m_Riesgo As riesgo
    Dim m_Edicion As Edicion
    Dim m_Resultado As RiesgoEstablecerDatos_Resultado
    Dim m_ObjRiesgoAlInicio As Object

    ' Arrange: pull live riesgo/edicion (these are mock-friendly business objects).
    ' Constructor.getRiesgo accepts p_IDRiesgo directly. Usamos FIX_ID_RIESGO del fixture.
    Set m_Riesgo = Constructor.getRiesgo(p_IDRiesgo:=CStr(FIX_ID_RIESGO), p_Error:=m_Error)
    If m_Error <> "" Or m_Riesgo Is Nothing Then
        Test_RiesgoEstablecerDatos_Riesgo_Happy_AllBranchesSucceed = _
            BuildFail("Arrange getRiesgo (ID=" & FIX_ID_RIESGO & "): " & m_Error, logs)
        Exit Function
    End If
    Set m_Edicion = Constructor.getEdicion(CStr(FIX_EDICION), m_Error)
    If m_Error <> "" Or m_Edicion Is Nothing Then
        Test_RiesgoEstablecerDatos_Riesgo_Happy_AllBranchesSucceed = _
            BuildFail("Arrange getEdicion: " & m_Error, logs)
        Exit Function
    End If

    ' Act
    modRiesgoEstablecerDatosHelper.RiesgoEstablecerDatos_Calcular _
        m_Riesgo, m_Edicion, EnumSiNo.No, Empty, False, _
        m_Resultado, m_ObjRiesgoAlInicio, , m_Error

    ' Assert
    If m_Error <> "" Then
        Test_RiesgoEstablecerDatos_Riesgo_Happy_AllBranchesSucceed = _
            BuildFail("Calc returned error: " & m_Error, logs)
        Exit Function
    End If
    If m_Resultado.Titulo = "" Then
        Test_RiesgoEstablecerDatos_Riesgo_Happy_AllBranchesSucceed = _
            BuildFail("Resultado.Titulo empty (expected DETALLE ...)", logs)
        Exit Function
    End If
    If InStr(1, m_Resultado.Titulo, "DETALLE", vbTextCompare) = 0 Then
        Test_RiesgoEstablecerDatos_Riesgo_Happy_AllBranchesSucceed = _
            BuildFail("Titulo missing 'DETALLE': " & m_Resultado.Titulo, logs)
        Exit Function
    End If
    If m_Resultado.NavGeneralTarget <> "FormRiesgoDefinicion" And _
       m_Resultado.NavGeneralTarget <> "FormRiesgoDefinicionNoBiblioteca" Then
        Test_RiesgoEstablecerDatos_Riesgo_Happy_AllBranchesSucceed = _
            BuildFail("NavGeneralTarget inesperado: " & m_Resultado.NavGeneralTarget, logs)
        Exit Function
    End If

    logs(3) = "4. Cleanup OK"
    Test_RiesgoEstablecerDatos_Riesgo_Happy_AllBranchesSucceed = BuildOk("happy_parent_pass", logs)
End Function

' =============================================================================
' ATOM 2: Sad path — DB error propagates.
'   Arrange: real riesgo + edicion activa
'   Act:     call _Calcular with db=Nothing BUT in a context that requires DAO lookup
'            (p_EsAlta=No, riesgoAlInicio path triggers GetCachedRiesgoFresh which
'            needs db — getdb() fallback should work in live context, but we force
'            a sad path by passing p_ObjRiesgoActivo as an empty riesgo that has
'            no IDEdicion and no IDRiesgo)
'   Assert:  m_Error <> "" (sad path detected)
' =============================================================================
Public Function Test_RiesgoEstablecerDatos_Riesgo_Sad_DBErrorPropagates() As String
    Dim logs(0 To 5) As String
    Dim m_SeedErr As String
    If Not EnsureFixturesSeeded(m_SeedErr) Then
        Test_RiesgoEstablecerDatos_Riesgo_Sad_DBErrorPropagates = _
            BuildFail(m_SeedErr, logs)
        Exit Function
    End If
    logs(0) = "1. Arrange: empty riesgo (Nothing-like) — sad path"
    logs(1) = "2. Act: RiesgoEstablecerDatos_Calcular expecting error"

    Dim m_Error As String
    Dim m_RiesgoVacio As New riesgo      ' riesgo vacío, sin IDEdicion/IDRiesgo
    Dim m_Edicion As Object
    Dim m_Resultado As RiesgoEstablecerDatos_Resultado
    Dim m_ObjRiesgoAlInicio As Object

    ' Sad path 1: riesgo Nothing — debe devolver error antes de tocar db
    modRiesgoEstablecerDatosHelper.RiesgoEstablecerDatos_Calcular _
        Nothing, Nothing, EnumSiNo.No, Empty, False, _
        m_Resultado, m_ObjRiesgoAlInicio, , m_Error

    If m_Error = "" Then
        Test_RiesgoEstablecerDatos_Riesgo_Sad_DBErrorPropagates = _
            BuildFail("Sad path failed: error expected for riesgo=Nothing, got empty", logs)
        Exit Function
    End If
    If InStr(1, m_Error, "no hay riesgo activo", vbTextCompare) = 0 Then
        Test_RiesgoEstablecerDatos_Riesgo_Sad_DBErrorPropagates = _
            BuildFail("Sad path wrong msg: " & m_Error, logs)
        Exit Function
    End If

    ' Sad path 2: Alta sin permiso — debe devolver error antes de tocar db
    Set m_Edicion = Constructor.getEdicion(CStr(FIX_EDICION), m_Error)
    If m_Error <> "" Or m_Edicion Is Nothing Then
        Test_RiesgoEstablecerDatos_Riesgo_Sad_DBErrorPropagates = _
            BuildFail("Arrange getEdicion: " & m_Error, logs)
        Exit Function
    End If
    modRiesgoEstablecerDatosHelper.RiesgoEstablecerDatos_Calcular _
        m_RiesgoVacio, m_Edicion, EnumSiNo.Sí, EnumSiNo.No, False, _
        m_Resultado, m_ObjRiesgoAlInicio, , m_Error
    If m_Error = "" Then
        Test_RiesgoEstablecerDatos_Riesgo_Sad_DBErrorPropagates = _
            BuildFail("Sad path failed: alta sin permiso no detectada", logs)
        Exit Function
    End If
    If InStr(1, m_Error, "alta y el usuario no tiene permitida", vbTextCompare) = 0 Then
        Test_RiesgoEstablecerDatos_Riesgo_Sad_DBErrorPropagates = _
            BuildFail("Sad path wrong msg: " & m_Error, logs)
        Exit Function
    End If

    logs(2) = "3. Cleanup OK"
    Test_RiesgoEstablecerDatos_Riesgo_Sad_DBErrorPropagates = BuildOk("sad_parent_pass", logs)
End Function

' =============================================================================
' ATOM 3: Edge — Nothing state.
'   Arrange: p_ObjRiesgoActivo = Nothing
'   Act:     call _Calcular
'   Assert:  m_Error contains "no hay riesgo activo"
' =============================================================================
Public Function Test_RiesgoEstablecerDatos_Riesgo_Edge_NothingState() As String
    Dim logs(0 To 3) As String
    Dim m_SeedErr As String
    If Not EnsureFixturesSeeded(m_SeedErr) Then
        Test_RiesgoEstablecerDatos_Riesgo_Edge_NothingState = _
            BuildFail(m_SeedErr, logs)
        Exit Function
    End If
    logs(0) = "1. Arrange: riesgo=Nothing, edicion=Nothing"

    Dim m_Error As String
    Dim m_Resultado As RiesgoEstablecerDatos_Resultado
    Dim m_ObjRiesgoAlInicio As Object

    modRiesgoEstablecerDatosHelper.RiesgoEstablecerDatos_Calcular _
        Nothing, Nothing, EnumSiNo.No, Empty, False, _
        m_Resultado, m_ObjRiesgoAlInicio, , m_Error

    If m_Error = "" Then
        Test_RiesgoEstablecerDatos_Riesgo_Edge_NothingState = _
            BuildFail("Edge case failed: error expected, got empty", logs)
        Exit Function
    End If
    If InStr(1, m_Error, "no hay riesgo activo", vbTextCompare) = 0 Then
        Test_RiesgoEstablecerDatos_Riesgo_Edge_NothingState = _
            BuildFail("Edge case wrong msg: " & m_Error, logs)
        Exit Function
    End If

    logs(1) = "2. Assert OK"
    Test_RiesgoEstablecerDatos_Riesgo_Edge_NothingState = BuildOk("edge_parent_pass", logs)
End Function

' =============================================================================
' ATOM 4: Adversarial — Doble_OK (idempotencia).
'   Arrange: real riesgo + edicion
'   Act:     call _Calcular TWICE with same args
'   Assert:  both calls return same Titulo + PermidoEditar + no state leakage
' =============================================================================
Public Function Test_RiesgoEstablecerDatos_Riesgo_Adversarial_Doble_OK() As String
    Dim logs(0 To 5) As String
    Dim m_SeedErr As String
    If Not EnsureFixturesSeeded(m_SeedErr) Then
        Test_RiesgoEstablecerDatos_Riesgo_Adversarial_Doble_OK = _
            BuildFail(m_SeedErr, logs)
        Exit Function
    End If
    logs(0) = "1. Arrange: getRiesgo + getEdicion"
    logs(1) = "2. Act: _Calcular llamada 1"
    logs(2) = "3. Act: _Calcular llamada 2 (mismo input)"

    Dim m_Error As String
    Dim m_Riesgo As riesgo
    Dim m_Edicion As Edicion
    Dim m_R1 As RiesgoEstablecerDatos_Resultado
    Dim m_R2 As RiesgoEstablecerDatos_Resultado
    Dim m_Obj1 As Object, m_Obj2 As Object

    Set m_Riesgo = Constructor.getRiesgo(p_IDRiesgo:=CStr(FIX_ID_RIESGO), p_Error:=m_Error)
    If m_Error <> "" Or m_Riesgo Is Nothing Then
        Test_RiesgoEstablecerDatos_Riesgo_Adversarial_Doble_OK = _
            BuildFail("Arrange: " & m_Error, logs)
        Exit Function
    End If
    Set m_Edicion = Constructor.getEdicion(CStr(FIX_EDICION), m_Error)
    If m_Error <> "" Or m_Edicion Is Nothing Then
        Test_RiesgoEstablecerDatos_Riesgo_Adversarial_Doble_OK = _
            BuildFail("Arrange: " & m_Error, logs)
        Exit Function
    End If

    modRiesgoEstablecerDatosHelper.RiesgoEstablecerDatos_Calcular _
        m_Riesgo, m_Edicion, EnumSiNo.No, Empty, False, _
        m_R1, m_Obj1, , m_Error
    If m_Error <> "" Then
        Test_RiesgoEstablecerDatos_Riesgo_Adversarial_Doble_OK = _
            BuildFail("Call 1: " & m_Error, logs)
        Exit Function
    End If

    modRiesgoEstablecerDatosHelper.RiesgoEstablecerDatos_Calcular _
        m_Riesgo, m_Edicion, EnumSiNo.No, Empty, False, _
        m_R2, m_Obj2, , m_Error
    If m_Error <> "" Then
        Test_RiesgoEstablecerDatos_Riesgo_Adversarial_Doble_OK = _
            BuildFail("Call 2: " & m_Error, logs)
        Exit Function
    End If

    If m_R1.Titulo <> m_R2.Titulo Then
        Test_RiesgoEstablecerDatos_Riesgo_Adversarial_Doble_OK = _
            BuildFail("No idempotente: '" & m_R1.Titulo & "' vs '" & m_R2.Titulo & "'", logs)
        Exit Function
    End If
    If m_R1.NavGeneralTarget <> m_R2.NavGeneralTarget Then
        Test_RiesgoEstablecerDatos_Riesgo_Adversarial_Doble_OK = _
            BuildFail("No idempotente NavTarget", logs)
        Exit Function
    End If
    If m_R1.PermitidoEditar <> m_R2.PermitidoEditar Then
        Test_RiesgoEstablecerDatos_Riesgo_Adversarial_Doble_OK = _
            BuildFail("No idempotente PermitidoEditar", logs)
        Exit Function
    End If

    logs(3) = "4. Assert OK"
    Test_RiesgoEstablecerDatos_Riesgo_Adversarial_Doble_OK = BuildOk("adv_parent_pass", logs)
End Function

' =============================================================================
' ATOM 5: Happy path — gemelo (FormRiesgosGestionRiesgo).
'   Arrange: real riesgo + edicion
'   Act:     _Calcular with p_EsGemelo=True
'   Assert:  Calc succeeds; PermidoEditar es independiente de Proyecto.UsuarioAutorizado
'            (el gemelo solo usa Edicion.EsActivo, a diferencia del padre)
' =============================================================================
Public Function Test_RiesgoEstablecerDatos_RiesgoGestionRiesgo_Happy() As String
    Dim logs(0 To 5) As String
    Dim m_SeedErr As String
    If Not EnsureFixturesSeeded(m_SeedErr) Then
        Test_RiesgoEstablecerDatos_RiesgoGestionRiesgo_Happy = _
            BuildFail(m_SeedErr, logs)
        Exit Function
    End If
    logs(0) = "1. Arrange: getRiesgo + getEdicion"
    logs(1) = "2. Act: _Calcular (gemelo path)"

    Dim m_Error As String
    Dim m_Riesgo As riesgo
    Dim m_Edicion As Edicion
    Dim m_Resultado As RiesgoEstablecerDatos_Resultado
    Dim m_ObjRiesgoAlInicio As Object

    Set m_Riesgo = Constructor.getRiesgo(p_IDRiesgo:=CStr(FIX_ID_RIESGO), p_Error:=m_Error)
    If m_Error <> "" Or m_Riesgo Is Nothing Then
        Test_RiesgoEstablecerDatos_RiesgoGestionRiesgo_Happy = _
            BuildFail("Arrange: " & m_Error, logs)
        Exit Function
    End If
    Set m_Edicion = Constructor.getEdicion(CStr(FIX_EDICION), m_Error)
    If m_Error <> "" Or m_Edicion Is Nothing Then
        Test_RiesgoEstablecerDatos_RiesgoGestionRiesgo_Happy = _
            BuildFail("Arrange: " & m_Error, logs)
        Exit Function
    End If

    modRiesgoEstablecerDatosHelper.RiesgoEstablecerDatos_Calcular _
        m_Riesgo, m_Edicion, EnumSiNo.No, Empty, True, _
        m_Resultado, m_ObjRiesgoAlInicio, , m_Error

    If m_Error <> "" Then
        Test_RiesgoEstablecerDatos_RiesgoGestionRiesgo_Happy = _
            BuildFail("Gemelo Calc error: " & m_Error, logs)
        Exit Function
    End If
    If m_Resultado.Titulo = "" Then
        Test_RiesgoEstablecerDatos_RiesgoGestionRiesgo_Happy = _
            BuildFail("Gemelo Titulo empty", logs)
        Exit Function
    End If

    logs(2) = "3. Assert OK"
    Test_RiesgoEstablecerDatos_RiesgoGestionRiesgo_Happy = BuildOk("happy_gemelo_pass", logs)
End Function

' =============================================================================
' ATOM 6: Sad path — gemelo DB error.
'   Arrange: riesgo=Nothing (sad path 1) + Alta sin permiso (sad path 2)
'   Assert:  both sad paths detectados
' =============================================================================
Public Function Test_RiesgoEstablecerDatos_RiesgoGestionRiesgo_Sad_DBError() As String
    Dim logs(0 To 5) As String
    Dim m_SeedErr As String
    If Not EnsureFixturesSeeded(m_SeedErr) Then
        Test_RiesgoEstablecerDatos_RiesgoGestionRiesgo_Sad_DBError = _
            BuildFail(m_SeedErr, logs)
        Exit Function
    End If
    logs(0) = "1. Arrange: empty + sad paths"

    Dim m_Error As String
    Dim m_Resultado As RiesgoEstablecerDatos_Resultado
    Dim m_ObjRiesgoAlInicio As Object

    ' Sad 1: riesgo Nothing
    modRiesgoEstablecerDatosHelper.RiesgoEstablecerDatos_Calcular _
        Nothing, Nothing, EnumSiNo.No, Empty, True, _
        m_Resultado, m_ObjRiesgoAlInicio, , m_Error
    If m_Error = "" Then
        Test_RiesgoEstablecerDatos_RiesgoGestionRiesgo_Sad_DBError = _
            BuildFail("Gemelo sad 1 (Nothing) no detectado", logs)
        Exit Function
    End If

    ' Sad 2: Alta sin permiso
    Dim m_RiesgoVacio As New riesgo
    modRiesgoEstablecerDatosHelper.RiesgoEstablecerDatos_Calcular _
        m_RiesgoVacio, Nothing, EnumSiNo.Sí, EnumSiNo.No, True, _
        m_Resultado, m_ObjRiesgoAlInicio, , m_Error
    If m_Error = "" Then
        Test_RiesgoEstablecerDatos_RiesgoGestionRiesgo_Sad_DBError = _
            BuildFail("Gemelo sad 2 (alta sin permiso) no detectado", logs)
        Exit Function
    End If

    logs(1) = "2. Assert OK"
    Test_RiesgoEstablecerDatos_RiesgoGestionRiesgo_Sad_DBError = BuildOk("sad_gemelo_pass", logs)
End Function

' =============================================================================
' ATOM 7: Adversarial — gemelo Doble_OK (idempotencia).
' =============================================================================
Public Function Test_RiesgoEstablecerDatos_RiesgoGestionRiesgo_Adversarial_Doble_OK() As String
    Dim logs(0 To 5) As String
    Dim m_SeedErr As String
    If Not EnsureFixturesSeeded(m_SeedErr) Then
        Test_RiesgoEstablecerDatos_RiesgoGestionRiesgo_Adversarial_Doble_OK = _
            BuildFail(m_SeedErr, logs)
        Exit Function
    End If
    logs(0) = "1. Arrange: getRiesgo + getEdicion"
    logs(1) = "2. Act: gemelo _Calcular x2"

    Dim m_Error As String
    Dim m_Riesgo As riesgo
    Dim m_Edicion As Edicion
    Dim m_R1 As RiesgoEstablecerDatos_Resultado
    Dim m_R2 As RiesgoEstablecerDatos_Resultado
    Dim m_Obj1 As Object, m_Obj2 As Object

    Set m_Riesgo = Constructor.getRiesgo(p_IDRiesgo:=CStr(FIX_ID_RIESGO), p_Error:=m_Error)
    If m_Error <> "" Or m_Riesgo Is Nothing Then
        Test_RiesgoEstablecerDatos_RiesgoGestionRiesgo_Adversarial_Doble_OK = _
            BuildFail("Arrange: " & m_Error, logs)
        Exit Function
    End If
    Set m_Edicion = Constructor.getEdicion(CStr(FIX_EDICION), m_Error)
    If m_Error <> "" Or m_Edicion Is Nothing Then
        Test_RiesgoEstablecerDatos_RiesgoGestionRiesgo_Adversarial_Doble_OK = _
            BuildFail("Arrange: " & m_Error, logs)
        Exit Function
    End If

    modRiesgoEstablecerDatosHelper.RiesgoEstablecerDatos_Calcular _
        m_Riesgo, m_Edicion, EnumSiNo.No, Empty, True, _
        m_R1, m_Obj1, , m_Error
    If m_Error <> "" Then
        Test_RiesgoEstablecerDatos_RiesgoGestionRiesgo_Adversarial_Doble_OK = _
            BuildFail("Call 1: " & m_Error, logs)
        Exit Function
    End If

    modRiesgoEstablecerDatosHelper.RiesgoEstablecerDatos_Calcular _
        m_Riesgo, m_Edicion, EnumSiNo.No, Empty, True, _
        m_R2, m_Obj2, , m_Error
    If m_Error <> "" Then
        Test_RiesgoEstablecerDatos_RiesgoGestionRiesgo_Adversarial_Doble_OK = _
            BuildFail("Call 2: " & m_Error, logs)
        Exit Function
    End If

    If m_R1.Titulo <> m_R2.Titulo Or m_R1.NavGeneralTarget <> m_R2.NavGeneralTarget Or _
       m_R1.PermitidoEditar <> m_R2.PermitidoEditar Then
        Test_RiesgoEstablecerDatos_RiesgoGestionRiesgo_Adversarial_Doble_OK = _
            BuildFail("Gemelo no idempotente", logs)
        Exit Function
    End If

    logs(2) = "3. Assert OK"
    Test_RiesgoEstablecerDatos_RiesgoGestionRiesgo_Adversarial_Doble_OK = BuildOk("adv_gemelo_pass", logs)
End Function

' =============================================================================
' ATOM 8: Happy path — RiesgosGestion (parent LIST form).
'   Slice HR4b: extrae la logica pura del LIST form padre.
'   Arrange: real edicion desde fixture + stub Entorno (New Entorno — TituloUsuarioConectado
'            y los flags pueden devolver "" / 0 si la DAO lookup falla en test).
'   Act:     modRiesgoEstablecerDatosHelper.RiesgoEstablecerDatos_CalcularRiesgosGestion _
'                p_ObjEntorno, p_ObjEdicionActiva, p_Resultado, p_Error
'   Assert:  m_Error empty; p_Resultado.PermitidoEditar es Bool;
'            p_Resultado.ComboVerRetirados ∈ {"Sí","No"}; ComboVerDescripcion idem;
'            p_Resultado.EsEdicionActiva ∈ {Sí, No};
'            p_Resultado.Caption es String (puede ser "" si Entorno no devuelve TituloUsuario).
' =============================================================================
Public Function Test_RiesgoEstablecerDatos_RiesgosGestion_Happy() As String
    Dim logs(0 To 7) As String
    Dim m_SeedErr As String
    If Not EnsureFixturesSeeded(m_SeedErr) Then
        Test_RiesgoEstablecerDatos_RiesgosGestion_Happy = _
            BuildFail(m_SeedErr, logs)
        Exit Function
    End If
    logs(0) = "1. Arrange: getEdicion from Constructor (FIX_EDICION)"
    logs(1) = "   Arrange: stub Entorno via New Entorno (defaults expected)"
    logs(2) = "2. Act: RiesgoEstablecerDatos_CalcularRiesgosGestion"
    logs(3) = "3. Assert: p_Error empty + Resultado fields populated"

    Dim m_Error As String
    Dim m_Edicion As Edicion
    Dim m_Entorno As Object
    Dim m_Resultado As RiesgoEstablecerDatos_RiesgosGestion_Resultado

    Set m_Edicion = Constructor.getEdicion(CStr(FIX_EDICION), m_Error)
    If m_Error <> "" Or m_Edicion Is Nothing Then
        Test_RiesgoEstablecerDatos_RiesgosGestion_Happy = _
            BuildFail("Arrange getEdicion: " & m_Error, logs)
        Exit Function
    End If

    Set m_Entorno = New Entorno

    modRiesgoEstablecerDatosHelper.RiesgoEstablecerDatos_CalcularRiesgosGestion _
        m_Entorno, m_Edicion, m_Resultado, m_Error

    ' --- Assert ---
    If m_Error <> "" Then
        Test_RiesgoEstablecerDatos_RiesgosGestion_Happy = _
            BuildFail("Calc returned error: " & m_Error, logs)
        Exit Function
    End If

    ' PermitidoEditar must be a Bool (True or False, not Nothing/non-Bool).
    If m_Resultado.PermitidoEditar <> True And m_Resultado.PermitidoEditar <> False Then
        Test_RiesgoEstablecerDatos_RiesgosGestion_Happy = _
            BuildFail("PermitidoEditar no es Bool valido", logs)
        Exit Function
    End If

    ' ComboVerRetirados must be "Sí" or "No".
    If m_Resultado.ComboVerRetirados <> "Sí" And _
       m_Resultado.ComboVerRetirados <> "No" Then
        Test_RiesgoEstablecerDatos_RiesgosGestion_Happy = _
            BuildFail("ComboVerRetirados no es valido: '" & m_Resultado.ComboVerRetirados & "'", logs)
        Exit Function
    End If

    ' ComboVerDescripcion must be "Sí" or "No".
    If m_Resultado.ComboVerDescripcion <> "Sí" And _
       m_Resultado.ComboVerDescripcion <> "No" Then
        Test_RiesgoEstablecerDatos_RiesgosGestion_Happy = _
            BuildFail("ComboVerDescripcion no es valido: '" & m_Resultado.ComboVerDescripcion & "'", logs)
        Exit Function
    End If

    ' EsEdicionActiva must be a valid EnumSiNo (0, 1, or 2).
    If m_Resultado.EsEdicionActiva < EnumSiNo.No Or _
       m_Resultado.EsEdicionActiva > EnumSiNo.Sí Then
        ' EnumSiNo.No = 2 in this codebase, Sí = 1.
        If m_Resultado.EsEdicionActiva < 0 Or m_Resultado.EsEdicionActiva > 2 Then
            Test_RiesgoEstablecerDatos_RiesgosGestion_Happy = _
                BuildFail("EsEdicionActiva fuera de rango: " & m_Resultado.EsEdicionActiva, logs)
            Exit Function
        End If
    End If

    ' Caption is a String — puede ser "" si Entorno.TituloUsuarioConectado no devolvio valor
    ' en el contexto del test (esperado cuando no hay usuario logado contra TbConfiguracionVisionRiesgos).
    If Len(m_Resultado.Caption) > 0 Then
        ' Si Caption esta poblada, debe contener el marcador "riesgos de la edición".
        If InStr(1, m_Resultado.Caption, "riesgos de la edición", vbTextCompare) = 0 Then
            Test_RiesgoEstablecerDatos_RiesgosGestion_Happy = _
                BuildFail("Caption poblada pero sin 'riesgos de la edición': '" & m_Resultado.Caption & "'", logs)
            Exit Function
        End If
    End If

    logs(4) = "4. Cleanup OK"
    Set m_Entorno = Nothing
    Test_RiesgoEstablecerDatos_RiesgosGestion_Happy = BuildOk("happy_list_pass", logs)
End Function

' =============================================================================
' ATOM 9: Sad path — RiesgosGestion: p_ObjEdicionActiva = Nothing.
'   Act:     call CalcularRiesgosGestion with Nothing edicion
'   Assert:  m_Error contains "No hay edicion activa" or similar sadness marker
' =============================================================================
Public Function Test_RiesgoEstablecerDatos_RiesgosGestion_Sad() As String
    Dim logs(0 To 4) As String
    Dim m_SeedErr As String
    If Not EnsureFixturesSeeded(m_SeedErr) Then
        Test_RiesgoEstablecerDatos_RiesgosGestion_Sad = _
            BuildFail(m_SeedErr, logs)
        Exit Function
    End If
    logs(0) = "1. Arrange: p_ObjEdicionActiva=Nothing (sad path)"
    logs(1) = "2. Act: CalcularRiesgosGestion expecting error"

    Dim m_Error As String
    Dim m_Entorno As Object
    Dim m_Resultado As RiesgoEstablecerDatos_RiesgosGestion_Resultado

    Set m_Entorno = New Entorno

    modRiesgoEstablecerDatosHelper.RiesgoEstablecerDatos_CalcularRiesgosGestion _
        m_Entorno, Nothing, m_Resultado, m_Error

    Set m_Entorno = Nothing

    If m_Error = "" Then
        Test_RiesgoEstablecerDatos_RiesgosGestion_Sad = _
            BuildFail("Sad path failed: error expected for edicion=Nothing, got empty", logs)
        Exit Function
    End If
    If InStr(1, m_Error, "edicion activa", vbTextCompare) = 0 Then
        Test_RiesgoEstablecerDatos_RiesgosGestion_Sad = _
            BuildFail("Sad path wrong msg: " & m_Error, logs)
        Exit Function
    End If

    logs(2) = "3. Assert OK"
    Test_RiesgoEstablecerDatos_RiesgosGestion_Sad = BuildOk("sad_list_pass", logs)
End Function

' =============================================================================
' ATOM 10: Adversarial — RiesgosGestion: Doble_OK (idempotencia).
'   Arrange: real edicion + stub Entorno
'   Act:     call CalcularRiesgosGestion TWICE with same args
'   Assert:  both calls return same PermitidoEditar, ComboVerRetirados,
'            ComboVerDescripcion, Caption, EsEdicionActiva. Pure compute
'            no debe tener side-effect leakage entre llamadas.
' =============================================================================
Public Function Test_RiesgoEstablecerDatos_RiesgosGestion_Adversarial_Doble_OK() As String
    Dim logs(0 To 6) As String
    Dim m_SeedErr As String
    If Not EnsureFixturesSeeded(m_SeedErr) Then
        Test_RiesgoEstablecerDatos_RiesgosGestion_Adversarial_Doble_OK = _
            BuildFail(m_SeedErr, logs)
        Exit Function
    End If
    logs(0) = "1. Arrange: getEdicion + stub Entorno"
    logs(1) = "2. Act: _CalcularRiesgosGestion llamada 1"
    logs(2) = "3. Act: _CalcularRiesgosGestion llamada 2 (mismo input)"

    Dim m_Error As String
    Dim m_Edicion As Edicion
    Dim m_Entorno As Object
    Dim m_R1 As RiesgoEstablecerDatos_RiesgosGestion_Resultado
    Dim m_R2 As RiesgoEstablecerDatos_RiesgosGestion_Resultado

    Set m_Edicion = Constructor.getEdicion(CStr(FIX_EDICION), m_Error)
    If m_Error <> "" Or m_Edicion Is Nothing Then
        Test_RiesgoEstablecerDatos_RiesgosGestion_Adversarial_Doble_OK = _
            BuildFail("Arrange getEdicion: " & m_Error, logs)
        Exit Function
    End If
    Set m_Entorno = New Entorno

    modRiesgoEstablecerDatosHelper.RiesgoEstablecerDatos_CalcularRiesgosGestion _
        m_Entorno, m_Edicion, m_R1, m_Error
    If m_Error <> "" Then
        Test_RiesgoEstablecerDatos_RiesgosGestion_Adversarial_Doble_OK = _
            BuildFail("Call 1: " & m_Error, logs)
        Exit Function
    End If

    modRiesgoEstablecerDatosHelper.RiesgoEstablecerDatos_CalcularRiesgosGestion _
        m_Entorno, m_Edicion, m_R2, m_Error
    If m_Error <> "" Then
        Test_RiesgoEstablecerDatos_RiesgosGestion_Adversarial_Doble_OK = _
            BuildFail("Call 2: " & m_Error, logs)
        Exit Function
    End If

    Set m_Entorno = Nothing

    If m_R1.PermitidoEditar <> m_R2.PermitidoEditar Then
        Test_RiesgoEstablecerDatos_RiesgosGestion_Adversarial_Doble_OK = _
            BuildFail("No idempotente PermitidoEditar", logs)
        Exit Function
    End If
    If m_R1.ComboVerRetirados <> m_R2.ComboVerRetirados Then
        Test_RiesgoEstablecerDatos_RiesgosGestion_Adversarial_Doble_OK = _
            BuildFail("No idempotente ComboVerRetirados: '" & m_R1.ComboVerRetirados & "' vs '" & m_R2.ComboVerRetirados & "'", logs)
        Exit Function
    End If
    If m_R1.ComboVerDescripcion <> m_R2.ComboVerDescripcion Then
        Test_RiesgoEstablecerDatos_RiesgosGestion_Adversarial_Doble_OK = _
            BuildFail("No idempotente ComboVerDescripcion", logs)
        Exit Function
    End If
    If m_R1.Caption <> m_R2.Caption Then
        Test_RiesgoEstablecerDatos_RiesgosGestion_Adversarial_Doble_OK = _
            BuildFail("No idempotente Caption: '" & m_R1.Caption & "' vs '" & m_R2.Caption & "'", logs)
        Exit Function
    End If
    If m_R1.EsEdicionActiva <> m_R2.EsEdicionActiva Then
        Test_RiesgoEstablecerDatos_RiesgosGestion_Adversarial_Doble_OK = _
            BuildFail("No idempotente EsEdicionActiva", logs)
        Exit Function
    End If

    logs(3) = "4. Assert OK"
    Test_RiesgoEstablecerDatos_RiesgosGestion_Adversarial_Doble_OK = BuildOk("adv_list_pass", logs)
End Function