Attribute VB_Name = "Test_Subcontratistas"
Option Compare Database
Option Explicit

' ============================================================
' Tests para SuministradoresHelper — Skill access-vba-tdd v1.7
'
' Fixture: SeedAll crea grafo completo en rango FIX_ID_BASE (=900000)
'   - Expediente fixture ? Proyecto fixture ? Edición fixture
'   - TbSuministradores: hierarchy (parent + 2 children + grandchild)
'   - TbExpedientesSuministradores: vincula hierarchy al expediente
'   - TbProyectosEdicionesSuministradores: children en la edición
'
' Fixture IDs:
'   FIX_ROOT      = 900100 (ConsorcioPropio='Sí')
'   FIX_CHILD1   = 900110 (ConsorcioPropio='Sí', child of ROOT)
'   FIX_CHILD2   = 900120 (ConsorcioPropio='Sí', child of ROOT)
'   FIX_GRANDCHILD = 900111 (ConsorcioPropio='Sí', child of CHILD1)
'
' Reglas skill v1.7 aplicadas:
'   §1: Public Function + JSON String retorno, cero UI
'   §2: BuildJsonOk/Fail vía Test_Helper_JSON, EscapeJsonString
'   §3.4: GoTo Teardown en assert fallido, teardown defensivo
'   §3.5: SeedAll/TeardownAll en RunAll, no seed individual por test
'   §2 idempotencia: NO ProbeFixtureId, NO datos preexistentes
' ============================================================

' --- Constants de IDs del fixture (matching Test_Fixtures) ---
Private Const FIX_EXPEDIENTE  As Long = 900100  ' fixture expediente
Private Const FIX_PROYECTO    As Long = 900101  ' fixture proyecto
Private Const FIX_EDICION     As Long = 900102  ' fixture edición
Private Const FIX_ROOT        As Long = 900110  ' supplier raíz (ConsorcioPropio=Sí)
Private Const FIX_CHILD1      As Long = 900120  ' hijo directo de ROOT
Private Const FIX_CHILD2      As Long = 900130  ' hijo directo de ROOT
Private Const FIX_GRANDCHILD  As Long = 900121  ' nieto (child of CHILD1)

' --- Helpers JSON (delegación a Test_Helper_JSON público) ---

Private Function EscapeJsonString(ByVal s As String) As String
    EscapeJsonString = Test_Helper.EscapeJsonString(s)
End Function

Private Function BuildOk(ByVal value As Variant, ByRef logs() As String) As String
    BuildOk = Test_Helper.BuildJsonOk(value, logs)
End Function

Private Function BuildFail(ByVal msg As String, ByRef logs() As String) As String
    BuildFail = Test_Helper.BuildJsonFail(msg, logs)
End Function

' --- Helpers privados ---

Private Function KeysCsv(ByVal dic As Scripting.Dictionary) As String
    Dim key As Variant
    Dim result As String
    result = ""
    If Not dic Is Nothing Then
        For Each key In dic.keys
            If result <> "" Then result = result & ","
            result = result & CStr(key)
        Next
    End If
    KeysCsv = result
End Function

Private Function CountDict(ByVal dic As Scripting.Dictionary) As String
    If dic Is Nothing Then
        CountDict = "NULL"
    Else
        CountDict = CStr(dic.Count)
    End If
End Function

Private Function BoolStr(ByVal b As Boolean) As String
    BoolStr = IIf(b, "true", "false")
End Function

Private Function DictHas(ByVal dic As Scripting.Dictionary, ByVal key As String) As Boolean
    If dic Is Nothing Then Exit Function
    DictHas = dic.Exists(key)
End Function

Private Function JsonOk(ByVal json As String) As Boolean
    JsonOk = (InStr(json, """ok"":true") > 0)
End Function

' --- AssertLocalBackend guard (skill v2.0 §3) ---
' Verifica que el backend local de test esté configurado.
' Delegado a Test_Helper.ForceLocalBackend que hace EVE en modo LOCAL.
Private Function AssertLocalBackend(Optional ByRef p_Error As String = "") As Boolean
    Dim errMsg As String
    AssertLocalBackend = Test_Helper.ForceLocalBackend(errMsg)
    p_Error = errMsg
End Function

' ============================================================
' TEST 1: getSubcontratistasPorExpediente — excluye segundo nivel
'
' Fixture hierarchy: ROOT(900110) ? CHILD1(900120), CHILD2(900130)
'                                         ?
'                                   GRANDCHILD(900121)
'
' getSubcontratistasPorExpediente devuelve SOLO hijos directos
' (CHILD1, CHILD2) — NO el nieto (GRANDCHILD).
'
' Arrange: SeedAll ejecutó SeedSubcatGraph ? grafo completo listo
' Act: getSubcontratistasPorExpediente(con expediente_fixture)
' Assert: count=2, keys=[CHILD1, CHILD2], NO GRANDCHILD
' ============================================================
Public Function Test_Subcon_Helper_ConsorcioMasUno_ExcluyeSegundoNivel() As String
    Dim logs(0 To 3) As String
    Dim idExp As String
    Dim sError As String
    Dim dicHijos As Scripting.Dictionary
    Dim ok As Boolean
    On Error GoTo EH

    ' Arrange — SuiteSetup ya configuró el backend; verificar con guard
    If Not AssertLocalBackend() Then
        Test_Subcon_Helper_ConsorcioMasUno_ExcluyeSegundoNivel = BuildFail("TESTS BLOCKED: backend local no configurado", logs)
        Exit Function
    End If

    ' Poblar fixture de subcontratistas antes de testear
    Test_Fixtures.SeedSubcatGraph

    idExp = CStr(FIX_EXPEDIENTE)
    logs(0) = "1. Arrange: expediente fixture=" & idExp

    ' Act
    Set dicHijos = SuministradoresHelper.getSubcontratistasPorExpediente(idExp, sError)

    If sError <> "" Then
        Test_Fixtures.TeardownSubcatGraph
        Test_Subcon_Helper_ConsorcioMasUno_ExcluyeSegundoNivel = BuildFail("getSubcontratistasPorExpediente: " & sError, logs)
        Exit Function
    End If

    logs(1) = "2. Act: dicHijos.Count=" & CountDict(dicHijos)

    ' Assert: exactamente 2 hijos directos, no el nieto
    If Not dicHijos Is Nothing Then
        ok = dicHijos.Count = 2 _
            And DictHas(dicHijos, CStr(FIX_CHILD1)) _
            And DictHas(dicHijos, CStr(FIX_CHILD2)) _
            And Not DictHas(dicHijos, CStr(FIX_GRANDCHILD))
    End If

    logs(2) = "3. Assert: " & IIf(ok, "OK", "FAIL") & " keys=[" & KeysCsv(dicHijos) & "]"

    If Not ok Then
        Test_Fixtures.TeardownSubcatGraph
        Test_Subcon_Helper_ConsorcioMasUno_ExcluyeSegundoNivel = BuildFail( _
            "Count=" & CountDict(dicHijos) & " expected=2" & _
            " missing FIX_CHILD1?" & BoolStr(Not DictHas(dicHijos, CStr(FIX_CHILD1))) & _
            " missing FIX_CHILD2?" & BoolStr(Not DictHas(dicHijos, CStr(FIX_CHILD2))) & _
            " has FIX_GRANDCHILD?" & BoolStr(DictHas(dicHijos, CStr(FIX_GRANDCHILD))), logs)
        Exit Function
    End If

    Test_Fixtures.TeardownSubcatGraph
    Test_Subcon_Helper_ConsorcioMasUno_ExcluyeSegundoNivel = BuildOk(2, logs)
    Exit Function

EH:
    On Error Resume Next
    Test_Fixtures.TeardownSubcatGraph
    Test_Subcon_Helper_ConsorcioMasUno_ExcluyeSegundoNivel = BuildFail(Err.Description, logs)
End Function

' ============================================================
' TEST 2: getSuministradoresFaltantesEnEdicion — happy path
'
' Fixture: SeedSubcatGraph pobló edición con CHILD1 y CHILD2.
' Ambos son hijos directos de ROOT (ConsorcioPropio='Sí').
'
' Arrange: edición tiene [CHILD1, CHILD2] — exactamente los subcontratistas
' Act: getSuministradoresFaltantesEnEdicion
' Assert: faltan=0 (ningún subcontratista falta en edición)
' ============================================================
Public Function Test_Subcon_Helper_Diff_CaminoFeliz() As String
    Dim logs(0 To 4) As String
    On Error GoTo EH

    ' Arrange — SuiteSetup ya configuró el backend
    If Not AssertLocalBackend() Then
        Test_Subcon_Helper_Diff_CaminoFeliz = BuildFail("TESTS BLOCKED: backend local no configurado", logs)
        Exit Function
    End If

    ' Poblar fixture de subcontratistas
    Test_Fixtures.SeedSubcatGraph

    Dim idExp As String: idExp = CStr(FIX_EXPEDIENTE)
    Dim IDEdicion As String: IDEdicion = CStr(FIX_EDICION)
    logs(0) = "1. Arrange: exp=" & idExp & " edicion=" & IDEdicion & " (poblada por SeedSubcatGraph)"

    ' Act
    Dim sError As String
    Dim dicFaltan As Scripting.Dictionary

    logs(1) = "2. Act: getSuministradoresFaltantesEnEdicion"
    Set dicFaltan = SuministradoresHelper.getSuministradoresFaltantesEnEdicion(idExp, IDEdicion, sError)
    If sError <> "" Then
        Test_Fixtures.TeardownSubcatGraph
        Test_Subcon_Helper_Diff_CaminoFeliz = BuildFail("Faltantes: " & sError, logs)
        Exit Function
    End If

    logs(2) = "3. Assert: faltan=" & CountDict(dicFaltan)

    ' Assert: ningún subcontratista falta
    If Not (dicFaltan Is Nothing Or dicFaltan.Count = 0) Then
        logs(3) = "4. Fail: esperados 0 faltantes, obtenido " & CountDict(dicFaltan) & " keys=[" & KeysCsv(dicFaltan) & "]"
        Test_Fixtures.TeardownSubcatGraph
        Test_Subcon_Helper_Diff_CaminoFeliz = BuildFail("Faltan expected=0 got=" & CountDict(dicFaltan), logs)
        Exit Function
    End If

    logs(3) = "4. Assert OK: faltan=0"
    Test_Fixtures.TeardownSubcatGraph
    Test_Subcon_Helper_Diff_CaminoFeliz = BuildOk("faltan=0", logs)
    Exit Function

EH:
    On Error Resume Next
    Test_Fixtures.TeardownSubcatGraph
    Test_Subcon_Helper_Diff_CaminoFeliz = BuildFail(Err.Description, logs)
End Function

' ============================================================
' TEST 3: getSuministradoresSobrantesEnEdicion — edition con extra
'
' Fixture: SeedSubcatGraph puso CHILD1+CHILD2 en la edición.
' Agregamos ROOT (que no es subcontratista) a la edición.
'
' Arrange:SeedSubcatGraph + agregar ROOT a la edición
' Act: getSuministradoresSobrantesEnEdicion
' Assert: sobran=1 (ROOT no es subcontratista, está de más)
' Teardown: quitar ROOT de la edición
' ============================================================
Public Function Test_Subcon_Helper_Diff_CaminoTriste() As String
    Dim logs(0 To 7) As String
    Dim IDEdicion As String
    Dim idExp As String
    Dim nextRowId As Long
    Dim rs As DAO.Recordset
    Dim sError As String
    Dim dicSobran As Scripting.Dictionary
    Dim ok As Boolean
    Dim dbErr As String
    Dim db As Variant
    Dim tErr As String
    Dim dbT As Variant
    On Error GoTo EH

    ' Arrange — SuiteSetup ya configuró el backend
    If Not AssertLocalBackend() Then
        Test_Subcon_Helper_Diff_CaminoTriste = BuildFail("TESTS BLOCKED: backend local no configurado", logs)
        Exit Function
    End If

    ' Poblar fixture de subcontratistas
    Test_Fixtures.SeedSubcatGraph

    ' Obtener conexión al backend fixture
    Set db = GetTestDbSubcat(dbErr)
    If VarType(db) = vbString Then
        Test_Subcon_Helper_Diff_CaminoTriste = BuildFail("GetTestDbSubcat: " & dbErr, logs)
        GoTo Teardown
    End If
    If db Is Nothing Then
        Test_Subcon_Helper_Diff_CaminoTriste = BuildFail("GetTestDbSubcat: Nothing", logs)
        GoTo Teardown
    End If

    IDEdicion = CStr(FIX_EDICION)
    idExp = CStr(FIX_EXPEDIENTE)

    ' Agregar ROOT (900110) a la edición — no es subcontratista, así que será "sobrante"
    Set rs = db.OpenRecordset("SELECT MAX(ID) FROM TbProyectosEdicionesSuministradores", dbOpenSnapshot)
    If Not rs.EOF And Not IsNull(rs.Fields(0).value) Then nextRowId = rs.Fields(0).value + 10 Else nextRowId = 900200
    rs.Close: Set rs = Nothing

    db.Execute "INSERT INTO TbProyectosEdicionesSuministradores (ID, IDEdicion, IDSuministrador) " & _
               "VALUES (" & nextRowId & ", " & IDEdicion & ", " & FIX_ROOT & ")"
    logs(0) = "1. Arrange: ROOT(" & FIX_ROOT & ") agregado a edición " & IDEdicion

    ' Act
    logs(1) = "2. Act: getSuministradoresSobrantesEnEdicion"
    Set dicSobran = SuministradoresHelper.getSuministradoresSobrantesEnEdicion(idExp, IDEdicion, sError)
    If sError <> "" Then
        Test_Subcon_Helper_Diff_CaminoTriste = BuildFail("Sobrantes: " & sError, logs)
        GoTo Teardown
    End If

    logs(2) = "3. Assert: sobran=" & CountDict(dicSobran) & " keys=[" & KeysCsv(dicSobran) & "]"

    ' Assert: ROOT (900110) es sobrante (no es subcontratista)
    If Not dicSobran Is Nothing Then
        ok = dicSobran.Count = 1 And DictHas(dicSobran, CStr(FIX_ROOT))
    End If

    logs(3) = "4. Assert: " & IIf(ok, "OK", "FAIL")
    If Not ok Then
        Test_Subcon_Helper_Diff_CaminoTriste = BuildFail("Sobran expected=1 (ROOT=" & FIX_ROOT & ") got=" & CountDict(dicSobran) & " keys=[" & KeysCsv(dicSobran) & "]", logs)
        GoTo Teardown
    End If

    logs(4) = "5. Assert OK"
    Test_Subcon_Helper_Diff_CaminoTriste = BuildOk("sobran=1 (ROOT)", logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    ' Guard: solo borrar si IDEdicion tiene valor válido
    If IDEdicion <> "" And CLng(Nz(IDEdicion, 0)) <> 0 Then
        If VarType(db) <> vbString Then
            If Not db Is Nothing Then
                db.Execute "DELETE FROM TbProyectosEdicionesSuministradores WHERE IDEdicion=" & IDEdicion & " AND IDSuministrador=" & FIX_ROOT
            End If
        End If
    End If
    Test_Fixtures.TeardownSubcatGraph
    Set db = Nothing
    Exit Function

EH:
    Test_Subcon_Helper_Diff_CaminoTriste = BuildFail(Err.Description, logs)
    On Error Resume Next
    ' Guard defensivo en error handler
    If IDEdicion <> "" And CLng(Nz(IDEdicion, 0)) <> 0 Then
        Set dbT = GetTestDbSubcat(tErr)
        If VarType(dbT) <> vbString Then
            If Not dbT Is Nothing Then
                dbT.Execute "DELETE FROM TbProyectosEdicionesSuministradores WHERE IDEdicion=" & IDEdicion & " AND IDSuministrador=" & FIX_ROOT
                Set dbT = Nothing
            End If
        End If
    End If
    Test_Fixtures.TeardownSubcatGraph
End Function

' ============================================================
' TEST 4: getSuministradoresSobrantesEnEdicion — grandchild en edición
'
' Fixture: GRANDCHILD no está en la edición (SeedSubcatGraph no lo agregó).
' Agregamos GRANDCHILD a la edición manualmente.
'
' Arrange:SeedSubcatGraph + agregar GRANDCHILD a la edición
' Act: getSuministradoresSobrantesEnEdicion
' Assert: sobran=1 (GRANDCHILD no es subcontratista directo)
' Teardown: quitar GRANDCHILD de la edición
' ============================================================
Public Function Test_Subcon_Helper_Diff_NoIncluyeNieto() As String
    Dim logs(0 To 7) As String
    Dim IDEdicion As String
    Dim idExp As String
    Dim nextRowId As Long
    Dim rs As DAO.Recordset
    Dim sError As String
    Dim dicSobran As Scripting.Dictionary
    Dim ok As Boolean
    Dim dbErr As String
    Dim db As Variant
    Dim tErr As String
    Dim dbT As Variant
    On Error GoTo EH

    ' Arrange — SuiteSetup ya configuró el backend
    If Not AssertLocalBackend() Then
        Test_Subcon_Helper_Diff_NoIncluyeNieto = BuildFail("TESTS BLOCKED: backend local no configurado", logs)
        Exit Function
    End If

    ' Poblar fixture de subcontratistas
    Test_Fixtures.SeedSubcatGraph

    Set db = GetTestDbSubcat(dbErr)
    If VarType(db) = vbString Then
        Test_Subcon_Helper_Diff_NoIncluyeNieto = BuildFail("GetTestDbSubcat: " & dbErr, logs)
        GoTo Teardown
    End If
    If db Is Nothing Then
        Test_Subcon_Helper_Diff_NoIncluyeNieto = BuildFail("GetTestDbSubcat: Nothing", logs)
        GoTo Teardown
    End If

    IDEdicion = CStr(FIX_EDICION)
    idExp = CStr(FIX_EXPEDIENTE)

    ' Agregar GRANDCHILD a la edición (no es hijo directo de ROOT, es nieto)
    Set rs = db.OpenRecordset("SELECT MAX(ID) FROM TbProyectosEdicionesSuministradores", dbOpenSnapshot)
    If Not rs.EOF And Not IsNull(rs.Fields(0).value) Then nextRowId = rs.Fields(0).value + 10 Else nextRowId = 900210
    rs.Close: Set rs = Nothing

    db.Execute "INSERT INTO TbProyectosEdicionesSuministradores (ID, IDEdicion, IDSuministrador) " & _
               "VALUES (" & nextRowId & ", " & IDEdicion & ", " & FIX_GRANDCHILD & ")"
    logs(0) = "1. Arrange: GRANDCHILD(" & FIX_GRANDCHILD & ") agregado a edición"

    ' Act
    logs(1) = "2. Act: getSuministradoresSobrantesEnEdicion"
    Set dicSobran = SuministradoresHelper.getSuministradoresSobrantesEnEdicion(idExp, IDEdicion, sError)
    If sError <> "" Then
        Test_Subcon_Helper_Diff_NoIncluyeNieto = BuildFail("Sobrantes: " & sError, logs)
        GoTo Teardown
    End If

    logs(2) = "3. Assert: sobran=" & CountDict(dicSobran) & " keys=[" & KeysCsv(dicSobran) & "]"

    ' Assert: GRANDCHILD (900121) es sobrante
    If Not dicSobran Is Nothing Then
        ok = dicSobran.Count = 1 And DictHas(dicSobran, CStr(FIX_GRANDCHILD))
    End If

    logs(3) = "4. Assert: " & IIf(ok, "OK", "FAIL")
    If Not ok Then
        Test_Subcon_Helper_Diff_NoIncluyeNieto = BuildFail("Sobran expected=1 (GRANDCHILD=" & FIX_GRANDCHILD & ") got=" & CountDict(dicSobran) & " keys=[" & KeysCsv(dicSobran) & "]", logs)
        GoTo Teardown
    End If

    logs(4) = "5. Assert OK"
    Test_Subcon_Helper_Diff_NoIncluyeNieto = BuildOk("sobran=1 (GRANDCHILD)", logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    If IDEdicion <> "" And CLng(Nz(IDEdicion, 0)) <> 0 Then
        If VarType(db) <> vbString Then
            If Not db Is Nothing Then
                db.Execute "DELETE FROM TbProyectosEdicionesSuministradores WHERE IDEdicion=" & IDEdicion & " AND IDSuministrador=" & FIX_GRANDCHILD
            End If
        End If
    End If
    Test_Fixtures.TeardownSubcatGraph
    Set db = Nothing
    Exit Function

EH:
    Test_Subcon_Helper_Diff_NoIncluyeNieto = BuildFail(Err.Description, logs)
    On Error Resume Next
    If IDEdicion <> "" And CLng(Nz(IDEdicion, 0)) <> 0 Then
        Set dbT = GetTestDbSubcat(tErr)
        If VarType(dbT) <> vbString Then
            If Not dbT Is Nothing Then
                dbT.Execute "DELETE FROM TbProyectosEdicionesSuministradores WHERE IDEdicion=" & IDEdicion & " AND IDSuministrador=" & FIX_GRANDCHILD
                Set dbT = Nothing
            End If
        End If
    End If
    Test_Fixtures.TeardownSubcatGraph
End Function

' ============================================================
' TEST 5: Alta y eliminar evidencia
'
' Fixture: SeedSubcatGraph no agregó GRANDCHILD a la edición.
' Hacemos alta y verificar que aparece, luego eliminar y verificar que no existe.
'
' Arrange: GRANDCHILD no está en la edición
' Act: AltaSuministradorEvidenciaEnEdicion(GRANDCHILD)
' Assert: getSuministradorEnEdicion devuelve GRANDCHILD
' Act: EliminarSuministradorEvidenciaEnEdicion(GRANDCHILD)
' Assert: getSuministradorEnEdicion devuelve Nothing
' ============================================================
Public Function Test_Subcon_Helper_AltaYEliminarEvidencia() As String
    Dim logs(0 To 8) As String
    Dim IDEdicion As String
    Dim idExp As String
    Dim sError As String
    Dim objBefore As SuministradorParaEvidencias
    Dim objEv As SuministradorParaEvidencias
    Dim objAfter As SuministradorParaEvidencias
    Dim objDeleted As SuministradorParaEvidencias
    Dim delResult As String
    Dim dbErr As String
    Dim dbErr2 As String
    Dim tErr As String
    Dim db As Variant
    Dim dbT As Variant
    On Error GoTo EH

    ' Arrange — SuiteSetup ya configuró el backend
    If Not AssertLocalBackend() Then
        Test_Subcon_Helper_AltaYEliminarEvidencia = BuildFail("TESTS BLOCKED: backend local no configurado", logs)
        Exit Function
    End If

    ' Poblar fixture de subcontratistas
    Test_Fixtures.SeedSubcatGraph

    IDEdicion = CStr(FIX_EDICION)
    idExp = CStr(FIX_EXPEDIENTE)

    ' Verificar pre-condición: GRANDCHILD NO está en la edición
    Set objBefore = SuministradoresHelper.getSuministradorEnEdicion( _
        p_IDEdicion:=IDEdicion, p_IDSuministrador:=CStr(FIX_GRANDCHILD), p_Error:=sError)
    If Not objBefore Is Nothing Then
        logs(0) = "0. Pre: GRANDCHILD ya existe — cleanup"
        Set db = GetTestDbSubcat(dbErr)
        If VarType(db) <> vbString Then
            If Not db Is Nothing Then
                db.Execute "DELETE FROM TbProyectosEdicionesSuministradores WHERE IDEdicion=" & IDEdicion & " AND IDSuministrador=" & FIX_GRANDCHILD
                Set db = Nothing
            End If
        End If
    Else
        logs(0) = "0. Pre: GRANDCHILD no está (OK)"
    End If

    ' Act: Alta
    logs(1) = "1. Act: AltaSuministradorEvidenciaEnEdicion(GRANDCHILD)"
    Set objEv = SuministradoresHelper.AltaSuministradorEvidenciaEnEdicion(CStr(FIX_GRANDCHILD), IDEdicion, sError)
    If sError <> "" Then
        Test_Subcon_Helper_AltaYEliminarEvidencia = BuildFail("Alta: " & sError, logs)
        GoTo Teardown
    End If
    If objEv Is Nothing Then
        logs(2) = "2. Fail: objEv es Nothing"
        Test_Subcon_Helper_AltaYEliminarEvidencia = BuildFail("Alta devolvió Nothing", logs)
        GoTo Teardown
    End If
    logs(2) = "2. Alta OK: IDSuministrador=" & objEv.idSuministrador

    ' Assert: verificar que existe tras alta
    Set objAfter = SuministradoresHelper.getSuministradorEnEdicion( _
        p_IDEdicion:=IDEdicion, p_IDSuministrador:=CStr(FIX_GRANDCHILD), p_Error:=sError)
    If sError <> "" Then
        logs(3) = "3. Fail: getSuministradorEnEdicion error: " & sError
        Test_Subcon_Helper_AltaYEliminarEvidencia = BuildFail("Get tras alta: " & sError, logs)
        GoTo Teardown
    End If
    If objAfter Is Nothing Then
        logs(3) = "3. Fail: getSuministradorEnEdicion devolvió Nothing"
        Test_Subcon_Helper_AltaYEliminarEvidencia = BuildFail("GetSuministradorEnEdicion returned Nothing after alta", logs)
        GoTo Teardown
    End If
    If objAfter.idSuministrador <> CStr(FIX_GRANDCHILD) Then
        logs(3) = "3. Fail: IDSuministrador=" & objAfter.idSuministrador & " esperado=" & FIX_GRANDCHILD
        Test_Subcon_Helper_AltaYEliminarEvidencia = BuildFail("Wrong IDSuministrador: " & objAfter.idSuministrador & " expected=" & FIX_GRANDCHILD, logs)
        GoTo Teardown
    End If
    logs(3) = "3. Assert: GRANDCHILD existe en edición (OK)"

    ' Act: Eliminar
    logs(4) = "4. Act: EliminarSuministradorEvidenciaEnEdicion(GRANDCHILD)"
    delResult = SuministradoresHelper.EliminarSuministradorEvidenciaEnEdicion(CStr(FIX_GRANDCHILD), IDEdicion, sError)
    If sError <> "" Then
        logs(5) = "5. Fail: Eliminar error: " & sError
        Test_Subcon_Helper_AltaYEliminarEvidencia = BuildFail("Eliminar: " & sError, logs)
        GoTo Teardown
    End If
    logs(5) = "5. Eliminar OK"

    ' Assert: verificar que fue eliminado
    Set objDeleted = SuministradoresHelper.getSuministradorEnEdicion( _
        p_IDEdicion:=IDEdicion, p_IDSuministrador:=CStr(FIX_GRANDCHILD), p_Error:=sError)
    If Not objDeleted Is Nothing Then
        logs(6) = "6. Fail: GRANDCHILD aún existe tras eliminar"
        Test_Subcon_Helper_AltaYEliminarEvidencia = BuildFail("GRANDCHILD still exists after delete", logs)
        GoTo Teardown
    End If
    logs(6) = "6. Assert: GRANDCHILD eliminado (Nothing returned — OK)"

    logs(7) = "7. Teardown: restore fixture baseline"
    Test_Subcon_Helper_AltaYEliminarEvidencia = BuildOk("alta+eliminar OK", logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    If IDEdicion <> "" And CLng(Nz(IDEdicion, 0)) <> 0 Then
        Set dbT = GetTestDbSubcat(dbErr2)
        If VarType(dbT) <> vbString Then
            If Not dbT Is Nothing Then
                dbT.Execute "DELETE FROM TbProyectosEdicionesSuministradores WHERE IDEdicion=" & IDEdicion & " AND IDSuministrador=" & FIX_GRANDCHILD
                Set dbT = Nothing
            End If
        End If
    End If
    Test_Fixtures.TeardownSubcatGraph
    Exit Function

EH:
    Test_Subcon_Helper_AltaYEliminarEvidencia = BuildFail(Err.Description, logs)
    On Error Resume Next
    If IDEdicion <> "" And CLng(Nz(IDEdicion, 0)) <> 0 Then
        Set dbT = GetTestDbSubcat(tErr)
        If VarType(dbT) <> vbString Then
            If Not dbT Is Nothing Then
                dbT.Execute "DELETE FROM TbProyectosEdicionesSuministradores WHERE IDEdicion=" & IDEdicion & " AND IDSuministrador=" & FIX_GRANDCHILD
                Set dbT = Nothing
            End If
        End If
    End If
    Test_Fixtures.TeardownSubcatGraph
End Function

' ============================================================
' RUNALL — Agregador de suite
'
' Skill v1.7 §3.5: SeedAll antes, TeardownAll después.
' El grafo base (Fixture IDs = 900000) se crea en SeedAll y se
' limpia en TeardownAll.
' ============================================================
Public Function Test_Subcontratistas_RunAll() As String
    Dim runError As String
    Dim setupLogs(0 To 0) As String
    Dim r As String
    Dim results As String
    Dim passed As Long
    Dim failed As Long
    Dim okFlag As Boolean
    Dim runLogs(0 To 1) As String
    Dim runPayload As String
    Dim ehLogs(0 To 0) As String

    ' -- SuiteSetup (v2.1): configurar modo testing ------------------
    If Not ForceLocalBackend(runError) Then
        setupLogs(0) = "SuiteSetup falló"
        Test_Subcontratistas_RunAll = BuildJsonFail("TESTS BLOCKED: " & runError, setupLogs)
        Exit Function
    End If

    ' -- SeedAll: grafo base — UNA VEZ --------------------------------
    Test_Fixtures.SeedAll

    ' Test 1
    r = Test_Subcon_Helper_ConsorcioMasUno_ExcluyeSegundoNivel()
    results = "{""name"":""Test_Subcon_Helper_ConsorcioMasUno_ExcluyeSegundoNivel"",""result"":" & r & "}"
    passed = 0
    failed = 0
    If JsonOk(r) Then passed = passed + 1 Else failed = failed + 1

    ' Test 2
    r = Test_Subcon_Helper_Diff_CaminoFeliz()
    results = results & ",{""name"":""Test_Subcon_Helper_Diff_CaminoFeliz"",""result"":" & r & "}"
    If JsonOk(r) Then passed = passed + 1 Else failed = failed + 1

    ' Test 3
    r = Test_Subcon_Helper_Diff_CaminoTriste()
    results = results & ",{""name"":""Test_Subcon_Helper_Diff_CaminoTriste"",""result"":" & r & "}"
    If JsonOk(r) Then passed = passed + 1 Else failed = failed + 1

    ' Test 4
    r = Test_Subcon_Helper_Diff_NoIncluyeNieto()
    results = results & ",{""name"":""Test_Subcon_Helper_Diff_NoIncluyeNieto"",""result"":" & r & "}"
    If JsonOk(r) Then passed = passed + 1 Else failed = failed + 1

    ' Test 5
    r = Test_Subcon_Helper_AltaYEliminarEvidencia()
    results = results & ",{""name"":""Test_Subcon_Helper_AltaYEliminarEvidencia"",""result"":" & r & "}"
    If JsonOk(r) Then passed = passed + 1 Else failed = failed + 1

    ' -- TeardownAll -------------------------------------------------
    Test_Fixtures.TeardownAll

    ' -- SuiteTeardown -----------------------------------------------
    Test_Helper.ResetTestSession

    okFlag = (failed = 0)
    runLogs(0) = "passed=" & passed & ";failed=" & failed
    runPayload = "total=5;passed=" & passed & ";failed=" & failed
    If okFlag Then
        runLogs(1) = "Suite completa OK"
        Test_Subcontratistas_RunAll = BuildJsonOk(runPayload, runLogs)
    Else
        runLogs(1) = "results=[" & results & "]"
        Test_Subcontratistas_RunAll = BuildJsonFail( _
            "Subcontratistas suite failed: " & CStr(failed), _
            runLogs)
    End If
    Exit Function

EH:
    On Error Resume Next
    Test_Fixtures.TeardownAll
    Test_Helper.ResetTestSession
    ehLogs(0) = "Error en suite"
    Test_Subcontratistas_RunAll = BuildJsonFail(Err.Description, ehLogs)
End Function

' ============================================================
' GetTestDbSubcat — conexión directa al backend local
'
' Skill v1.7 §2: errores devueltos como JSON, nunca como Err.Raise.
' Retorna Variant: DAO.Database o String JSON de error.
' ============================================================
Public Function GetTestDbSubcat(ByRef p_Error As String) As Variant
    Dim localPath As String
    Dim ws As DAO.Workspace
    Dim db As DAO.Database

    Set GetTestDbSubcat = Nothing
    p_Error = ""

    ' v2.0 §6: GetTestDb NO llama ForceLocalBackend/ResetGlobals.
    ' SuiteSetup ya configuró m_BackendSandboxURL y m_PasswordBackend.
    localPath = m_BackendSandboxURL

    ' Skill v1.7 §2: verificar inline que el path existe y no es "".
    If localPath = "" Then
        p_Error = "TESTS BLOCKED: BackendSandbox is empty. SuiteSetup did not run."
        Exit Function
    End If
    If Not CreateObject("Scripting.FileSystemObject").FileExists(localPath) Then
        p_Error = "TESTS BLOCKED: BackendSandbox not found: " & EscapeJsonString(localPath)
        Exit Function
    End If

    Set ws = DBEngine(0)
    On Error Resume Next
    Set db = ws.OpenDatabase(localPath, False, False, ";PWD=" & m_PasswordBackend)
    If Err.Number <> 0 Then
        p_Error = "TESTS BLOCKED: Cannot open backend: " & EscapeJsonString(Err.Description)
        Set db = Nothing
        Set GetTestDbSubcat = Nothing
        Exit Function
    End If
    On Error GoTo 0

    Set GetTestDbSubcat = db
    p_Error = ""
End Function



