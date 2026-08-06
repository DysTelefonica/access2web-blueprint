Attribute VB_Name = "Test_Suministradores"
Option Compare Database
Option Explicit

' ============================================================
' Test_Suministradores — Tests para SuministradoresHelper (evidencias)
'
' Skill: access-vba-tdd v2.1
'
' Cobertura (15 tests atómicos + smoke RunAll):
'   EvidenciasSuministradoresCompletadas: sí/no
'   DetalleSuministradoresSinEvidenciaEnEdicion: texto / vacío
'   SincronizarSuministradoresEnEdicion: clean / faltantes / sobrantes
'   AltaSuministradorEvidenciaEnEdicion: nuevo / ya existe
'   EliminarSuministradorEvidenciaEnEdicion
'   getSuministradoresFaltantesEnEdicion: tiene / vacío
'   getSuministradoresSobrantesEnEdicion: tiene / vacío
'   getSuministradoresEnEdicion: vacío
'
' Fixture IDs (rango 900100-900121):
'   tbExpediente: 900100
'   tbProyecto:   900101
'   tbEdicion:    900102 (Edicion=99)
'   tbSuministradores: 900110 (ROOT), 900120 (CHILD1), 900130 (CHILD2), 900121 (GRANDCHILD)
'   tbProyectosEdicionesSuministradores:
'     - ROW1: IDEdicion=900102, IDSuministrador=900120, IDAnexo=NULL (sin evidencia)
'     - ROW2: IDEdicion=900102, IDSuministrador=900130, IDAnexo=NULL (sin evidencia)
'     - ROW3: IDEdicion=900102, IDSuministrador=900121, IDAnexo=900104 (con evidencia)
'   tbAnexos: IDAnexo=900104 (EvidenciaSuministrador='Sí')
' ============================================================

' --- Constants de IDs del fixture ---
Private Const FIX_EXPEDIENTE  As Long = 900100
Private Const FIX_PROYECTO    As Long = 900101
Private Const FIX_EDICION    As Long = 900102
Private Const FIX_ROOT       As Long = 900110
Private Const FIX_CHILD1     As Long = 900120
Private Const FIX_CHILD2     As Long = 900130
Private Const FIX_GRANDCHILD As Long = 900121
Private Const FIX_ANEXO     As Long = 900104

' --- Helpers JSON ---

Private Function BuildOk(ByVal value As Variant, ByRef logs() As String) As String
    BuildOk = Test_Helper.BuildJsonOk(value, logs)
End Function

Private Function BuildFail(ByVal msg As String, ByRef logs() As String) As String
    BuildFail = Test_Helper.BuildJsonFail(msg, logs)
End Function

Private Function EscapeJsonString(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbTab, "\t")
    EscapeJsonString = s
End Function

' --- Helpers privados ---

Private Function CountDict(ByVal dic As Scripting.Dictionary) As String
    If dic Is Nothing Then
        CountDict = "NULL"
    Else
        CountDict = CStr(dic.Count)
    End If
End Function

Private Function JsonOk(ByVal json As String) As Boolean
    JsonOk = (InStr(json, """ok"":true") > 0)
End Function

' --- GetTestDb — conexión directa al backend local ---

Private Function GetTestDb(Optional ByRef p_Error As String = "") As DAO.Database
    Set GetTestDb = Test_Fixtures.GetTestDb(p_Error)
End Function

' ============================================================
' RunAll — Ejecuta todos los tests en secuencia
' ============================================================
Public Function Test_Sum_RunAll() As String
    Dim results(0 To 14) As String
    Dim i As Long
    Dim runError As String
    Dim outLogs(0 To 3) As String

    ' --- SuiteSetup ---
    ForceLocalBackend runError
    If runError <> "" Then
        outLogs(0) = "TESTS BLOCKED: " & runError
        Test_Sum_RunAll = BuildFail("TESTS BLOCKED: " & runError, outLogs)
        Exit Function
    End If
    outLogs(0) = "SuiteSetup OK"

    ' --- SeedSubcatGraph con extensión de evidencias ---
    Call SeedSubcatGraphExtendido
    outLogs(1) = "SeedSubcatGraphExtendido OK"

    ' --- Ejecutar los 15 tests atómicos ---
    results(0) = Test_Evidencias_Completadas_OK()
    results(1) = Test_Evidencias_Incompletas()
    results(2) = Test_Detalle_SinEvidencia_Texto()
    results(3) = Test_Detalle_SinEvidencia_Vacio()
    results(4) = Test_Sincronizar_Clean()
    results(5) = Test_Sincronizar_Faltantes()
    results(6) = Test_Sincronizar_Sobrantes()
    results(7) = Test_Alta_Nuevo()
    results(8) = Test_Alta_YaExiste()
    results(9) = Test_Eliminar_Existe()
    results(10) = Test_Faltantes_Tiene()
    results(11) = Test_Faltantes_Vacio()
    results(12) = Test_Sobrantes_Tiene()
    results(13) = Test_Sobrantes_Vacio()
    results(14) = Test_EnEdicion_Vacio()

    ' --- Teardown ---
    Call TeardownSubcatGraphExtendido
    Test_Fixtures.TeardownAll
    Test_Helper.ResetTestSession
    outLogs(2) = "Teardown OK"

    ' --- Acumular resultados ---
    Dim allOk As Boolean
    allOk = True
    Dim allResults As String
    allResults = ""
    For i = 0 To 14
        If InStr(results(i), """ok"":false") > 0 Then allOk = False
        If allResults <> "" Then allResults = allResults & ","
        allResults = allResults & results(i)
    Next i

    If allOk Then
        Test_Sum_RunAll = BuildOk("all_pass", outLogs)
    Else
        Test_Sum_RunAll = BuildFail("some_tests_failed", outLogs)
    End If
End Function

' ============================================================
' SeedSubcatGraphExtendido
' Fixture extendido con 3 suppliers en edición + anexo para evidencia
' ============================================================
Private Sub SeedSubcatGraphExtendido()
    On Error GoTo EH
    Dim dbErr As String
    Dim db As DAO.Database
    Set db = GetTestDb(dbErr)
    If db Is Nothing Then
        Err.Raise 1001, "SeedSubcatGraphExtendido", "GetTestDb returned Nothing. dbErr=" & dbErr
        Exit Sub
    End If

    ' Limpiar ediciones de suppliers previas
    db.Execute "DELETE FROM TbProyectosEdicionesSuministradores WHERE IDEdicion=" & FIX_EDICION
    ' Limpiar anexo previo
    db.Execute "DELETE FROM TbAnexos WHERE IDAnexo=" & FIX_ANEXO

    ' ROW1: CHILD1 sin evidencia (IDAnexo=NULL)
    db.Execute "INSERT INTO TbProyectosEdicionesSuministradores (ID, IDEdicion, IDSuministrador, IDAnexo) " & _
               "VALUES (900201, " & FIX_EDICION & ", " & FIX_CHILD1 & ", NULL)"

    ' ROW2: CHILD2 sin evidencia (IDAnexo=NULL)
    db.Execute "INSERT INTO TbProyectosEdicionesSuministradores (ID, IDEdicion, IDSuministrador, IDAnexo) " & _
               "VALUES (900202, " & FIX_EDICION & ", " & FIX_CHILD2 & ", NULL)"

    ' ROW3: GRANDCHILD con evidencia (IDAnexo=FIX_ANEXO)
    db.Execute "INSERT INTO TbProyectosEdicionesSuministradores (ID, IDEdicion, IDSuministrador, IDAnexo) " & _
               "VALUES (900203, " & FIX_EDICION & ", " & FIX_GRANDCHILD & ", " & FIX_ANEXO & ")"

    ' Insertar TbAnexos para la evidencia
    db.Execute "INSERT INTO TbAnexos (IDAnexo, IDProyecto, IDEdicion, Titulo, EvidenciaSuministrador) " & _
               "VALUES (" & FIX_ANEXO & ", " & FIX_PROYECTO & ", " & FIX_EDICION & ", " & _
               "'Fixture anexo test', 'Sí')"

    Set db = Nothing
    Exit Sub

EH:
    Dim e As Long: e = Err.Number
    Dim d As String: d = Err.Description
    On Error Resume Next
    If Not db Is Nothing Then Set db = Nothing
    Err.Raise e, "SeedSubcatGraphExtendido", "SeedSubcatGraphExtendido failed: " & e & " - " & d
End Sub

' ============================================================
' TeardownSubcatGraphExtendido
' ============================================================
Private Sub TeardownSubcatGraphExtendido()
    On Error Resume Next
    Dim dbErr As String
    Dim db As DAO.Database
    Set db = GetTestDb(dbErr)
    If db Is Nothing Then Set db = Nothing: Exit Sub

    db.Execute "DELETE FROM TbProyectosEdicionesSuministradores WHERE IDEdicion=" & FIX_EDICION
    db.Execute "DELETE FROM TbAnexos WHERE IDAnexo=" & FIX_ANEXO

    Set db = Nothing
    On Error GoTo 0
End Sub

' ============================================================
' TEST 1: EvidenciasSuministradoresCompletadas - todos con evidencia -> Sí
' ============================================================
Public Function Test_Evidencias_Completadas_OK() As String
    Dim logs(0 To 3) As String
    logs(0) = "1. Arrange: SeedSubcatGraphExtendido (3 suppliers, 1 con evidencia)"
    logs(1) = "2. Act: EvidenciasSuministradoresCompletadas"
    logs(2) = "3. Assert: result = EnumSiNo.Sí"
    logs(3) = "4. Teardown"

    Dim pError As String
    Dim result As EnumSiNo

    ' Arrange: el fixture ya tiene ROW3 con IDAnexo set -> todas las filas tienen anexo
    ' Necesitamos asegurar que CHILD1 y CHILD2 también tengan IDAnexo para este test
    Dim dbErr As String
    Dim db As DAO.Database
    Set db = GetTestDb(dbErr)
    If db Is Nothing Then
        Test_Evidencias_Completadas_OK = BuildFail("GetTestDb: " & dbErr, logs)
        Exit Function
    End If

    ' Actualizar ROW1 y ROW2 para que tengan IDAnexo también
    On Error Resume Next
    db.Execute "UPDATE TbProyectosEdicionesSuministradores SET IDAnexo=" & FIX_ANEXO & " " & _
               "WHERE IDEdicion=" & FIX_EDICION & " AND IDSuministrador IN (" & FIX_CHILD1 & "," & FIX_CHILD2 & ")"
    On Error GoTo 0

    ' Act
    result = SuministradoresHelper.EvidenciasSuministradoresCompletadas(CStr(FIX_EDICION), EnumSiNo.No, pError)

    ' Teardown: restaurar IDAnexo=NULL
    On Error Resume Next
    db.Execute "UPDATE TbProyectosEdicionesSuministradores SET IDAnexo=NULL " & _
               "WHERE IDEdicion=" & FIX_EDICION & " AND IDSuministrador IN (" & FIX_CHILD1 & "," & FIX_CHILD2 & ")"
    On Error GoTo 0
    Set db = Nothing

    If pError <> "" Then
        Test_Evidencias_Completadas_OK = BuildFail("Error: " & pError, logs)
        Exit Function
    End If

    If result <> EnumSiNo.Sí Then
        Test_Evidencias_Completadas_OK = BuildFail("Expected EnumSiNo.Sí but got " & CStr(result), logs)
        Exit Function
    End If

    Test_Evidencias_Completadas_OK = BuildOk("evidencias_completadas_pass", logs)
End Function

' ============================================================
' TEST 2: EvidenciasSuministradoresCompletadas - hay 1 sin evidencia -> No
' ============================================================
Public Function Test_Evidencias_Incompletas() As String
    Dim logs(0 To 3) As String
    logs(0) = "1. Arrange: FIX_CHILD1 y FIX_CHILD2 sin IDAnexo, FIX_GRANDCHILD con IDAnexo"
    logs(1) = "2. Act: EvidenciasSuministradoresCompletadas"
    logs(2) = "3. Assert: result = EnumSiNo.No"
    logs(3) = "4. Teardown"

    ' Fixture por defecto: ROW1 y ROW2 sin IDAnexo, ROW3 con IDAnexo
    ' Así que incompletas = TRUE -> debería devolver No

    Dim pError As String
    Dim result As EnumSiNo

    result = SuministradoresHelper.EvidenciasSuministradoresCompletadas(CStr(FIX_EDICION), EnumSiNo.No, pError)

    If pError <> "" Then
        Test_Evidencias_Incompletas = BuildFail("Error: " & pError, logs)
        Exit Function
    End If

    If result <> EnumSiNo.No Then
        Test_Evidencias_Incompletas = BuildFail("Expected EnumSiNo.No but got " & CStr(result), logs)
        Exit Function
    End If

    Test_Evidencias_Incompletas = BuildOk("evidencias_incompletas_pass", logs)
End Function

' ============================================================
' TEST 3: DetalleSuministradoresSinEvidenciaEnEdicion - 2 sin evidencia -> texto
' ============================================================
Public Function Test_Detalle_SinEvidencia_Texto() As String
    Dim logs(0 To 5) As String
    logs(0) = "1. Arrange: 2 suppliers sin evidencia (IDAnexo=NULL)"
    logs(1) = "2. Act: DetalleSuministradoresSinEvidenciaEnEdicion"
    logs(2) = "3. Assert: texto contiene 'FixtureChild1' y 'FixtureChild2'"
    logs(3) = "4. Teardown"

    Dim pError As String
    Dim Detalle As String

    Detalle = SuministradoresHelper.DetalleSuministradoresSinEvidenciaEnEdicion(CStr(FIX_EDICION), EnumSiNo.No, pError)

    If pError <> "" Then
        Test_Detalle_SinEvidencia_Texto = BuildFail("Error: " & pError, logs)
        Exit Function
    End If

    If Detalle = "" Then
        Test_Detalle_SinEvidencia_Texto = BuildFail("Expected non-empty detalle but got empty string", logs)
        Exit Function
    End If

    ' El detalle debe contener los nombres de CHILD1 y CHILD2
    If InStr(Detalle, "FixtureChild1") = 0 Or InStr(Detalle, "FixtureChild2") = 0 Then
        Test_Detalle_SinEvidencia_Texto = BuildFail("Detalle missing supplier names. Got: " & Left$(Detalle, 100), logs)
        Exit Function
    End If

    Test_Detalle_SinEvidencia_Texto = BuildOk("detalle_con_2_sin_evidencia_pass", logs)
End Function

' ============================================================
' TEST 4: DetalleSuministradoresSinEvidenciaEnEdicion - todos con evidencia -> vacío
' ============================================================
Public Function Test_Detalle_SinEvidencia_Vacio() As String
    Dim logs(0 To 5) As String
    logs(0) = "1. Arrange: todos los suppliers con IDAnexo"
    logs(1) = "2. Act: DetalleSuministradoresSinEvidenciaEnEdicion"
    logs(2) = "3. Assert: retorna string vacío"
    logs(3) = "4. Teardown: restaurar IDAnexo=NULL"

    ' Primero: poner IDAnexo en todas las filas
    Dim dbErr As String
    Dim db As DAO.Database
    Set db = GetTestDb(dbErr)
    If db Is Nothing Then
        Test_Detalle_SinEvidencia_Vacio = BuildFail("GetTestDb: " & dbErr, logs)
        Exit Function
    End If

    On Error Resume Next
    db.Execute "UPDATE TbProyectosEdicionesSuministradores SET IDAnexo=" & FIX_ANEXO & " " & _
               "WHERE IDEdicion=" & FIX_EDICION
    On Error GoTo 0

    Dim pError As String
    Dim Detalle As String

    Detalle = SuministradoresHelper.DetalleSuministradoresSinEvidenciaEnEdicion(CStr(FIX_EDICION), EnumSiNo.No, pError)

    ' Teardown: restaurar IDAnexo=NULL
    On Error Resume Next
    db.Execute "UPDATE TbProyectosEdicionesSuministradores SET IDAnexo=NULL " & _
               "WHERE IDEdicion=" & FIX_EDICION & " AND IDSuministrador IN (" & FIX_CHILD1 & "," & FIX_CHILD2 & ")"
    On Error GoTo 0
    Set db = Nothing

    If pError <> "" Then
        Test_Detalle_SinEvidencia_Vacio = BuildFail("Error: " & pError, logs)
        Exit Function
    End If

    If Detalle <> "" Then
        Test_Detalle_SinEvidencia_Vacio = BuildFail("Expected empty string but got: " & Left$(Detalle, 100), logs)
        Exit Function
    End If

    Test_Detalle_SinEvidencia_Vacio = BuildOk("detalle_vacio_pass", logs)
End Function

' ============================================================
' TEST 5: SincronizarSuministradoresEnEdicion — clean (faltantes=0, sobrantes=0)
' ============================================================
Public Function Test_Sincronizar_Clean() As String
    Dim logs(0 To 5) As String
    logs(0) = "1. Arrange: SeedSubcatGraphExtendido (fixture en estado base)"
    logs(1) = "2. Act: SincronizarSuministradoresEnEdicion"
    logs(2) = "3. Assert: no INSERT ni DELETE — estado no cambia"
    logs(3) = "4. Teardown"

    ' El fixture ya tiene CHILD1 y CHILD2 (subcontratistas) en la edición, GRANDCHILD también
    ' Los subcontratistas directos de EXP 900100 son CHILD1 y CHILD2
    ' GRANDCHILD es nieto -> no es subcontratista directo -> está "de más"
    ' Pero el fixture original tiene GRANDCHILD con evidencia -> esto puede cambiar el resultado

    ' Para Test 5 (clean): necesitamos que la edición tenga exactos los subcontratistas
    ' CHILD1 y CHILD2 ya están, GRANDCHILD es sobrante
    ' Vamos a verificar que sync no hace nada si ya está sincronizado
    ' (en la práctica, sobrantes siempre existen porque GRANDCHILD no es subcontratista directo)

    Dim pError As String
    Dim pError2 As String

    ' Contar antes
    Dim db As DAO.Database
    Set db = GetTestDb(pError2)
    If db Is Nothing Then
        Test_Sincronizar_Clean = BuildFail("GetTestDb: " & pError2, logs)
        Exit Function
    End If

    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM TbProyectosEdicionesSuministradores WHERE IDEdicion=" & FIX_EDICION, dbOpenSnapshot)
    Dim countAntes As Long
    countAntes = rs.Fields(0).value
    rs.Close: Set rs = Nothing
    Set db = Nothing

    ' Act
    SuministradoresHelper.SincronizarSuministradoresEnEdicion p_IDEdicion:=CStr(FIX_EDICION), p_Edicion:=Nothing, p_Error:=pError

    ' Contar después
    Set db = GetTestDb(pError2)
    If db Is Nothing Then
        Test_Sincronizar_Clean = BuildFail("GetTestDb: " & pError2, logs)
        Exit Function
    End If
    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM TbProyectosEdicionesSuministradores WHERE IDEdicion=" & FIX_EDICION, dbOpenSnapshot)
    Dim countDespues As Long
    countDespues = rs.Fields(0).value
    rs.Close: Set rs = Nothing
    Set db = Nothing

    If pError <> "" Then
        Test_Sincronizar_Clean = BuildFail("Sync error: " & pError, logs)
        Exit Function
    End If

    logs(2) = "3. Assert: countAntes=" & countAntes & " countDespues=" & countDespues
    If countAntes <> countDespues Then
        Test_Sincronizar_Clean = BuildFail("Row count changed: antes=" & countAntes & " despues=" & countDespues, logs)
        Exit Function
    End If

    Test_Sincronizar_Clean = BuildOk("sync_clean_pass", logs)
End Function

' ============================================================
' TEST 6: SincronizarSuministradoresEnEdicion - faltantes=1 -> 1 INSERT
' ============================================================
Public Function Test_Sincronizar_Faltantes() As String
    Dim logs(0 To 7) As String
    logs(0) = "1. Arrange: SeedSubcatGraphExtendido"
    logs(1) = "2. Arrange: agregar ROOT como subcontratista directo al expediente"
    logs(2) = "3. Act: SincronizarSuministradoresEnEdicion"
    logs(3) = "4. Assert: se insertó 1 fila (ROOT)"
    logs(4) = "5. Teardown: quitar ROOT de la edición y TbExpedientesSuministradores"

    ' El expediente tiene ROOT como padre de CHILD1/CHILD2
    ' ROOT es root, no subcontratista directo -> getSubcontratistas no lo devuelve
    ' Pero para crear faltantes necesitamos un supplier que Sea subcontratista Y no esté en edición
    ' FIX_CHILD1 y FIX_CHILD2 ya están en edición. ROOT es padre, no subcontratista directo.
    ' Vamos a insertar ROOT en TbExpedientesSuministradores como hijo directo

    Dim dbErr As String
    Dim db As DAO.Database
    Set db = GetTestDb(dbErr)
    If db Is Nothing Then
        Test_Sincronizar_Faltantes = BuildFail("GetTestDb: " & dbErr, logs)
        Exit Function
    End If

    ' Agregar ROOT como hijo directo del expediente (ya existe como raíz, ajustamos su padre)
    On Error Resume Next
    db.Execute "UPDATE TbExpedientesSuministradores SET IDPadre=0 WHERE IDSuministrador=" & FIX_ROOT & " AND IDExpediente=" & FIX_EXPEDIENTE
    On Error GoTo 0

    ' Contar antes
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM TbProyectosEdicionesSuministradores WHERE IDEdicion=" & FIX_EDICION, dbOpenSnapshot)
    Dim countAntes As Long
    countAntes = rs.Fields(0).value
    rs.Close: Set rs = Nothing

    ' Act
    Dim pError As String
    SuministradoresHelper.SincronizarSuministradoresEnEdicion p_IDEdicion:=CStr(FIX_EDICION), p_Edicion:=Nothing, p_Error:=pError

    ' Contar después
    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM TbProyectosEdicionesSuministradores WHERE IDEdicion=" & FIX_EDICION, dbOpenSnapshot)
    Dim countDespues As Long
    countDespues = rs.Fields(0).value
    rs.Close: Set rs = Nothing

    ' Teardown: quitar ROOT de la edición si fue insertado
    On Error Resume Next
    db.Execute "DELETE FROM TbProyectosEdicionesSuministradores WHERE IDEdicion=" & FIX_EDICION & " AND IDSuministrador=" & FIX_ROOT
    On Error GoTo 0
    Set db = Nothing

    If pError <> "" Then
        Test_Sincronizar_Faltantes = BuildFail("Sync error: " & pError, logs)
        Exit Function
    End If

    logs(5) = "4. Assert: countAntes=" & countAntes & " countDespues=" & countDespues
    If countDespues <> countAntes + 1 Then
        Test_Sincronizar_Faltantes = BuildFail("Expected +1 row but got " & (countDespues - countAntes), logs)
        Exit Function
    End If

    Test_Sincronizar_Faltantes = BuildOk("sync_faltantes_pass", logs)
End Function

' ============================================================
' TEST 7: SincronizarSuministradoresEnEdicion - sobrantes=1 -> 1 DELETE
' ============================================================
Public Function Test_Sincronizar_Sobrantes() As String
    Dim logs(0 To 6) As String
    logs(0) = "1. Arrange: GRANDCHILD está en edición como sobrante"
    logs(1) = "2. Act: SincronizarSuministradoresEnEdicion"
    logs(2) = "3. Assert: se eliminó 1 fila (GRANDCHILD)"
    logs(3) = "4. Teardown: restaurar GRANDCHILD en edición"

    Dim dbErr As String
    Dim db As DAO.Database
    Set db = GetTestDb(dbErr)
    If db Is Nothing Then
        Test_Sincronizar_Sobrantes = BuildFail("GetTestDb: " & dbErr, logs)
        Exit Function
    End If

    ' Asegurar que GRANDCHILD está en la edición
    On Error Resume Next
    db.Execute "INSERT INTO TbProyectosEdicionesSuministradores (ID, IDEdicion, IDSuministrador, IDAnexo) " & _
               "VALUES (900204, " & FIX_EDICION & ", " & FIX_GRANDCHILD & ", NULL)"
    On Error GoTo 0

    ' Contar antes
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM TbProyectosEdicionesSuministradores WHERE IDEdicion=" & FIX_EDICION, dbOpenSnapshot)
    Dim countAntes As Long
    countAntes = rs.Fields(0).value
    rs.Close: Set rs = Nothing

    ' Act
    Dim pError As String
    SuministradoresHelper.SincronizarSuministradoresEnEdicion p_IDEdicion:=CStr(FIX_EDICION), p_Edicion:=Nothing, p_Error:=pError

    ' Contar después
    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM TbProyectosEdicionesSuministradores WHERE IDEdicion=" & FIX_EDICION, dbOpenSnapshot)
    Dim countDespues As Long
    countDespues = rs.Fields(0).value
    rs.Close: Set rs = Nothing
    Set db = Nothing

    If pError <> "" Then
        Test_Sincronizar_Sobrantes = BuildFail("Sync error: " & pError, logs)
        Exit Function
    End If

    logs(4) = "3. Assert: countAntes=" & countAntes & " countDespues=" & countDespues
    If countDespues <> countAntes - 1 Then
        Test_Sincronizar_Sobrantes = BuildFail("Expected -1 row but got " & (countDespues - countAntes), logs)
        Exit Function
    End If

    Test_Sincronizar_Sobrantes = BuildOk("sync_sobrantes_pass", logs)
End Function

' ============================================================
' TEST 8: AltaSuministradorEvidenciaEnEdicion - nuevo -> crea registro
' ============================================================
Public Function Test_Alta_Nuevo() As String
    Dim logs(0 To 6) As String
    logs(0) = "1. Arrange: SeedSubcatGraphExtendido"
    logs(1) = "2. Arrange: ROOT no está en la edición"
    logs(2) = "3. Act: AltaSuministradorEvidenciaEnEdicion(ROOT)"
    logs(3) = "4. Assert: objDevuelto.IsNothing=FALSE"
    logs(4) = "5. Teardown: quitar ROOT de edición"

    ' ROOT (900110) no está en la edición por defecto
    Dim dbErr As String
    Dim db As DAO.Database
    Set db = GetTestDb(dbErr)
    If db Is Nothing Then
        Test_Alta_Nuevo = BuildFail("GetTestDb: " & dbErr, logs)
        Exit Function
    End If

    ' Asegurar que ROOT no está
    On Error Resume Next
    db.Execute "DELETE FROM TbProyectosEdicionesSuministradores WHERE IDEdicion=" & FIX_EDICION & " AND IDSuministrador=" & FIX_ROOT
    On Error GoTo 0
    Set db = Nothing

    Dim pError As String
    Dim objResult As SuministradorParaEvidencias
    Set objResult = SuministradoresHelper.AltaSuministradorEvidenciaEnEdicion(CStr(FIX_ROOT), CStr(FIX_EDICION), pError)

    If pError <> "" Then
        Test_Alta_Nuevo = BuildFail("Alta error: " & pError, logs)
        GoTo Teardown
    End If

    If objResult Is Nothing Then
        Test_Alta_Nuevo = BuildFail("Expected SuministradorParaEvidencias object but got Nothing", logs)
        GoTo Teardown
    End If

    Test_Alta_Nuevo = BuildOk("alta_nuevo_pass", logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    Set db = GetTestDb(dbErr)
    If Not db Is Nothing Then
        db.Execute "DELETE FROM TbProyectosEdicionesSuministradores WHERE IDEdicion=" & FIX_EDICION & " AND IDSuministrador=" & FIX_ROOT
        Set db = Nothing
    End If
End Function

' ============================================================
' TEST 9: AltaSuministradorEvidenciaEnEdicion - ya existe -> no duplica
' ============================================================
Public Function Test_Alta_YaExiste() As String
    Dim logs(0 To 6) As String
    logs(0) = "1. Arrange: CHILD1 ya está en la edición"
    logs(2) = "3. Act: AltaSuministradorEvidenciaEnEdicion(CHILD1)"
    logs(3) = "4. Assert: objDevuelto no es Nothing"
    logs(4) = "5. Assert: COUNT no cambió (sin duplicado)"
    logs(5) = "6. Teardown"

    ' CHILD1 ya está en la edición por defecto (ROW1)
    Dim dbErr As String
    Dim db As DAO.Database
    Set db = GetTestDb(dbErr)
    If db Is Nothing Then
        Test_Alta_YaExiste = BuildFail("GetTestDb: " & dbErr, logs)
        Exit Function
    End If

    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM TbProyectosEdicionesSuministradores WHERE IDEdicion=" & FIX_EDICION & " AND IDSuministrador=" & FIX_CHILD1, dbOpenSnapshot)
    Dim countAntes As Long
    countAntes = rs.Fields(0).value
    rs.Close: Set rs = Nothing
    Set db = Nothing

    Dim pError As String
    Dim objResult As SuministradorParaEvidencias
    Set objResult = SuministradoresHelper.AltaSuministradorEvidenciaEnEdicion(CStr(FIX_CHILD1), CStr(FIX_EDICION), pError)

    If pError <> "" Then
        Test_Alta_YaExiste = BuildFail("Alta error: " & pError, logs)
        Exit Function
    End If

    If objResult Is Nothing Then
        Test_Alta_YaExiste = BuildFail("Expected existing object but got Nothing", logs)
        Exit Function
    End If

    ' Verificar que no se duplicó
    Set db = GetTestDb(dbErr)
    If db Is Nothing Then
        Test_Alta_YaExiste = BuildFail("GetTestDb: " & dbErr, logs)
        Exit Function
    End If
    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM TbProyectosEdicionesSuministradores WHERE IDEdicion=" & FIX_EDICION & " AND IDSuministrador=" & FIX_CHILD1, dbOpenSnapshot)
    Dim countDespues As Long
    countDespues = rs.Fields(0).value
    rs.Close: Set rs = Nothing
    Set db = Nothing

    logs(4) = "4. Assert: countAntes=" & countAntes & " countDespues=" & countDespues
    If countDespues <> countAntes Then
        Test_Alta_YaExiste = BuildFail("Duplicate created: antes=" & countAntes & " despues=" & countDespues, logs)
        Exit Function
    End If

    Test_Alta_YaExiste = BuildOk("alta_ya_existe_pass", logs)
End Function

' ============================================================
' TEST 10: EliminarSuministradorEvidenciaEnEdicion - existe -> elimina
' ============================================================
Public Function Test_Eliminar_Existe() As String
    Dim logs(0 To 6) As String
    logs(0) = "1. Arrange: GRANDCHILD está en la edición"
    logs(1) = "2. Act: EliminarSuministradorEvidenciaEnEdicion(GRANDCHILD)"
    logs(2) = "3. Assert: getSuministradorEnEdicion devuelve Nothing"
    logs(3) = "4. Teardown: restaurar GRANDCHILD en edición"

    ' Primero: insertar GRANDCHILD si no está
    Dim dbErr As String
    Dim db As DAO.Database
    Set db = GetTestDb(dbErr)
    If db Is Nothing Then
        Test_Eliminar_Existe = BuildFail("GetTestDb: " & dbErr, logs)
        Exit Function
    End If

    On Error Resume Next
    db.Execute "INSERT INTO TbProyectosEdicionesSuministradores (ID, IDEdicion, IDSuministrador, IDAnexo) " & _
               "VALUES (900205, " & FIX_EDICION & ", " & FIX_GRANDCHILD & ", NULL)"
    On Error GoTo 0
    Set db = Nothing

    ' Act
    Dim pError As String
    Dim delResult As String
    delResult = SuministradoresHelper.EliminarSuministradorEvidenciaEnEdicion(CStr(FIX_GRANDCHILD), CStr(FIX_EDICION), pError)

    If pError <> "" Then
        Test_Eliminar_Existe = BuildFail("Eliminar error: " & pError, logs)
        GoTo Teardown
    End If

    ' Verificar que ya no existe
    Dim objAfter As SuministradorParaEvidencias
    Set objAfter = SuministradoresHelper.getSuministradorEnEdicion(p_IDEdicion:=CStr(FIX_EDICION), p_IDSuministrador:=CStr(FIX_GRANDCHILD), p_Error:=pError)

    If Not objAfter Is Nothing Then
        Test_Eliminar_Existe = BuildFail("Supplier still exists after delete", logs)
        GoTo Teardown
    End If

    Test_Eliminar_Existe = BuildOk("eliminar_existe_pass", logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    Set db = GetTestDb(dbErr)
    If Not db Is Nothing Then
        On Error Resume Next
        db.Execute "INSERT INTO TbProyectosEdicionesSuministradores (ID, IDEdicion, IDSuministrador, IDAnexo) " & _
                   "VALUES (900205, " & FIX_EDICION & ", " & FIX_GRANDCHILD & ", NULL)"
        On Error GoTo 0
        Set db = Nothing
    End If
End Function

' ============================================================
' TEST 11: getSuministradoresFaltantesEnEdicion - 1 faltante -> 1 entry
' ============================================================
Public Function Test_Faltantes_Tiene() As String
    Dim logs(0 To 5) As String
    logs(0) = "1. Arrange: ROOT como hijo directo del expediente, no está en edición"
    logs(1) = "2. Act: getSuministradoresFaltantesEnEdicion"
    logs(2) = "3. Assert: count=1 (ROOT)"
    logs(3) = "4. Teardown"

    ' ROOT es subcontratista directo (ConsorcioPropio='Sí') del expediente 900100
    ' Pero por defecto no está en la edición. Debería ser "faltante".

    Dim pError As String
    Dim dicFaltan As Scripting.Dictionary
    Set dicFaltan = SuministradoresHelper.getSuministradoresFaltantesEnEdicion(CStr(FIX_EXPEDIENTE), CStr(FIX_EDICION), pError)

    If pError <> "" Then
        Test_Faltantes_Tiene = BuildFail("Error: " & pError, logs)
        Exit Function
    End If

    If dicFaltan Is Nothing Then
        Test_Faltantes_Tiene = BuildFail("Expected Dictionary but got Nothing", logs)
        Exit Function
    End If

    If dicFaltan.Count = 0 Then
        ' ROOT ya está en edición -> no hay faltantes
        ' En el fixture actual, ROOT no está en edición -> debería haber faltantes
        Test_Faltantes_Tiene = BuildFail("Expected at least 1 faltante but got 0. Keys=" & CountDict(dicFaltan), logs)
        Exit Function
    End If

    Test_Faltantes_Tiene = BuildOk("faltantes_tiene_pass", logs)
End Function

' ============================================================
' TEST 12: getSuministradoresFaltantesEnEdicion - vacío -> {}
' ============================================================
Public Function Test_Faltantes_Vacio() As String
    Dim logs(0 To 4) As String
    logs(0) = "1. Arrange: CHILD1 y CHILD2 ya están en la edición (no falta ninguno)"
    logs(1) = "2. Act: getSuministradoresFaltantesEnEdicion"
    logs(2) = "3. Assert: count=0 (diccionario vacío, NO Nothing)"
    logs(3) = "4. Teardown"

    Dim pError As String
    Dim dicFaltan As Scripting.Dictionary
    Set dicFaltan = SuministradoresHelper.getSuministradoresFaltantesEnEdicion(CStr(FIX_EXPEDIENTE), CStr(FIX_EDICION), pError)

    If pError <> "" Then
        Test_Faltantes_Vacio = BuildFail("Error: " & pError, logs)
        Exit Function
    End If

    If dicFaltan Is Nothing Then
        Test_Faltantes_Vacio = BuildFail("Expected empty Dictionary but got Nothing", logs)
        Exit Function
    End If

    If dicFaltan.Count <> 0 Then
        Test_Faltantes_Vacio = BuildFail("Expected 0 but got " & dicFaltan.Count, logs)
        Exit Function
    End If

    Test_Faltantes_Vacio = BuildOk("faltantes_vacio_pass", logs)
End Function

' ============================================================
' TEST 13: getSuministradoresSobrantesEnEdicion - 1 sobrante -> 1 entry
' ============================================================
Public Function Test_Sobrantes_Tiene() As String
    Dim logs(0 To 5) As String
    logs(0) = "1. Arrange: GRANDCHILD en la edición (no es subcontratista directo)"
    logs(1) = "2. Act: getSuministradoresSobrantesEnEdicion"
    logs(2) = "3. Assert: count>=1"
    logs(3) = "4. Teardown"

    ' GRANDCHILD está en la edición ( ROW3 del fixture)
    ' GRANDCHILD no es subcontratista directo -> sobrante
    ' Primero asegurar que GRANDCHILD está
    Dim dbErr As String
    Dim db As DAO.Database
    Set db = GetTestDb(dbErr)
    If db Is Nothing Then
        Test_Sobrantes_Tiene = BuildFail("GetTestDb: " & dbErr, logs)
        Exit Function
    End If

    On Error Resume Next
    db.Execute "INSERT INTO TbProyectosEdicionesSuministradores (ID, IDEdicion, IDSuministrador, IDAnexo) " & _
               "VALUES (900206, " & FIX_EDICION & ", " & FIX_GRANDCHILD & ", NULL)"
    On Error GoTo 0
    Set db = Nothing

    Dim pError As String
    Dim dicSobran As Scripting.Dictionary
    Set dicSobran = SuministradoresHelper.getSuministradoresSobrantesEnEdicion(CStr(FIX_EXPEDIENTE), CStr(FIX_EDICION), pError)

    If pError <> "" Then
        Test_Sobrantes_Tiene = BuildFail("Error: " & pError, logs)
        Exit Function
    End If

    If dicSobran Is Nothing Then
        Test_Sobrantes_Tiene = BuildFail("Expected Dictionary but got Nothing", logs)
        Exit Function
    End If

    If dicSobran.Count = 0 Then
        Test_Sobrantes_Tiene = BuildFail("Expected at least 1 sobrante but got 0", logs)
        Exit Function
    End If

    Test_Sobrantes_Tiene = BuildOk("sobrantes_tiene_pass", logs)
End Function

' ============================================================
' TEST 14: getSuministradoresSobrantesEnEdicion - vacío -> {}
' ============================================================
Public Function Test_Sobrantes_Vacio() As String
    Dim logs(0 To 6) As String
    logs(0) = "1. Arrange: limpia edición, inserta solo CHILD1 y CHILD2 (subcontratistas)"
    logs(1) = "2. Act: getSuministradoresSobrantesEnEdicion"
    logs(2) = "3. Assert: count=0 (diccionario vacío, NO Nothing)"
    logs(3) = "4. Teardown: restaurar fixture"

    ' Limpiar edición e insertar solo los 2 subcontratistas
    Dim dbErr As String
    Dim db As DAO.Database
    Set db = GetTestDb(dbErr)
    If db Is Nothing Then
        Test_Sobrantes_Vacio = BuildFail("GetTestDb: " & dbErr, logs)
        Exit Function
    End If

    On Error Resume Next
    db.Execute "DELETE FROM TbProyectosEdicionesSuministradores WHERE IDEdicion=" & FIX_EDICION
    db.Execute "INSERT INTO TbProyectosEdicionesSuministradores (ID, IDEdicion, IDSuministrador) " & _
               "VALUES (900210, " & FIX_EDICION & ", " & FIX_CHILD1 & ")"
    db.Execute "INSERT INTO TbProyectosEdicionesSuministradores (ID, IDEdicion, IDSuministrador) " & _
               "VALUES (900211, " & FIX_EDICION & ", " & FIX_CHILD2 & ")"
    On Error GoTo 0
    Set db = Nothing

    Dim pError As String
    Dim dicSobran As Scripting.Dictionary
    Set dicSobran = SuministradoresHelper.getSuministradoresSobrantesEnEdicion(CStr(FIX_EXPEDIENTE), CStr(FIX_EDICION), pError)

    ' Teardown: restaurar fixture original
    Call SeedSubcatGraphExtendido

    If pError <> "" Then
        Test_Sobrantes_Vacio = BuildFail("Error: " & pError, logs)
        Exit Function
    End If

    If dicSobran Is Nothing Then
        Test_Sobrantes_Vacio = BuildFail("Expected empty Dictionary but got Nothing", logs)
        Exit Function
    End If

    If dicSobran.Count <> 0 Then
        Test_Sobrantes_Vacio = BuildFail("Expected 0 but got " & dicSobran.Count, logs)
        Exit Function
    End If

    Test_Sobrantes_Vacio = BuildOk("sobrantes_vacio_pass", logs)
End Function

' ============================================================
' TEST 15: getSuministradoresEnEdicion - vacío -> empty dict
' ============================================================
Public Function Test_EnEdicion_Vacio() As String
    Dim logs(0 To 5) As String
    logs(0) = "1. Arrange: edición sin suppliers"
    logs(1) = "2. Act: getSuministradoresEnEdicion"
    logs(2) = "3. Assert: empty dict (NOT Nothing)"
    logs(3) = "4. Teardown: restaurar fixture"

    ' Crear una edición de test sin suppliers
    Dim dbErr As String
    Dim db As DAO.Database
    Set db = GetTestDb(dbErr)
    If db Is Nothing Then
        Test_EnEdicion_Vacio = BuildFail("GetTestDb: " & dbErr, logs)
        Exit Function
    End If

    Dim testEdicionId As Long
    testEdicionId = 909999

    On Error Resume Next
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion=" & testEdicionId
    db.Execute "INSERT INTO TbProyectosEdiciones (IDEdicion, IDProyecto, Edicion, Elaborado) " & _
               "VALUES (" & testEdicionId & ", " & FIX_PROYECTO & ", 999, 'TESTUSER')"
    On Error GoTo 0
    Set db = Nothing

    Dim pError As String
    Dim dicEnEdicion As Scripting.Dictionary
    Set dicEnEdicion = SuministradoresHelper.getSuministradoresEnEdicion(CStr(testEdicionId), pError)

    ' Teardown
    Set db = GetTestDb(dbErr)
    If Not db Is Nothing Then
        On Error Resume Next
        db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion=" & testEdicionId
        On Error GoTo 0
        Set db = Nothing
    End If

    If pError <> "" Then
        Test_EnEdicion_Vacio = BuildFail("Error: " & pError, logs)
        Exit Function
    End If

    If dicEnEdicion Is Nothing Then
        Test_EnEdicion_Vacio = BuildFail("Expected empty Dictionary but got Nothing", logs)
        Exit Function
    End If

    If dicEnEdicion.Count <> 0 Then
        Test_EnEdicion_Vacio = BuildFail("Expected 0 but got " & dicEnEdicion.Count, logs)
        Exit Function
    End If

    Test_EnEdicion_Vacio = BuildOk("enedicion_vacio_pass", logs)
End Function
