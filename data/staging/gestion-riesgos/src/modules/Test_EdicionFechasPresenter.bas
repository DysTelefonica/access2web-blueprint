Attribute VB_Name = "Test_EdicionFechasPresenter"
Option Compare Database
Option Explicit

' ============================================================
' Test_EdicionFechasPresenter - TDD atoms for modEdicionFechasPresenter
'
' Helper: ObtenerEtiquetaFecha
'   Signature: Public Function ObtenerEtiquetaFecha( _
'                 ByVal p_IDEdicion As String
'                 Optional ByRef db As DAO.Database = Nothing
'
' Helper es un presenter: dado un (p_IDEdicion, p_NombreCampo) más el
' estado persistido en TbProyectosEdiciones.FechaPublicacion, devuelve la
' etiqueta legible en español. Es semi-puro (lee BD), por lo que el
' parámetro `db` es REQUERIDO por skill (access-vba-e2e-methodology
' §hard rule 2): quitarlo cambiaría el output.
'
' Architecture:
'   - DB-driven helper: helper acepta `db` ByRef opcional; átomos TDD
'     inyectan sandbox via Test_Fixtures.GetTestDb().
'   - Schema-first: TbProyectosEdiciones.FechaPublicacion es dbDate (Date
'     type). Puede ser Null (edición aún no publicada) o Date válido
'     (edición publicada). El helper distingue los dos casos.
'   - El átomo adversarial (atom 4) usa DOS ediciones con estados
'     distintos para probar que la etiqueta varía según el estado de
'     CADA edición.
'
' Fixture IDs: 906001-906099 (sin overlap con:
'   - Bloque 2: 900500-900599, 50010+
'   - Bloque 3 / REQ-CAL-07: 901000-901099
'   - Bloque 3 / REQ-CAL-09/10: 902000-902099
'   - Bloque 3 / REQ-CAL-12: 904000-904099
'   - Bloque 4 / REQ-CAL-04: helper puro, sin fixtures DB
'   - Bloque 4 / REQ-CAL-05: 905001-905099
' )
'
' Skill: access-vba-tdd v2.4.3
' SDD:    e2e-form-by-form-2026-06-22 - Bloque 4 - REQ-CAL-08
' ============================================================

' --- Module-level constants for REQ-CAL-08 fixtures ---
'     El helper lee TbProyectosEdiciones.FechaPublicacion. Para cubrir
'     los 4 átomos necesitamos:
'       - 1 edición SIN FechaPublicacion (atoms 1, 2, sad)
'       - 1 edición CON FechaPublicacion (atom 3, edge)
'       - 2 ediciones en el atom 4 (adversarial): una con FechaPublicacion,
'         otra sin.
Private Const FIX_ID_EXPEDIENTE As Long = 906001
Private Const FIX_ID_PROYECTO   As Long = 906002

Private Const FIX_ID_EDICION_SIN_PUBLICAR  As Long = 906010
Private Const FIX_ID_EDICION_CON_PUBLICAR  As Long = 906011
Private Const FIX_ID_EDICION_ADV_SIN_PUB   As Long = 906012
Private Const FIX_ID_EDICION_ADV_CON_PUB   As Long = 906013

Private Const FIX_FECHA_PUBLICACION As String = "2026-06-15"   ' yyyy-mm-dd (date literal)

' --- JSON helpers wrapper (project convention) ---
Private Function BuildOk(ByVal value As Variant, ByRef logs() As String) As String
    BuildOk = Test_Helper.BuildJsonOk(value, logs)
End Function

Private Function BuildFail(ByVal msg As String, ByRef logs() As String) As String
    BuildFail = Test_Helper.BuildJsonFail(msg, logs)
End Function

' --- EnsureTestConfigLoaded delegation ---
Private Function EnsureTestConfigLoaded(ByRef p_Error As String) As Boolean
    EnsureTestConfigLoaded = Test_Helper.EnsureTestConfigLoaded(p_Error)
End Function

' --- Siembra del grafo padre (Expediente, Proyecto) compartido por
'     todos los átomos. Una sola vez.
Private Sub SeedE08ParentGraph(ByRef db As DAO.Database)
    On Error GoTo EH_Seed

    ' Limpieza previa (FK-inversa) — solo nuestras fixtures
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion IN (" & _
               FIX_ID_EDICION_SIN_PUBLICAR & "," & _
               FIX_ID_EDICION_CON_PUBLICAR & "," & _
               FIX_ID_EDICION_ADV_SIN_PUB & "," & _
               FIX_ID_EDICION_ADV_CON_PUB & ")", dbFailOnError
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto=" & FIX_ID_PROYECTO, dbFailOnError
    db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & FIX_ID_EXPEDIENTE, dbFailOnError

    ' 1. Expediente (padre raíz)
    db.Execute "INSERT INTO TbExpedientes (IDExpediente, Nemotecnico, Titulo, Tipo, IDEstado) " & _
               "VALUES (" & FIX_ID_EXPEDIENTE & ", 'TEST08', 'Fixture REQ-CAL-08', 'Test', 1)", _
               dbFailOnError

    ' 2. Proyecto (hijo de Expediente)
    db.Execute "INSERT INTO TbProyectos (IDProyecto, IDExpediente, Proyecto) " & _
               "VALUES (" & FIX_ID_PROYECTO & ", " & FIX_ID_EXPEDIENTE & ", 'TESTPROJ08')", _
               dbFailOnError

    Exit Sub

EH_Seed:
    Dim eN As Long: eN = Err.Number
    Dim ed As String: ed = Err.description
    Err.Raise eN, "SeedE08ParentGraph", "Seed failed: " & eN & " - " & ed
End Sub

' --- Siembra de una edición con FechaPublicacion opcional (Null o Date).
'     Schema-first: solo IDEdicion, IDProyecto, Edicion, Elaborado son
'     NOT NULL; FechaPublicacion puede ser Null (sin publicar) o Date.
Private Sub SeedE08Edicion(ByRef db As DAO.Database, _
                           ByVal p_IDEdicion As Long, _
                           ByVal p_ConPublicacion As Boolean)
    Dim m_SQL As String
    If p_ConPublicacion Then
        m_SQL = "INSERT INTO TbProyectosEdiciones " & _
                "(IDProyecto, IDEdicion, Edicion, Elaborado, FechaPublicacion) " & _
                "VALUES (" & FIX_ID_PROYECTO & ", " & p_IDEdicion & ", 1, 'test_user', #" & _
                FIX_FECHA_PUBLICACION & "#)"
    Else
        m_SQL = "INSERT INTO TbProyectosEdiciones " & _
                "(IDProyecto, IDEdicion, Edicion, Elaborado) " & _
                "VALUES (" & FIX_ID_PROYECTO & ", " & p_IDEdicion & ", 1, 'test_user')"
    End If
    db.Execute m_SQL, dbFailOnError
End Sub

' --- Teardown FK-inverso: limpia solo nuestras fixtures ---
Private Sub TeardownE08(ByRef db As DAO.Database)
    On Error Resume Next
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion IN (" & _
               FIX_ID_EDICION_SIN_PUBLICAR & "," & _
               FIX_ID_EDICION_CON_PUBLICAR & "," & _
               FIX_ID_EDICION_ADV_SIN_PUB & "," & _
               FIX_ID_EDICION_ADV_CON_PUB & ")", dbFailOnError
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto=" & FIX_ID_PROYECTO, dbFailOnError
    db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & FIX_ID_EXPEDIENTE, dbFailOnError
    On Error GoTo 0
End Sub

' ============================================================
' ATOM 1 — Happy: FechaPreparadaParaPublicar ? etiqueta fija
' GIVEN sandbox + edición (cualquier estado; la etiqueta es fija)
' WHEN  ObtenerEtiquetaFecha(FIX_ID_EDICION_SIN_PUBLICAR,
'                            "FechaPreparadaParaPublicar", db, err)
' THEN  retorna "Fecha propuesta para publicar" (literal),
'       p_Error vacío
'
' Esta etiqueta NO depende del estado de la edición. Es la fecha que
' Calidad/Técnico propusieron para publicar.
' ============================================================
Public Function Test_EdicionFechasPresenter_Happy_FechaPreparadaParaPublicar_RetornaEtiquetaCorrecta() As String
    Dim logs(0 To 6) As String
    Test_EdicionFechasPresenter_Happy_FechaPreparadaParaPublicar_RetornaEtiquetaCorrecta = _
        BuildFail("test did not complete", logs)
    On Error GoTo HandleError

    logs(0) = "1. Arrange: EnsureTestConfigLoaded"
    logs(1) = "2. Arrange: SeedE08ParentGraph + SeedE08Edicion(SIN publicacion)"
    logs(2) = "3. Act: ObtenerEtiquetaFecha('FechaPreparadaParaPublicar')"
    logs(3) = "4. Assert: retorno no vacío"
    logs(4) = "5. Assert: retorno = 'Fecha propuesta para publicar' (literal)"
    logs(5) = "6. Assert: p_Error vacío"

    Dim cfgErr As String
    If Not EnsureTestConfigLoaded(cfgErr) Then
        Test_EdicionFechasPresenter_Happy_FechaPreparadaParaPublicar_RetornaEtiquetaCorrecta = _
            BuildFail(cfgErr, logs)
        GoTo Teardown
    End If

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_EdicionFechasPresenter_Happy_FechaPreparadaParaPublicar_RetornaEtiquetaCorrecta = _
            BuildFail("GetTestDb Nothing: " & dbErr, logs)
        GoTo Teardown
    End If

    SeedE08ParentGraph db
    SeedE08Edicion db, FIX_ID_EDICION_SIN_PUBLICAR, False

    Dim m_Result As String
    Dim m_Err As String
    m_Result = ObtenerEtiquetaFecha(CStr(FIX_ID_EDICION_SIN_PUBLICAR), _
                                     "FechaPreparadaParaPublicar", db, m_Err)

    If Len(m_Err) <> 0 Then
        logs(5) = "6. Assert FAIL: p_Error no esperado: " & m_Err
        Test_EdicionFechasPresenter_Happy_FechaPreparadaParaPublicar_RetornaEtiquetaCorrecta = _
            BuildFail("p_Error no esperado: " & m_Err, logs)
        GoTo Teardown
    End If
    If Len(m_Result) = 0 Then
        logs(4) = "5. Assert FAIL: retorno vacío, esperaba 'Fecha propuesta para publicar'"
        Test_EdicionFechasPresenter_Happy_FechaPreparadaParaPublicar_RetornaEtiquetaCorrecta = _
            BuildFail("retorno vacío, esperaba 'Fecha propuesta para publicar'", logs)
        GoTo Teardown
    End If
    If m_Result <> "Fecha propuesta para publicar" Then
        logs(4) = "5. Assert FAIL: esperaba 'Fecha propuesta para publicar', obtuvo: '" & m_Result & "'"
        Test_EdicionFechasPresenter_Happy_FechaPreparadaParaPublicar_RetornaEtiquetaCorrecta = _
            BuildFail("esperaba 'Fecha propuesta para publicar', obtuvo: '" & m_Result & "'", logs)
        GoTo Teardown
    End If

    logs(5) = "6. Assert PASS: 'Fecha propuesta para publicar' (literal)"
    Test_EdicionFechasPresenter_Happy_FechaPreparadaParaPublicar_RetornaEtiquetaCorrecta = _
        BuildOk(m_Result, logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    If Not db Is Nothing Then TeardownE08 db
    Set db = Nothing
    Test_Helper.ResetTestSession
    On Error GoTo 0
    Exit Function

HandleError:
    Test_EdicionFechasPresenter_Happy_FechaPreparadaParaPublicar_RetornaEtiquetaCorrecta = _
        BuildFail("unexpected: " & Err.description & " (source: " & Err.Source & ")", logs)
    Resume Teardown
End Function

' ============================================================
' ATOM 2 — Sad: FechaCierre sin publicación ? "Fecha última publicación"
' GIVEN sandbox + edición SIN FechaPublicacion (Null)
' WHEN  ObtenerEtiquetaFecha(FIX_ID_EDICION_SIN_PUBLICAR,
'                            "FechaCierre", db, err)
' THEN  retorna "Fecha última publicación" (genérico, sin publicación
'       efectiva), p_Error vacío
'
' Sad porque la edición aún no fue publicada; el helper degrada
' elegantemente a un label genérico en lugar de afirmar algo falso.
' ============================================================
Public Function Test_EdicionFechasPresenter_Sad_FechaCierreSinPublicacion_RetornaGenerico() As String
    Dim logs(0 To 6) As String
    Test_EdicionFechasPresenter_Sad_FechaCierreSinPublicacion_RetornaGenerico = _
        BuildFail("test did not complete", logs)
    On Error GoTo HandleError

    logs(0) = "1. Arrange: EnsureTestConfigLoaded"
    logs(1) = "2. Arrange: SeedE08ParentGraph + SeedE08Edicion(SIN publicacion)"
    logs(2) = "3. Act: ObtenerEtiquetaFecha('FechaCierre') sobre edición sin publicar"
    logs(3) = "4. Assert: retorno = 'Fecha última publicación' (genérico)"
    logs(4) = "5. Assert: p_Error vacío"

    Dim cfgErr As String
    If Not EnsureTestConfigLoaded(cfgErr) Then
        Test_EdicionFechasPresenter_Sad_FechaCierreSinPublicacion_RetornaGenerico = _
            BuildFail(cfgErr, logs)
        GoTo Teardown
    End If

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_EdicionFechasPresenter_Sad_FechaCierreSinPublicacion_RetornaGenerico = _
            BuildFail("GetTestDb Nothing: " & dbErr, logs)
        GoTo Teardown
    End If

    SeedE08ParentGraph db
    SeedE08Edicion db, FIX_ID_EDICION_SIN_PUBLICAR, False

    Dim m_Result As String
    Dim m_Err As String
    m_Result = ObtenerEtiquetaFecha(CStr(FIX_ID_EDICION_SIN_PUBLICAR), _
                                     "FechaCierre", db, m_Err)

    If Len(m_Err) <> 0 Then
        logs(4) = "5. Assert FAIL: p_Error no esperado: " & m_Err
        Test_EdicionFechasPresenter_Sad_FechaCierreSinPublicacion_RetornaGenerico = _
            BuildFail("p_Error no esperado: " & m_Err, logs)
        GoTo Teardown
    End If
    If m_Result <> "Fecha última publicación" Then
        logs(3) = "4. Assert FAIL: esperaba 'Fecha última publicación', obtuvo: '" & m_Result & "'"
        Test_EdicionFechasPresenter_Sad_FechaCierreSinPublicacion_RetornaGenerico = _
            BuildFail("esperaba 'Fecha última publicación', obtuvo: '" & m_Result & "'", logs)
        GoTo Teardown
    End If

    logs(4) = "5. Assert PASS: 'Fecha última publicación' (genérico sin publicación efectiva)"
    Test_EdicionFechasPresenter_Sad_FechaCierreSinPublicacion_RetornaGenerico = _
        BuildOk(m_Result, logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    If Not db Is Nothing Then TeardownE08 db
    Set db = Nothing
    Test_Helper.ResetTestSession
    On Error GoTo 0
    Exit Function

HandleError:
    Test_EdicionFechasPresenter_Sad_FechaCierreSinPublicacion_RetornaGenerico = _
        BuildFail("unexpected: " & Err.description & " (source: " & Err.Source & ")", logs)
    Resume Teardown
End Function

' ============================================================
' ATOM 3 — Edge: FechaCierre CON publicación ? "Fecha publicación edición actual"
' GIVEN sandbox + edición CON FechaPublicacion = FIX_FECHA_PUBLICACION
' WHEN  ObtenerEtiquetaFecha(FIX_ID_EDICION_CON_PUBLICAR,
'                            "FechaCierre", db, err)
' THEN  retorna "Fecha publicación edición actual", p_Error vacío
'
' Edge: el campo del form se llama "FechaCierre" pero su contenido es
' la fecha de la última publicación. La etiqueta refleja este
' contrato que el spec confirmó.
' ============================================================
Public Function Test_EdicionFechasPresenter_Edge_FechaCierreConPublicacion_RetornaPublicacionActual() As String
    Dim logs(0 To 6) As String
    Test_EdicionFechasPresenter_Edge_FechaCierreConPublicacion_RetornaPublicacionActual = _
        BuildFail("test did not complete", logs)
    On Error GoTo HandleError

    logs(0) = "1. Arrange: EnsureTestConfigLoaded"
    logs(1) = "2. Arrange: SeedE08ParentGraph + SeedE08Edicion(CON publicacion=" & FIX_FECHA_PUBLICACION & ")"
    logs(2) = "3. Act: ObtenerEtiquetaFecha('FechaCierre') sobre edición publicada"
    logs(3) = "4. Assert: retorno = 'Fecha publicación edición actual'"
    logs(4) = "5. Assert: p_Error vacío"

    Dim cfgErr As String
    If Not EnsureTestConfigLoaded(cfgErr) Then
        Test_EdicionFechasPresenter_Edge_FechaCierreConPublicacion_RetornaPublicacionActual = _
            BuildFail(cfgErr, logs)
        GoTo Teardown
    End If

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_EdicionFechasPresenter_Edge_FechaCierreConPublicacion_RetornaPublicacionActual = _
            BuildFail("GetTestDb Nothing: " & dbErr, logs)
        GoTo Teardown
    End If

    SeedE08ParentGraph db
    SeedE08Edicion db, FIX_ID_EDICION_CON_PUBLICAR, True

    Dim m_Result As String
    Dim m_Err As String
    m_Result = ObtenerEtiquetaFecha(CStr(FIX_ID_EDICION_CON_PUBLICAR), _
                                     "FechaCierre", db, m_Err)

    If Len(m_Err) <> 0 Then
        logs(4) = "5. Assert FAIL: p_Error no esperado: " & m_Err
        Test_EdicionFechasPresenter_Edge_FechaCierreConPublicacion_RetornaPublicacionActual = _
            BuildFail("p_Error no esperado: " & m_Err, logs)
        GoTo Teardown
    End If
    If m_Result <> "Fecha publicación edición actual" Then
        logs(3) = "4. Assert FAIL: esperaba 'Fecha publicación edición actual', obtuvo: '" & m_Result & "'"
        Test_EdicionFechasPresenter_Edge_FechaCierreConPublicacion_RetornaPublicacionActual = _
            BuildFail("esperaba 'Fecha publicación edición actual', obtuvo: '" & m_Result & "'", logs)
        GoTo Teardown
    End If

    logs(4) = "5. Assert PASS: 'Fecha publicación edición actual' (refleja fecha efectiva)"
    Test_EdicionFechasPresenter_Edge_FechaCierreConPublicacion_RetornaPublicacionActual = _
        BuildOk(m_Result, logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    If Not db Is Nothing Then TeardownE08 db
    Set db = Nothing
    Test_Helper.ResetTestSession
    On Error GoTo 0
    Exit Function

HandleError:
    Test_EdicionFechasPresenter_Edge_FechaCierreConPublicacion_RetornaPublicacionActual = _
        BuildFail("unexpected: " & Err.description & " (source: " & Err.Source & ")", logs)
    Resume Teardown
End Function

' ============================================================
' ATOM 4 — Adversarial: ColumnaTablaHistorico varía por estado
'                       de publicación de CADA edición
' GIVEN sandbox + DOS ediciones:
'       FIX_ID_EDICION_ADV_CON_PUB: CON FechaPublicacion
'       FIX_ID_EDICION_ADV_SIN_PUB: SIN FechaPublicacion
' WHEN  ObtenerEtiquetaFecha(<cada edicion>, "ColumnaTablaHistorico",
'                            db, err)
' THEN
'       - Edición CON publicación ? "Fecha publicación"
'       - Edición SIN publicación ? "Fecha creación edición"
'       - p_Error vacío en ambos casos
'
' Lock-in: la etiqueta del header de la columna histórica VARÍA por
' fila según el estado de esa edición específica. Esto cierra el
' contrato que el usuario reportó como bug: la columna "Fecha" no
' distinguía si el valor era FechaPublicacion o FechaEdicion.
' ============================================================
Public Function Test_EdicionFechasPresenter_Adversarial_ColumnaTablaHistorico_VariaPorEstadoPublicacion() As String
    Dim logs(0 To 9) As String
    Test_EdicionFechasPresenter_Adversarial_ColumnaTablaHistorico_VariaPorEstadoPublicacion = _
        BuildFail("test did not complete", logs)
    On Error GoTo HandleError

    logs(0) = "1. Arrange: EnsureTestConfigLoaded"
    logs(1) = "2. Arrange: SeedE08ParentGraph + 2 ediciones (CON pub + SIN pub)"
    logs(2) = "3. Act: ObtenerEtiquetaFecha('ColumnaTablaHistorico') para cada edición"
    logs(3) = "4. Assert: edición CON pub ? 'Fecha publicación'"
    logs(4) = "5. Assert: edición SIN pub ? 'Fecha creación edición'"
    logs(5) = "6. Assert: p_Error vacío en ambos casos"
    logs(6) = "7. Lock-in: la etiqueta varía por estado (cierra bug del usuario)"

    Dim cfgErr As String
    If Not EnsureTestConfigLoaded(cfgErr) Then
        Test_EdicionFechasPresenter_Adversarial_ColumnaTablaHistorico_VariaPorEstadoPublicacion = _
            BuildFail(cfgErr, logs)
        GoTo Teardown
    End If

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_EdicionFechasPresenter_Adversarial_ColumnaTablaHistorico_VariaPorEstadoPublicacion = _
            BuildFail("GetTestDb Nothing: " & dbErr, logs)
        GoTo Teardown
    End If

    SeedE08ParentGraph db
    SeedE08Edicion db, FIX_ID_EDICION_ADV_CON_PUB, True
    SeedE08Edicion db, FIX_ID_EDICION_ADV_SIN_PUB, False

    Dim m_ResultCon As String
    Dim m_ErrCon As String
    m_ResultCon = ObtenerEtiquetaFecha(CStr(FIX_ID_EDICION_ADV_CON_PUB), _
                                       "ColumnaTablaHistorico", db, m_ErrCon)

    If Len(m_ErrCon) <> 0 Then
        logs(5) = "6. Assert FAIL: p_Error no esperado (CON pub): " & m_ErrCon
        Test_EdicionFechasPresenter_Adversarial_ColumnaTablaHistorico_VariaPorEstadoPublicacion = _
            BuildFail("p_Error no esperado (CON pub): " & m_ErrCon, logs)
        GoTo Teardown
    End If
    If m_ResultCon <> "Fecha publicación" Then
        logs(3) = "4. Assert FAIL: edición CON pub esperaba 'Fecha publicación', obtuvo: '" & m_ResultCon & "'"
        Test_EdicionFechasPresenter_Adversarial_ColumnaTablaHistorico_VariaPorEstadoPublicacion = _
            BuildFail("CON pub esperaba 'Fecha publicación', obtuvo: '" & m_ResultCon & "'", logs)
        GoTo Teardown
    End If

    Dim m_ResultSin As String
    Dim m_ErrSin As String
    m_ResultSin = ObtenerEtiquetaFecha(CStr(FIX_ID_EDICION_ADV_SIN_PUB), _
                                       "ColumnaTablaHistorico", db, m_ErrSin)

    If Len(m_ErrSin) <> 0 Then
        logs(5) = "6. Assert FAIL: p_Error no esperado (SIN pub): " & m_ErrSin
        Test_EdicionFechasPresenter_Adversarial_ColumnaTablaHistorico_VariaPorEstadoPublicacion = _
            BuildFail("p_Error no esperado (SIN pub): " & m_ErrSin, logs)
        GoTo Teardown
    End If
    If m_ResultSin <> "Fecha creación edición" Then
        logs(4) = "5. Assert FAIL: edición SIN pub esperaba 'Fecha creación edición', obtuvo: '" & m_ResultSin & "'"
        Test_EdicionFechasPresenter_Adversarial_ColumnaTablaHistorico_VariaPorEstadoPublicacion = _
            BuildFail("SIN pub esperaba 'Fecha creación edición', obtuvo: '" & m_ResultSin & "'", logs)
        GoTo Teardown
    End If

    ' --- Lock-in: las dos etiquetas son distintas (cierra el bug) ---
    If m_ResultCon = m_ResultSin Then
        logs(6) = "7. Assert FAIL: etiquetas iguales; bug NO cerrado"
        Test_EdicionFechasPresenter_Adversarial_ColumnaTablaHistorico_VariaPorEstadoPublicacion = _
            BuildFail("etiquetas iguales; bug NO cerrado", logs)
        GoTo Teardown
    End If

    logs(6) = "7. Assert PASS: CON pub='" & m_ResultCon & "', SIN pub='" & m_ResultSin & "'"
    logs(7) = "8. Contrato cerrado: la columna 'Fecha' del histórico distingue por estado de publicación"
    Test_EdicionFechasPresenter_Adversarial_ColumnaTablaHistorico_VariaPorEstadoPublicacion = _
        BuildOk("adversarial_columna_varia_por_estado", logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    If Not db Is Nothing Then TeardownE08 db
    Set db = Nothing
    Test_Helper.ResetTestSession
    On Error GoTo 0
    Exit Function

HandleError:
    Test_EdicionFechasPresenter_Adversarial_ColumnaTablaHistorico_VariaPorEstadoPublicacion = _
        BuildFail("unexpected: " & Err.description & " (source: " & Err.Source & ")", logs)
    Resume Teardown
End Function


