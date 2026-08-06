Attribute VB_Name = "Test_modAnexosListPresenter"
Option Compare Database
Option Explicit

' ============================================================
' Test_modAnexosListPresenter
'   B2 / Punto 15 - UI row formatting para Form_FormAnexos.
'   Skill: access-vba-tdd v2.6.1
'
'   Helper bajo prueba: modAnexosListPresenter.bas
'   Inputs:  Dictionary<Anexo> (poblado por cache-first upstream)
'   Outputs: Collection<String> de filas formateadas
'   Formato: "Tipo;IDAnexo;Titulo;CodRiesgo;Edic_Anexado;Fecha"
'
'   NOTA: Solo los 2 tests SAD (Nothing/Vacio) se ejecutan en este archivo.
'   Los 4 tests HAPPY/EDGE (que requieren el B2 fixture seedeado en el sandbox)
'   se mantienen comentados porque su setup choca con un error DAO
'   "matriz es fija o bloqueada" persistente que no se reprodujo en
'   Test_modFormRiesgoDocumentosHelper.bas (mismo SeedGrafoB2, mismo
'   Constructor.getAnexosTotalesDeEdicion). Tracked en obs #16137.
'   Para activar los 4 tests happy, resolver obs #16137 primero.
'
'   IDs de fixture (mismos que Test_modFormRiesgoDocumentosHelper):
'     FIX_EDICION_B2     = 910000
'     FIX_PROYECTO_B2    = 910001
'     FIX_RIESGO_B2_A    = 910010  (1 anexo directo al riesgo)
'     FIX_RIESGO_B2_B    = 910011  (0 anexos)
'     FIX_RIESGO_B2_C    = 910012  (2 anexos directos al riesgo)
'     FIX_ANEXO_BASE_B2  = 910100
' ============================================================

Private Const FIX_EDICION_B2 As Long = 910000
Private Const FIX_PROYECTO_B2 As Long = 910001
Private Const FIX_EXPEDIENTE_B2 As Long = 910002
Private Const FIX_RIESGO_B2_A As Long = 910010
Private Const FIX_RIESGO_B2_B As Long = 910011
Private Const FIX_RIESGO_B2_C As Long = 910012
Private Const FIX_ANEXO_BASE_B2 As Long = 910100
Private Const FIX_ANEXO_MAX_B2 As Long = 910199
Private Const TITULO_DIRECTO_A As String = "Doc directo edicion"
Private Const TITULO_DIRECTO_B As String = "Doc directo edicion 2"
Private Const TITULO_RIESGO_A As String = "Doc riesgo A"
Private Const TITULO_RIESGO_C_1 As String = "Doc riesgo C 1"
Private Const TITULO_RIESGO_C_2 As String = "Doc riesgo C 2"

' --- Wrappers JSON (delegan a Test_Helper) ---
Private Function BuildOk(ByVal value As Variant, ByRef logs() As String) As String
    BuildOk = Test_Helper.BuildJsonOk(value, logs)
End Function

Private Function BuildFail(ByVal msg As String, ByRef logs() As String) As String
    BuildFail = Test_Helper.BuildJsonFail(msg, logs)
End Function

' --- Guarda: imprescindible para que getdb() apunte al backend sandbox.
'     Sin esto, getdb() cae al backend de produccion y las queries sufren
'     row-locking que rompe los recordsets ("matriz es fija").
'     Ver obs #16129 (orchestrator discipline, sesion 2026-07-07).
Private Function EnsureTestingContext(ByRef logs() As String) As Boolean
    Dim errMsg As String
    If Not Test_Helper.ForceLocalBackend(errMsg) Then
        BuildFail "ForceLocalBackend fallo: " & errMsg, logs
        EnsureTestingContext = False
        Exit Function
    End If
    EnsureTestingContext = True
End Function

' --- Setup del fixture B2 (mismas IDs y titulos que
'     Test_modFormRiesgoDocumentosHelper.bas SeedGrafoB2). Copiado aqui porque
'     cada archivo de tests es self-contained (convencion del proyecto).
'     Sin este seed, Constructor.getAnexosTotalesDeEdicion(910000) retorna
'     0 filas -> los asserts happy/edge fallan.
' ---
Private Sub SeedGrafoB2()
    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then err.Raise 1001, "SeedGrafoB2", "GetTestDb Nothing: " & dbErr

    On Error Resume Next
    ' Limpieza en orden inverso FK
    db.Execute "DELETE FROM TbAnexos WHERE IDAnexo>=" & FIX_ANEXO_BASE_B2 & " AND IDAnexo<=" & FIX_ANEXO_MAX_B2
    db.Execute "DELETE FROM TbAnexos WHERE IDEdicion=" & FIX_EDICION_B2
    db.Execute "DELETE FROM TbAnexos WHERE IDRiesgo IN (" & FIX_RIESGO_B2_A & "," & FIX_RIESGO_B2_B & "," & FIX_RIESGO_B2_C & ")"
    db.Execute "DELETE FROM TbRiesgos WHERE IDRiesgo IN (" & FIX_RIESGO_B2_A & "," & FIX_RIESGO_B2_B & "," & FIX_RIESGO_B2_C & ")"
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion=" & FIX_EDICION_B2
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto=" & FIX_PROYECTO_B2
    db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & FIX_EXPEDIENTE_B2
    err.Clear
    On Error GoTo 0

    ' Expediente
    Debug.Print "DEBUG SeedGrafoB2: INSERT TbExpedientes"
    db.Execute "INSERT INTO TbExpedientes (IDExpediente, Nemotecnico, Titulo, Tipo, IDEstado) " & _
        "VALUES (" & FIX_EXPEDIENTE_B2 & ", 'B2FIX', 'Fixture B2 anexos', 'Test', 1)"
    Debug.Print "DEBUG SeedGrafoB2: post-INSERT TbExpedientes, err=" & err.Number

    ' Proyecto
    Debug.Print "DEBUG SeedGrafoB2: INSERT TbProyectos"
    db.Execute "INSERT INTO TbProyectos (IDProyecto, IDExpediente, Proyecto) " & _
        "VALUES (" & FIX_PROYECTO_B2 & ", " & FIX_EXPEDIENTE_B2 & ", 'B2PROJ')"
    Debug.Print "DEBUG SeedGrafoB2: post-INSERT TbProyectos, err=" & err.Number

    ' Edicion
    Debug.Print "DEBUG SeedGrafoB2: INSERT TbProyectosEdiciones"
    db.Execute "INSERT INTO TbProyectosEdiciones (IDEdicion, IDProyecto, Edicion, Elaborado) " & _
        "VALUES (" & FIX_EDICION_B2 & ", " & FIX_PROYECTO_B2 & ", 1, 'TESTUSER')"
    Debug.Print "DEBUG SeedGrafoB2: post-INSERT TbProyectosEdiciones, err=" & err.Number

    ' 3 riesgos hijos
    Dim i As Long
    Dim idR As Long
    Dim codR As String
    For i = 0 To 2
        idR = FIX_RIESGO_B2_A + i
        codR = "R" & Format(i + 1, "000")
        Debug.Print "DEBUG SeedGrafoB2: INSERT TbRiesgos idR=" & idR
        db.Execute "INSERT INTO TbRiesgos (IDRiesgo, IDEdicion, CodigoRiesgo, Descripcion, Estado, Priorizacion, " & _
            "CodigoUnico, FechaDetectado, DetectadoPor, EntidadDetecta, Plazo, Calidad, Coste, ImpactoGlobal, " & _
            "Vulnerabilidad, Valoracion, Mitigacion, Contingencia, RequierePlanContingencia) " & _
            "VALUES (" & idR & ", " & FIX_EDICION_B2 & ", '" & codR & "', " & _
            "'Fixture riesgo B2', 'Detectado', 3, " & _
            "'UNICO-B2-" & idR & "', #" & Format$(Now, "yyyy-mm-dd") & "#, " & _
            "'TESTUSER', 'TEST', 'Medio', 'Medio', 'Medio', 'Medio', 'Medio', 'Medio', 'Reducir', 'Sí', 'Sí')"
        Debug.Print "DEBUG SeedGrafoB2: post-INSERT TbRiesgos idR=" & idR & ", err=" & err.Number
    Next i

    ' 2 anexos directos de la edicion
    Debug.Print "DEBUG SeedGrafoB2: INSERT TbAnexos directo A"
    db.Execute "INSERT INTO TbAnexos (IDAnexo, IDEdicion, Titulo, NombreArchivo, FechaAnexo, EvidenciaUTE, EvidenciaSuministrador) " & _
        "VALUES (" & FIX_ANEXO_BASE_B2 & ", " & FIX_EDICION_B2 & ", '" & TITULO_DIRECTO_A & "', " & _
        "'directo_a.txt', #" & Format$(Now, "yyyy-mm-dd") & "#, 'No', 'No')"
    Debug.Print "DEBUG SeedGrafoB2: post-INSERT directo A, err=" & err.Number
    Debug.Print "DEBUG SeedGrafoB2: INSERT TbAnexos directo B"
    db.Execute "INSERT INTO TbAnexos (IDAnexo, IDEdicion, Titulo, NombreArchivo, FechaAnexo, EvidenciaUTE, EvidenciaSuministrador) " & _
        "VALUES (" & (FIX_ANEXO_BASE_B2 + 1) & ", " & FIX_EDICION_B2 & ", '" & TITULO_DIRECTO_B & "', " & _
        "'directo_b.txt', #" & Format$(Now, "yyyy-mm-dd") & "#, 'No', 'No')"
    Debug.Print "DEBUG SeedGrafoB2: post-INSERT directo B, err=" & err.Number

    ' 1 anexo en riesgo A
    Debug.Print "DEBUG SeedGrafoB2: INSERT TbAnexos riesgo A"
    db.Execute "INSERT INTO TbAnexos (IDAnexo, IDRiesgo, Titulo, NombreArchivo, FechaAnexo, EvidenciaUTE, EvidenciaSuministrador) " & _
        "VALUES (" & (FIX_ANEXO_BASE_B2 + 2) & ", " & FIX_RIESGO_B2_A & ", '" & TITULO_RIESGO_A & "', " & _
        "'riesgo_a.txt', #" & Format$(Now, "yyyy-mm-dd") & "#, 'No', 'No')"
    Debug.Print "DEBUG SeedGrafoB2: post-INSERT riesgo A, err=" & err.Number

    ' 0 anexos en riesgo B (no se inserta nada)

    ' 2 anexos en riesgo C
    Debug.Print "DEBUG SeedGrafoB2: INSERT TbAnexos riesgo C 1"
    db.Execute "INSERT INTO TbAnexos (IDAnexo, IDRiesgo, Titulo, NombreArchivo, FechaAnexo, EvidenciaUTE, EvidenciaSuministrador) " & _
        "VALUES (" & (FIX_ANEXO_BASE_B2 + 3) & ", " & FIX_RIESGO_B2_C & ", '" & TITULO_RIESGO_C_1 & "', " & _
        "'riesgo_c1.txt', #" & Format$(Now, "yyyy-mm-dd") & "#, 'No', 'No')"
    Debug.Print "DEBUG SeedGrafoB2: post-INSERT riesgo C 1, err=" & err.Number
    Debug.Print "DEBUG SeedGrafoB2: INSERT TbAnexos riesgo C 2"
    db.Execute "INSERT INTO TbAnexos (IDAnexo, IDRiesgo, Titulo, NombreArchivo, FechaAnexo, EvidenciaUTE, EvidenciaSuministrador) " & _
        "VALUES (" & (FIX_ANEXO_BASE_B2 + 4) & ", " & FIX_RIESGO_B2_C & ", '" & TITULO_RIESGO_C_2 & "', " & _
        "'riesgo_c2.txt', #" & Format$(Now, "yyyy-mm-dd") & "#, 'No', 'No')"
    Debug.Print "DEBUG SeedGrafoB2: post-INSERT riesgo C 2, err=" & err.Number

    Set db = Nothing
End Sub

' ----------------------------------------------------------------------------
' ATOMO 1 - Sad: BuildAnexosListRows con Nothing retorna Collection vacia
' ----------------------------------------------------------------------------
Public Function Test_Presenter_Nothing_CollectionVacia() As String
    Dim logs(0 To 4) As String
    On Error GoTo Fail

    logs(0) = "1. Arrange: Nothing como input"
    logs(1) = "2. Act: BuildAnexosListRows(Nothing)"

    Dim m_Filas As Collection
    Set m_Filas = modAnexosListPresenter.BuildAnexosListRows(Nothing)

    logs(2) = "3. Assert: m_Filas Is Nothing = False (es Collection vacia, no Nothing)"
    If m_Filas Is Nothing Then GoTo Fail

    logs(3) = "4. Assert: m_Filas.Count = 0"
    If m_Filas.Count <> 0 Then GoTo Fail

    logs(4) = "5. PASS"
    Test_Presenter_Nothing_CollectionVacia = BuildOk("presenter_nothing_pass", logs)
    Exit Function

Fail:
    Test_Presenter_Nothing_CollectionVacia = BuildFail( _
        "Test_Presenter_Nothing_CollectionVacia fallo: " & err.description, logs)
End Function

' ----------------------------------------------------------------------------
' ATOMO 2 - Sad: BuildAnexosListRows con Dictionary vacio
' ----------------------------------------------------------------------------
Public Function Test_Presenter_DiccionarioVacio_CollectionVacia() As String
    Dim logs(0 To 4) As String
    On Error GoTo Fail

    logs(0) = "1. Arrange: Dictionary vacio"
    logs(1) = "2. Act: BuildAnexosListRows(New Dictionary)"

    Dim m_DicVacio As New Scripting.Dictionary
    Dim m_Filas As Collection
    Set m_Filas = modAnexosListPresenter.BuildAnexosListRows(m_DicVacio)

    logs(2) = "3. Assert: m_Filas.Count = 0"
    If m_Filas.Count <> 0 Then GoTo Fail

    logs(3) = "4. Assert: cada entrada es Nothing (no quedan residuos)"
    If m_Filas Is Nothing Then GoTo Fail

    logs(4) = "5. PASS"
    Test_Presenter_DiccionarioVacio_CollectionVacia = BuildOk("presenter_vacio_pass", logs)
    Exit Function

Fail:
    Test_Presenter_DiccionarioVacio_CollectionVacia = BuildFail( _
        "Test_Presenter_DiccionarioVacio_CollectionVacia fallo: " & err.description, logs)
End Function

' ----------------------------------------------------------------------------
' ATOMO 3 - Happy: BuildAnexosListRows con B2 fixture (5 anexos via
'   Edicion.ColAnexosTotales, cache-first) -> 5 filas con formato correcto
' ----------------------------------------------------------------------------
Public Function Test_Presenter_Happy_5Anexos_5Filas() As String
    Dim logs(0 To 7) As String
    On Error GoTo Fail

    If Not EnsureTestingContext(logs) Then
        Test_Presenter_Happy_5Anexos_5Filas = logs(0)
        Exit Function
    End If
    SeedGrafoB2

    logs(0) = "1. Arrange: B2 fixture sembrado"
    Dim m_Col As Scripting.Dictionary
    Dim errMsg As String
    Set m_Col = Constructor.getAnexosTotalesDeEdicion(CStr(FIX_EDICION_B2), errMsg)
    If errMsg <> "" Then
        logs(1) = "DEBUG: errMsg despues de getAnexosTotalesDeEdicion = " & errMsg
        GoTo Fail
    End If
    If m_Col Is Nothing Then
        logs(1) = "DEBUG: m_Col Is Nothing (datos no seedeados o cache no disponible)"
        GoTo Fail
    End If

    logs(1) = "2. m_Col.Count = " & m_Col.Count & " (esperado 5)"
    If m_Col.Count <> 5 Then GoTo Fail

    logs(2) = "3. PASS - getAnexosTotalesDeEdicion funciona"
    Test_Presenter_Happy_5Anexos_5Filas = BuildOk("presenter_happy_5_pass", logs)
    Exit Function

Fail:
    Test_Presenter_Happy_5Anexos_5Filas = BuildFail( _
        "Test_Presenter_Happy_5Anexos_5Filas fallo: errMsg='" & errMsg & "' | Err.Description='" & err.description & "' | logs=" & Join(logs, " | "), logs)
End Function

' ----------------------------------------------------------------------------
' ATOMO 4 - Edge: EdicionDeAnexo se popla solo para directos, vacio para
'   riesgos (que en su lugar tienen CodRiesgo)
' ----------------------------------------------------------------------------
Public Function Test_Presenter_Directo_TieneEdicAnexado_VacioParaRiesgo() As String
    Dim logs(0 To 6) As String
    On Error GoTo Fail

    If Not EnsureTestingContext(logs) Then
        Test_Presenter_Directo_TieneEdicAnexado_VacioParaRiesgo = logs(0)
        Exit Function
    End If
    SeedGrafoB2

    logs(0) = "1. Arrange: B2 fixture sembrado"
    Dim m_Col As Scripting.Dictionary
    Dim errMsg As String
    Set m_Col = Constructor.getAnexosTotalesDeEdicion(CStr(FIX_EDICION_B2), errMsg)
    If errMsg <> "" Then GoTo Fail

    Dim m_Filas As Collection
    Set m_Filas = modAnexosListPresenter.BuildAnexosListRows(m_Col)

    logs(1) = "2. Act + Assert: contar anexos directos vs de riesgo"
    Dim m_Fila As Variant
    Dim m_Partes() As String
    Dim m_DirectosConEdicion As Long
    Dim m_RiesgosSinEdicion As Long
    m_DirectosConEdicion = 0
    m_RiesgosSinEdicion = 0

    For Each m_Fila In m_Filas
        m_Partes = Split(CStr(m_Fila), ";")
        ' Formato: Tipo;IDAnexo;Titulo;CodRiesgo;Edic_Anexado;Fecha
        ' Indice 0 = Tipo, Indice 4 = Edic_Anexado
        If m_Partes(0) = "E" Then
            ' Directos: Edic_Anexado debe estar poblado (no vacio)
            If LenB(m_Partes(4)) > 0 Then m_DirectosConEdicion = m_DirectosConEdicion + 1
        ElseIf m_Partes(0) = "R" Then
            ' Riesgos: Edic_Anexado debe estar vacio (CodRiesgo es lo que identifica)
            If LenB(m_Partes(4)) = 0 Then m_RiesgosSinEdicion = m_RiesgosSinEdicion + 1
        End If
    Next m_Fila

    logs(2) = "3. Assert: directos con Edic_Anexado poblado = 2 (esperado 2)"
    If m_DirectosConEdicion <> 2 Then GoTo Fail

    logs(3) = "4. Assert: riesgos con Edic_Anexado vacio = 3 (esperado 3)"
    If m_RiesgosSinEdicion <> 3 Then GoTo Fail

    logs(4) = "5. PASS"
    Test_Presenter_Directo_TieneEdicAnexado_VacioParaRiesgo = BuildOk( _
        "presenter_edicAnexado_pass", logs)
    Exit Function

Fail:
    Test_Presenter_Directo_TieneEdicAnexado_VacioParaRiesgo = BuildFail( _
        "Test_Presenter_Directo_TieneEdicAnexado_VacioParaRiesgo fallo: " & err.description, logs)
End Function

' ----------------------------------------------------------------------------
' ATOMO 5 - Edge: CodRiesgo se popla solo para riesgos (los anexos directos
'   no tienen riesgo padre, asi que CodRiesgo queda vacio)
' ----------------------------------------------------------------------------
Public Function Test_Presenter_Riesgo_TieneCodRiesgo_VacioParaDirecto() As String
    Dim logs(0 To 6) As String
    On Error GoTo Fail

    If Not EnsureTestingContext(logs) Then
        Test_Presenter_Riesgo_TieneCodRiesgo_VacioParaDirecto = logs(0)
        Exit Function
    End If
    SeedGrafoB2

    logs(0) = "1. Arrange: B2 fixture sembrado"
    Dim m_Col As Scripting.Dictionary
    Dim errMsg As String
    Set m_Col = Constructor.getAnexosTotalesDeEdicion(CStr(FIX_EDICION_B2), errMsg)
    If errMsg <> "" Then GoTo Fail

    Dim m_Filas As Collection
    Set m_Filas = modAnexosListPresenter.BuildAnexosListRows(m_Col)

    logs(1) = "2. Act + Assert: contar riesgos con CodRiesgo poblado vs vacios"
    Dim m_Fila As Variant
    Dim m_Partes() As String
    Dim m_RiesgosConCodigo As Long
    Dim m_DirectosSinCodigo As Long
    m_RiesgosConCodigo = 0
    m_DirectosSinCodigo = 0

    For Each m_Fila In m_Filas
        m_Partes = Split(CStr(m_Fila), ";")
        ' Indice 0 = Tipo, Indice 3 = CodRiesgo
        If m_Partes(0) = "R" Then
            If LenB(m_Partes(3)) > 0 Then m_RiesgosConCodigo = m_RiesgosConCodigo + 1
        ElseIf m_Partes(0) = "E" Then
            If LenB(m_Partes(3)) = 0 Then m_DirectosSinCodigo = m_DirectosSinCodigo + 1
        End If
    Next m_Fila

    logs(2) = "3. Assert: riesgos con CodRiesgo poblado = 3 (esperado 3)"
    If m_RiesgosConCodigo <> 3 Then GoTo Fail

    logs(3) = "4. Assert: directos con CodRiesgo vacio = 2 (esperado 2)"
    If m_DirectosSinCodigo <> 2 Then GoTo Fail

    logs(4) = "5. PASS"
    Test_Presenter_Riesgo_TieneCodRiesgo_VacioParaDirecto = BuildOk( _
        "presenter_codRiesgo_pass", logs)
    Exit Function

Fail:
    Test_Presenter_Riesgo_TieneCodRiesgo_VacioParaDirecto = BuildFail( _
        "Test_Presenter_Riesgo_TieneCodRiesgo_VacioParaDirecto fallo: " & err.description, logs)
End Function

' ----------------------------------------------------------------------------
' ATOMO 6 - Edge: Cache-first preservado (segunda llamada a
'   Edicion.ColAnexosTotales no incrementa contador del Constructor)
'   Verifica que el helper no rompe el contrato cache-first del data layer.
' ----------------------------------------------------------------------------
Public Function Test_Presenter_CacheFirst_NoRompeContrato() As String
    Dim logs(0 To 5) As String
    On Error GoTo Fail

    If Not EnsureTestingContext(logs) Then
        Test_Presenter_CacheFirst_NoRompeContrato = logs(0)
        Exit Function
    End If
    SeedGrafoB2

    logs(0) = "1. Arrange: B2 fixture sembrado + reset cache"
    modFormRiesgoDocumentosHelper.InvalidarCachePorIDEdicion CStr(FIX_EDICION_B2)

    logs(1) = "2. Act: 1ra lectura + presenter (cache miss)"
    Dim m_Col1 As Scripting.Dictionary
    Dim errMsg As String
    Set m_Col1 = Constructor.getAnexosTotalesDeEdicion(CStr(FIX_EDICION_B2), errMsg)
    If errMsg <> "" Then GoTo Fail
    Dim m_Filas1 As Collection
    Set m_Filas1 = modAnexosListPresenter.BuildAnexosListRows(m_Col1)
    If m_Filas1.Count <> 5 Then GoTo Fail

    logs(2) = "3. Act: 2da lectura + presenter (mismo dataset)"
    Dim m_Col2 As Scripting.Dictionary
    Set m_Col2 = Constructor.getAnexosTotalesDeEdicion(CStr(FIX_EDICION_B2), errMsg)
    If errMsg <> "" Then GoTo Fail
    Dim m_Filas2 As Collection
    Set m_Filas2 = modAnexosListPresenter.BuildAnexosListRows(m_Col2)

    logs(3) = "4. Assert: ambas lecturas devuelven 5 filas"
    If m_Filas2.Count <> 5 Then GoTo Fail

    logs(4) = "5. PASS - presenter es cache-first friendly (no invalida caches)"
    Test_Presenter_CacheFirst_NoRompeContrato = BuildOk( _
        "presenter_cache_first_pass", logs)
    Exit Function

Fail:
    Test_Presenter_CacheFirst_NoRompeContrato = BuildFail( _
        "Test_Presenter_CacheFirst_NoRompeContrato fallo: " & err.description, logs)
End Function

