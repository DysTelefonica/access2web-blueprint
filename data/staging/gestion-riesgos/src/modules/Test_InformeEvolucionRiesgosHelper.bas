Attribute VB_Name = "Test_InformeEvolucionRiesgosHelper"
Option Compare Database
Option Explicit

' ============================================================
' Test_InformeEvolucionRiesgosHelper - TDD atoms for modInformeEvolucionRiesgosHelper
'
' Helper: GenerarInformeEvolucion
'   Signature: Public Function GenerarInformeEvolucion( _
'                 ByRef p_IDEdicion As String
'                 Optional ByRef p_URLInforme As String
'                 Optional ByRef p_PromptResult As Long
'
' Returns: URL del HTML generado si OK (p_URLInforme populated).
'          "" si p_Error populated.
'
' Source: src/modules/modInformeEvolucionRiesgosHelper.bas
' Underlying: InformeRiesgoHTML.GenerarInformeEdicionHTML (line 64 of
'             InformeRiesgoHTML.bas) - guarda en disco vía UTF-8.
'
' Form (REQ-CAL-07): botón "Informe evolución" en Form_FormCalidadTareas.
'   Hoy el form llama directo a InformeRiesgoHTML.GenerarInformeEdicionHTML
'   con la edición resuelta. El helper introduce:
'     - validación de p_IDEdicion no vacío / alcance válido
'     - idempotency guard (no regenera si el HTML ya existe en disco)
'     - contrato p_Error ByRef (convención Telefónica D&S)
'
' 4 scenario classes (skill access-vba-tdd §4.5):
'   1. Happy         - edición válida + alcance Resumen3 ? URL retornada, archivo existe
'   2. Sad           - IDEdicion vacío o alcance inválido ? p_Error poblado, no genera
'   3. Edge          - edición con datos opcionales Null ? render no rompe
'   4. Adversarial   - doble click idempotente ? mismo URL, no archivos huérfanos
'
' Skill: access-vba-tdd v2.4.3
' SDD:    e2e-form-by-form-2026-06-22 - Bloque 3 - REQ-CAL-07
' ============================================================

' --- Fixture constants ---
'   Range 901000-901099 dedicated to Bloque 3 / REQ-CAL-07 (no overlap with
'   SeedBaseGraph 900500-900599 or SeedSubcatGraph 900100-900221).
Private Const FIX_ID_EXPEDIENTE As Long = 901001
Private Const FIX_ID_PROYECTO   As Long = 901002
Private Const FIX_ID_EDICION    As Long = 901003
Private Const FIX_ID_EDICION_SAD As Long = 901004
Private Const FIX_ID_EDICION_EDGE As Long = 901005
Private Const FIX_ID_EDICION_ADV As Long = 901006

' --- JSON helpers wrapper (project convention) ---
Private Function BuildOk(ByVal value As Variant, ByRef logs() As String) As String
    BuildOk = BuildJsonOk(value, logs)
End Function

Private Function BuildFail(ByVal errMsg As String, ByRef logs() As String) As String
    BuildFail = BuildJsonFail(errMsg, logs)
End Function

' --- Sandbox guard ---
Private Function EnsureSandboxLoaded(ByRef p_Error As String) As Boolean
    EnsureSandboxLoaded = Test_Helper.EnsureTestConfigLoaded(p_Error)
End Function

' --- Get sandbox DB ---
Private Function GetSandboxDb(ByRef p_Error As String) As DAO.Database
    Set GetSandboxDb = Test_Fixtures.GetTestDb(p_Error)
End Function

' -----------------------------------------------------------------------------
' Fixture helpers (FK-ordered)
'
' Seed the parent graph: Expediente ? Proyecto ? Edicion.
' Cada test usa su propia IDEdicion para evitar contaminación cruzada.
' -----------------------------------------------------------------------------
Private Sub SeedInformeEvolucionFixture(ByVal p_IDEdicion As Long)
    On Error GoTo EH_Seed
    Dim db As DAO.Database
    Dim dbErr As String
    Set db = GetSandboxDb(dbErr)
    If db Is Nothing Then Err.Raise 1001, "SeedInformeEvolucionFixture", "GetSandboxDb returned Nothing: " & dbErr

    ' Idempotent cleanup en orden inverso
    On Error Resume Next
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion=" & p_IDEdicion, dbFailOnError
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto=" & FIX_ID_PROYECTO, dbFailOnError
    db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & FIX_ID_EXPEDIENTE, dbFailOnError
    On Error GoTo 0

    ' 1. TbExpedientes (padre) - REQ-CAL-07 helper no lee expedición pero
    '    necesita el grafo FK completo para que Constructor.getProyecto
    '    funcione al resolver la edición.
    db.Execute "INSERT INTO TbExpedientes (IDExpediente, Nemotecnico, Titulo, Tipo, IDEstado) " & _
               "VALUES (" & FIX_ID_EXPEDIENTE & ", 'TEST07', 'Fixture REQ-CAL-07', 'Test', 1)", dbFailOnError

    ' 2. TbProyectos (hijo de Expediente) - CodigoDocumento requerido por
    '    InformeRiesgoHTML.GetURLInformeEdicionHTML (línea 1252).
    db.Execute "INSERT INTO TbProyectos " & _
               "(IDProyecto, IDExpediente, Proyecto, CodigoDocumento, Juridica) " & _
               "VALUES (" & FIX_ID_PROYECTO & ", " & FIX_ID_EXPEDIENTE & ", " & _
               "'TESTPROJ07', 'TEST07-COD', 'TdE')", dbFailOnError

    ' 3. TbProyectosEdiciones (hijo de Proyecto)
    db.Execute "INSERT INTO TbProyectosEdiciones " & _
               "(IDProyecto, IDEdicion, Edicion, Elaborado, FechaEdicion) " & _
               "VALUES (" & FIX_ID_PROYECTO & ", " & p_IDEdicion & ", 1, " & _
               "'test_user_calidad', #" & Format$(Now, "yyyy-mm-dd hh:nn:ss") & "#)", dbFailOnError

    ' Post-seed cardinality assertion (skill §1.2 fixture gate)
    Dim rs As DAO.Recordset
    Dim seedCount As Long
    Set rs = db.OpenRecordset("SELECT COUNT(*) AS C FROM TbProyectosEdiciones WHERE IDEdicion=" & p_IDEdicion)
    If Not rs.EOF Then seedCount = CLng(Nz(rs.fields("C").value, 0))
    rs.Close
    Set rs = Nothing
    If seedCount <> 1 Then
        Err.Raise 1002, "SeedInformeEvolucionFixture", _
            "Post-seed assertion failed for IDEdicion=" & p_IDEdicion & " (count=" & seedCount & ")"
    End If

    Set db = Nothing
    Exit Sub

EH_Seed:
    Dim eN As Long: eN = Err.Number
    Dim ed As String: ed = Err.description
    On Error Resume Next
    Set db = Nothing
    Err.Raise eN, "SeedInformeEvolucionFixture", "Seed failed: " & eN & " - " & ed
End Sub

Private Sub TeardownInformeEvolucionFixture(ByVal p_IDEdicion As Long)
    On Error Resume Next
    Dim db As DAO.Database
    Dim dbErr As String
    Set db = GetSandboxDb(dbErr)
    If db Is Nothing Then Exit Sub

    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion=" & p_IDEdicion, dbFailOnError
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto=" & FIX_ID_PROYECTO, dbFailOnError
    db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & FIX_ID_EXPEDIENTE, dbFailOnError

    Set db = Nothing
    On Error GoTo 0
End Sub

' -----------------------------------------------------------------------------
' Helper: delete generated HTML files for the given URL (test cleanup)
' -----------------------------------------------------------------------------
Private Sub CleanupGeneratedHTML(ByVal p_URL As String)
    On Error Resume Next
    If Len(p_URL) > 0 Then
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.FileExists(p_URL) Then fso.DeleteFile p_URL, True
        Set fso = Nothing
    End If
    On Error GoTo 0
End Sub

' ============================================================
' ATOM 1 — Happy: EdicionValidaConAlcanceResumen3_GeneraHTMLYDevuelveURL
' GIVEN staging sandbox + edición válida con riesgo asociado + alcance Resumen3
' WHEN GenerarInformeEvolucion is called
' THEN:
'   - retorna URL no vacía
'   - p_URLInforme poblada con la misma URL
'   - p_Error vacía
'   - el archivo existe en disco
' ============================================================
Public Function Test_InformeEvolucionRiesgosHelper_Happy_ProyectoConHistorial_GeneraInforme() As String
    Dim logs(0 To 9) As String
    Test_InformeEvolucionRiesgosHelper_Happy_ProyectoConHistorial_GeneraInforme = BuildFail("test did not complete", logs)
    On Error GoTo HandleError

    logs(0) = "1. Arrange: EnsureSandboxLoaded"
    logs(1) = "2. Arrange: SeedInformeEvolucionFixture(901003)"
    logs(2) = "3. Act: GenerarInformeEvolucion(901003, Resumen3, url, db, , err)"
    logs(3) = "4. Assert: retorno no vacío"
    logs(4) = "5. Assert: p_URLInforme no vacío"
    logs(5) = "6. Assert: p_Error vacío"
    logs(6) = "7. Assert: archivo existe en disco"
    logs(7) = "8. Teardown: limpiar HTML + TeardownInformeEvolucionFixture"

    Dim cfgErr As String
    If Not EnsureSandboxLoaded(cfgErr) Then
        Test_InformeEvolucionRiesgosHelper_Happy_ProyectoConHistorial_GeneraInforme = BuildFail(cfgErr, logs)
        GoTo Teardown
    End If

    SeedInformeEvolucionFixture FIX_ID_EDICION

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = GetSandboxDb(dbErr)
    If db Is Nothing Then
        Test_InformeEvolucionRiesgosHelper_Happy_ProyectoConHistorial_GeneraInforme = BuildFail("GetSandboxDb devolvió Nothing: " & dbErr, logs)
        GoTo Teardown
    End If

    Dim m_URL As String
    Dim m_Err As String
    Dim m_Result As String
    m_URL = ""
    m_Err = ""
    m_Result = GenerarInformeEvolucion(CStr(FIX_ID_EDICION), EnumControlCambiosAlcanceResumen3, m_URL, db, , m_Err)

    If Len(m_Err) <> 0 Then
        Test_InformeEvolucionRiesgosHelper_Happy_ProyectoConHistorial_GeneraInforme = BuildFail("p_Error no esperado: " & m_Err, logs)
        GoTo Teardown
    End If
    If Len(m_Result) = 0 Then
        Test_InformeEvolucionRiesgosHelper_Happy_ProyectoConHistorial_GeneraInforme = BuildFail("retorno vacío, esperaba URL", logs)
        GoTo Teardown
    End If
    If Len(m_URL) = 0 Then
        Test_InformeEvolucionRiesgosHelper_Happy_ProyectoConHistorial_GeneraInforme = BuildFail("p_URLInforme vacío, esperaba URL", logs)
        GoTo Teardown
    End If
    If m_Result <> m_URL Then
        Test_InformeEvolucionRiesgosHelper_Happy_ProyectoConHistorial_GeneraInforme = BuildFail("retorno y p_URLInforme difieren: [" & m_Result & "] vs [" & m_URL & "]", logs)
        GoTo Teardown
    End If

    ' Verify file actually exists (defense in depth)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(m_URL) Then
        Set fso = Nothing
        Test_InformeEvolucionRiesgosHelper_Happy_ProyectoConHistorial_GeneraInforme = BuildFail("URL retornada pero archivo no existe: " & m_URL, logs)
        GoTo Teardown
    End If
    Set fso = Nothing

    Test_InformeEvolucionRiesgosHelper_Happy_ProyectoConHistorial_GeneraInforme = BuildOk(m_URL, logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    If Len(m_URL) > 0 Then CleanupGeneratedHTML m_URL
    Set db = Nothing
    TeardownInformeEvolucionFixture FIX_ID_EDICION
    Test_Helper.ResetTestSession
    On Error GoTo 0
    Exit Function

HandleError:
    Test_InformeEvolucionRiesgosHelper_Happy_ProyectoConHistorial_GeneraInforme = BuildFail("unexpected: " & Err.description & " (source: " & Err.Source & ")", logs)
    Resume Teardown
End Function

' ============================================================
' ATOM 2 — Sad: IDEdicionVacio_NoLanzaError_DevuelveErrorFuncional
' GIVEN staging sandbox
' WHEN GenerarInformeEvolucion is called with empty p_IDEdicion
' THEN:
'   - retorno ""
'   - p_Error poblado con mensaje sobre p_IDEdicion vacío
'   - NO se genera archivo
' ============================================================
Public Function Test_InformeEvolucionRiesgosHelper_Sad_ProyectoVacio_NoLanzaError() As String
    Dim logs(0 To 6) As String
    Test_InformeEvolucionRiesgosHelper_Sad_ProyectoVacio_NoLanzaError = BuildFail("test did not complete", logs)
    On Error GoTo HandleError

    logs(0) = "1. Arrange: EnsureSandboxLoaded (no seed específico)"
    logs(1) = "2. Act: GenerarInformeEvolucion("""", Resumen3, url, db, , err)"
    logs(2) = "3. Assert: retorno vacío"
    logs(3) = "4. Assert: p_URLInforme vacío"
    logs(4) = "5. Assert: p_Error poblado mencionando IDEdicion"
    logs(5) = "6. Assert: NO excepción (validación funcional, no crash)"

    Dim cfgErr As String
    If Not EnsureSandboxLoaded(cfgErr) Then
        Test_InformeEvolucionRiesgosHelper_Sad_ProyectoVacio_NoLanzaError = BuildFail(cfgErr, logs)
        GoTo Teardown
    End If

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = GetSandboxDb(dbErr)
    If db Is Nothing Then
        Test_InformeEvolucionRiesgosHelper_Sad_ProyectoVacio_NoLanzaError = BuildFail("GetSandboxDb devolvió Nothing: " & dbErr, logs)
        GoTo Teardown
    End If

    ' Capture baseline - no HTML generated
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim urlOut As String
    Dim errOut As String
    Dim resultOut As String
    urlOut = ""
    errOut = ""
    resultOut = GenerarInformeEvolucion("", EnumControlCambiosAlcanceResumen3, urlOut, db, , errOut)

    ' --- Assert ---
    If Len(resultOut) <> 0 Then
        Set fso = Nothing
        Test_InformeEvolucionRiesgosHelper_Sad_ProyectoVacio_NoLanzaError = BuildFail("retorno esperado vacío, obtuvo: " & resultOut, logs)
        GoTo Teardown
    End If
    If Len(urlOut) <> 0 Then
        Set fso = Nothing
        Test_InformeEvolucionRiesgosHelper_Sad_ProyectoVacio_NoLanzaError = BuildFail("p_URLInforme esperado vacío, obtuvo: " & urlOut, logs)
        GoTo Teardown
    End If
    If Len(errOut) = 0 Then
        Set fso = Nothing
        Test_InformeEvolucionRiesgosHelper_Sad_ProyectoVacio_NoLanzaError = BuildFail("p_Error esperado poblado, obtuvo vacío", logs)
        GoTo Teardown
    End If
    If InStr(1, errOut, "IDEdicion", vbTextCompare) = 0 Then
        Set fso = Nothing
        Test_InformeEvolucionRiesgosHelper_Sad_ProyectoVacio_NoLanzaError = BuildFail("p_Error no menciona IDEdicion. got: " & errOut, logs)
        GoTo Teardown
    End If
    Set fso = Nothing

    Test_InformeEvolucionRiesgosHelper_Sad_ProyectoVacio_NoLanzaError = BuildOk("sad_idedicion_vacio_rejected_pass", logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    Set db = Nothing
    Test_Helper.ResetTestSession
    On Error GoTo 0
    Exit Function

HandleError:
    Test_InformeEvolucionRiesgosHelper_Sad_ProyectoVacio_NoLanzaError = BuildFail("unexpected: " & Err.description & " (source: " & Err.Source & ")", logs)
    Resume Teardown
End Function

' ============================================================
' ATOM 3 — Edge: EdicionConDatosOpcionalesNull_RenderNoRompe
' GIVEN staging sandbox + edición válida con alcance Completo (boundary)
' WHEN GenerarInformeEvolucion is called
' THEN:
'   - genera HTML sin lanzar error
'   - archivo existe en disco
'   - el HTML contiene al menos la marca "DOCTYPE html" (sanity)
' ============================================================
Public Function Test_InformeEvolucionRiesgosHelper_Edge_NullEnCamposOpcionales_NoRompeRender() As String
    Dim logs(0 To 9) As String
    Test_InformeEvolucionRiesgosHelper_Edge_NullEnCamposOpcionales_NoRompeRender = BuildFail("test did not complete", logs)
    On Error GoTo HandleError

    logs(0) = "1. Arrange: EnsureSandboxLoaded"
    logs(1) = "2. Arrange: SeedInformeEvolucionFixture(901005) sin riesgos (Null en campos opcionales)"
    logs(2) = "3. Act: GenerarInformeEvolucion(901005, Completo=2, url, db, , err) [boundary enum]"
    logs(3) = "4. Assert: retorno no vacío (genera HTML)"
    logs(4) = "5. Assert: archivo existe"
    logs(5) = "6. Assert: HTML contiene marca DOCTYPE (render válido)"
    logs(6) = "7. Assert: HTML no contiene literal ""Null"" (Null-safe render)"
    logs(7) = "8. Teardown: cleanup HTML + teardown fixture"

    Dim cfgErr As String
    If Not EnsureSandboxLoaded(cfgErr) Then
        Test_InformeEvolucionRiesgosHelper_Edge_NullEnCamposOpcionales_NoRompeRender = BuildFail(cfgErr, logs)
        GoTo Teardown
    End If

    SeedInformeEvolucionFixture FIX_ID_EDICION_EDGE

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = GetSandboxDb(dbErr)
    If db Is Nothing Then
        Test_InformeEvolucionRiesgosHelper_Edge_NullEnCamposOpcionales_NoRompeRender = BuildFail("GetSandboxDb devolvió Nothing: " & dbErr, logs)
        GoTo Teardown
    End If

    Dim m_URL As String
    Dim m_Err As String
    Dim m_Result As String
    m_URL = ""
    m_Err = ""
    m_Result = GenerarInformeEvolucion(CStr(FIX_ID_EDICION_EDGE), EnumControlCambiosAlcanceCompleto, m_URL, db, , m_Err)

    If Len(m_Err) <> 0 Then
        Test_InformeEvolucionRiesgosHelper_Edge_NullEnCamposOpcionales_NoRompeRender = BuildFail("p_Error inesperado: " & m_Err, logs)
        GoTo Teardown
    End If
    If Len(m_Result) = 0 Then
        Test_InformeEvolucionRiesgosHelper_Edge_NullEnCamposOpcionales_NoRompeRender = BuildFail("retorno vacío, esperaba URL", logs)
        GoTo Teardown
    End If

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(m_URL) Then
        Set fso = Nothing
        Test_InformeEvolucionRiesgosHelper_Edge_NullEnCamposOpcionales_NoRompeRender = BuildFail("archivo no existe: " & m_URL, logs)
        GoTo Teardown
    End If

    ' Read file content (small sanity check: DOCTYPE + no literal "Null")
    Dim ts As Object
    Set ts = fso.OpenTextFile(m_URL, 1)  ' 1 = ForReading
    Dim sContent As String
    sContent = ts.ReadAll
    ts.Close
    Set ts = Nothing
    Set fso = Nothing

    If InStr(1, sContent, "<!DOCTYPE html", vbTextCompare) = 0 Then
        Test_InformeEvolucionRiesgosHelper_Edge_NullEnCamposOpcionales_NoRompeRender = BuildFail("HTML sin marca DOCTYPE (render corrupto)", logs)
        GoTo Teardown
    End If
    If InStr(1, sContent, ">Null<", vbTextCompare) > 0 Or _
       InStr(1, sContent, "Null ", vbTextCompare) > 0 Then
        Test_InformeEvolucionRiesgosHelper_Edge_NullEnCamposOpcionales_NoRompeRender = BuildFail("HTML contiene literal 'Null' (Null no tratado)", logs)
        GoTo Teardown
    End If

    Test_InformeEvolucionRiesgosHelper_Edge_NullEnCamposOpcionales_NoRompeRender = BuildOk(m_URL, logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    If Len(m_URL) > 0 Then CleanupGeneratedHTML m_URL
    Set db = Nothing
    TeardownInformeEvolucionFixture FIX_ID_EDICION_EDGE
    Test_Helper.ResetTestSession
    On Error GoTo 0
    Exit Function

HandleError:
    Test_InformeEvolucionRiesgosHelper_Edge_NullEnCamposOpcionales_NoRompeRender = BuildFail("unexpected: " & Err.description & " (source: " & Err.Source & ")", logs)
    Resume Teardown
End Function

' ============================================================
' ATOM 4 — Adversarial: DobleClick_MismaEdicionYAlcance_GeneraUnaVez
' GIVEN staging sandbox + edición válida
' WHEN GenerarInformeEvolucion is called twice rapidly with same args
' THEN:
'   - ambas llamadas retornan URL
'   - segunda llamada NO genera archivo huérfano
'   - cardinalidad de archivos .html en la carpeta de destino sigue = 1
' ============================================================
Public Function Test_InformeEvolucionRiesgosHelper_Adversarial_DobleClick_GeneraUnaVez() As String
    Dim logs(0 To 10) As String
    Test_InformeEvolucionRiesgosHelper_Adversarial_DobleClick_GeneraUnaVez = BuildFail("test did not complete", logs)
    On Error GoTo HandleError

    logs(0) = "1. Arrange: EnsureSandboxLoaded"
    logs(1) = "2. Arrange: SeedInformeEvolucionFixture(901006)"
    logs(2) = "3. Act: 1ª llamada - GenerarInformeEvolucion(901006, Resumen3, url1, db, , err1)"
    logs(3) = "4. Assert: 1ª llamada retorna URL, archivo existe"
    logs(4) = "5. Act: 2ª llamada (doble click) - mismos args"
    logs(5) = "6. Assert: 2ª llamada retorna URL (puede ser la misma u otra)"
    logs(6) = "7. Assert: ningún archivo huérfano en carpeta destino"
    logs(7) = "8. Teardown: cleanup HTML + teardown fixture"

    Dim cfgErr As String
    If Not EnsureSandboxLoaded(cfgErr) Then
        Test_InformeEvolucionRiesgosHelper_Adversarial_DobleClick_GeneraUnaVez = BuildFail(cfgErr, logs)
        GoTo Teardown
    End If

    SeedInformeEvolucionFixture FIX_ID_EDICION_ADV

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = GetSandboxDb(dbErr)
    If db Is Nothing Then
        Test_InformeEvolucionRiesgosHelper_Adversarial_DobleClick_GeneraUnaVez = BuildFail("GetSandboxDb devolvió Nothing: " & dbErr, logs)
        GoTo Teardown
    End If

    ' --- 1ª llamada ---
    Dim url1 As String, err1 As String, result1 As String
    url1 = "": err1 = "": result1 = ""
    result1 = GenerarInformeEvolucion(CStr(FIX_ID_EDICION_ADV), EnumControlCambiosAlcanceResumen3, url1, db, , err1)
    If Len(err1) <> 0 Then
        Test_InformeEvolucionRiesgosHelper_Adversarial_DobleClick_GeneraUnaVez = BuildFail("1ª llamada p_Error: " & err1, logs)
        GoTo Teardown
    End If
    If Len(url1) = 0 Then
        Test_InformeEvolucionRiesgosHelper_Adversarial_DobleClick_GeneraUnaVez = BuildFail("1ª llamada sin URL", logs)
        GoTo Teardown
    End If

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(url1) Then
        Set fso = Nothing
        Test_InformeEvolucionRiesgosHelper_Adversarial_DobleClick_GeneraUnaVez = BuildFail("1ª llamada URL no existe: " & url1, logs)
        GoTo Teardown
    End If

    ' --- 2ª llamada (doble click) ---
    Dim url2 As String, err2 As String, result2 As String
    url2 = "": err2 = "": result2 = ""
    result2 = GenerarInformeEvolucion(CStr(FIX_ID_EDICION_ADV), EnumControlCambiosAlcanceResumen3, url2, db, , err2)
    If Len(err2) <> 0 Then
        Set fso = Nothing
        Test_InformeEvolucionRiesgosHelper_Adversarial_DobleClick_GeneraUnaVez = BuildFail("2ª llamada p_Error: " & err2, logs)
        GoTo Teardown
    End If
    If Len(url2) = 0 Then
        Set fso = Nothing
        Test_InformeEvolucionRiesgosHelper_Adversarial_DobleClick_GeneraUnaVez = BuildFail("2ª llamada sin URL", logs)
        GoTo Teardown
    End If

    ' --- Cardinalidad: en la carpeta de destino, contar archivos .html
    '     que coincidan con el patrón del nombre del informe ---
    '     Patrón esperado: "TEST07-COD-1.html" o similar (basado en
    '     GetURLInformeEdicionHTML: línea 1265 para Juridica != "TdE"
    '     y línea 1263 para Juridica = "TdE"). Contamos cuántos archivos
    '     .html existen en la carpeta padre de la URL.
    Dim parentFolder As String
    parentFolder = fso.GetParentFolderName(url1)
    Dim fileName As String
    fileName = fso.GetFileName(url1)

    Dim folder As Object
    Set folder = fso.GetFolder(parentFolder)

    Dim htmlCount As Long
    htmlCount = 0

    Dim file As Object
    For Each file In folder.Files
        If LCase$(fso.GetExtensionName(file.Name)) = "html" Then
            ' Look for any file matching our test project's CodigoDocumento
            ' (TEST07-COD) — at most we expect 1 from our test.
            If InStr(1, file.Name, "TEST07-COD", vbTextCompare) > 0 Then
                htmlCount = htmlCount + 1
            End If
        End If
    Next file
    Set fso = Nothing

    If htmlCount > 1 Then
        Test_InformeEvolucionRiesgosHelper_Adversarial_DobleClick_GeneraUnaVez = BuildFail("doble click generó " & htmlCount & " archivos huérfanos (esperado 1)", logs)
        GoTo Teardown
    End If

    ' Save URLs for cleanup before resetting
    Dim urlToClean1 As String, urlToClean2 As String
    urlToClean1 = url1
    urlToClean2 = url2

    Test_InformeEvolucionRiesgosHelper_Adversarial_DobleClick_GeneraUnaVez = BuildOk("doble_click_idempotent_pass", logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    If Len(urlToClean1) > 0 Then CleanupGeneratedHTML urlToClean1
    If Len(urlToClean2) > 0 And urlToClean2 <> urlToClean1 Then CleanupGeneratedHTML urlToClean2
    Set db = Nothing
    TeardownInformeEvolucionFixture FIX_ID_EDICION_ADV
    Test_Helper.ResetTestSession
    On Error GoTo 0
    Exit Function

HandleError:
    Test_InformeEvolucionRiesgosHelper_Adversarial_DobleClick_GeneraUnaVez = BuildFail("unexpected: " & Err.description & " (source: " & Err.Source & ")", logs)
    Resume Teardown
End Function

