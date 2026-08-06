Attribute VB_Name = "Test_InformeHTMLPresenter"
Option Compare Database
Option Explicit

' ============================================================
' Test_InformeHTMLPresenter - TDD atoms for modInformeHTMLPresenter
'
' Helpers:
'   ObtenerEncabezadoFecha(ByVal p_TipoInforme As String) As String
'   ObtenerPrefijoFechaCelda(ByVal p_TipoInforme As String,
'                             ByVal p_FechaPublicacion As String,
'                             ByVal p_FechaEdicion As String) As String
'
' Helper es PURE (no DAO, no CurrentDb, no recordsets). Per
' access-vba-e2e-methodology §hard rule 2: la firma no incluye `db` porque
' removerlo no cambiaría el output. Por la misma razón, los átomos no
' invocan GetTestDb ni sembrado de BD.
'
' Skill: access-vba-tdd v2.4.3
' SDD:    e2e-form-by-form-2026-06-22 - Bloque 4 - Refactor inline - T-R.1
'
' Fixture IDs: N/A (helper pure, sin fixtures DB).
'
' Cobertura 4/4 escenarios (skill §4.7):
'   1. Happy       — ObtenerEncabezadoFecha("EdicionHistorico") retorna
'                    "Fecha pub. / creación".
'   2. Sad         — ObtenerEncabezadoFecha("TipoDesconocido") retorna
'                    default genérico "Fecha" sin fallar.
'   3. Edge        — ObtenerPrefijoFechaCelda con FechaPublicacion poblada
'                    retorna "Pub. ".
'   4. Adversarial — ObtenerPrefijoFechaCelda con AMBAS fechas vacías /
'                    no-fecha retorna prefijo "" (no error).
' ============================================================

' --- Constantes de contrato (literales que el helper DEBE retornar) ---
'     Documentadas como constantes porque cualquier refactor que las cambie
'     rompe la compatibilidad con los call sites (InformeRiesgoHTML.bas)
'     y con los UAT cards de aceptación.
'
'     Históricamente también las consumía ExcelInforme.bas, pero ese módulo
'     se eliminó: la generación de Excel para publicación ya no existe.
Private Const ENCABEZADO_EDICION_HISTORICO As String = "Fecha pub. / creación"
Private Const ENCABEZADO_GENERICO As String = "Fecha"
Private Const PREFIJO_PUB As String = "Pub. "
Private Const PREFIJO_CREAC As String = "Creac. "

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

' ============================================================
' ATOM 1 — Happy: encabezado del Cuadro de Control (EdicionHistorico)
' GIVEN helper PURE
' WHEN  ObtenerEncabezadoFecha("EdicionHistorico")
' THEN  retorna "Fecha pub. / creación" (literal del contrato)
'
' Mapea al call site InformeRiesgoHTML.bas línea ~437
' ("Fecha pub. / creación").
'
' (Históricamente también mapeaba a ExcelInforme.bas línea 755, que era el
' upper-case del mismo encabezado; ese módulo se eliminó.)
' ============================================================
Public Function Test_InformeHTMLPresenter_Happy_EdicionHistorico_RetornaEncabezadoCompuesto() As String
    Dim logs(0 To 4) As String
    Test_InformeHTMLPresenter_Happy_EdicionHistorico_RetornaEncabezadoCompuesto = _
        BuildFail("test did not complete", logs)
    On Error GoTo HandleError

    logs(0) = "1. Arrange: helper PURE (sin DB)"
    logs(1) = "2. Act: ObtenerEncabezadoFecha(""EdicionHistorico"")"
    logs(2) = "3. Assert: retorno no vacío"
    logs(3) = "4. Assert: retorno = ""Fecha pub. / creación"" (literal del contrato)"

    Dim m_Result As String
    m_Result = ObtenerEncabezadoFecha("EdicionHistorico")

    If Len(m_Result) = 0 Then
        logs(3) = "4. Assert FAIL: retorno vacío, esperaba 'Fecha pub. / creación'"
        Test_InformeHTMLPresenter_Happy_EdicionHistorico_RetornaEncabezadoCompuesto = _
            BuildFail("retorno vacío, esperaba 'Fecha pub. / creación'", logs)
        Exit Function
    End If

    If m_Result <> ENCABEZADO_EDICION_HISTORICO Then
        logs(3) = "4. Assert FAIL: esperaba '" & ENCABEZADO_EDICION_HISTORICO & _
                  "', obtuvo: '" & m_Result & "'"
        Test_InformeHTMLPresenter_Happy_EdicionHistorico_RetornaEncabezadoCompuesto = _
            BuildFail("esperaba '" & ENCABEZADO_EDICION_HISTORICO & _
                      "', obtuvo: '" & m_Result & "'", logs)
        Exit Function
    End If

    logs(3) = "4. Assert PASS: header 'Fecha pub. / creación' = literal del contrato"
    Test_InformeHTMLPresenter_Happy_EdicionHistorico_RetornaEncabezadoCompuesto = _
        BuildOk(m_Result, logs)
    Exit Function

HandleError:
    Test_InformeHTMLPresenter_Happy_EdicionHistorico_RetornaEncabezadoCompuesto = _
        BuildFail("unexpected: " & Err.description & " (source: " & Err.Source & ")", logs)
End Function

' ============================================================
' ATOM 2 — Sad: tipo de informe desconocido retorna default genérico
' GIVEN helper PURE
' WHEN  ObtenerEncabezadoFecha("TipoInventado")
' THEN  retorna "Fecha" (default defensivo, no error)
'
' Defensa contra typos en los call sites. El helper es contractualmente
' "nunca vacío": si llega un tipo desconocido, entrega el header genérico
' (que es el mismo que retorna "RiesgoEstado"). Mapea al call site de la
' tabla principal de riesgos.
' ============================================================
Public Function Test_InformeHTMLPresenter_Sad_TipoDesconocido_RetornaGenerico() As String
    Dim logs(0 To 4) As String
    Test_InformeHTMLPresenter_Sad_TipoDesconocido_RetornaGenerico = _
        BuildFail("test did not complete", logs)
    On Error GoTo HandleError

    logs(0) = "1. Arrange: helper PURE (sin DB)"
    logs(1) = "2. Act: ObtenerEncabezadoFecha(""TipoInventado"")"
    logs(2) = "3. Assert: retorno no vacío (defensa explícita)"
    logs(3) = "4. Assert: retorno = ""Fecha"" (default defensivo, no error)"

    Dim m_Result As String
    m_Result = ObtenerEncabezadoFecha("TipoInventado")

    If Len(m_Result) = 0 Then
        logs(3) = "4. Assert FAIL: retorno vacío, default defensivo debería ser 'Fecha'"
        Test_InformeHTMLPresenter_Sad_TipoDesconocido_RetornaGenerico = _
            BuildFail("retorno vacío, default defensivo debería ser 'Fecha'", logs)
        Exit Function
    End If

    If m_Result <> ENCABEZADO_GENERICO Then
        logs(3) = "4. Assert FAIL: default esperaba '" & ENCABEZADO_GENERICO & _
                  "', obtuvo: '" & m_Result & "'"
        Test_InformeHTMLPresenter_Sad_TipoDesconocido_RetornaGenerico = _
            BuildFail("default esperaba '" & ENCABEZADO_GENERICO & _
                      "', obtuvo: '" & m_Result & "'", logs)
        Exit Function
    End If

    logs(3) = "4. Assert PASS: default defensivo 'Fecha' para tipo desconocido"
    Test_InformeHTMLPresenter_Sad_TipoDesconocido_RetornaGenerico = _
        BuildOk(m_Result, logs)
    Exit Function

HandleError:
    Test_InformeHTMLPresenter_Sad_TipoDesconocido_RetornaGenerico = _
        BuildFail("unexpected: " & Err.description & " (source: " & Err.Source & ")", logs)
End Function

' ============================================================
' ATOM 3 — Edge: FechaPublicacion poblada ? prefijo "Pub. "
' GIVEN helper PURE
' WHEN  ObtenerPrefijoFechaCelda("EdicionHistorico",
'                                "15/06/2026",
'                                "01/06/2026")
' THEN  retorna "Pub. " (FechaPublicacion tiene prioridad sobre FechaEdicion)
'
' Mapea al call site de InformeRiesgoHTML.bas línea ~448-451 (celda con
' ambas fechas pobladas: prefiere Pub.).
'
' (Históricamente también mapeaba a ExcelInforme.bas línea 835, rama de la
' edición actual publicada; ese módulo se eliminó.)
' ============================================================
Public Function Test_InformeHTMLPresenter_Edge_FechaPublicacionPoblada_RetornaPubPrefijo() As String
    Dim logs(0 To 5) As String
    Test_InformeHTMLPresenter_Edge_FechaPublicacionPoblada_RetornaPubPrefijo = _
        BuildFail("test did not complete", logs)
    On Error GoTo HandleError

    logs(0) = "1. Arrange: helper PURE"
    logs(1) = "2. Arrange: p_FechaPublicacion='15/06/2026' (poblada), " & _
              "p_FechaEdicion='01/06/2026' (también poblada)"
    logs(2) = "3. Act: ObtenerPrefijoFechaCelda(""EdicionHistorico"", " & _
              """15/06/2026"", ""01/06/2026"")"
    logs(3) = "4. Assert: retorno no vacío"
    logs(4) = "5. Assert: retorno = ""Pub. "" (FechaPublicacion tiene prioridad)"

    Dim m_Result As String
    m_Result = ObtenerPrefijoFechaCelda("EdicionHistorico", "15/06/2026", "01/06/2026")

    If Len(m_Result) = 0 Then
        logs(4) = "5. Assert FAIL: retorno vacío, esperaba 'Pub. '"
        Test_InformeHTMLPresenter_Edge_FechaPublicacionPoblada_RetornaPubPrefijo = _
            BuildFail("retorno vacío, esperaba 'Pub. '", logs)
        Exit Function
    End If

    If m_Result <> PREFIJO_PUB Then
        logs(4) = "5. Assert FAIL: esperaba '" & PREFIJO_PUB & _
                  "', obtuvo: '" & m_Result & "'"
        Test_InformeHTMLPresenter_Edge_FechaPublicacionPoblada_RetornaPubPrefijo = _
            BuildFail("esperaba '" & PREFIJO_PUB & "', obtuvo: '" & m_Result & "'", logs)
        Exit Function
    End If

    logs(4) = "5. Assert PASS: 'Pub. ' — FechaPublicacion prioriza sobre FechaEdicion"
    Test_InformeHTMLPresenter_Edge_FechaPublicacionPoblada_RetornaPubPrefijo = _
        BuildOk(m_Result, logs)
    Exit Function

HandleError:
    Test_InformeHTMLPresenter_Edge_FechaPublicacionPoblada_RetornaPubPrefijo = _
        BuildFail("unexpected: " & Err.description & " (source: " & Err.Source & ")", logs)
End Function

' ============================================================
' ATOM 4 — Adversarial: AMBAS fechas vacías/no-fecha ? prefijo ""
' GIVEN helper PURE
' WHEN  ObtenerPrefijoFechaCelda("EdicionHistorico", "", "no_es_fecha")
' THEN  retorna "" (prefijo vacío: la celda mostrará solo la fecha
'       formateada, o quedará vacía si Format$ recibe "")
'
' Defensa contra entradas hostiles: cadena vacía, Null implícito (""),
' string no-fecha. El helper NO debe retornar "Pub. " ni "Creac. " ni
' explotar. Mapea al call site de InformeRiesgoHTML.bas donde algunas
' ediciones pueden tener ambos campos en Null antes de ser publicadas.
' ============================================================
Public Function Test_InformeHTMLPresenter_Adversarial_AmbasFechasVacias_RetornaPrefijoVacio() As String
    Dim logs(0 To 6) As String
    Test_InformeHTMLPresenter_Adversarial_AmbasFechasVacias_RetornaPrefijoVacio = _
        BuildFail("test did not complete", logs)
    On Error GoTo HandleError

    logs(0) = "1. Arrange: helper PURE"
    logs(1) = "2. Arrange: p_FechaPublicacion='' (vacía), " & _
              "p_FechaEdicion='no_es_fecha' (string no-fecha)"
    logs(2) = "3. Act: ObtenerPrefijoFechaCelda(""EdicionHistorico"", """", ""no_es_fecha"")"
    logs(3) = "4. Assert: helper NO explotó (admite entradas vacías/no-fecha)"
    logs(4) = "5. Assert: retorno = """" (prefijo vacío)"
    logs(5) = "6. Assert: retorno NO es ""Pub. "" ni ""Creac. """

    Dim m_Result As String
    m_Result = ObtenerPrefijoFechaCelda("EdicionHistorico", "", "no_es_fecha")

    If Len(m_Result) <> 0 Then
        logs(5) = "6. Assert FAIL: esperaba prefijo vacío, obtuvo: '" & m_Result & "'"
        Test_InformeHTMLPresenter_Adversarial_AmbasFechasVacias_RetornaPrefijoVacio = _
            BuildFail("esperaba prefijo vacío, obtuvo: '" & m_Result & "'", logs)
        Exit Function
    End If

    If m_Result = PREFIJO_PUB Then
        logs(5) = "6. Assert FAIL: helper asignó 'Pub. ' a entrada vacía (debería ser '')"
        Test_InformeHTMLPresenter_Adversarial_AmbasFechasVacias_RetornaPrefijoVacio = _
            BuildFail("helper asignó 'Pub. ' a entrada vacía (debería ser '')", logs)
        Exit Function
    End If

    If m_Result = PREFIJO_CREAC Then
        logs(5) = "6. Assert FAIL: helper asignó 'Creac. ' a entrada no-fecha (debería ser '')"
        Test_InformeHTMLPresenter_Adversarial_AmbasFechasVacias_RetornaPrefijoVacio = _
            BuildFail("helper asignó 'Creac. ' a entrada no-fecha (debería ser '')", logs)
        Exit Function
    End If

    logs(4) = "5. Assert PASS: prefijo vacío para entradas vacías/no-fecha"
    logs(5) = "6. Assert PASS: NO es 'Pub. ' ni 'Creac. '"
    Test_InformeHTMLPresenter_Adversarial_AmbasFechasVacias_RetornaPrefijoVacio = _
        BuildOk("prefijo_vacio_ok", logs)
    Exit Function

HandleError:
    Test_InformeHTMLPresenter_Adversarial_AmbasFechasVacias_RetornaPrefijoVacio = _
        BuildFail("unexpected: " & Err.description & " (source: " & Err.Source & ")", logs)
End Function

