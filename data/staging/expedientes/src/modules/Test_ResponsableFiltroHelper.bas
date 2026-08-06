Attribute VB_Name = "Test_ResponsableFiltroHelper"
Option Compare Database
Option Explicit

' -----------------------------------------------------------------------------
' Test_ResponsableFiltroHelper
'
' Atomos TDD para `modResponsableFiltroHelper`.  Como el helper es PURO
' (sin SQL, sin DAO.Database, sin Forms), no necesita temp `.accdb`.  Los
' tests construyen `Usuario` reales via New y los meten en un Dictionary
' para probar el mapeo.
'
' Cubre:
'   * BC=1 (combo expone solo el ID) vs BC=2 (combo expone "id;nombre")
'   * "Todos" sentinel (sin filtro)
'   * Empty / Null
'   * "0;N/A" sentinel UX de Seguridad
'   * "0" cuando BC=1 y el sentinel cae sobre la columna ID
'   * No-match (ID que no esta en el Dictionary)
' -----------------------------------------------------------------------------

' -----------------------------------------------------------------------------
' Atomos publicos (entrypoints del runner)
' -----------------------------------------------------------------------------

Public Function Test_ResponsableFiltro_EsTodos_TrueParaEmpty() As String
    Dim m_Logs(0 To 3) As String
    m_Logs(0) = "1. Arrange: empty string ''"
    m_Logs(1) = "2. Act: EsTodos('') -> esperado True"
    If Not modResponsableFiltroHelper.ResponsableFiltro_EsTodos("") Then
        Test_ResponsableFiltro_EsTodos_TrueParaEmpty = BuildFail("'' debio ser Todos", m_Logs, "")
        Exit Function
    End If
    m_Logs(2) = "3. Act: EsTodos(Null) -> esperado True"
    If Not modResponsableFiltroHelper.ResponsableFiltro_EsTodos(Null) Then
        Test_ResponsableFiltro_EsTodos_TrueParaEmpty = BuildFail("Null debio ser Todos", m_Logs, "")
        Exit Function
    End If
    m_Logs(3) = "4. Act: EsTodos('Todos') -> esperado True"
    If Not modResponsableFiltroHelper.ResponsableFiltro_EsTodos("Todos") Then
        Test_ResponsableFiltro_EsTodos_TrueParaEmpty = BuildFail("'Todos' debio ser Todos", m_Logs, "")
        Exit Function
    End If
    Test_ResponsableFiltro_EsTodos_TrueParaEmpty = BuildOk("ok", m_Logs)
End Function


Public Function Test_ResponsableFiltro_EsTodos_TrueParaSentinelSeguridad() As String
    Dim m_Logs(0 To 2) As String
    m_Logs(0) = "1. Arrange: sentinel '0;N/A' y '0'"
    m_Logs(1) = "2. Act: EsTodos('0;N/A') -> esperado True (UX Seguridad)"
    If Not modResponsableFiltroHelper.ResponsableFiltro_EsTodos("0;N/A") Then
        Test_ResponsableFiltro_EsTodos_TrueParaSentinelSeguridad = _
            BuildFail("'0;N/A' debio ser Todos", m_Logs, "")
        Exit Function
    End If
    m_Logs(2) = "3. Act: EsTodos('0') -> esperado True (BC=1 sentinel ID)"
    If Not modResponsableFiltroHelper.ResponsableFiltro_EsTodos("0") Then
        Test_ResponsableFiltro_EsTodos_TrueParaSentinelSeguridad = _
            BuildFail("'0' debio ser Todos", m_Logs, "")
        Exit Function
    End If
    Test_ResponsableFiltro_EsTodos_TrueParaSentinelSeguridad = BuildOk("ok", m_Logs)
End Function


Public Function Test_ResponsableFiltro_EsTodos_FalseParaIDReal() As String
    Dim m_Logs(0 To 2) As String
    m_Logs(0) = "1. Arrange: ID real '119'"
    m_Logs(1) = "2. Act: EsTodos('119') -> esperado False"
    If modResponsableFiltroHelper.ResponsableFiltro_EsTodos("119") Then
        Test_ResponsableFiltro_EsTodos_FalseParaIDReal = _
            BuildFail("'119' NO debio ser Todos", m_Logs, "")
        Exit Function
    End If
    m_Logs(2) = "3. Act: EsTodos('0;Algo') -> esperado False (no es el sentinel exacto)"
    If modResponsableFiltroHelper.ResponsableFiltro_EsTodos("0;Algo") Then
        Test_ResponsableFiltro_EsTodos_FalseParaIDReal = _
            BuildFail("'0;Algo' NO debio ser Todos", m_Logs, "")
        Exit Function
    End If
    Test_ResponsableFiltro_EsTodos_FalseParaIDReal = BuildOk("ok", m_Logs)
End Function


Public Function Test_ResponsableFiltro_ValorAId_BC1_IDPuro() As String
    Dim m_Logs(0 To 3) As String
    Dim m_Error As String
    Dim m_Result As String
    m_Logs(0) = "1. Arrange: combo expone solo ID '119' (BC=1)"
    m_Result = modResponsableFiltroHelper.ResponsableFiltro_ValorAId("119", m_Error)
    If m_Error <> "" Then
        Test_ResponsableFiltro_ValorAId_BC1_IDPuro = BuildFail("unexpected p_Error", m_Logs, m_Error)
        Exit Function
    End If
    m_Logs(1) = "2. Assert: result = '119'"
    If m_Result <> "119" Then
        Test_ResponsableFiltro_ValorAId_BC1_IDPuro = BuildFail("result='" & m_Result & "' esperaba '119'", m_Logs, "")
        Exit Function
    End If
    m_Logs(2) = "3. Arrange: combo expone '0;N/A' (sentinel UX)"
    m_Result = modResponsableFiltroHelper.ResponsableFiltro_ValorAId("0;N/A", m_Error)
    If m_Error <> "" Then
        Test_ResponsableFiltro_ValorAId_BC1_IDPuro = BuildFail("unexpected p_Error (sentinel)", m_Logs, m_Error)
        Exit Function
    End If
    m_Logs(3) = "4. Assert: result = '' (sentinel -> sin filtro)"
    If m_Result <> "" Then
        Test_ResponsableFiltro_ValorAId_BC1_IDPuro = BuildFail("sentinel result='" & m_Result & "' esperaba ''", m_Logs, "")
        Exit Function
    End If
    Test_ResponsableFiltro_ValorAId_BC1_IDPuro = BuildOk("ok", m_Logs)
End Function


Public Function Test_ResponsableFiltro_ValorAId_BC2_IDPuntoYComaNombre() As String
    Dim m_Logs(0 To 2) As String
    Dim m_Error As String
    Dim m_Result As String
    m_Logs(0) = "1. Arrange: combo expone '119;Ana Rubio Canales' (BC=2)"
    m_Result = modResponsableFiltroHelper.ResponsableFiltro_ValorAId( _
                    "119;Ana Rubio Canales", m_Error)
    If m_Error <> "" Then
        Test_ResponsableFiltro_ValorAId_BC2_IDPuntoYComaNombre = _
            BuildFail("unexpected p_Error", m_Logs, m_Error)
        Exit Function
    End If
    m_Logs(1) = "2. Assert: result = '119' (solo el ID)"
    If m_Result <> "119" Then
        Test_ResponsableFiltro_ValorAId_BC2_IDPuntoYComaNombre = _
            BuildFail("result='" & m_Result & "' esperaba '119'", m_Logs, "")
        Exit Function
    End If
    m_Logs(2) = "3. Assert: edge case '162;Torralba;Rodriguez' -> '162' (primer ';' como split)"
    m_Result = modResponsableFiltroHelper.ResponsableFiltro_ValorAId( _
                    "162;Torralba;Rodriguez", m_Error)
    If m_Result <> "162" Then
        Test_ResponsableFiltro_ValorAId_BC2_IDPuntoYComaNombre = _
            BuildFail("multi-; result='" & m_Result & "' esperaba '162'", m_Logs, "")
        Exit Function
    End If
    Test_ResponsableFiltro_ValorAId_BC2_IDPuntoYComaNombre = BuildOk("ok", m_Logs)
End Function


Public Function Test_ResponsableFiltro_ValorAId_SinFiltroParaTodos() As String
    Dim m_Logs(0 To 2) As String
    Dim m_Error As String
    Dim m_Result As String
    m_Logs(0) = "1. Arrange: 'Todos' / '' / Null"
    m_Logs(1) = "2. Act: 'Todos' -> ''"
    m_Result = modResponsableFiltroHelper.ResponsableFiltro_ValorAId("Todos", m_Error)
    If m_Result <> "" Then
        Test_ResponsableFiltro_ValorAId_SinFiltroParaTodos = _
            BuildFail("'Todos' result='" & m_Result & "' esperaba ''", m_Logs, "")
        Exit Function
    End If
    m_Logs(2) = "3. Act: '' -> ''"
    m_Result = modResponsableFiltroHelper.ResponsableFiltro_ValorAId("", m_Error)
    If m_Result <> "" Then
        Test_ResponsableFiltro_ValorAId_SinFiltroParaTodos = _
            BuildFail("'' result='" & m_Result & "' esperaba ''", m_Logs, "")
        Exit Function
    End If
    Test_ResponsableFiltro_ValorAId_SinFiltroParaTodos = BuildOk("ok", m_Logs)
End Function


Public Function Test_ResponsableFiltro_IdACmbRow_MatchDevuelveIDPuntoYComaNombre() As String
    Dim dic As Scripting.Dictionary
    Dim u As Usuario
    Dim m_Logs(0 To 3) As String
    Dim m_Result As String

    On Error GoTo errores
    Set dic = New Scripting.Dictionary
    dic.CompareMode = TextCompare
    Set u = New Usuario
    u.ID = "119"
    u.Nombre = "Ana Rubio Canales"
    dic.Add "119", u
    Set u = New Usuario
    u.ID = "181"
    u.Nombre = "Beatriz Noval Gutierrez"
    dic.Add "181", u
    Set u = Nothing

    m_Logs(0) = "1. Arrange: dict con 119+181"
    m_Logs(1) = "2. Act: IdACmbRow('119', dict, False) -> '119;Ana Rubio Canales'"
    m_Result = modResponsableFiltroHelper.ResponsableFiltro_IdACmbRow( _
                    "119", dic, False)
    If m_Result <> "119;Ana Rubio Canales" Then
        Test_ResponsableFiltro_IdACmbRow_MatchDevuelveIDPuntoYComaNombre = _
            BuildFail("match result='" & m_Result & "' esperaba '119;Ana Rubio Canales'", m_Logs, "")
        GoTo Teardown
    End If
    m_Logs(2) = "3. Act: IdACmbRow('999', dict, False) -> '' (no match)"
    m_Result = modResponsableFiltroHelper.ResponsableFiltro_IdACmbRow( _
                    "999", dic, False)
    If m_Result <> "" Then
        Test_ResponsableFiltro_IdACmbRow_MatchDevuelveIDPuntoYComaNombre = _
            BuildFail("no-match result='" & m_Result & "' esperaba ''", m_Logs, "")
        GoTo Teardown
    End If
    m_Logs(3) = "4. Act: IdACmbRow('', dict, True) -> '0;N/A' (Seguridad sentinel)"
    m_Result = modResponsableFiltroHelper.ResponsableFiltro_IdACmbRow( _
                    "", dic, True)
    If m_Result <> "0;N/A" Then
        Test_ResponsableFiltro_IdACmbRow_MatchDevuelveIDPuntoYComaNombre = _
            BuildFail("vacio+sentinel result='" & m_Result & "' esperaba '0;N/A'", m_Logs, "")
        GoTo Teardown
    End If

    Test_ResponsableFiltro_IdACmbRow_MatchDevuelveIDPuntoYComaNombre = BuildOk("ok", m_Logs)
    GoTo Teardown

errores:
    Test_ResponsableFiltro_IdACmbRow_MatchDevuelveIDPuntoYComaNombre = _
        BuildFail("Err " & Err.Number & ": " & Err.Description, m_Logs, "")
Teardown:
    Set dic = Nothing
    Set u = Nothing
End Function


Public Function Test_ResponsableFiltro_IdACmbRow_CalidadSinSentinel() As String
    ' Calidad NO admite "ninguno" -- un ID vacio con p_ConSentinel=False
    ' debe devolver "".
    Dim dic As Scripting.Dictionary
    Dim m_Logs(0 To 2) As String
    Dim m_Result As String
    Set dic = New Scripting.Dictionary
    dic.CompareMode = TextCompare

    m_Logs(0) = "1. Arrange: dict vacio, ID='' , p_ConSentinel=False (Calidad)"
    m_Logs(1) = "2. Act: IdACmbRow('', dict, False) -> '' (no sentinel para Calidad)"
    m_Result = modResponsableFiltroHelper.ResponsableFiltro_IdACmbRow( _
                    "", dic, False)
    If m_Result <> "" Then
        Test_ResponsableFiltro_IdACmbRow_CalidadSinSentinel = _
            BuildFail("Calidad vacio result='" & m_Result & "' esperaba ''", m_Logs, "")
        Set dic = Nothing
        Exit Function
    End If
    m_Logs(2) = "3. Act: IdACmbRow(Nothing dict, False) -> '' (dict vacio)"
    m_Result = modResponsableFiltroHelper.ResponsableFiltro_IdACmbRow( _
                    "119", Nothing, False)
    If m_Result <> "" Then
        Test_ResponsableFiltro_IdACmbRow_CalidadSinSentinel = _
            BuildFail("dict Nothing result='" & m_Result & "' esperaba ''", m_Logs, "")
        Set dic = Nothing
        Exit Function
    End If
    Test_ResponsableFiltro_IdACmbRow_CalidadSinSentinel = BuildOk("ok", m_Logs)
    Set dic = Nothing
End Function


Public Function Test_ResponsableFiltro_ValorEsId_DetectaCambio() As String
    Dim m_Logs(0 To 4) As String

    m_Logs(0) = "1. Arrange: '119' vs ID '119' -> True"
    If Not modResponsableFiltroHelper.ResponsableFiltro_ValorEsId("119", "119") Then
        Test_ResponsableFiltro_ValorEsId_DetectaCambio = _
            BuildFail("'119' vs '119' debio ser True", m_Logs, "")
        Exit Function
    End If
    m_Logs(1) = "2. '119;Ana' vs '119' -> True (BC=2)"
    If Not modResponsableFiltroHelper.ResponsableFiltro_ValorEsId( _
                "119;Ana Rubio", "119") Then
        Test_ResponsableFiltro_ValorEsId_DetectaCambio = _
            BuildFail("'119;Ana' vs '119' debio ser True", m_Logs, "")
        Exit Function
    End If
    m_Logs(2) = "3. 'Todos' vs '' -> True (ambos sin filtro)"
    If Not modResponsableFiltroHelper.ResponsableFiltro_ValorEsId("Todos", "") Then
        Test_ResponsableFiltro_ValorEsId_DetectaCambio = _
            BuildFail("'Todos' vs '' debio ser True", m_Logs, "")
        Exit Function
    End If
    m_Logs(3) = "4. '119' vs '' -> False (filtro vs sin filtro)"
    If modResponsableFiltroHelper.ResponsableFiltro_ValorEsId("119", "") Then
        Test_ResponsableFiltro_ValorEsId_DetectaCambio = _
            BuildFail("'119' vs '' NO debio ser True", m_Logs, "")
        Exit Function
    End If
    m_Logs(4) = "5. '0;N/A' vs '' -> True (sentinel UX == sin filtro)"
    If Not modResponsableFiltroHelper.ResponsableFiltro_ValorEsId("0;N/A", "") Then
        Test_ResponsableFiltro_ValorEsId_DetectaCambio = _
            BuildFail("'0;N/A' vs '' debio ser True", m_Logs, "")
        Exit Function
    End If

    Test_ResponsableFiltro_ValorEsId_DetectaCambio = BuildOk("ok", m_Logs)
End Function


' -----------------------------------------------------------------------------
' Wrappers JSON (runner contract §2)
' -----------------------------------------------------------------------------

Private Function BuildOk(ByVal p_Value As String, ByRef p_Logs() As String) As String
    BuildOk = "{""ok"":true,""value"":""" & EscapeJson(p_Value) & """,""payload"":null,""error"":null,""logs"":[" & LogsToJson(p_Logs) & "]}"
End Function


Private Function BuildFail(ByVal p_Msg As String, ByRef p_Logs() As String, ByVal p_Det As String) As String
    Dim m_ErrJson As String
    Dim m_MsgCompleto As String
    m_MsgCompleto = p_Msg
    If Len(p_Det) > 0 Then m_MsgCompleto = m_MsgCompleto & " | " & p_Det
    m_ErrJson = """" & EscapeJson(m_MsgCompleto) & """"
    BuildFail = "{""ok"":false,""value"":null,""payload"":null,""error"":" & m_ErrJson & ",""logs"":[" & LogsToJson(p_Logs) & "]}"
End Function


Private Function LogsToJson(ByRef p_Logs() As String) As String
    Dim i As Long, s As String
    For i = LBound(p_Logs) To UBound(p_Logs)
        If Len(p_Logs(i)) > 0 Then
            If Len(s) > 0 Then s = s & ","
            s = s & """" & EscapeJson(p_Logs(i)) & """"
        End If
    Next
    LogsToJson = s
End Function


Private Function EscapeJson(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbCr, "\n")
    EscapeJson = s
End Function