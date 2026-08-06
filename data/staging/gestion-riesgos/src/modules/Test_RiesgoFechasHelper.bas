Attribute VB_Name = "Test_RiesgoFechasHelper"
Option Compare Database
Option Explicit

' ============================================================
' Test_RiesgoFechasHelper - TDD atoms for modRiesgoFechasHelper
'
' Helper: ValidarFechaCampo
'   Signature: Public Function ValidarFechaCampo( _
'                 ByVal p_NombreCampo As String
'                 Optional ByRef db As DAO.Database = Nothing
'
' Architecture: pure function over (p_NombreCampo, p_Valor). The helper
' centralizes REQ-CAL-04 (no future dates except for FechaInicio and
' FechaFinPrevista; past dates accepted where they make sense).
'
' The helper preserves the semantic of Riesgo.ValidarFechaMaterializacionPermitida
' (Riesgo.cls:6219) for FechaMaterializado so that the 4 existing atoms in
' Test_RiesgoMaterializacion.bas stay GREEN:
'   Test_RiesgoMaterializacion_FechaFuturaBloqueada
'   Test_RiesgoMaterializacion_FechaHoyPermitida
'   Test_RiesgoMaterializacion_FechaPasadaPermitida
'   Test_RiesgoMaterializacion_RegistrarFechaFuturaNoPersiste
'
' Skill: access-vba-tdd v2.4.3
' SDD:    e2e-form-by-form-2026-06-22 - Bloque 4 - REQ-CAL-04
' ============================================================

' --- JSON helpers wrapper (project convention) ---
Private Function BuildOk(ByVal value As Variant, ByRef logs() As String) As String
    BuildOk = Test_Helper.BuildJsonOk(value, logs)
End Function

Private Function BuildFail(ByVal msg As String, ByRef logs() As String) As String
    BuildFail = Test_Helper.BuildJsonFail(msg, logs)
End Function

' --- Helpers to format date inputs the way the helper expects (dd/mm/yyyy) ---
Private Function FormatDate(ByVal p_Date As Date) As String
    FormatDate = Format$(p_Date, "dd/mm/yyyy")
End Function

' ----------------------------------------------------------------
' ATOM 1: Happy - FechaInicio futura permitida (+30 días)
' Expected:
'   - result = True
'   - p_Error = ""
' RED ? GREEN via whitelist for FECHAINICIO.
' ----------------------------------------------------------------
Public Function Test_RiesgoFechasHelper_Happy_FechaInicioFuturaAceptada() As String
    On Error GoTo EH

    Dim logs(0 To 4) As String
    logs(0) = "1. Arrange: p_NombreCampo=""FechaInicio"", p_Valor=+30 días"
    logs(1) = "2. Act: ValidarFechaCampo"
    logs(2) = "3. Assert: result = True (whitelist FECHAINICIO)"

    Dim m_Nombre As String
    Dim m_FechaFutura As String
    Dim m_Resultado As Boolean
    Dim m_Error As String

    m_Nombre = "FechaInicio"
    m_FechaFutura = FormatDate(DateAdd("d", 30, Date))

    m_Resultado = ValidarFechaCampo(m_Nombre, m_FechaFutura, Nothing, m_Error)

    If m_Resultado <> True Then
        logs(3) = "3. Assert FAIL: expected True, got " & CStr(m_Resultado) & _
            " (error=" & m_Error & ")"
        Test_RiesgoFechasHelper_Happy_FechaInicioFuturaAceptada = _
            BuildFail("FechaInicio futura +30d debe aceptarse", logs)
        Exit Function
    End If

    If m_Error <> "" Then
        logs(3) = "3. Assert FAIL: p_Error poblado: " & m_Error
        Test_RiesgoFechasHelper_Happy_FechaInicioFuturaAceptada = _
            BuildFail("p_Error debe estar vacío cuando el resultado es True", logs)
        Exit Function
    End If

    logs(3) = "4. Assert PASS: True sin error"
    Test_RiesgoFechasHelper_Happy_FechaInicioFuturaAceptada = _
        BuildOk(CStr(m_Resultado), logs)
    Exit Function

EH:
    logs(2) = "2. Act: raised " & Err.Number & " - " & Err.description
    Test_RiesgoFechasHelper_Happy_FechaInicioFuturaAceptada = _
        BuildFail("ValidarFechaCampo raised: " & Err.description, logs)
End Function

' ----------------------------------------------------------------
' ATOM 2: Sad - FechaDetectado futura rechazada (+7 días)
' Expected:
'   - result = False
'   - p_Error contiene "fecha de detectado"
' RED ? GREEN via regla "resto > hoy ? False".
' ----------------------------------------------------------------
Public Function Test_RiesgoFechasHelper_Sad_FechaDetectadoFuturaRechazada() As String
    On Error GoTo EH

    Dim logs(0 To 4) As String
    logs(0) = "1. Arrange: p_NombreCampo=""FechaDetectado"", p_Valor=+7 días"
    logs(1) = "2. Act: ValidarFechaCampo"
    logs(2) = "3. Assert: result = False"
    logs(3) = "4. Assert: p_Error contiene ""fecha de detectado"" y ""posterior"""

    Dim m_Nombre As String
    Dim m_FechaFutura As String
    Dim m_Resultado As Boolean
    Dim m_Error As String

    m_Nombre = "FechaDetectado"
    m_FechaFutura = FormatDate(DateAdd("d", 7, Date))

    m_Resultado = ValidarFechaCampo(m_Nombre, m_FechaFutura, Nothing, m_Error)

    If m_Resultado <> False Then
        logs(2) = "3. Assert FAIL: expected False, got " & CStr(m_Resultado)
        Test_RiesgoFechasHelper_Sad_FechaDetectadoFuturaRechazada = _
            BuildFail("FechaDetectado futura +7d debe rechazarse", logs)
        Exit Function
    End If

    If InStr(1, m_Error, "fecha de detectado", vbTextCompare) = 0 Then
        logs(3) = "4. Assert FAIL: p_Error no contiene ""fecha de detectado"": " & m_Error
        Test_RiesgoFechasHelper_Sad_FechaDetectadoFuturaRechazada = _
            BuildFail("p_Error debe mencionar la etiqueta del campo", logs)
        Exit Function
    End If

    If InStr(1, m_Error, "posterior", vbTextCompare) = 0 Then
        logs(3) = "4. Assert FAIL: p_Error no contiene ""posterior"": " & m_Error
        Test_RiesgoFechasHelper_Sad_FechaDetectadoFuturaRechazada = _
            BuildFail("p_Error debe indicar que no puede ser posterior a hoy", logs)
        Exit Function
    End If

    logs(3) = "4. Assert PASS: False con etiqueta y motivo"
    Test_RiesgoFechasHelper_Sad_FechaDetectadoFuturaRechazada = _
        BuildOk(CStr(m_Resultado), logs)
    Exit Function

EH:
    logs(2) = "2. Act: raised " & Err.Number & " - " & Err.description
    Test_RiesgoFechasHelper_Sad_FechaDetectadoFuturaRechazada = _
        BuildFail("ValidarFechaCampo raised: " & Err.description, logs)
End Function

' ----------------------------------------------------------------
' ATOM 3: Edge - FechaDetectado pasada aceptada (-30 días)
' Expected:
'   - result = True
'   - p_Error = ""
' RED ? GREEN: las fechas pasadas son válidas (lectura operativa
' propuesta, pendiente de refrendar con Calidad).
' ----------------------------------------------------------------
Public Function Test_RiesgoFechasHelper_Edge_FechaPasadaAceptadaEnCualquierCampo() As String
    On Error GoTo EH

    Dim logs(0 To 4) As String
    logs(0) = "1. Arrange: p_NombreCampo=""FechaDetectado"", p_Valor=-30 días"
    logs(1) = "2. Act: ValidarFechaCampo"
    logs(2) = "3. Assert: result = True (fecha pasada válida)"

    Dim m_Nombre As String
    Dim m_FechaPasada As String
    Dim m_Resultado As Boolean
    Dim m_Error As String

    m_Nombre = "FechaDetectado"
    m_FechaPasada = FormatDate(DateAdd("d", -30, Date))

    m_Resultado = ValidarFechaCampo(m_Nombre, m_FechaPasada, Nothing, m_Error)

    If m_Resultado <> True Then
        logs(3) = "3. Assert FAIL: expected True, got " & CStr(m_Resultado) & _
            " (error=" & m_Error & ")"
        Test_RiesgoFechasHelper_Edge_FechaPasadaAceptadaEnCualquierCampo = _
            BuildFail("FechaDetectado pasada -30d debe aceptarse", logs)
        Exit Function
    End If

    If m_Error <> "" Then
        logs(3) = "3. Assert FAIL: p_Error poblado: " & m_Error
        Test_RiesgoFechasHelper_Edge_FechaPasadaAceptadaEnCualquierCampo = _
            BuildFail("p_Error debe estar vacío cuando el resultado es True", logs)
        Exit Function
    End If

    logs(3) = "4. Assert PASS: True sin error (fecha pasada válida)"
    Test_RiesgoFechasHelper_Edge_FechaPasadaAceptadaEnCualquierCampo = _
        BuildOk(CStr(m_Resultado), logs)
    Exit Function

EH:
    logs(2) = "2. Act: raised " & Err.Number & " - " & Err.description
    Test_RiesgoFechasHelper_Edge_FechaPasadaAceptadaEnCualquierCampo = _
        BuildFail("ValidarFechaCampo raised: " & Err.description, logs)
End Function

' ----------------------------------------------------------------
' ATOM 4: Adversarial - FechaMaterializado futura delega en
'            Riesgo.ValidarFechaMaterializacionPermitida
' Expected:
'   - result = False
'   - p_Error contiene "materialización" y "posterior"
' Este átomo FIJA el contrato: las fechas futuras en FechaMaterializado
' siguen bloqueándose, preservando los 4 átomos verdes de
' Test_RiesgoMaterializacion.bas.
' ----------------------------------------------------------------
Public Function Test_RiesgoFechasHelper_Adversarial_DelegateAFuncionExistente_ParaMaterializado() As String
    On Error GoTo EH

    Dim logs(0 To 5) As String
    logs(0) = "1. Arrange: p_NombreCampo=""FechaMaterializado"", p_Valor=+5 días"
    logs(1) = "2. Act: ValidarFechaCampo (debe bloquear futuras)"
    logs(2) = "3. Assert: result = False"
    logs(3) = "4. Assert: p_Error contiene ""materialización"" y ""posterior"""

    Dim m_Nombre As String
    Dim m_FechaFutura As String
    Dim m_Resultado As Boolean
    Dim m_Error As String

    m_Nombre = "FechaMaterializado"
    m_FechaFutura = FormatDate(DateAdd("d", 5, Date))

    m_Resultado = ValidarFechaCampo(m_Nombre, m_FechaFutura, Nothing, m_Error)

    If m_Resultado <> False Then
        logs(2) = "3. Assert FAIL: expected False, got " & CStr(m_Resultado)
        Test_RiesgoFechasHelper_Adversarial_DelegateAFuncionExistente_ParaMaterializado = _
            BuildFail("FechaMaterializado futura +5d debe bloquearse (preserva Test_RiesgoMaterializacion)", logs)
        Exit Function
    End If

    If InStr(1, m_Error, "materialización", vbTextCompare) = 0 Then
        logs(3) = "4. Assert FAIL: p_Error no contiene ""materialización"": " & m_Error
        Test_RiesgoFechasHelper_Adversarial_DelegateAFuncionExistente_ParaMaterializado = _
            BuildFail("p_Error debe mencionar materialización", logs)
        Exit Function
    End If

    If InStr(1, m_Error, "posterior", vbTextCompare) = 0 Then
        logs(3) = "4. Assert FAIL: p_Error no contiene ""posterior"": " & m_Error
        Test_RiesgoFechasHelper_Adversarial_DelegateAFuncionExistente_ParaMaterializado = _
            BuildFail("p_Error debe indicar que no puede ser posterior a hoy", logs)
        Exit Function
    End If

    logs(3) = "4. Assert PASS: False con etiqueta materialización + motivo"
    logs(4) = "5. Contrato preservado: Test_RiesgoMaterializacion_FechaFutura* siguen verdes"
    Test_RiesgoFechasHelper_Adversarial_DelegateAFuncionExistente_ParaMaterializado = _
        BuildOk(CStr(m_Resultado), logs)
    Exit Function

EH:
    logs(2) = "2. Act: raised " & Err.Number & " - " & Err.description
    Test_RiesgoFechasHelper_Adversarial_DelegateAFuncionExistente_ParaMaterializado = _
        BuildFail("ValidarFechaCampo raised: " & Err.description, logs)
End Function

