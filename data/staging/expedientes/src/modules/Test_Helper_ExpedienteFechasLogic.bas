Attribute VB_Name = "Test_Helper_ExpedienteFechasLogic"
Option Compare Database
Option Explicit
' Tests para Helper_ExpedienteFechasLogic y Form_FormExpedienteFechas
' (PRUEBA-004 REFAC-4a slice 1).
' Cobertura: BR-19-01..08.
'   BR-19-01..06: calculo derivado de FechaFinGarantia a partir
'     de GarantiaMeses, FechaCertificacion y FechaFinContrato.
'   BR-19-07: Form_FormExpedienteFechas delegates to
'     Helper_ExpedienteFechasLogic.CalcularFechaFinGarantia and contains
'     zero DateAdd( in event handlers (static-source atom).
'   BR-19-08: Form_FormExpedienteFechas contains zero MsgBox(, zero
'     'pregunta = ' writes, and at least 1 FormInteraction_Mensaje(
'     call (static-source atom).
' La logica de fechas estaba duplicada en 4 event handlers de
' Form_FormExpedienteFechas.cls (ComandoFECHACERTIFICACION_Click,
' FECHACERTIFICACION_BeforeUpdate, FechaFinContrato_BeforeUpdate,
' GARANTIAMESES_BeforeUpdate).
'
' Convenciones:
'   - 0-arg Public Function (atoms)
'   - JSON return via local BuildOk/BuildFail wrappers (per access-vba-tdd
'     §1.8) which delegate to Test_Helper.BuildJsonOk / Test_Helper.BuildJsonFail.
'   - Pure function for BR-19-01..06 (no DAO, no UI) -> sin fixture graph.
'   - BR-19-07/08: static-source atoms (read Form_FormExpedienteFechas.cls
'     via FileSystemObject; forms cannot be opened in COM).

' --- JSON wrappers delegate to modTestingCoreHelper (per e2e rule #10) -------------
' Locally aliased to BuildOk/BuildFail so existing atom bodies compile unchanged.
Private Function BuildOk(ByVal value As Variant, ByRef logs() As String) As String
    BuildOk = TestingCore_BuildOk(value, logs)
End Function

Private Function BuildFail(ByVal msg As String, ByRef logs() As String) As String
    BuildFail = TestingCore_BuildFail(msg, logs)
End Function

' --- BR-19-01: GarantiaMeses no numerico -> FechaFinGarantia = Null (sin calculo)
Public Function Test_Helper_ExpedienteFechasLogic_CalcularFechaFinGarantia_GarantiaMesesNoNumerico_DevuelveNull() As String
    Dim logs(0 To 2) As String
    Dim fechaFin As Variant
    Dim errMsg As String
    Dim result As String
    logs(0) = "Arrange: GarantiaMeses=Null, FechaCertificacion=#2024-01-15#, FechaFinContrato=#2024-01-01#"
    result = Helper_ExpedienteFechasLogic.CalcularFechaFinGarantia(Null, #1/15/2024#, #1/1/2024#, fechaFin, errMsg)
    logs(1) = "Act: called"
    If result <> "" Then
        Test_Helper_ExpedienteFechasLogic_CalcularFechaFinGarantia_GarantiaMesesNoNumerico_DevuelveNull = BuildFail("expected success, got: " & result & " err: " & errMsg, logs)
        Exit Function
    End If
    If Not IsNull(fechaFin) Then
        Test_Helper_ExpedienteFechasLogic_CalcularFechaFinGarantia_GarantiaMesesNoNumerico_DevuelveNull = BuildFail("expected Null, got: " & CStr(fechaFin), logs)
        Exit Function
    End If
    logs(2) = "Assert: fechaFin = Null (sin calculo)"
    Test_Helper_ExpedienteFechasLogic_CalcularFechaFinGarantia_GarantiaMesesNoNumerico_DevuelveNull = BuildOk("ok", logs)
End Function

' --- BR-19-02: FechaFinContrato no es fecha -> FechaFinGarantia = Null
Public Function Test_Helper_ExpedienteFechasLogic_CalcularFechaFinGarantia_FechaFinContratoVacia_DevuelveNull() As String
    Dim logs(0 To 2) As String
    Dim fechaFin As Variant
    Dim errMsg As String
    Dim result As String
    logs(0) = "Arrange: GarantiaMeses=12, FechaCertificacion=#2024-01-15#, FechaFinContrato=Null"
    result = Helper_ExpedienteFechasLogic.CalcularFechaFinGarantia(12, #1/15/2024#, Null, fechaFin, errMsg)
    logs(1) = "Act: called"
    If result <> "" Then
        Test_Helper_ExpedienteFechasLogic_CalcularFechaFinGarantia_FechaFinContratoVacia_DevuelveNull = BuildFail("expected success, got: " & result & " err: " & errMsg, logs)
        Exit Function
    End If
    If Not IsNull(fechaFin) Then
        Test_Helper_ExpedienteFechasLogic_CalcularFechaFinGarantia_FechaFinContratoVacia_DevuelveNull = BuildFail("expected Null, got: " & CStr(fechaFin), logs)
        Exit Function
    End If
    logs(2) = "Assert: fechaFin = Null (sin FechaFinContrato no hay base para calcular)"
    Test_Helper_ExpedienteFechasLogic_CalcularFechaFinGarantia_FechaFinContratoVacia_DevuelveNull = BuildOk("ok", logs)
End Function

' --- BR-19-03: FechaCertificacion valida -> DateAdd sobre FechaCertificacion
Public Function Test_Helper_ExpedienteFechasLogic_CalcularFechaFinGarantia_FechaCertificacionValida_DevuelveDateAddMeses() As String
    Dim logs(0 To 2) As String
    Dim fechaFin As Variant
    Dim errMsg As String
    Dim result As String
    logs(0) = "Arrange: GarantiaMeses=12, FechaCertificacion=#2024-01-15#, FechaFinContrato=#2024-01-01# (distinto)"
    result = Helper_ExpedienteFechasLogic.CalcularFechaFinGarantia(12, #1/15/2024#, #1/1/2024#, fechaFin, errMsg)
    logs(1) = "Act: called"
    If result <> "" Then
        Test_Helper_ExpedienteFechasLogic_CalcularFechaFinGarantia_FechaCertificacionValida_DevuelveDateAddMeses = BuildFail("expected success, got: " & result & " err: " & errMsg, logs)
        Exit Function
    End If
    If CDate(fechaFin) <> #1/15/2025# Then
        Test_Helper_ExpedienteFechasLogic_CalcularFechaFinGarantia_FechaCertificacionValida_DevuelveDateAddMeses = BuildFail("expected #2025-01-15# (12 meses sobre FechaCertificacion), got: " & CStr(fechaFin), logs)
        Exit Function
    End If
    logs(2) = "Assert: fechaFin = #2025-01-15# (DateAdd m, 12, sobre FechaCertificacion)"
    Test_Helper_ExpedienteFechasLogic_CalcularFechaFinGarantia_FechaCertificacionValida_DevuelveDateAddMeses = BuildOk("ok", logs)
End Function

' --- BR-19-04: FechaCertificacion vacia -> cae a FechaFinContrato
Public Function Test_Helper_ExpedienteFechasLogic_CalcularFechaFinGarantia_FechaCertificacionVacia_UsaFechaFinContrato() As String
    Dim logs(0 To 2) As String
    Dim fechaFin As Variant
    Dim errMsg As String
    Dim result As String
    logs(0) = "Arrange: GarantiaMeses=24, FechaCertificacion=Null, FechaFinContrato=#2023-06-01#"
    result = Helper_ExpedienteFechasLogic.CalcularFechaFinGarantia(24, Null, #6/1/2023#, fechaFin, errMsg)
    logs(1) = "Act: called"
    If result <> "" Then
        Test_Helper_ExpedienteFechasLogic_CalcularFechaFinGarantia_FechaCertificacionVacia_UsaFechaFinContrato = BuildFail("expected success, got: " & result & " err: " & errMsg, logs)
        Exit Function
    End If
    If CDate(fechaFin) <> #6/1/2025# Then
        Test_Helper_ExpedienteFechasLogic_CalcularFechaFinGarantia_FechaCertificacionVacia_UsaFechaFinContrato = BuildFail("expected #2025-06-01# (24 meses sobre FechaFinContrato), got: " & CStr(fechaFin), logs)
        Exit Function
    End If
    logs(2) = "Assert: fechaFin = #2025-06-01# (fallback a FechaFinContrato)"
    Test_Helper_ExpedienteFechasLogic_CalcularFechaFinGarantia_FechaCertificacionVacia_UsaFechaFinContrato = BuildOk("ok", logs)
End Function

' --- BR-19-05: GarantiaMeses = 0 -> devuelve la misma fecha (edge: sin extension)
Public Function Test_Helper_ExpedienteFechasLogic_CalcularFechaFinGarantia_GarantiaMesesCero_DevuelveIgual() As String
    Dim logs(0 To 2) As String
    Dim fechaFin As Variant
    Dim errMsg As String
    Dim result As String
    logs(0) = "Arrange: GarantiaMeses=0, FechaCertificacion=#2024-03-10#"
    result = Helper_ExpedienteFechasLogic.CalcularFechaFinGarantia(0, #3/10/2024#, #1/1/2024#, fechaFin, errMsg)
    logs(1) = "Act: called"
    If result <> "" Then
        Test_Helper_ExpedienteFechasLogic_CalcularFechaFinGarantia_GarantiaMesesCero_DevuelveIgual = BuildFail("expected success, got: " & result & " err: " & errMsg, logs)
        Exit Function
    End If
    If CDate(fechaFin) <> #3/10/2024# Then
        Test_Helper_ExpedienteFechasLogic_CalcularFechaFinGarantia_GarantiaMesesCero_DevuelveIgual = BuildFail("expected #2024-03-10# (0 meses = misma fecha), got: " & CStr(fechaFin), logs)
        Exit Function
    End If
    logs(2) = "Assert: fechaFin = #2024-03-10# (edge: cero meses)"
    Test_Helper_ExpedienteFechasLogic_CalcularFechaFinGarantia_GarantiaMesesCero_DevuelveIgual = BuildOk("ok", logs)
End Function

' --- BR-19-06: GarantiaMeses negativo -> DateAdd hacia atras (adversarial)
Public Function Test_Helper_ExpedienteFechasLogic_CalcularFechaFinGarantia_GarantiaMesesNegativo_DevuelveHaciaAtras() As String
    Dim logs(0 To 2) As String
    Dim fechaFin As Variant
    Dim errMsg As String
    Dim result As String
    logs(0) = "Arrange: GarantiaMeses=-3, FechaCertificacion=#2024-03-10#"
    result = Helper_ExpedienteFechasLogic.CalcularFechaFinGarantia(-3, #3/10/2024#, #1/1/2024#, fechaFin, errMsg)
    logs(1) = "Act: called"
    If result <> "" Then
        Test_Helper_ExpedienteFechasLogic_CalcularFechaFinGarantia_GarantiaMesesNegativo_DevuelveHaciaAtras = BuildFail("expected success, got: " & result & " err: " & errMsg, logs)
        Exit Function
    End If
    If CDate(fechaFin) <> #12/10/2023# Then
        Test_Helper_ExpedienteFechasLogic_CalcularFechaFinGarantia_GarantiaMesesNegativo_DevuelveHaciaAtras = BuildFail("expected #2023-12-10# (-3 meses desde #2024-03-10#), got: " & CStr(fechaFin), logs)
        Exit Function
    End If
    logs(2) = "Assert: fechaFin = #2023-12-10# (adversarial: DateAdd acepta negativos)"
    Test_Helper_ExpedienteFechasLogic_CalcularFechaFinGarantia_GarantiaMesesNegativo_DevuelveHaciaAtras = BuildOk("ok", logs)
End Function

' --- BR-19-07: Form_FormExpedienteFechas delegates date calculation to helper
' Static-source atom: reads Form_FormExpedienteFechas.cls, asserts the file
' contains at least 4 calls to Helper_ExpedienteFechasLogic.CalcularFechaFinGarantia
' (one per date handler) AND zero DateAdd( matches (date arithmetic lives
' in the helper now). Forms cannot be opened in COM, so the static audit is
' the only viable assertion mechanism per e2e rule #1.
Public Function Test_Helper_ExpedienteFechasLogic_BR19_07_FormDelegaEnHelper_NoLogicaEnHandler() As String
    Dim logs(0 To 3) As String
    logs(0) = "Arrange: read Form_FormExpedienteFechas.cls"

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim formPath As String
    formPath = CurrentProject.Path & "\src\forms\Form_FormExpedienteFechas.cls"

    If Not fso.FileExists(formPath) Then
        Test_Helper_ExpedienteFechasLogic_BR19_07_FormDelegaEnHelper_NoLogicaEnHandler = BuildFail("form file not found at: " & formPath, logs)
        Exit Function
    End If

    Dim formContent As String
    With fso.OpenTextFile(formPath, 1)
        formContent = .ReadAll
        .Close
    End With
    logs(1) = "Act: file read (" & Len(formContent) & " bytes)"

    Dim helperToken As String
    helperToken = "Helper_ExpedienteFechasLogic.CalcularFechaFinGarantia"
    Dim helperCount As Long
    helperCount = (Len(formContent) - Len(Replace(formContent, helperToken, ""))) / Len(helperToken)

    Dim dateAddToken As String
    dateAddToken = "DateAdd("
    Dim dateAddCount As Long
    dateAddCount = (Len(formContent) - Len(Replace(formContent, dateAddToken, ""))) / Len(dateAddToken)

    ' Count FormInteraction_EstablecerPropiedad( occurrences (cross-form control property
    ' access via helper, e2e rule #11). Form_Load uses it twice for .Visible + .Enabled.
    Dim propiedadToken As String
    propiedadToken = "FormInteraction_EstablecerPropiedad("
    Dim propiedadCount As Long
    propiedadCount = (Len(formContent) - Len(Replace(formContent, propiedadToken, ""))) / Len(propiedadToken)

    If helperCount < 4 Then
        Test_Helper_ExpedienteFechasLogic_BR19_07_FormDelegaEnHelper_NoLogicaEnHandler = BuildFail("expected at least 4 helper calls (one per date handler), found: " & helperCount, logs)
        Exit Function
    End If
    If dateAddCount > 0 Then
        Test_Helper_ExpedienteFechasLogic_BR19_07_FormDelegaEnHelper_NoLogicaEnHandler = BuildFail("expected zero DateAdd( in form, found: " & dateAddCount, logs)
        Exit Function
    End If
    If propiedadCount < 2 Then
        Test_Helper_ExpedienteFechasLogic_BR19_07_FormDelegaEnHelper_NoLogicaEnHandler = BuildFail("expected at least 2 FormInteraction_EstablecerPropiedad( calls (Form_Load .Visible + .Enabled via helper), found: " & propiedadCount, logs)
        Exit Function
    End If

    logs(2) = "Assert: helper calls >= 4, DateAdd count = 0, FormInteraction_EstablecerPropiedad >= 2"
    logs(3) = "helperCount=" & helperCount & ", dateAddCount=" & dateAddCount & ", propiedadCount=" & propiedadCount
    Test_Helper_ExpedienteFechasLogic_BR19_07_FormDelegaEnHelper_NoLogicaEnHandler = BuildOk("ok", logs)
End Function

' --- BR-19-08: Form_FormExpedienteFechas contains no MsgBox, no 'pregunta =' writes,
' and uses FormInteraction_Mensaje in error branches.
' Static-source atom: reads Form_FormExpedienteFechas.cls, asserts zero
' MsgBox( matches, zero 'pregunta = ' writes (the global Variables Globales.pregunta
' is preserved for 14 other forms but this form stops writing to it), and
' at least 1 FormInteraction_Mensaje( call (e2e rule #5: helpers receive
' Optional ByRef p_Resultado so atoms assert the prompt without a real modal).
Public Function Test_Helper_ExpedienteFechasLogic_BR19_08_NoMsgBoxEnHandler_MensajeInyectado() As String
    Dim logs(0 To 3) As String
    logs(0) = "Arrange: read Form_FormExpedienteFechas.cls"

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim formPath As String
    formPath = CurrentProject.Path & "\src\forms\Form_FormExpedienteFechas.cls"

    If Not fso.FileExists(formPath) Then
        Test_Helper_ExpedienteFechasLogic_BR19_08_NoMsgBoxEnHandler_MensajeInyectado = BuildFail("form file not found at: " & formPath, logs)
        Exit Function
    End If

    Dim formContent As String
    With fso.OpenTextFile(formPath, 1)
        formContent = .ReadAll
        .Close
    End With
    logs(1) = "Act: file read (" & Len(formContent) & " bytes)"

    Dim msgBoxToken As String
    msgBoxToken = "MsgBox("
    Dim msgBoxCount As Long
    msgBoxCount = (Len(formContent) - Len(Replace(formContent, msgBoxToken, ""))) / Len(msgBoxToken)

    Dim preguntaToken As String
    preguntaToken = "pregunta = "
    Dim preguntaCount As Long
    preguntaCount = (Len(formContent) - Len(Replace(formContent, preguntaToken, ""))) / Len(preguntaToken)

    Dim mensajeToken As String
    mensajeToken = "FormInteraction_Mensaje("
    Dim mensajeCount As Long
    mensajeCount = (Len(formContent) - Len(Replace(formContent, mensajeToken, ""))) / Len(mensajeToken)

    If msgBoxCount > 0 Then
        Test_Helper_ExpedienteFechasLogic_BR19_08_NoMsgBoxEnHandler_MensajeInyectado = BuildFail("expected zero MsgBox( in form, found: " & msgBoxCount, logs)
        Exit Function
    End If
    If preguntaCount > 0 Then
        Test_Helper_ExpedienteFechasLogic_BR19_08_NoMsgBoxEnHandler_MensajeInyectado = BuildFail("expected zero 'pregunta = ' writes in form, found: " & preguntaCount, logs)
        Exit Function
    End If
    If mensajeCount < 1 Then
        Test_Helper_ExpedienteFechasLogic_BR19_08_NoMsgBoxEnHandler_MensajeInyectado = BuildFail("expected at least 1 FormInteraction_Mensaje( call in form, found: " & mensajeCount, logs)
        Exit Function
    End If

    logs(2) = "Assert: MsgBox=0, pregunta=0, mensaje>=1"
    logs(3) = "msgBoxCount=" & msgBoxCount & ", preguntaCount=" & preguntaCount & ", mensajeCount=" & mensajeCount
    Test_Helper_ExpedienteFechasLogic_BR19_08_NoMsgBoxEnHandler_MensajeInyectado = BuildOk("ok", logs)
End Function
