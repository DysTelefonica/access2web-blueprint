Attribute VB_Name = "Test_Helper_ExpedienteAlta"
Option Compare Database
Option Explicit
' Tests para Helper_ExpedienteAlta (PRUEBA-003 REFAC-3a).
' Cobertura: BR-01-02, BR-01-03, BR-01-04 (PR-C del coverage matrix).
' Pendientes: BR-01-01 happy path, BR-01-05 rollback on FK.
'
' Convenciones:
'   - 0-arg Public Function
'   - JSON return: BuildJsonOk / BuildJsonFail
'   - Mock object via CreateObject("Scripting.Dictionary") (los helpers usan CallByName,
'     asi que un Dictionary con las keys correctas funciona como mock - el helper solo
'     LEE properties via VbGet, no escribe)

' BR-01-02: CodExp vacio en un Expediente (no Lote) -> NO + motivo
Public Function Test_Helper_ExpedienteAlta_ValidarAlta_CodExpVacio_DevuelveNO() As String
    Dim logs(0 To 2) As String
    Dim exp As MockExpediente
    Set exp = New MockExpediente
    exp.CodExp = ""
    exp.EsLote = "No"
    Dim motivo As String
    Dim errMsg As String
    Dim result As String
    logs(0) = "Arrange: MockExpediente con CodExp='' EsLote='No'"
    result = Helper_ExpedienteAlta.ValidarAlta(exp, motivo, errMsg)
    logs(1) = "Act: ValidarAlta called"
    If result <> "NO" Then
        Test_Helper_ExpedienteAlta_ValidarAlta_CodExpVacio_DevuelveNO = BuildJsonFail("expected NO, got: " & result & " err: " & errMsg, logs)
        Exit Function
    End If
    If InStr(motivo, "Codigo") = 0 Then
        Test_Helper_ExpedienteAlta_ValidarAlta_CodExpVacio_DevuelveNO = BuildJsonFail("expected motivo about Codigo, got: " & motivo, logs)
        Exit Function
    End If
    If errMsg <> "" Then
        Test_Helper_ExpedienteAlta_ValidarAlta_CodExpVacio_DevuelveNO = BuildJsonFail("expected no error, got: " & errMsg, logs)
        Exit Function
    End If
    logs(2) = "Assert: NO + motivo about Codigo"
    Test_Helper_ExpedienteAlta_ValidarAlta_CodExpVacio_DevuelveNO = BuildJsonOk("ok", logs)
End Function

' BR-01-03: EsLote='Sí' sin IDExpedientePadre -> NO + motivo
Public Function Test_Helper_ExpedienteAlta_ValidarAlta_LoteSinPadre_DevuelveNO() As String
    Dim logs(0 To 2) As String
    Dim exp As MockExpediente
    Set exp = New MockExpediente
    exp.CodExp = "LOTE-001"
    exp.EsLote = "Sí"
    exp.IDExpedientePadre = ""
    Dim motivo As String
    Dim errMsg As String
    Dim result As String
    logs(0) = "Arrange: MockExpediente con EsLote='Sí' sin padre"
    result = Helper_ExpedienteAlta.ValidarAlta(exp, motivo, errMsg)
    logs(1) = "Act: ValidarAlta called"
    If result <> "NO" Then
        Test_Helper_ExpedienteAlta_ValidarAlta_LoteSinPadre_DevuelveNO = BuildJsonFail("expected NO, got: " & result & " err: " & errMsg, logs)
        Exit Function
    End If
    If InStr(motivo, "Lote") = 0 Or InStr(motivo, "padre") = 0 Then
        Test_Helper_ExpedienteAlta_ValidarAlta_LoteSinPadre_DevuelveNO = BuildJsonFail("expected motivo about Lote+padre, got: " & motivo, logs)
        Exit Function
    End If
    logs(2) = "Assert: NO + motivo about Lote+padre"
    Test_Helper_ExpedienteAlta_ValidarAlta_LoteSinPadre_DevuelveNO = BuildJsonOk("ok", logs)
End Function

' BR-01-04: ImporteLicitacion negativo -> NO + motivo
Public Function Test_Helper_ExpedienteAlta_ValidarAlta_ImporteNegativo_DevuelveNO() As String
    Dim logs(0 To 2) As String
    Dim exp As MockExpediente
    Set exp = New MockExpediente
    exp.CodExp = "EXP-001"
    exp.EsLote = "No"
    exp.IDExpedientePadre = ""
    exp.ImporteLicitacion = -100
    Dim motivo As String
    Dim errMsg As String
    Dim result As String
    logs(0) = "Arrange: MockExpediente con ImporteLicitacion=-100"
    result = Helper_ExpedienteAlta.ValidarAlta(exp, motivo, errMsg)
    logs(1) = "Act: ValidarAlta called"
    If result <> "NO" Then
        Test_Helper_ExpedienteAlta_ValidarAlta_ImporteNegativo_DevuelveNO = BuildJsonFail("expected NO, got: " & result & " err: " & errMsg, logs)
        Exit Function
    End If
    If InStr(motivo, "negativo") = 0 Then
        Test_Helper_ExpedienteAlta_ValidarAlta_ImporteNegativo_DevuelveNO = BuildJsonFail("expected motivo about negativo, got: " & motivo, logs)
        Exit Function
    End If
    logs(2) = "Assert: NO + motivo about negativo"
    Test_Helper_ExpedienteAlta_ValidarAlta_ImporteNegativo_DevuelveNO = BuildJsonOk("ok", logs)
End Function

' BR-01-02..04: Happy path - Expediente valido -> OK
Public Function Test_Helper_ExpedienteAlta_ValidarAlta_ExpedienteValido_DevuelveOK() As String
    Dim logs(0 To 2) As String
    Dim exp As MockExpediente
    Set exp = New MockExpediente
    exp.CodExp = "EXP-2024-001"
    exp.EsLote = "No"
    exp.IDExpedientePadre = ""
    exp.ImporteLicitacion = 50000
    Dim motivo As String
    Dim errMsg As String
    Dim result As String
    logs(0) = "Arrange: MockExpediente valido (CodExp, EsLote=No, sin padre, Importe=50000)"
    result = Helper_ExpedienteAlta.ValidarAlta(exp, motivo, errMsg)
    logs(1) = "Act: ValidarAlta called"
    If result <> "OK" Then
        Test_Helper_ExpedienteAlta_ValidarAlta_ExpedienteValido_DevuelveOK = BuildJsonFail("expected OK, got: " & result & " motivo: " & motivo & " err: " & errMsg, logs)
        Exit Function
    End If
    If motivo <> "" Then
        Test_Helper_ExpedienteAlta_ValidarAlta_ExpedienteValido_DevuelveOK = BuildJsonFail("expected motivo vacio, got: " & motivo, logs)
        Exit Function
    End If
    If errMsg <> "" Then
        Test_Helper_ExpedienteAlta_ValidarAlta_ExpedienteValido_DevuelveOK = BuildJsonFail("expected no error, got: " & errMsg, logs)
        Exit Function
    End If
    logs(2) = "Assert: OK + motivo vacio + sin error"
    Test_Helper_ExpedienteAlta_ValidarAlta_ExpedienteValido_DevuelveOK = BuildJsonOk("ok", logs)
End Function

' BR-01-02: Lote con CodExp vacio -> OK (los Lotes se generan automaticamente)
Public Function Test_Helper_ExpedienteAlta_ValidarAlta_LoteSinCodExp_DevuelveOK() As String
    Dim logs(0 To 2) As String
    Dim exp As MockExpediente
    Set exp = New MockExpediente
    exp.CodExp = ""
    exp.EsLote = "Sí"
    exp.IDExpedientePadre = "12345"  ' tiene padre
    Dim motivo As String
    Dim errMsg As String
    Dim result As String
    logs(0) = "Arrange: Lote sin CodExp (se genera auto) con padre"
    result = Helper_ExpedienteAlta.ValidarAlta(exp, motivo, errMsg)
    logs(1) = "Act: ValidarAlta called"
    If result <> "OK" Then
        Test_Helper_ExpedienteAlta_ValidarAlta_LoteSinCodExp_DevuelveOK = BuildJsonFail("expected OK (Lote sin CodExp es valido), got: " & result & " motivo: " & motivo & " err: " & errMsg, logs)
        Exit Function
    End If
    logs(2) = "Assert: OK (Lote sin CodExp se valida OK)"
    Test_Helper_ExpedienteAlta_ValidarAlta_LoteSinCodExp_DevuelveOK = BuildJsonOk("ok", logs)
End Function
