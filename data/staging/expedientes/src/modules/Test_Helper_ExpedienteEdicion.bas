Attribute VB_Name = "Test_Helper_ExpedienteEdicion"
Option Compare Database
Option Explicit
' Tests para Helper_ExpedienteEdicion (PRUEBA-003 REFAC-3a slice 2).
' Cobertura: BR-02-01 (campos sucios), BR-04-01 (cambio tipo), BR-05-01 (estado calculado).
' BR-02-02 (bloqueado) es DB-dependent y se difiere a un slice con DAO injection.
'
' Convenciones:
'   - 0-arg Public Function
'   - JSON return: BuildJsonOk / BuildJsonFail
'   - MockExpedienteEdicion para tests con properties reales

' BR-02-01: SoloCamposSucios detecta los campos que cambiaron entre 2 entidades
Public Function Test_Helper_ExpedienteEdicion_SoloCamposSucios_DosCambios_DevuelveTituloYEstado() As String
    Dim logs(0 To 3) As String
    Dim anterior As MockExpedienteEdicion
    Dim nuevo As MockExpedienteEdicion
    Set anterior = New MockExpedienteEdicion
    Set nuevo = New MockExpedienteEdicion
    anterior.CodExp = "EXP-001"
    anterior.Titulo = "Titulo Original"
    anterior.Nemotecnico = "Nemo"
    anterior.ImporteLicitacion = 50000
    anterior.ESTADO = "Vigente"
    nuevo.CodExp = "EXP-001"           ' sin cambio
    nuevo.Titulo = "Titulo Modificado" ' cambio
    nuevo.Nemotecnico = "Nemo"          ' sin cambio
    nuevo.ImporteLicitacion = 75000    ' cambio
    nuevo.ESTADO = "Vigente"            ' sin cambio
    Dim camposSucios As Variant
    Dim errMsg As String
    Dim result As String
    logs(0) = "Arrange: anterior Titulo='Original' Importe=50000, nuevo Titulo='Modificado' Importe=75000"
    result = Helper_ExpedienteEdicion.SoloCamposSucios(anterior, nuevo, Array("CodExp", "Titulo", "Nemotecnico", "ImporteLicitacion", "ESTADO"), camposSucios, errMsg)
    logs(1) = "Act: SoloCamposSucios called"
    If result <> "" Then
        Test_Helper_ExpedienteEdicion_SoloCamposSucios_DosCambios_DevuelveTituloYEstado = BuildJsonFail("expected success, got: " & result & " err: " & errMsg, logs)
        Exit Function
    End If
    Dim m_Campos As String
    m_Campos = CStr(camposSucios)
    If InStr(m_Campos, "Titulo") = 0 Then
        Test_Helper_ExpedienteEdicion_SoloCamposSucios_DosCambios_DevuelveTituloYEstado = BuildJsonFail("expected Titulo en sucios, got: " & m_Campos, logs)
        Exit Function
    End If
    If InStr(m_Campos, "ImporteLicitacion") = 0 Then
        Test_Helper_ExpedienteEdicion_SoloCamposSucios_DosCambios_DevuelveTituloYEstado = BuildJsonFail("expected ImporteLicitacion en sucios, got: " & m_Campos, logs)
        Exit Function
    End If
    If InStr(m_Campos, "CodExp") > 0 Then
        Test_Helper_ExpedienteEdicion_SoloCamposSucios_DosCambios_DevuelveTituloYEstado = BuildJsonFail("CodExp NO deberia estar en sucios (no cambio), got: " & m_Campos, logs)
        Exit Function
    End If
    logs(2) = "Assert: Titulo + ImporteLicitacion en sucios; CodExp NO esta"
    Test_Helper_ExpedienteEdicion_SoloCamposSucios_DosCambios_DevuelveTituloYEstado = BuildJsonOk("ok", logs)
End Function

' BR-02-01: SoloCamposSucios sin cambios -> string vacio
Public Function Test_Helper_ExpedienteEdicion_SoloCamposSucios_SinCambios_DevuelveVacio() As String
    Dim logs(0 To 2) As String
    Dim anterior As MockExpedienteEdicion
    Dim nuevo As MockExpedienteEdicion
    Set anterior = New MockExpedienteEdicion
    Set nuevo = New MockExpedienteEdicion
    anterior.CodExp = "EXP-001" : nuevo.CodExp = "EXP-001"
    anterior.Titulo = "T" : nuevo.Titulo = "T"
    Dim camposSucios As Variant
    Dim errMsg As String
    Dim result As String
    logs(0) = "Arrange: anterior y nuevo identicos"
    result = Helper_ExpedienteEdicion.SoloCamposSucios(anterior, nuevo, Array("CodExp", "Titulo"), camposSucios, errMsg)
    logs(1) = "Act: called"
    If result <> "" Then
        Test_Helper_ExpedienteEdicion_SoloCamposSucios_SinCambios_DevuelveVacio = BuildJsonFail("expected success, got: " & result, logs)
        Exit Function
    End If
    If CStr(camposSucios) <> "" Then
        Test_Helper_ExpedienteEdicion_SoloCamposSucios_SinCambios_DevuelveVacio = BuildJsonFail("expected vacio, got: " & CStr(camposSucios), logs)
        Exit Function
    End If
    logs(2) = "Assert: vacio"
    Test_Helper_ExpedienteEdicion_SoloCamposSucios_SinCambios_DevuelveVacio = BuildJsonOk("ok", logs)
End Function

' BR-04-01: Cambiar a Lote sin padre -> NO + motivo
Public Function Test_Helper_ExpedienteEdicion_CambiarTipo_RegularALote_SinPadre_DevuelveNO() As String
    Dim logs(0 To 2) As String
    Dim motivo As String
    Dim errMsg As String
    Dim result As String
    logs(0) = "Arrange: EsExpediente='Sí' (Regular), TipoNuevo='Lote', sin padre"
    result = Helper_ExpedienteEdicion.CambiarTipoRequierePadre("No", "No", "Sí", "No", "Lote", "", motivo, errMsg)
    logs(1) = "Act: called"
    If result <> "NO" Then
        Test_Helper_ExpedienteEdicion_CambiarTipo_RegularALote_SinPadre_DevuelveNO = BuildJsonFail("expected NO, got: " & result, logs)
        Exit Function
    End If
    If InStr(motivo, "Lote") = 0 Then
        Test_Helper_ExpedienteEdicion_CambiarTipo_RegularALote_SinPadre_DevuelveNO = BuildJsonFail("expected motivo about Lote, got: " & motivo, logs)
        Exit Function
    End If
    logs(2) = "Assert: NO + motivo about Lote+padre"
    Test_Helper_ExpedienteEdicion_CambiarTipo_RegularALote_SinPadre_DevuelveNO = BuildJsonOk("ok", logs)
End Function

' BR-04-01: Cambiar a Lote CON padre -> OK
Public Function Test_Helper_ExpedienteEdicion_CambiarTipo_RegularALote_ConPadre_DevuelveOK() As String
    Dim logs(0 To 2) As String
    Dim motivo As String
    Dim errMsg As String
    Dim result As String
    logs(0) = "Arrange: EsExpediente='Sí' (Regular), TipoNuevo='Lote', con padre"
    result = Helper_ExpedienteEdicion.CambiarTipoRequierePadre("No", "No", "Sí", "No", "Lote", "12345", motivo, errMsg)
    logs(1) = "Act: called"
    If result <> "OK" Then
        Test_Helper_ExpedienteEdicion_CambiarTipo_RegularALote_ConPadre_DevuelveOK = BuildJsonFail("expected OK, got: " & result, logs)
        Exit Function
    End If
    logs(2) = "Assert: OK (cambio a Lote con padre es valido)"
    Test_Helper_ExpedienteEdicion_CambiarTipo_RegularALote_ConPadre_DevuelveOK = BuildJsonOk("ok", logs)
End Function

' BR-05-01: EstadoCalculado sin FechaInicio -> "No iniciado"
Public Function Test_Helper_ExpedienteEdicion_EstadoCalculado_SinFechaInicio_DevuelveNoIniciado() As String
    Dim logs(0 To 2) As String
    Dim estado As String
    Dim errMsg As String
    Dim result As String
    logs(0) = "Arrange: FechaInicio=Null, FechaFin=Null, EstadoManual=''"
    result = Helper_ExpedienteEdicion.EstadoCalculadoTexto(Null, Null, "", estado, errMsg)
    logs(1) = "Act: called"
    If result <> "OK" Then
        Test_Helper_ExpedienteEdicion_EstadoCalculado_SinFechaInicio_DevuelveNoIniciado = BuildJsonFail("expected OK, got: " & result, logs)
        Exit Function
    End If
    If estado <> "No iniciado" Then
        Test_Helper_ExpedienteEdicion_EstadoCalculado_SinFechaInicio_DevuelveNoIniciado = BuildJsonFail("expected 'No iniciado', got: " & estado, logs)
        Exit Function
    End If
    logs(2) = "Assert: 'No iniciado'"
    Test_Helper_ExpedienteEdicion_EstadoCalculado_SinFechaInicio_DevuelveNoIniciado = BuildJsonOk("ok", logs)
End Function

' BR-05-01: EstadoCalculado con EstadoManual set -> override gana
Public Function Test_Helper_ExpedienteEdicion_EstadoCalculado_EstadoManualOverride_DevuelveEstadoManual() As String
    Dim logs(0 To 2) As String
    Dim estado As String
    Dim errMsg As String
    Dim result As String
    logs(0) = "Arrange: sin fechas pero EstadoManual='Vigente'"
    result = Helper_ExpedienteEdicion.EstadoCalculadoTexto(Null, Null, "Vigente", estado, errMsg)
    logs(1) = "Act: called"
    If result <> "OK" Then
        Test_Helper_ExpedienteEdicion_EstadoCalculado_EstadoManualOverride_DevuelveEstadoManual = BuildJsonFail("expected OK, got: " & result, logs)
        Exit Function
    End If
    If estado <> "Vigente" Then
        Test_Helper_ExpedienteEdicion_EstadoCalculado_EstadoManualOverride_DevuelveEstadoManual = BuildJsonFail("expected 'Vigente' (manual override), got: " & estado, logs)
        Exit Function
    End If
    logs(2) = "Assert: 'Vigente' (EstadoManual gana sobre el calculo)"
    Test_Helper_ExpedienteEdicion_EstadoCalculado_EstadoManualOverride_DevuelveEstadoManual = BuildJsonOk("ok", logs)
End Function
