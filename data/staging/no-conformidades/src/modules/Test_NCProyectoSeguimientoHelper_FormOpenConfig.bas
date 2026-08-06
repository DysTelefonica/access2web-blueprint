Attribute VB_Name = "Test_NCProyectoSeguimientoHelper_FormOpenConfig"
Option Compare Database
Option Explicit

' ============================================
' MÓDULO DE TEST — NCProyectoSeguimientoHelper.InicializarConfigParaFormOpen (Issue #85)
' ============================================
' Test scenario: NCProyectoSeguimientoHelper.InicializarConfigParaFormOpen() returns a
' Dictionary with the three configuration keys the form's Form_Open should apply:
'   - CargaInicialPendiente: True  (loads indicadores on first Form_Timer tick)
'   - IndicadoresReiniciando: EnumSino.No
'   - TimerIntervalMs: 100
'
' Helper contract: stateless (no module-level state), no DB, no UI, no Me/DoCmd.
' Pure function — same input (no args) → same Dictionary shape.
' Fixture strategy: none required (no DB/UI touched). Tests assert dictionary shape.
' ============================================

Public Function Test_NCProyectoSeguimientoHelper_InicializarConfigParaFormOpen_HasExpectedValues_Atomic() As String
    On Error GoTo EH

    Dim logs As Collection
    Dim config As Object
    Dim assertError As String

    Set logs = TestHelper.NewLogs

    ' Act
    Set config = NCProyectoSeguimientoHelper.InicializarConfigParaFormOpen()

    ' Assert 1: config is not Nothing
    If Not TestHelper.AssertTrue(Not config Is Nothing, _
             "Assert1: InicializarConfigParaFormOpen debe devolver un Dictionary (no Nothing)", _
             logs, assertError) Then
        Test_NCProyectoSeguimientoHelper_InicializarConfigParaFormOpen_HasExpectedValues_Atomic = TestHelper.BuildJsonFail(assertError, logs)
        Exit Function
    End If

    ' Assert 2: CargaInicialPendiente = True
    If Not TestHelper.AssertTrue(config.Exists("CargaInicialPendiente"), _
             "Assert2: config debe contener la clave 'CargaInicialPendiente'", _
             logs, assertError) Then
        Test_NCProyectoSeguimientoHelper_InicializarConfigParaFormOpen_HasExpectedValues_Atomic = TestHelper.BuildJsonFail(assertError, logs)
        Exit Function
    End If
    If Not TestHelper.AssertTrue(config("CargaInicialPendiente") = True, _
             "Assert2b: 'CargaInicialPendiente' debe ser True; got: " & config("CargaInicialPendiente"), _
             logs, assertError) Then
        Test_NCProyectoSeguimientoHelper_InicializarConfigParaFormOpen_HasExpectedValues_Atomic = TestHelper.BuildJsonFail(assertError, logs)
        Exit Function
    End If
    TestHelper.AddLog logs, "Assert2: CargaInicialPendiente=True"

    ' Assert 3: IndicadoresReiniciando = EnumSino.No
    If Not TestHelper.AssertTrue(config.Exists("IndicadoresReiniciando"), _
             "Assert3: config debe contener la clave 'IndicadoresReiniciando'", _
             logs, assertError) Then
        Test_NCProyectoSeguimientoHelper_InicializarConfigParaFormOpen_HasExpectedValues_Atomic = TestHelper.BuildJsonFail(assertError, logs)
        Exit Function
    End If
    If Not TestHelper.AssertTrue(config("IndicadoresReiniciando") = EnumSino.No, _
             "Assert3b: 'IndicadoresReiniciando' debe ser EnumSino.No (" & EnumSino.No & "); got: " & config("IndicadoresReiniciando"), _
             logs, assertError) Then
        Test_NCProyectoSeguimientoHelper_InicializarConfigParaFormOpen_HasExpectedValues_Atomic = TestHelper.BuildJsonFail(assertError, logs)
        Exit Function
    End If
    TestHelper.AddLog logs, "Assert3: IndicadoresReiniciando=EnumSino.No"

    ' Assert 4: TimerIntervalMs = 100
    If Not TestHelper.AssertTrue(config.Exists("TimerIntervalMs"), _
             "Assert4: config debe contener la clave 'TimerIntervalMs'", _
             logs, assertError) Then
        Test_NCProyectoSeguimientoHelper_InicializarConfigParaFormOpen_HasExpectedValues_Atomic = TestHelper.BuildJsonFail(assertError, logs)
        Exit Function
    End If
    If Not TestHelper.AssertTrue(CLng(config("TimerIntervalMs")) = 100, _
             "Assert4b: 'TimerIntervalMs' debe ser 100; got: " & config("TimerIntervalMs"), _
             logs, assertError) Then
        Test_NCProyectoSeguimientoHelper_InicializarConfigParaFormOpen_HasExpectedValues_Atomic = TestHelper.BuildJsonFail(assertError, logs)
        Exit Function
    End If
    TestHelper.AddLog logs, "Assert4: TimerIntervalMs=100"

    Test_NCProyectoSeguimientoHelper_InicializarConfigParaFormOpen_HasExpectedValues_Atomic = TestHelper.BuildJsonOk(logs, "form_open_config_ok")
    Exit Function

EH:
    TestHelper.AddLog logs, "Error: " & Err.Description
    Test_NCProyectoSeguimientoHelper_InicializarConfigParaFormOpen_HasExpectedValues_Atomic = TestHelper.BuildJsonFail("EH: " & Err.Description, logs)
End Function
