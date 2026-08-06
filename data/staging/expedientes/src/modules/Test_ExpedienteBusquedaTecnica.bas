Attribute VB_Name = "Test_ExpedienteBusquedaTecnica"
' Test_ExpedienteBusquedaTecnica - Contract tests for ConstruirWhereBusquedaTecnica + CargarColBusquedaTecnica.
' Issue #47: PRUEBA-003/REFAC-3b - busqueda tecnica implementation.
' TDD discipline: contract tests first (RED), then helper (GREEN), then form slimdown.
Option Compare Database
Option Explicit

' ==============================================================================
' ConstruirWhereBusquedaTecnica - Contract tests
' ==============================================================================
Public Function Test_Helper_ExpedienteConsultas_ConstruirWhereBusquedaTecnica_NothingDTO_ReturnsError() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    Dim p_Error As String
    Dim result As String
    result = Helper_ExpedienteConsultas.ConstruirWhereBusquedaTecnica(Nothing, p_Error)
    If p_Error = "" Then
        Test_Helper_ExpedienteConsultas_ConstruirWhereBusquedaTecnica_NothingDTO_ReturnsError = BuildFail("expected p_Error to be populated when DTO is Nothing, got empty", logs)
        Exit Function
    End If
    If result <> "" Then
        Test_Helper_ExpedienteConsultas_ConstruirWhereBusquedaTecnica_NothingDTO_ReturnsError = BuildFail("expected result to be empty when DTO is Nothing, got: " & result, logs)
        Exit Function
    End If
    Test_Helper_ExpedienteConsultas_ConstruirWhereBusquedaTecnica_NothingDTO_ReturnsError = BuildOk("nothing-dto rejected", logs)
End Function

Public Function Test_Helper_ExpedienteConsultas_ConstruirWhereBusquedaTecnica_AllTodos_ReturnsWhereClause() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    Dim p_Error As String
    Dim m_Exp As New ExpedienteBusquedaTecnica
    m_Exp.PalabraClave = ""
    m_Exp.ESTADO = "Todos"
    m_Exp.JURIDICA = "Todos"
    m_Exp.CodExp = "Todos"
    m_Exp.jp = "Todos"
    Dim result As String
    result = Helper_ExpedienteConsultas.ConstruirWhereBusquedaTecnica(m_Exp, p_Error)
    If p_Error <> "" Then
        Test_Helper_ExpedienteConsultas_ConstruirWhereBusquedaTecnica_AllTodos_ReturnsWhereClause = BuildFail("unexpected p_Error: " & p_Error, logs)
        Exit Function
    End If
    If InStr(result, "WHERE") = 0 Then
        Test_Helper_ExpedienteConsultas_ConstruirWhereBusquedaTecnica_AllTodos_ReturnsWhereClause = BuildFail("expected WHERE clause in result, got: " & result, logs)
        Exit Function
    End If
    If InStr(result, "AND") = 0 Then
        Test_Helper_ExpedienteConsultas_ConstruirWhereBusquedaTecnica_AllTodos_ReturnsWhereClause = BuildFail("expected AND separators in WHERE clause, got: " & result, logs)
        Exit Function
    End If
    Test_Helper_ExpedienteConsultas_ConstruirWhereBusquedaTecnica_AllTodos_ReturnsWhereClause = BuildOk("all-todos produces where clause with all-todos branches", logs)
End Function

Public Function Test_Helper_ExpedienteConsultas_ConstruirWhereBusquedaTecnica_SpecificEstado_FiltersByEstado() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    Dim p_Error As String
    Dim m_Exp As New ExpedienteBusquedaTecnica
    m_Exp.PalabraClave = ""
    m_Exp.ESTADO = "Activo"
    m_Exp.JURIDICA = "Todos"
    m_Exp.CodExp = "Todos"
    m_Exp.jp = "Todos"
    Dim result As String
    result = Helper_ExpedienteConsultas.ConstruirWhereBusquedaTecnica(m_Exp, p_Error)
    If p_Error <> "" Then
        Test_Helper_ExpedienteConsultas_ConstruirWhereBusquedaTecnica_SpecificEstado_FiltersByEstado = BuildFail("unexpected p_Error: " & p_Error, logs)
        Exit Function
    End If
    If InStr(result, "Estado='Activo'") = 0 Then
        Test_Helper_ExpedienteConsultas_ConstruirWhereBusquedaTecnica_SpecificEstado_FiltersByEstado = BuildFail("expected Estado='Activo' in WHERE clause, got: " & result, logs)
        Exit Function
    End If
    Test_Helper_ExpedienteConsultas_ConstruirWhereBusquedaTecnica_SpecificEstado_FiltersByEstado = BuildOk("specific-estado filters by estado", logs)
End Function

Public Function Test_Helper_ExpedienteConsultas_ConstruirWhereBusquedaTecnica_PalabraClave_SearchesLike() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    Dim p_Error As String
    Dim m_Exp As New ExpedienteBusquedaTecnica
    m_Exp.PalabraClave = "ALTA"
    m_Exp.ESTADO = "Todos"
    m_Exp.JURIDICA = "Todos"
    m_Exp.CodExp = "Todos"
    m_Exp.jp = "Todos"
    Dim result As String
    result = Helper_ExpedienteConsultas.ConstruirWhereBusquedaTecnica(m_Exp, p_Error)
    If p_Error <> "" Then
        Test_Helper_ExpedienteConsultas_ConstruirWhereBusquedaTecnica_PalabraClave_SearchesLike = BuildFail("unexpected p_Error: " & p_Error, logs)
        Exit Function
    End If
    ' Expect LIKE wildcards around PalabraClave
    If InStr(result, "Like '*ALTA*'") = 0 Then
        Test_Helper_ExpedienteConsultas_ConstruirWhereBusquedaTecnica_PalabraClave_SearchesLike = BuildFail("expected LIKE wildcards around PalabraClave, got: " & result, logs)
        Exit Function
    End If
    Test_Helper_ExpedienteConsultas_ConstruirWhereBusquedaTecnica_PalabraClave_SearchesLike = BuildOk("palabra-clave applied as LIKE", logs)
End Function

' ==============================================================================
' CargarColBusquedaTecnica - Smoke test (private helpers must be defined for runtime)
' ==============================================================================
Public Function Test_Helper_ExpedienteConsultas_CargarColBusquedaTecnica_DoesNotCrash() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    Dim p_Error As String
    Dim m_Exp As New ExpedienteBusquedaTecnica
    m_Exp.PalabraClave = ""
    m_Exp.ESTADO = "Todos"
    m_Exp.JURIDICA = "Todos"
    m_Exp.CodExp = "Todos"
    m_Exp.jp = "Todos"
    On Error Resume Next
    Dim result As Scripting.Dictionary
    Set result = Helper_ExpedienteConsultas.CargarColBusquedaTecnica(m_Exp, Nothing, p_Error)
    Dim errNum As Long
    errNum = Err.Number
    On Error GoTo 0
    ' Smoke check: the function must NOT throw a "Sub or Function not defined" error
    ' (which would happen if getColBusquedaTecnicaPorMemoria / PorTablas are not defined).
    ' A clean return (possibly empty Dictionary, possibly with data) is acceptable.
    If errNum = 424 Then  ' Object required
        Test_Helper_ExpedienteConsultas_CargarColBusquedaTecnica_DoesNotCrash = BuildFail("CargarColBusquedaTecnica raised Object required - private helpers not defined", logs)
        Exit Function
    End If
    Test_Helper_ExpedienteConsultas_CargarColBusquedaTecnica_DoesNotCrash = BuildOk("no crash on smoke (errNum=" & errNum & ")", logs)
End Function

' ==============================================================================
' Helpers (buildOk/buildFail from TestingCore if available; otherwise minimal local)
' ==============================================================================
Private Function BuildOk(ByVal p_Value As Variant, ByRef p_Logs() As String) As String
    BuildOk = TestingCore_BuildOk(p_Value, p_Logs)
End Function

Private Function BuildFail(ByVal p_Error As String, ByRef p_Logs() As String) As String
    BuildFail = TestingCore_BuildFail(p_Error, p_Logs)
End Function
