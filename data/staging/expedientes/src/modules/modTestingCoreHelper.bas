Attribute VB_Name = "modTestingCoreHelper"
Option Compare Database
Option Explicit

' modTestingCoreHelper — canonical testing-core helpers for TDD atoms.
'
' Implements access-vba-e2e-methodology rule #10 (shared test infrastructure
' must exist before any feature helper that runs atoms) and rule #9
' (per-module prefix on all public names: TestingCore_*).
'
' Architecture:
'   - Public API: TestingCore_BuildJsonOk, TestingCore_BuildJsonFail,
'     TestingCore_BuildOk, TestingCore_BuildFail, TestingCore_RaiseError,
'     TestingCore_HandleError, TestingCore_InitLogs, TestingCore_EscapeJsonString,
'     TestingCore_SqlText.
'   - Implementation: delegates to the legacy TestHelper.bas (BuildJsonOk,
'     BuildJsonFail, EscapeJsonString, JsonStringArray, SqlStr). Existing
'     19 Test modules continue to call TestHelper.* until migrated in a
'     later PR (preflight-audit.md: "deferred churn-only PR with no behavior
'     change").
'   - This module adds: InitLogs (String(N) zero-fill), RaiseError
'     (Err.Raise wrapper with LoggingContext source), HandleError
'     (returns canonical JSON failure given an Err context).
'
' Signatures MUST match spec SDD forms-thin-coverage design.md.
' NO ByRef db As DAO.Database (pure string/JSON constructors, e2e rule #2 honesty).
' Pure helpers; do NOT introduce DAO here.
' On Error GoTo EH pattern (vba-access §1.4). Split guards per §1.6.1.

' --- TestingCore_BuildJsonOk ------------------------------------------------------------
' Builds the canonical success JSON envelope:
'   {"ok":true,"value":<p_Value>,"payload":null,"error":null,"logs":[...]}
' Delegates to TestHelper.BuildJsonOk for the actual serialization.
Public Function TestingCore_BuildJsonOk(ByVal p_Value As Variant, ByRef p_Logs() As String) As String
    On Error GoTo EH
    TestingCore_BuildJsonOk = TestHelper.BuildJsonOk(p_Value, p_Logs)
    Exit Function
EH:
    ' Last-resort fail JSON. Should never fire because TestHelper.BuildJsonOk is pure.
    TestingCore_BuildJsonOk = "{""ok"":false,""value"":null,""payload"":null,""error"":""TestingCore_BuildJsonOk: " & Err.Description & """,""logs"":[]}"
End Function

' --- TestingCore_BuildJsonFail ----------------------------------------------------------
' Builds the canonical failure JSON envelope:
'   {"ok":false,"value":null,"payload":null,"error":"<p_Error>","logs":[...]}
' Delegates to TestHelper.BuildJsonFail for the actual serialization.
Public Function TestingCore_BuildJsonFail(ByVal p_Error As String, ByRef p_Logs() As String) As String
    On Error GoTo EH
    TestingCore_BuildJsonFail = TestHelper.BuildJsonFail(p_Error, p_Logs)
    Exit Function
EH:
    TestingCore_BuildJsonFail = "{""ok"":false,""value"":null,""payload"":null,""error"":""TestingCore_BuildJsonFail: " & Err.Description & """,""logs"":[]}"
End Function

' --- TestingCore_BuildOk ----------------------------------------------------------------
' Convenience wrapper for atoms that only need {ok, value, logs} (no payload).
' Returns the canonical success envelope with payload:null.
Public Function TestingCore_BuildOk(ByVal p_Value As Variant, ByRef p_Logs() As String) As String
    TestingCore_BuildOk = TestingCore_BuildJsonOk(p_Value, p_Logs)
End Function

' --- TestingCore_BuildFail --------------------------------------------------------------
' Convenience wrapper for atoms that only need {ok:false, error, logs} (no payload).
' Returns the canonical failure envelope with value:null and payload:null.
Public Function TestingCore_BuildFail(ByVal p_Error As String, ByRef p_Logs() As String) As String
    TestingCore_BuildFail = TestingCore_BuildJsonFail(p_Error, p_Logs)
End Function

' --- TestingCore_RaiseError -------------------------------------------------------------
' Raises a runtime error with the given code and description.
' Use inside a Public Function's EH block when an atom must propagate failure.
' Source attribute is "TestingCore_RaiseError" so the runtime trace is greppable.
Public Sub TestingCore_RaiseError(ByVal p_Code As Long, ByVal p_Desc As String)
    Err.Raise p_Code, "TestingCore_RaiseError", p_Desc
End Sub

' --- TestingCore_HandleError ------------------------------------------------------------
' Builds the canonical failure JSON for a propagated error.
' Returns the envelope: {"ok":false,"value":null,"payload":null,"error":"<p_Procedure>: <p_Module>: <Err.Description>","logs":[...]}
' The caller is expected to have already raised or cleared Err before calling this.
Public Function TestingCore_HandleError(ByVal p_Procedure As String, ByVal p_Module As String, ByRef p_Logs() As String) As String
    Dim errMsg As String
    errMsg = p_Procedure & ": " & p_Module & ": " & Err.Description
    TestingCore_HandleError = TestingCore_BuildJsonFail(errMsg, p_Logs)
End Function

' --- TestingCore_InitLogs ----------------------------------------------------------------
' Returns a zero-length String() array of size p_Size, pre-filled with "".
' Use as the canonical logs accumulator: Dim logs() As String : logs = TestingCore_InitLogs(8).
Public Function TestingCore_InitLogs(ByVal p_Size As Long) As String()
    Dim logs() As String
    If p_Size < 0 Then p_Size = 0
    ReDim logs(0 To p_Size)
    TestingCore_InitLogs = logs
End Function

' --- TestingCore_EscapeJsonString -------------------------------------------------------
' Escapes a string for safe inclusion inside a JSON string literal.
' Delegates to TestHelper.EscapeJsonString (legacy canonical implementation).
Public Function TestingCore_EscapeJsonString(ByVal p_S As String) As String
    TestingCore_EscapeJsonString = TestHelper.EscapeJsonString(p_S)
End Function

' --- TestingCore_SqlText -----------------------------------------------------------------
' Returns a SQL-safe literal of p_S: doubles single-quotes for embedded apostrophes.
' Useful for building ad-hoc DAO SQL fragments in test fixtures.
' For parameterised DAO work, prefer DAO.QueryDef with .Parameters (see vba-access §9).
' Delegates to TestHelper.SqlStr (legacy canonical implementation).
Public Function TestingCore_SqlText(ByVal p_S As String) As String
    TestingCore_SqlText = TestHelper.SqlStr(p_S)
End Function