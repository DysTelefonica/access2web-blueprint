Attribute VB_Name = "Test_ModificadoHelper"
Option Compare Database
Option Explicit

' Test_ModificadoHelper — canonical JSON TDD atoms for Phase 3.4 / PR41.

Private Function BuildOk(ByVal p_Value As Variant, ByRef p_Logs() As String) As String
    BuildOk = TestingCore_BuildOk(p_Value, p_Logs)
End Function

Private Function BuildFail(ByVal p_Error As String, ByRef p_Logs() As String) As String
    BuildFail = TestingCore_BuildFail(p_Error, p_Logs)
End Function

Private Function PayloadOf(ByVal p_Json As String) As Object
    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(p_Json)
    Set PayloadOf = parsed("payload")
End Function

Private Function DictWithDescripcion(ByVal p_Value As String) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("DESCRIPCION") = p_Value
    d("NModificado") = "1"
    d("FechaFirmaModificado") = "2026-01-01"
    d("FechaFinModificado") = "2026-02-01"
    Set DictWithDescripcion = d
End Function

Public Function Test_ModificadoHelper_FormOpenState_AdminEnablesRegister() As String
    Dim logs() As String: logs = TestingCore_InitLogs(2)
    On Error GoTo EH
    Dim errMsg As String, payload As Object
    Set payload = PayloadOf(modModificadoHelper.Modificado_FormOpenState(True, errMsg))
    If CBool(payload("registrarEnabled")) <> True Then Err.Raise 1001, , "expected registrar enabled"
    Test_ModificadoHelper_FormOpenState_AdminEnablesRegister = BuildOk("admin-state", logs)
    Exit Function
EH:
    Test_ModificadoHelper_FormOpenState_AdminEnablesRegister = BuildFail(Err.Description, logs)
End Function

Public Function Test_ModificadoHelper_BuildFormValues_MapsFields() As String
    Dim logs() As String: logs = TestingCore_InitLogs(2)
    On Error GoTo EH
    Dim errMsg As String, payload As Object
    Set payload = PayloadOf(modModificadoHelper.Modificado_BuildFormValues("desc", "2", "f1", "f2", errMsg))
    If CStr(payload("NModificado")) <> "2" Then Err.Raise 1001, , "expected NModificado"
    Test_ModificadoHelper_BuildFormValues_MapsFields = BuildOk("values", logs)
    Exit Function
EH:
    Test_ModificadoHelper_BuildFormValues_MapsFields = BuildFail(Err.Description, logs)
End Function

Public Function Test_ModificadoHelper_HasChanges_SameValuesFalse() As String
    Dim logs() As String: logs = TestingCore_InitLogs(2)
    On Error GoTo EH
    Dim errMsg As String, payload As Object
    Set payload = PayloadOf(modModificadoHelper.Modificado_HasChanges(DictWithDescripcion("a"), DictWithDescripcion("a"), errMsg))
    If CBool(payload("hasChanges")) <> False Then Err.Raise 1001, , "expected no changes"
    Test_ModificadoHelper_HasChanges_SameValuesFalse = BuildOk("same-values", logs)
    Exit Function
EH:
    Test_ModificadoHelper_HasChanges_SameValuesFalse = BuildFail(Err.Description, logs)
End Function

Public Function Test_ModificadoHelper_RegisterDecision_SadNoChanges() As String
    Dim logs() As String: logs = TestingCore_InitLogs(2)
    On Error GoTo EH
    Dim errMsg As String, json As String
    json = modModificadoHelper.Modificado_RegisterDecision(False, errMsg)
    If InStr(json, """ok"":false") = 0 Then Err.Raise 1001, , "expected ok=false"
    Test_ModificadoHelper_RegisterDecision_SadNoChanges = BuildOk("no-changes-rejected", logs)
    Exit Function
EH:
    Test_ModificadoHelper_RegisterDecision_SadNoChanges = BuildFail(Err.Description, logs)
End Function
