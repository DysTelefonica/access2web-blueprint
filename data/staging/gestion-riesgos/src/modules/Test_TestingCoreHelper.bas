Attribute VB_Name = "Test_TestingCoreHelper"
Option Compare Database
Option Explicit

' ----------------------------------------------------------------------------
' Atomos TDD del helper modTestingCoreHelper.
'   Skill ref: v1.2-draft #10. Convencion: estos atomos son los unicos
'   que NO usan TestCore_BuildOk/BuildFail (porque estan probando
'   el helper que los provee). Usan asserts directos sobre la salida.
' ----------------------------------------------------------------------------

' Helper local: convierte una respuesta a booleano de "ok".
Private Function IsOk(ByVal p_Json As String) As Boolean
    IsOk = (InStr(1, p_Json, """ok"":true", vbTextCompare) > 0)
End Function

' Helper local: extrae el "error" del JSON.
Private Function ExtractError(ByVal p_Json As String) As String
    Dim m_Start As Long, m_End As Long
    m_Start = InStr(1, p_Json, """error"":""", vbTextCompare)
    If m_Start = 0 Then Exit Function
    m_Start = m_Start + Len("""error"":""")
    m_End = InStr(m_Start, p_Json, """", vbTextCompare)
    If m_End = 0 Then Exit Function
    ExtractError = Mid$(p_Json, m_Start, m_End - m_Start)
End Function

' ----------------------------------------------------------------------------
' ATOM 1 - Happy: TestCore_BuildJsonOk con valor string simple
' ----------------------------------------------------------------------------
Public Function Test_TestCore_BuildJsonOk_ValorSimple() As String
    On Error GoTo EH
    Dim logs(0 To 0) As String
    logs(0) = "1. Arrange"
    Dim m_Out As String
    m_Out = modTestingCoreHelper.TestCore_BuildJsonOk("ok", logs)
    If Not IsOk(m_Out) Then
        Test_TestCore_BuildJsonOk_ValorSimple = "{""ok"":false,""error"":""no retorno ok=true"",""logs"":[]}"
        Exit Function
    End If
    If InStr(1, m_Out, """value"":""ok""", vbTextCompare) = 0 Then
        Test_TestCore_BuildJsonOk_ValorSimple = "{""ok"":false,""error"":""valor no presente"",""logs"":[]}"
        Exit Function
    End If
    Test_TestCore_BuildJsonOk_ValorSimple = "{""ok"":true,""value"":""ok"",""logs"":[""1. Arrange""]}"
    Exit Function
EH:
    Test_TestCore_BuildJsonOk_ValorSimple = "{""ok"":false,""error"":""Test_TestCore_BuildJsonOk_ValorSimple: " & Err.Description & """,""logs"":[]}"
End Function

' ----------------------------------------------------------------------------
' ATOM 2 - Sad: TestCore_BuildJsonFail con mensaje simple
' ----------------------------------------------------------------------------
Public Function Test_TestCore_BuildJsonFail_Mensaje() As String
    On Error GoTo EH
    Dim logs(0 To 0) As String
    logs(0) = "1. Arrange"
    Dim m_Out As String
    m_Out = modTestingCoreHelper.TestCore_BuildJsonFail("algo fallo", logs)
    If IsOk(m_Out) Then
        Test_TestCore_BuildJsonFail_Mensaje = "{""ok"":false,""error"":""debio retornar ok=false"",""logs"":[]}"
        Exit Function
    End If
    If ExtractError(m_Out) <> "algo fallo" Then
        Test_TestCore_BuildJsonFail_Mensaje = "{""ok"":false,""error"":""mensaje no extraido bien: '" & ExtractError(m_Out) & "'"",""logs"":[]}"
        Exit Function
    End If
    Test_TestCore_BuildJsonFail_Mensaje = "{""ok"":true,""value"":""build_json_fail_mensaje_ok"",""logs"":[""1. Arrange""]}"
    Exit Function
EH:
    Test_TestCore_BuildJsonFail_Mensaje = "{""ok"":false,""error"":""Test_TestCore_BuildJsonFail_Mensaje: " & Err.Description & """,""logs"":[]}"
End Function

' ----------------------------------------------------------------------------
' ATOM 3 - Edge: TestCore_BuildOk con valor vacio
' ----------------------------------------------------------------------------
Public Function Test_TestCore_BuildOk_ValorVacio() As String
    On Error GoTo EH
    Dim logs(0 To 0) As String
    logs(0) = "1. Arrange"
    ' Test_Helper.BuildJsonOk acepta Empty, validamos que no crashea.
    Dim m_Out As String
    m_Out = modTestingCoreHelper.TestCore_BuildOk("", logs)
    ' El formato exacto lo define Test_Helper, solo verificamos que no falla
    ' y que el JSON es parseable (contiene "ok").
    If InStr(1, m_Out, """ok"":", vbTextCompare) = 0 Then
        Test_TestCore_BuildOk_ValorVacio = "{""ok"":false,""error"":""no devolvio JSON valido"",""logs"":[]}"
        Exit Function
    End If
    Test_TestCore_BuildOk_ValorVacio = "{""ok"":true,""value"":""build_ok_vacio_no_crashea"",""logs"":[""1. Arrange""]}"
    Exit Function
EH:
    Test_TestCore_BuildOk_ValorVacio = "{""ok"":false,""error"":""Test_TestCore_BuildOk_ValorVacio: " & Err.Description & """,""logs"":[]}"
End Function

' ----------------------------------------------------------------------------
' ATOM 4 - Edge: TestCore_BuildFail con Nothing como msg
' ----------------------------------------------------------------------------
Public Function Test_TestCore_BuildFail_Null() As String
    On Error GoTo EH
    Dim logs(0 To 0) As String
    logs(0) = "1. Arrange"
    Dim m_Out As String
    ' Test_Helper.BuildJsonFail acepta Null. Validamos que no crashea.
    m_Out = modTestingCoreHelper.TestCore_BuildFail(Nothing, logs)
    If InStr(1, m_Out, """ok"":", vbTextCompare) = 0 Then
        Test_TestCore_BuildFail_Null = "{""ok"":false,""error"":""no devolvio JSON"",""logs"":[]}"
        Exit Function
    End If
    Test_TestCore_BuildFail_Null = "{""ok"":true,""value"":""build_fail_null_no_crashea"",""logs"":[""1. Arrange""]}"
    Exit Function
EH:
    Test_TestCore_BuildFail_Null = "{""ok"":false,""error"":""Test_TestCore_BuildFail_Null: " & Err.Description & """,""logs"":[]}"
End Function

' ----------------------------------------------------------------------------
' ATOM 5 - Sad: TestCore_RaiseError pone Err.Number = 1000 cuando venia de 0
' ----------------------------------------------------------------------------
Public Function Test_TestCore_RaiseError_ConMensaje() As String
    On Error GoTo EH
    Dim m_Err As String
    Err.Clear
    m_Err = "fallo X"
    On Error Resume Next
    modTestingCoreHelper.TestCore_RaiseError m_Err, "MiFuncion"
    Dim m_Number As Long
    m_Number = Err.Number
    On Error GoTo EH
    If m_Number <> 1000 Then
        Test_TestCore_RaiseError_ConMensaje = "{""ok"":false,""error"":""Esperaba Err.Number=1000, dio " & CStr(m_Number) & """,""logs"":[]}"
        Exit Function
    End If
    Test_TestCore_RaiseError_ConMensaje = "{""ok"":true,""value"":""raise_error_numero_1000_ok"",""logs"":[]}"
    Exit Function
EH:
    Test_TestCore_RaiseError_ConMensaje = "{""ok"":false,""error"":""Test_TestCore_RaiseError_ConMensaje: " & Err.Description & """,""logs"":[]}"
End Function

' ----------------------------------------------------------------------------
' ATOM 6 - Happy: TestCore_InitLogs devuelve array de N strings vacios
' ----------------------------------------------------------------------------
Public Function Test_TestCore_InitLogs_Tamano3() As String
    On Error GoTo EH
    Dim m_Logs As Variant
    m_Logs = modTestingCoreHelper.TestCore_InitLogs(3)
    If UBound(m_Logs) - LBound(m_Logs) + 1 <> 3 Then
        Test_TestCore_InitLogs_Tamano3 = "{""ok"":false,""error"":""tamano esperado 3, dio " & CStr(UBound(m_Logs) - LBound(m_Logs) + 1) & """,""logs"":[]}"
        Exit Function
    End If
    If m_Logs(LBound(m_Logs)) <> "" Or m_Logs(UBound(m_Logs)) <> "" Then
        Test_TestCore_InitLogs_Tamano3 = "{""ok"":false,""error"":""elementos no vacios"",""logs"":[]}"
        Exit Function
    End If
    Test_TestCore_InitLogs_Tamano3 = "{""ok"":true,""value"":""init_logs_3_vacios_ok"",""logs"":[]}"
    Exit Function
EH:
    Test_TestCore_InitLogs_Tamano3 = "{""ok"":false,""error"":""Test_TestCore_InitLogs_Tamano3: " & Err.Description & """,""logs"":[]}"
End Function
