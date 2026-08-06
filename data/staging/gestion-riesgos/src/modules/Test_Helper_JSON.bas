Attribute VB_Name = "Test_Helper_JSON"
Option Compare Database
Option Explicit

' ============================================================
' Test_Helper_JSON — Helpers JSON para toda la suite
' BuildJsonOk / BuildJsonFail / EscapeJsonString
' Idempotentes, seguros para COM, sin concatenación manual
' ============================================================

Private Function BuildJsonOk(ByVal value As Variant, ByRef logs() As String) As String
    BuildJsonOk = "{""ok"":true,""value"":" & JsonValue(value) & ",""payload"":null,""error"":null,""logs"":" & JsonStringArray(logs) & "}"
End Function

Private Function BuildJsonFail(ByVal errorMsg As String, ByRef logs() As String) As String
    BuildJsonFail = "{""ok"":false,""value"":null,""payload"":null,""error"":""" & EscapeJsonString(errorMsg) & """,""logs"":" & JsonStringArray(logs) & "}"
End Function

Private Function EscapeJsonString(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbTab, "\t")
    EscapeJsonString = s
End Function

Private Function JsonValue(ByVal value As Variant) As String
    If IsNull(value) Or IsEmpty(value) Then
        JsonValue = "null"
        Exit Function
    End If
    Select Case VarType(value)
        Case vbBoolean
            JsonValue = LCase$(CStr(value))
        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
            JsonValue = Replace(CStr(value), ",", ".")
        Case Else
            JsonValue = """" & EscapeJsonString(CStr(value)) & """"
    End Select
End Function

Private Function JsonStringArray(ByRef logs() As String) As String
    On Error GoTo EmptyLogs
    Dim i As Long
    Dim parts As String
    Dim entry As String
    For i = LBound(logs) To UBound(logs)
        entry = logs(i)
        If Len(entry) > 0 Then
            If Len(parts) > 0 Then parts = parts & ","
            parts = parts & """" & EscapeJsonString(entry) & """"
        End If
    Next i
    JsonStringArray = "[" & parts & "]"
    Exit Function
EmptyLogs:
    JsonStringArray = "[]"
End Function
