Attribute VB_Name = "modFormulario1Helper"
Option Compare Database
Option Explicit

' Pure-data helpers for Form_Formulario1.
' Public prefix: Formulario1_*

Private Function BuildJsonPayload(ByVal p_Ok As Boolean, ByVal p_Payload As Object, ByVal p_ErrorMsg As String, ByRef p_Logs() As String) As String
    Dim payloadJson As String
    If p_Payload Is Nothing Then
        payloadJson = "null"
    Else
        payloadJson = JsonConverter.ConvertToJson(p_Payload)
    End If

    Dim errorJson As String
    If p_Ok Then
        errorJson = "null"
    Else
        errorJson = """" & TestHelper.EscapeJsonString(p_ErrorMsg) & """"
    End If

    BuildJsonPayload = "{""ok"":" & LCase$(CStr(p_Ok)) & ",""value"":null,""payload"":" & payloadJson & ",""error"":" & errorJson & ",""logs"":" & TestHelper.JsonStringArray(p_Logs) & "}"
End Function

Private Function JsStringLiteral(ByVal p_Value As String) As String
    JsStringLiteral = Replace(p_Value, "\", "\\")
    JsStringLiteral = Replace(JsStringLiteral, "'", "\'")
    JsStringLiteral = Replace(JsStringLiteral, vbCrLf, "\n")
    JsStringLiteral = Replace(JsStringLiteral, vbCr, "\n")
    JsStringLiteral = Replace(JsStringLiteral, vbLf, "\n")
End Function

Public Function Formulario1_BuildTempFilePath(ByVal p_TempDir As String, Optional ByRef p_Error As String) As String
    p_Error = ""
    If Len(Trim$(p_TempDir)) = 0 Then
        p_Error = "BuildTempFilePath: p_TempDir is required"
        Formulario1_BuildTempFilePath = ""
        Exit Function
    End If
    Formulario1_BuildTempFilePath = p_TempDir & "\firebaseRealtime.html"
End Function

Public Function Formulario1_BuildFirebaseHtml(Optional ByRef p_Error As String) As String
    p_Error = ""
    Formulario1_BuildFirebaseHtml = "<!DOCTYPE html>" & vbCrLf & _
        "<html><head>" & vbCrLf & _
        "<script src='https://www.gstatic.com/firebasejs/9.6.1/firebase-app.js'></script>" & vbCrLf & _
        "<script>" & vbCrLf & _
        "const firebaseConfig = {" & vbCrLf & _
        "  apiKey: 'AIzaSyCaPdLLadH3-naRZIwIm0He0ztm7DUFB9A'," & vbCrLf & _
        "  authDomain: 'defensayseguridad-c5a21.firebaseapp.com'," & vbCrLf & _
        "  databaseURL: 'https://defensayseguridad-c5a21-default-rtdb.europe-west1.firebasedatabase.app'," & vbCrLf & _
        "  projectId: 'defensayseguridad-c5a21'," & vbCrLf & _
        "  storageBucket: 'defensayseguridad-c5a21.firebasestorage.app'," & vbCrLf & _
        "  messagingSenderId: '333347675713'," & vbCrLf & _
        "  appId: '1:333347675713:web:da5d23cbc8113fa0de30c0'" & vbCrLf & _
        "};" & vbCrLf & _
        "firebase.initializeApp(firebaseConfig);" & vbCrLf & _
        "const db = firebase.database();" & vbCrLf & _
        "db.ref('notificaEsto').on('value', (snapshot) => {" & vbCrLf & _
        "  const value = snapshot.val();" & vbCrLf & _
        "  window.external.UpdateRealtime(value);" & vbCrLf & _
        "});" & vbCrLf & _
        "</script>" & vbCrLf & _
        "</head><body><h1>Escuchando cambios...</h1></body></html>"
End Function

Public Function Formulario1_BuildFirebaseSetScript(ByVal p_Message As String, Optional ByRef p_Error As String) As String
    p_Error = ""
    Formulario1_BuildFirebaseSetScript = "firebase.database().ref('notificaesto').set('" & JsStringLiteral(p_Message) & "');"
End Function

Public Function Formulario1_RealtimePayload(ByVal p_Value As String, Optional ByRef p_Error As String) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(1)
    p_Error = ""

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")
    payload("value") = p_Value
    Formulario1_RealtimePayload = BuildJsonPayload(True, payload, "", logs)
End Function
