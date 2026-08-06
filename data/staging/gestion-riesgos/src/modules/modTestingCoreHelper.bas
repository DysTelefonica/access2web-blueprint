Attribute VB_Name = "modTestingCoreHelper"
Option Compare Database
Option Explicit

' ----------------------------------------------------------------------------
' Helper de infraestructura para tests TDD en VBA.
'   Skill ref: access-vba-e2e-methodology v1.2-draft #10 (helpers de
'   infraestructura) y "Anti-patterns added 1.2-draft" (reinventar
'   BuildOk/BuildFail en cada Test_*.bas).
'
'   Que hace: contrato JSON entre atomo y runtest, inicializacion de
'   logs, raise del error de VBA con el patron 1000 del proyecto,
'   escape de strings para JSON. Es la base que todos los Test_*.bas
'   importan; sin esto, cada test file redefine BuildOk/BuildFail
'   (deuda tecnica que esta epica quiere evitar).
'
'   Que NO hace: asserts de un feature concreto. Esos viven en
'   Test_<Feature>Helper.bas. Este helper es generic.
'
'   Naming: TestCore_<Verbo> por convencion del proyecto.
' ----------------------------------------------------------------------------

' ----------------------------------------------------------------------------
' Contrato JSON de salida
'   Formato estable entre atomo y harness:
'     ok=true  -> {"ok":true, "value":<payload>, "logs":[...]}
'     ok=false -> {"ok":false, "error":"<msg>", "logs":[...]}
'   El harness (dysflow.test_vba) parsea este formato. Cambiar el
'   formato rompe el harness. No tocar sin actualizar tests.vba.json
'   y el harness a la vez.
' ----------------------------------------------------------------------------
Public Function TestCore_BuildJsonOk(ByVal p_Value As Variant, ByRef p_Logs() As String) As String
    Dim m_LogsJson As String
    m_LogsJson = TestCore_LogsArrayToJson(p_Logs)
    TestCore_BuildJsonOk = "{""ok"":true,""value"":""" & TestCore_EscapeJsonString(CStr(p_Value)) & """,""logs"":" & m_LogsJson & "}"
End Function

Public Function TestCore_BuildJsonFail(ByVal p_ErrorMsg As String, ByRef p_Logs() As String) As String
    Dim m_LogsJson As String
    m_LogsJson = TestCore_LogsArrayToJson(p_Logs)
    TestCore_BuildJsonFail = "{""ok"":false,""error"":""" & TestCore_EscapeJsonString(p_ErrorMsg) & """,""logs"":" & m_LogsJson & "}"
End Function

' ----------------------------------------------------------------------------
' Wrappers de conveniencia (compatibilidad con Test_Helper.bas legacy)
'   El proyecto ya tiene Test_Helper.bas con BuildJsonOk/BuildJsonFail
'   de signatura similar. Estos wrappers llaman al legacy para no
'   romper los atomos pre-existentes que usan Test_Helper.BuildJsonOk
'   directamente. Si un atomo nuevo quiere usar el helper, debe
'   llamar a TestCore_BuildJsonOk (con prefijo).
' ----------------------------------------------------------------------------
Public Function TestCore_BuildOk(ByVal p_Value As Variant, ByRef p_Logs() As String) As String
    TestCore_BuildOk = Test_Helper.BuildJsonOk(p_Value, p_Logs)
End Function

Public Function TestCore_BuildFail(ByVal p_Msg As Variant, ByRef p_Logs() As String) As String
    ' Acepta String, Null, Nothing, Empty, Error. Coerce a String para
    ' que Test_Helper.BuildJsonFail (que requiere As String) no falle.
    ' CStr(Nothing) lanza "Variable de objeto o bloque With no
    ' establecido", por eso se filtra por TypeName antes.
    Dim m_Msg As String
    Select Case TypeName(p_Msg)
        Case "Null", "Empty", "Nothing", "Error"
            m_Msg = ""
        Case Else
            m_Msg = CStr(p_Msg)
    End Select
    TestCore_BuildFail = Test_Helper.BuildJsonFail(m_Msg, p_Logs)
End Function

' ----------------------------------------------------------------------------
' RaiseError: encapsula el patron de error de VBA del proyecto
'   (Err.Raise 1000 si venia de nosotros, propagar el original si
'   es de un sistema subyacente). Llamar antes de Exit Function
'   para que el caller reciba un error 1000 consistente.
' ----------------------------------------------------------------------------
Public Sub TestCore_RaiseError(ByRef p_Error As String, ByVal p_Source As String)
    Dim m_ErrNumber As Long
    m_ErrNumber = IIf(Err.Number = 0, 1000, Err.Number)
    If p_Error = "" Then
        p_Error = "El metodo " & p_Source & " ha devuelto el error: " & vbCrLf & Err.Description
    End If
    Err.Raise m_ErrNumber, p_Source, p_Error
End Sub

' ----------------------------------------------------------------------------
' InitLogs: crea un array de logs pre-rellenado con espacios.
'   El atomo declara `Dim logs(0 To N) As String`, pero si quiere
'   que el array se rellene con "" (en lugar de variants vacios),
'   llama a InitLogs(N). Util para evitar "" vs Empty confusion.
' ----------------------------------------------------------------------------
Public Function TestCore_InitLogs(ByVal p_Count As Long) As Variant
    Dim m_Logs() As String
    ReDim m_Logs(0 To p_Count - 1)
    Dim i As Long
    For i = 0 To p_Count - 1
        m_Logs(i) = ""
    Next i
    TestCore_InitLogs = m_Logs
End Function

' ----------------------------------------------------------------------------
' Helpers privados (no parte del contrato publico)
' ----------------------------------------------------------------------------
Private Function TestCore_LogsArrayToJson(ByRef p_Logs() As String) As String
    Dim i As Long
    Dim m_Pieces() As String
    If (UBound(p_Logs) - LBound(p_Logs) + 1) <= 0 Then
        TestCore_LogsArrayToJson = "[]"
        Exit Function
    End If
    ReDim m_Pieces(LBound(p_Logs) To UBound(p_Logs))
    For i = LBound(p_Logs) To UBound(p_Logs)
        m_Pieces(i) = """" & TestCore_EscapeJsonString(p_Logs(i)) & """"
    Next i
    TestCore_LogsArrayToJson = "[" & Join(m_Pieces, ",") & "]"
End Function

Private Function TestCore_EscapeJsonString(ByVal p_S As String) As String
    ' Escape minimo: backslash, comilla doble, salto de linea.
    ' Suficiente para mensajes de logs y errores (no para binarios).
    Dim m_S As String
    m_S = Replace(p_S, "\", "\\")
    m_S = Replace(m_S, """", "\""")
    m_S = Replace(m_S, vbCrLf, "\n")
    m_S = Replace(m_S, vbLf, "\n")
    m_S = Replace(m_S, vbCr, "\n")
    m_S = Replace(m_S, vbTab, "\t")
    TestCore_EscapeJsonString = m_S
End Function
