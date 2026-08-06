Attribute VB_Name = "mod0BDOpcionesHelper"
Option Compare Database
Option Explicit

' Pure-data helpers for Form_Form0BDOpciones.
' Public prefix: Opciones_*

Private Function BuildJsonPayload(ByVal p_Ok As Boolean, ByVal p_Payload As Object, ByVal p_ErrorMsg As String, ByRef p_Logs() As String) As String
    Dim payloadJson As String
    If p_Payload Is Nothing Then payloadJson = "null" Else payloadJson = JsonConverter.ConvertToJson(p_Payload)
    Dim errorJson As String
    If p_Ok Then errorJson = "null" Else errorJson = """" & TestHelper.EscapeJsonString(p_ErrorMsg) & """"
    BuildJsonPayload = "{""ok"":" & LCase$(CStr(p_Ok)) & ",""value"":null,""payload"":" & payloadJson & ",""error"":" & errorJson & ",""logs"":" & TestHelper.JsonStringArray(p_Logs) & "}"
End Function

Private Function Envelope(ByVal p_Payload As Object, Optional ByVal p_ErrorMsg As String = "") As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    Envelope = BuildJsonPayload(Len(p_ErrorMsg) = 0, p_Payload, p_ErrorMsg, logs)
End Function

Public Function Opciones_FormOpenState(ByVal p_EsAdministrador As Boolean, ByVal p_Version As String, Optional ByRef p_Error As String) As String
    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")
    p_Error = ""
    payload("altaEnabled") = p_EsAdministrador
    payload("gestorEntidadesEnabled") = p_EsAdministrador
    payload("buscadorTecnicoVisible") = p_EsAdministrador
    payload("gestionEnabled") = True
    payload("versionCaption") = "Versión: " & p_Version
    payload("estadoVisible") = False
    Opciones_FormOpenState = Envelope(payload)
End Function

Public Function Opciones_GestionExpedientesTarget(ByVal p_EsTecnico As Boolean, Optional ByRef p_Error As String) As String
    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")
    p_Error = ""
    If p_EsTecnico Then payload("formName") = "FormExpedientesGestionTecnica" Else payload("formName") = "FormExpedientesGestion"
    payload("showProgress") = True
    Opciones_GestionExpedientesTarget = Envelope(payload)
End Function

Public Function Opciones_OpenFormAction(ByVal p_Action As String, Optional ByRef p_Error As String) As String
    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")
    p_Error = ""
    Select Case LCase$(Trim$(p_Action))
        Case "alta"
            payload("formName") = "FormExpedienteAltaTipo"
            payload("openArgs") = "DesdeInicio"
            payload("closeIfOpen") = False
        Case "buscador-tecnico"
            payload("formName") = "Form0BDOpcionesTecnicos"
            payload("openArgs") = ""
            payload("closeIfOpen") = True
        Case "gestor-entidades"
            payload("formName") = "Form0BDGestorEntidades"
            payload("openArgs") = ""
            payload("closeIfOpen") = True
        Case "batch-e2e"
            payload("formName") = "FormE2EGestionBatch"
            payload("openArgs") = ""
            payload("closeIfOpen") = True
        Case "tareas"
            payload("formName") = "FormTareas"
            payload("openArgs") = ""
            payload("closeIfOpen") = False
        Case Else
            p_Error = "Acción no reconocida"
            Opciones_OpenFormAction = Envelope(Nothing, p_Error)
            Exit Function
    End Select
    Opciones_OpenFormAction = Envelope(payload)
End Function

Public Function Opciones_LabelDispatch(ByVal p_Enabled As Boolean, ByVal p_ActionName As String, Optional ByRef p_Error As String) As String
    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")
    p_Error = ""
    If p_Enabled Then payload("action") = p_ActionName Else payload("action") = "none"
    Opciones_LabelDispatch = Envelope(payload)
End Function
