Attribute VB_Name = "modCPVAltaHelper"
Option Compare Database
Option Explicit

' modCPVAltaHelper — REWORK (2026-06-26)
' Pure-data helpers for Form_FormCPV (Alta/Edición form).
'
' Implements access-vba-e2e-methodology rule #1 (forms are thin UI wiring),
' rule #2 (honest helper signatures), rule #5 (per-module prefix: CPVAlta_*),
' and rule #9 (no `ByRef p_Form`).
'
' IMPORTANT: This module covers the Alta/Edición form (Form_FormCPV),
' which is structurally different from the Gestion list (Form_FormCPVsGestion,
' covered by modCPVHelper). Prefix disambiguation:
'   - `CPVAlta_*` — Alta/Edición form helpers (this module)
'   - `CPV_*` (no Alta suffix) — Gestion list helpers (modCPVHelper)
'
' New design (4 helpers, mirror of modComercialAltaHelper):
'   1. CPVAlta_Abrir_Inicializar(p_Modo, p_IDEntidad, ByRef p_Entidad, p_Error)
'   2. CPVAlta_VerificarCambios(p_ValoresActuales, p_ValoresOriginales, p_Error)
'   3. CPVAlta_Registrar(p_Valores, p_Error)
'   4. CPVAlta_Cerrar(p_Error)
'
' Telefonica D&S convention (vba-access §1.4.1): every Public Function ends with
' `Optional ByRef p_Error As String` as the LAST parameter.

' === Module-level constants (all at top per vba-access §10.1) ============================

Private Const CPV_ALTA_TITULO_ALTA As String = "ALTA DE CPV"
Private Const CPV_ALTA_TITULO_EDICION As String = "EDICIÓN DE CPV"

Private Const CPV_ALTA_MODO_ALTA As String = "alta"
Private Const CPV_ALTA_MODO_EDICION As String = "edicion"

Private Const CPV_ALTA_FIELD_NAME As String = "CPV"
Private Const CPV_ALTA_FIELD_DESCRIPCION As String = "DESCRIPCION"

Private Const CPV_ALTA_TEST_ID_BASE As Long = 900560


' === Local helpers (all at top per vba-access §10.1) ====================================

Private Function BuildJsonPayload( _
    ByVal p_Ok As Boolean, _
    ByVal p_Payload As Object, _
    ByVal p_ErrorMsg As String, _
    ByRef p_Logs() As String _
) As String
    Dim payloadJson As String
    If p_Payload Is Nothing Then
        payloadJson = "null"
    Else
        payloadJson = JsonConverter.ConvertToJson(p_Payload)
    End If

    Dim logsJson As String
    logsJson = TestHelper.JsonStringArray(p_Logs)

    Dim errorJson As String
    If p_Ok Then
        errorJson = "null"
    Else
        errorJson = """" & TestHelper.EscapeJsonString(p_ErrorMsg) & """"
    End If

    BuildJsonPayload = "{""ok"":" & LCase$(CStr(p_Ok)) & _
                       ",""value"":null" & _
                       ",""payload"":" & payloadJson & _
                       ",""error"":" & errorJson & _
                       ",""logs"":" & logsJson & "}"
End Function

Private Function BuildEmptyEntidad() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d(CPV_ALTA_FIELD_NAME) = ""
    d(CPV_ALTA_FIELD_DESCRIPCION) = ""
    Set BuildEmptyEntidad = d
End Function

Private Function NormalizarModo(ByVal p_Modo As String) As String
    Dim m As String
    m = LCase$("" & p_Modo)
    If m = CPV_ALTA_MODO_ALTA Then
        NormalizarModo = CPV_ALTA_MODO_ALTA
    ElseIf m = CPV_ALTA_MODO_EDICION Then
        NormalizarModo = CPV_ALTA_MODO_EDICION
    Else
        NormalizarModo = ""
    End If
End Function

Private Function ExtraerValor(ByVal p_Dict As Object, ByVal p_Key As String) As String
    If p_Dict Is Nothing Then
        ExtraerValor = ""
        Exit Function
    End If
    If p_Dict.Exists(p_Key) Then
        ExtraerValor = CStr(p_Dict(p_Key))
    Else
        ExtraerValor = ""
    End If
End Function


' === Public API =========================================================================

' --- CPVAlta_Abrir_Inicializar -----------------------------------------
Public Function CPVAlta_Abrir_Inicializar( _
    ByVal p_Modo As String, _
    ByVal p_IDEntidad As String, _
    ByRef p_Entidad As Object, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(5)
    p_Error = ""

    On Error GoTo errores

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")

    Dim modo As String
    modo = NormalizarModo(p_Modo)
    If Len(modo) = 0 Then
        p_Error = "CPVAlta_Abrir_Inicializar: p_Modo must be 'alta' or 'edicion', got '" & p_Modo & "'"
        Set p_Entidad = Nothing
        CPVAlta_Abrir_Inicializar = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim titulo As String
    If modo = CPV_ALTA_MODO_ALTA Then
        titulo = CPV_ALTA_TITULO_ALTA
        Set p_Entidad = BuildEmptyEntidad()
        payload("titulo") = titulo
        payload("modo") = modo
        payload("hasEntidad") = False
        payload("id") = ""
        Set payload("entidad") = Nothing

        logs(0) = "Abrir_Inicializar: mode=alta, p_Entidad=empty"
        CPVAlta_Abrir_Inicializar = BuildJsonPayload(True, payload, "", logs)
        Exit Function
    End If

    Dim idEntidad As String
    idEntidad = Trim$("" & p_IDEntidad)
    If Len(idEntidad) = 0 Then
        p_Error = "CPVAlta_Abrir_Inicializar: p_IDEntidad is required in edicion mode"
        Set p_Entidad = Nothing
        CPVAlta_Abrir_Inicializar = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim loaded As CPV
    Dim loadErr As String
    Set loaded = constructor.getCPV(p_IDCPV:=idEntidad, p_Error:=loadErr)
    If Len(loadErr) > 0 Then
        p_Error = "CPVAlta_Abrir_Inicializar: " & loadErr
        Set p_Entidad = Nothing
        CPVAlta_Abrir_Inicializar = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If
    If loaded Is Nothing Then
        p_Error = "CPVAlta_Abrir_Inicializar: No se ha podido encontrar el CPV registrado (id=" & idEntidad & ")"
        Set p_Entidad = Nothing
        CPVAlta_Abrir_Inicializar = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Set p_Entidad = BuildEmptyEntidad()
    If loaded.CPV <> "" Then
        p_Entidad(CPV_ALTA_FIELD_NAME) = CStr(loaded.CPV)
    End If
    If loaded.DESCRIPCION <> "" Then
        p_Entidad(CPV_ALTA_FIELD_DESCRIPCION) = CStr(loaded.DESCRIPCION)
    End If

    Set m_ObjCPVActivo = loaded

    titulo = CPV_ALTA_TITULO_EDICION
    payload("titulo") = titulo
    payload("modo") = modo
    payload("hasEntidad") = True
    payload("id") = idEntidad
    Set payload("entidad") = p_Entidad

    logs(0) = "Abrir_Inicializar: mode=edicion, id=" & idEntidad
    logs(1) = "Abrir_Inicializar: precargados populated"

    CPVAlta_Abrir_Inicializar = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "CPVAlta_Abrir_Inicializar: " & Err.Description
    End If
    Set p_Entidad = Nothing
    CPVAlta_Abrir_Inicializar = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- CPVAlta_VerificarCambios -------------------------------------------
Public Function CPVAlta_VerificarCambios( _
    ByVal p_ValoresActuales As Object, _
    ByVal p_ValoresOriginales As Object, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(4)
    p_Error = ""

    On Error GoTo errores

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")

    If p_ValoresActuales Is Nothing Then
        p_Error = "CPVAlta_VerificarCambios: p_ValoresActuales is Nothing"
        CPVAlta_VerificarCambios = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim diffs As Object
    Set diffs = CreateObject("Scripting.Dictionary")

    Dim currentName As String
    Dim currentDesc As String

    currentName = ExtraerValor(p_ValoresActuales, CPV_ALTA_FIELD_NAME)
    currentDesc = ExtraerValor(p_ValoresActuales, CPV_ALTA_FIELD_DESCRIPCION)

    Dim hayCambios As Boolean
    hayCambios = False

    If p_ValoresOriginales Is Nothing Then
        If Len(currentName) > 0 Then
            diffs(CPV_ALTA_FIELD_NAME) = True
            hayCambios = True
        Else
            diffs(CPV_ALTA_FIELD_NAME) = False
        End If
        If Len(currentDesc) > 0 Then
            diffs(CPV_ALTA_FIELD_DESCRIPCION) = True
            hayCambios = True
        Else
            diffs(CPV_ALTA_FIELD_DESCRIPCION) = False
        End If
        logs(0) = "VerificarCambios: alta mode, hayCambios=" & CStr(hayCambios)
    Else
        Dim originalName As String
        Dim originalDesc As String

        originalName = ExtraerValor(p_ValoresOriginales, CPV_ALTA_FIELD_NAME)
        originalDesc = ExtraerValor(p_ValoresOriginales, CPV_ALTA_FIELD_DESCRIPCION)

        If StrComp(currentName, originalName, vbTextCompare) <> 0 Then
            diffs(CPV_ALTA_FIELD_NAME) = True
            hayCambios = True
        Else
            diffs(CPV_ALTA_FIELD_NAME) = False
        End If

        If StrComp(currentDesc, originalDesc, vbTextCompare) <> 0 Then
            diffs(CPV_ALTA_FIELD_DESCRIPCION) = True
            hayCambios = True
        Else
            diffs(CPV_ALTA_FIELD_DESCRIPCION) = False
        End If

        logs(0) = "VerificarCambios: edicion mode, hayCambios=" & CStr(hayCambios)
    End If

    payload("hayCambios") = hayCambios
    Set payload("diffs") = diffs

    CPVAlta_VerificarCambios = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "CPVAlta_VerificarCambios: " & Err.Description
    End If
    CPVAlta_VerificarCambios = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- CPVAlta_Registrar -----------------------------------------------
Public Function CPVAlta_Registrar( _
    ByVal p_Valores As Object, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(6)
    p_Error = ""

    On Error GoTo errores

    If p_Valores Is Nothing Then
        p_Error = "CPVAlta_Registrar: p_Valores is Nothing"
        CPVAlta_Registrar = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")

    Dim currentName As String
    Dim currentDesc As String

    currentName = ExtraerValor(p_Valores, CPV_ALTA_FIELD_NAME)
    currentDesc = ExtraerValor(p_Valores, CPV_ALTA_FIELD_DESCRIPCION)

    Dim modo As String
    Dim idEntidad As String
    If m_ObjCPVActivo Is Nothing Then
        modo = CPV_ALTA_MODO_ALTA
        idEntidad = ""
    Else
        modo = CPV_ALTA_MODO_EDICION
        If m_ObjCPVActivo.IDCPV <> "" Then
            idEntidad = CStr(m_ObjCPVActivo.IDCPV)
        Else
            idEntidad = ""
        End If
    End If

    logs(0) = "Registrar: mode=" & modo & ", id=" & idEntidad

    Dim cm As CPV
    If m_ObjCPVActivo Is Nothing Then
        Set cm = New CPV
    Else
        Set cm = m_ObjCPVActivo
    End If

    cm.CPV = currentName
    cm.DESCRIPCION = currentDesc

    Set m_ObjCPVActivo = cm

    Dim m_CPVOp As New CPVOperaciones
    Dim m_AlInicio As CPV
    If modo = CPV_ALTA_MODO_EDICION Then
        Dim loadErr As String
        Set m_AlInicio = constructor.getCPV(p_IDCPV:=idEntidad, p_Error:=loadErr)
    Else
        Set m_AlInicio = Nothing
    End If

    With m_CPVOp
        Set .CPV = cm
        Dim registrarErr As String
        .Registrar m_AlInicio, registrarErr
        If Len(registrarErr) > 0 Then
            p_Error = "CPVAlta_Registrar: " & registrarErr
            CPVAlta_Registrar = BuildJsonPayload(False, Nothing, p_Error, logs)
            Exit Function
        End If
    End With

    Dim resultingId As String
    If cm.IDCPV <> "" Then
        resultingId = CStr(cm.IDCPV)
    Else
        resultingId = idEntidad
    End If

    payload("ok") = True
    payload("id") = resultingId
    payload("modo") = modo

    logs(1) = "Registrar: persisted, resultingId=" & resultingId

    CPVAlta_Registrar = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "CPVAlta_Registrar: " & Err.Description
    End If
    CPVAlta_Registrar = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- CPVAlta_Cerrar --------------------------------------------------
Public Function CPVAlta_Cerrar( _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    p_Error = ""

    On Error GoTo errores

    Set m_ObjCPVActivo = Nothing

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")
    payload("ok") = True

    logs(0) = "Cerrar: m_ObjCPVActivo cleared"

    CPVAlta_Cerrar = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "CPVAlta_Cerrar: " & Err.Description
    End If
    CPVAlta_Cerrar = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function
