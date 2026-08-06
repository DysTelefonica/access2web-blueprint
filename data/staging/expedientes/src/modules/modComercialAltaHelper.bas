Attribute VB_Name = "modComercialAltaHelper"
Option Compare Database
Option Explicit

' modComercialAltaHelper — REWORK (2026-06-26)
' Pure-data helpers for Form_FormComercial (Alta/Edición form).
'
' Implements access-vba-e2e-methodology rule #1 (forms are thin UI wiring),
' rule #2 (honest helper signatures), rule #5 (per-module prefix: ComercialAlta_*),
' and rule #9 (no `ByRef p_Form`).
'
' IMPORTANT: This module covers the Alta/Edición form (Form_FormComercial),
' which is structurally different from the Gestion list (Form_FormComercialesGestion,
' covered by modComercialHelper). Prefix disambiguation:
'   - `ComercialAlta_*` — Alta/Edición form helpers (this module)
'   - `Comercial_*` (no Alta suffix) — Gestion list helpers (modComercialHelper)
'
' New design (4 helpers):
'   1. ComercialAlta_Abrir_Inicializar(p_Modo, p_IDEntidad, ByRef p_Entidad, p_Error)
'      Returns JSON: {ok, payload:{titulo, modo, hasEntidad, id, entidad}, error, logs}
'   2. ComercialAlta_VerificarCambios(p_ValoresActuales, p_ValoresOriginales, p_Error)
'      Returns JSON: {ok, payload:{hayCambios, diffs}, error, logs}
'   3. ComercialAlta_Registrar(p_Valores, p_Error)
'      The helper decides alta vs edicion based on the m_ObjComercialActivo global.
'   4. ComercialAlta_Cerrar(p_Error)
'
' Telefonica D&S convention (vba-access §1.4.1): every Public Function ends with
' `Optional ByRef p_Error As String` as the LAST parameter.

' === Module-level constants (all at top per vba-access §10.1) ============================

Private Const COMERCIAL_ALTA_TITULO_ALTA As String = "ALTA DE COMERCIAL"
Private Const COMERCIAL_ALTA_TITULO_EDICION As String = "EDICIÓN DE COMERCIAL"

Private Const COMERCIAL_ALTA_MODO_ALTA As String = "alta"
Private Const COMERCIAL_ALTA_MODO_EDICION As String = "edicion"

Private Const COMERCIAL_ALTA_FIELD_NAME As String = "Comercial"
Private Const COMERCIAL_ALTA_FIELD_DESCRIPCION As String = "DESCRIPCION"

Private Const COMERCIAL_ALTA_TEST_ID_BASE As Long = 900550


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

' --- BuildEmptyEntidad --------------------------------------------------------------
Private Function BuildEmptyEntidad() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d(COMERCIAL_ALTA_FIELD_NAME) = ""
    d(COMERCIAL_ALTA_FIELD_DESCRIPCION) = ""
    Set BuildEmptyEntidad = d
End Function

' --- NormalizarModo --------------------------------------------------------------
Private Function NormalizarModo(ByVal p_Modo As String) As String
    Dim m As String
    m = LCase$("" & p_Modo)
    If m = COMERCIAL_ALTA_MODO_ALTA Then
        NormalizarModo = COMERCIAL_ALTA_MODO_ALTA
    ElseIf m = COMERCIAL_ALTA_MODO_EDICION Then
        NormalizarModo = COMERCIAL_ALTA_MODO_EDICION
    Else
        NormalizarModo = ""
    End If
End Function

' --- ExtraerValor ----------------------------------------------------------------
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

' --- ComercialAlta_Abrir_Inicializar -----------------------------------------
Public Function ComercialAlta_Abrir_Inicializar( _
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
        p_Error = "ComercialAlta_Abrir_Inicializar: p_Modo must be 'alta' or 'edicion', got '" & p_Modo & "'"
        Set p_Entidad = Nothing
        ComercialAlta_Abrir_Inicializar = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim titulo As String
    If modo = COMERCIAL_ALTA_MODO_ALTA Then
        titulo = COMERCIAL_ALTA_TITULO_ALTA
        Set p_Entidad = BuildEmptyEntidad()
        payload("titulo") = titulo
        payload("modo") = modo
        payload("hasEntidad") = False
        payload("id") = ""
        Set payload("entidad") = Nothing

        logs(0) = "Abrir_Inicializar: mode=alta, p_Entidad=empty"
        ComercialAlta_Abrir_Inicializar = BuildJsonPayload(True, payload, "", logs)
        Exit Function
    End If

    ' Edicion mode — load the entity by ID.
    Dim idEntidad As String
    idEntidad = Trim$("" & p_IDEntidad)
    If Len(idEntidad) = 0 Then
        p_Error = "ComercialAlta_Abrir_Inicializar: p_IDEntidad is required in edicion mode"
        Set p_Entidad = Nothing
        ComercialAlta_Abrir_Inicializar = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim loaded As Comercial
    Dim loadErr As String
    Set loaded = constructor.getComercial(p_IDComercial:=idEntidad, p_Error:=loadErr)
    If Len(loadErr) > 0 Then
        p_Error = "ComercialAlta_Abrir_Inicializar: " & loadErr
        Set p_Entidad = Nothing
        ComercialAlta_Abrir_Inicializar = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If
    If loaded Is Nothing Then
        p_Error = "ComercialAlta_Abrir_Inicializar: No se ha podido encontrar el Comercial registrado (id=" & idEntidad & ")"
        Set p_Entidad = Nothing
        ComercialAlta_Abrir_Inicializar = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Set p_Entidad = BuildEmptyEntidad()
    If loaded.Comercial <> "" Then
        p_Entidad(COMERCIAL_ALTA_FIELD_NAME) = CStr(loaded.Comercial)
    End If
    If loaded.DESCRIPCION <> "" Then
        p_Entidad(COMERCIAL_ALTA_FIELD_DESCRIPCION) = CStr(loaded.DESCRIPCION)
    End If

    ' Update the project-conventional global so Registrar picks it up later.
    Set m_ObjComercialActivo = loaded

    titulo = COMERCIAL_ALTA_TITULO_EDICION
    payload("titulo") = titulo
    payload("modo") = modo
    payload("hasEntidad") = True
    payload("id") = idEntidad
    Set payload("entidad") = p_Entidad

    logs(0) = "Abrir_Inicializar: mode=edicion, id=" & idEntidad
    logs(1) = "Abrir_Inicializar: precargados populated"

    ComercialAlta_Abrir_Inicializar = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "ComercialAlta_Abrir_Inicializar: " & Err.Description
    End If
    Set p_Entidad = Nothing
    ComercialAlta_Abrir_Inicializar = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- ComercialAlta_VerificarCambios -------------------------------------------
Public Function ComercialAlta_VerificarCambios( _
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
        p_Error = "ComercialAlta_VerificarCambios: p_ValoresActuales is Nothing"
        ComercialAlta_VerificarCambios = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim diffs As Object
    Set diffs = CreateObject("Scripting.Dictionary")

    Dim currentName As String
    Dim currentDesc As String

    currentName = ExtraerValor(p_ValoresActuales, COMERCIAL_ALTA_FIELD_NAME)
    currentDesc = ExtraerValor(p_ValoresActuales, COMERCIAL_ALTA_FIELD_DESCRIPCION)

    Dim hayCambios As Boolean
    hayCambios = False

    If p_ValoresOriginales Is Nothing Then
        ' Alta mode — any non-empty value counts as a change.
        If Len(currentName) > 0 Then
            diffs(COMERCIAL_ALTA_FIELD_NAME) = True
            hayCambios = True
        Else
            diffs(COMERCIAL_ALTA_FIELD_NAME) = False
        End If
        If Len(currentDesc) > 0 Then
            diffs(COMERCIAL_ALTA_FIELD_DESCRIPCION) = True
            hayCambios = True
        Else
            diffs(COMERCIAL_ALTA_FIELD_DESCRIPCION) = False
        End If
        logs(0) = "VerificarCambios: alta mode, hayCambios=" & CStr(hayCambios)
    Else
        ' Edicion mode — diff per field.
        Dim originalName As String
        Dim originalDesc As String

        originalName = ExtraerValor(p_ValoresOriginales, COMERCIAL_ALTA_FIELD_NAME)
        originalDesc = ExtraerValor(p_ValoresOriginales, COMERCIAL_ALTA_FIELD_DESCRIPCION)

        If StrComp(currentName, originalName, vbTextCompare) <> 0 Then
            diffs(COMERCIAL_ALTA_FIELD_NAME) = True
            hayCambios = True
        Else
            diffs(COMERCIAL_ALTA_FIELD_NAME) = False
        End If

        If StrComp(currentDesc, originalDesc, vbTextCompare) <> 0 Then
            diffs(COMERCIAL_ALTA_FIELD_DESCRIPCION) = True
            hayCambios = True
        Else
            diffs(COMERCIAL_ALTA_FIELD_DESCRIPCION) = False
        End If

        logs(0) = "VerificarCambios: edicion mode, hayCambios=" & CStr(hayCambios)
    End If

    payload("hayCambios") = hayCambios
    Set payload("diffs") = diffs

    ComercialAlta_VerificarCambios = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "ComercialAlta_VerificarCambios: " & Err.Description
    End If
    ComercialAlta_VerificarCambios = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- ComercialAlta_Registrar -----------------------------------------------
Public Function ComercialAlta_Registrar( _
    ByVal p_Valores As Object, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(6)
    p_Error = ""

    On Error GoTo errores

    If p_Valores Is Nothing Then
        p_Error = "ComercialAlta_Registrar: p_Valores is Nothing"
        ComercialAlta_Registrar = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")

    Dim currentName As String
    Dim currentDesc As String

    currentName = ExtraerValor(p_Valores, COMERCIAL_ALTA_FIELD_NAME)
    currentDesc = ExtraerValor(p_Valores, COMERCIAL_ALTA_FIELD_DESCRIPCION)

    Dim modo As String
    Dim idEntidad As String
    If m_ObjComercialActivo Is Nothing Then
        modo = COMERCIAL_ALTA_MODO_ALTA
        idEntidad = ""
    Else
        modo = COMERCIAL_ALTA_MODO_EDICION
        If m_ObjComercialActivo.IDComercial <> "" Then
            idEntidad = CStr(m_ObjComercialActivo.IDComercial)
        Else
            idEntidad = ""
        End If
    End If

    logs(0) = "Registrar: mode=" & modo & ", id=" & idEntidad

    Dim cm As Comercial
    If m_ObjComercialActivo Is Nothing Then
        Set cm = New Comercial
    Else
        Set cm = m_ObjComercialActivo
    End If

    cm.Comercial = currentName
    cm.DESCRIPCION = currentDesc

    Set m_ObjComercialActivo = cm

    Dim m_ComercialOp As New ComercialOperaciones
    Dim m_AlInicio As Comercial
    If modo = COMERCIAL_ALTA_MODO_EDICION Then
        Dim loadErr As String
        Set m_AlInicio = constructor.getComercial(p_IDComercial:=idEntidad, p_Error:=loadErr)
    Else
        Set m_AlInicio = Nothing
    End If

    With m_ComercialOp
        Set .Comercial = cm
        Dim registrarErr As String
        .Registrar m_AlInicio, registrarErr
        If Len(registrarErr) > 0 Then
            p_Error = "ComercialAlta_Registrar: " & registrarErr
            ComercialAlta_Registrar = BuildJsonPayload(False, Nothing, p_Error, logs)
            Exit Function
        End If
    End With

    Dim resultingId As String
    If cm.IDComercial <> "" Then
        resultingId = CStr(cm.IDComercial)
    Else
        resultingId = idEntidad
    End If

    payload("ok") = True
    payload("id") = resultingId
    payload("modo") = modo

    logs(1) = "Registrar: persisted, resultingId=" & resultingId

    ComercialAlta_Registrar = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "ComercialAlta_Registrar: " & Err.Description
    End If
    ComercialAlta_Registrar = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- ComercialAlta_Cerrar --------------------------------------------------
Public Function ComercialAlta_Cerrar( _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    p_Error = ""

    On Error GoTo errores

    Set m_ObjComercialActivo = Nothing

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")
    payload("ok") = True

    logs(0) = "Cerrar: m_ObjComercialActivo cleared"

    ComercialAlta_Cerrar = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "ComercialAlta_Cerrar: " & Err.Description
    End If
    ComercialAlta_Cerrar = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function
