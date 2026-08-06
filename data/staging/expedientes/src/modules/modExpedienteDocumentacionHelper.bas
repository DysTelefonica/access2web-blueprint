Attribute VB_Name = "modExpedienteDocumentacionHelper"
Option Compare Database
Option Explicit

' modExpedienteDocumentacionHelper — REWORK (2026-06-26)
' Pure-data helpers for Form_FormExpedienteDocumentacion
' (pestaña de documentación / anexos del expediente).
'
' Implements access-vba-e2e-methodology rule #1 (forms are thin UI wiring),
' rule #2 (honest helper signatures — accept only the data the helper needs),
' rule #5 (per-module prefix on Public names: ExpedienteDocumentacion_*),
' and rule #9 (no `ByRef p_Form` — helpers MUST NOT receive Form objects).
'
' Shape contract (Documentacion-specific):
'   The form wraps m_ObjExpedienteDTOActivo into a Scripting.Dictionary with
'   the keys "Expediente" (real Expediente class instance) and "ColAnexos"
'   (Scripting.Dictionary of anexos), then passes that Dictionary as p_DTO.
'   This decouples the helper from the ExpedienteDTO class shape and lets tests
'   pass plain Dictionary stubs.
'
' Anti-pattern removed (was in pre-PR-8 baseline):
'   - form did Application.Echo False/True in ListaDocumentos_DblClick
'   - form had On Error GoTo + CorreoAlAdministrador inline for every handler
'   - RellenarListas and EstablecerDatos did Me.Controls/For-Each inside the .cls
'
' New design (5 helpers, all pure-data, no Form refs):
'   1. ExpedienteDocumentacion_Form_Load(p_DTO, p_Titulo, p_Error)
'      -> {titulo, allowEdits, expedienteOK}
'   2. ExpedienteDocumentacion_EstablecerDatos(p_DTO, p_EsAdmin, p_Error)
'      -> {ejecutivosEnabled, expedienteOK} (form does the Tag="EJECUTIVO" loop)
'   3. ExpedienteDocumentacion_RellenarListas(p_DTO, p_Error)
'      -> {rowSource, count, items} (form assigns to ListaDocumentos.RowSource + lst.AddItem)
'   4. ExpedienteDocumentacion_AltaAnexo(p_DTO, p_URLLocal, p_PromptResult, p_Error)
'      -> {registrado, fileExists, error}
'   5. ExpedienteDocumentacion_EliminarAnexo(p_DTO, p_NombreDocumento, p_EsAdmin,
'                                            p_PromptResult, p_Error)
'      -> {eliminado, autorizado, promptAccepted, error}
'
' UI orchestration that stays in the form (rule #1):
'   - file picker (Seleccionar) — UI input, form calls it first then passes URL
'   - MsgBox "¿Desea realmente eliminar?" — UI prompt, form does it
'   - Me.ListaDocumentos.AddItem / RowSource assignment — UI wiring
'   - For Each ctl In Me.Controls (Tag="EJECUTIVO") — UI wiring
'   - Hourglass / DoEvents / MostrarPopupProgreso — UI guard
'
' Telefonica D&S convention (vba-access §1.4.1): every Public Function ends with
' `Optional ByRef p_Error As String` as the LAST parameter. On failure: set p_Error,
' return fail JSON. On success: return canonical JSON envelope.
'
' vba-access §10.1 declaration ordering: all Private Const / Private Function at top,
' Public Function atoms after. No mid-module consts.
'
' vba-access §1.6.1: split guards (no IIf/And short-circuit on the same object).

' === Module-level constants (all at top per vba-access §10.1) =====================

' Field separator used inside the rowSource CSV (matches Access listbox convention).
Private Const EXPEDIENTEDOC_FIELD_SEP As String = ";"

' Maximum rows emitted in the rowSource string. Defensive cap so a runaway test
' does not build a multi-MB CSV.
Private Const EXPEDIENTEDOC_MAX_ROWS As Long = 5000

' Sentinel value for "form should ask via MsgBox" (UI prompt path). Tests inject
' a non-zero value to skip the prompt.
Private Const EXPEDIENTEDOC_PROMPT_ASK As Long = 0

' Prompt-result values accepted in tests for the confirmation dialog.
' Mirrors vbYes (6) so tests can pass vbYes directly without a translation table.
Private Const EXPEDIENTEDOC_PROMPT_YES As Long = 6
Private Const EXPEDIENTEDOC_PROMPT_NO As Long = 7


' === Local helpers (all at top per vba-access §10.1) =============================

' --- BuildJsonPayload ------------------------------------------------------------
' Wraps a Dictionary payload in the canonical JSON envelope:
'   {"ok":true,"value":null,"payload":<payloadJson>,"error":null,"logs":[...]}
' Logs that are empty strings are stripped (consistent with TestHelper.JsonStringArray).
' Uses JsonConverter.ConvertToJson for safe payload serialization.
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

' --- ExpedienteDocumentacion_HasKeySafely ----------------------------------------
' Returns True if p_DTO is a Dictionary-like object that contains p_Key.
' Split guard: returns False if p_DTO is Nothing, not a Dictionary, or missing key.
Private Function ExpedienteDocumentacion_HasKeySafely( _
    ByVal p_DTO As Object, _
    ByVal p_Key As String _
) As Boolean
    On Error GoTo EH
    If p_DTO Is Nothing Then
        ExpedienteDocumentacion_HasKeySafely = False
        Exit Function
    End If
    ExpedienteDocumentacion_HasKeySafely = p_DTO.Exists(p_Key)
    Exit Function
EH:
    ExpedienteDocumentacion_HasKeySafely = False
End Function

' --- ExpedienteDocumentacion_GetValueSafely --------------------------------------
' Returns p_DTO(p_Key) if p_DTO is a Dictionary-like object with that key.
' Returns the provided default if p_DTO is Nothing, missing key, or accessing
' the key raised an error.
Private Function ExpedienteDocumentacion_GetValueSafely( _
    ByVal p_DTO As Object, _
    ByVal p_Key As String, _
    ByVal p_Default As Variant _
) As Variant
    On Error GoTo EH
    If p_DTO Is Nothing Then
        ExpedienteDocumentacion_GetValueSafely = p_Default
        Exit Function
    End If
    If Not p_DTO.Exists(p_Key) Then
        ExpedienteDocumentacion_GetValueSafely = p_Default
        Exit Function
    End If
    ExpedienteDocumentacion_GetValueSafely = p_DTO(p_Key)
    Exit Function
EH:
    ExpedienteDocumentacion_GetValueSafely = p_Default
End Function

' --- ExpedienteDocumentacion_BuildRowLine ----------------------------------------
' Builds a single rowSource line (semicolon-separated) for the listbox.
' Sanitizes embedded semicolons by replacing them with colons (Access listbox
' RowSource cannot contain field separators inside a value).
Private Function ExpedienteDocumentacion_BuildRowLine( _
    ByVal p_URLDocumento As String, _
    ByVal p_NombreArchivo As String _
) As String
    Dim safeURL As String
    safeURL = Replace(p_URLDocumento, EXPEDIENTEDOC_FIELD_SEP, ":")

    Dim safeNombre As String
    safeNombre = Replace(p_NombreArchivo, EXPEDIENTEDOC_FIELD_SEP, ":")

    ExpedienteDocumentacion_BuildRowLine = safeURL & EXPEDIENTEDOC_FIELD_SEP & safeNombre
End Function

' --- ExpedienteDocumentacion_LeerAnexoComoDic ------------------------------------
' Reads fields from an ExpedienteAnexo-like object that may be either:
'   (a) a real ExpedienteAnexo class instance (has .IDExpediente, .NombreDocumento,
'       .URLDocumento as properties)
'   (b) a Scripting.Dictionary stub with keys "IDExpediente", "NombreDocumento",
'       "URLDocumento"
'
' Returns "" for any field that is missing or raises an error.
Private Function ExpedienteDocumentacion_LeerAnexoComoDic( _
    ByVal p_Anexo As Object, _
    ByVal p_Key As String _
) As String
    On Error GoTo EH

    If p_Anexo Is Nothing Then
        ExpedienteDocumentacion_LeerAnexoComoDic = ""
        Exit Function
    End If

    ' Try the key (Dictionary access) first — covers both real classes that
    ' expose a default member and pure Dictionary stubs.
    If TypeName(p_Anexo) = "Dictionary" Then
        If p_Anexo.Exists(p_Key) Then
            ExpedienteDocumentacion_LeerAnexoComoDic = CStr(p_Anexo(p_Key))
            Exit Function
        End If
    End If

    ' Try direct property access (real ExpedienteAnexo class).
    Dim val As String
    val = ""
    If p_Key = "IDExpediente" Then
        val = p_Anexo.IDExpediente
    ElseIf p_Key = "NombreDocumento" Then
        val = p_Anexo.NombreDocumento
    ElseIf p_Key = "URLDocumento" Then
        val = p_Anexo.URLDocumento
    End If
    ExpedienteDocumentacion_LeerAnexoComoDic = val
    Exit Function
EH:
    ExpedienteDocumentacion_LeerAnexoComoDic = ""
End Function


' === Public API =================================================================

' --- ExpedienteDocumentacion_Form_Load -------------------------------------------
' Pure-data init for Form_Load: validates the DTO wrapper, returns the title
' and allowEdits flags the form should apply to its controls. The form does NOT
' receive a Form ref.
' Returns JSON: {ok, payload:{titulo, allowEdits, expedienteOK}, error, logs}.
Public Function ExpedienteDocumentacion_Form_Load( _
    ByVal p_DTO As Object, _
    ByVal p_Titulo As String, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(4)
    p_Error = ""

    On Error GoTo errores

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")

    Dim expedienteOK As Boolean
    expedienteOK = False

    If p_DTO Is Nothing Then
        p_Error = "El formulario se ha abierto con parámetros insuficientes (DTO Nothing)"
        ExpedienteDocumentacion_Form_Load = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    ' ExpedienteOK iff the wrapper has a non-Nothing "Expediente" entry.
    Dim expObj As Object
    Set expObj = ExpedienteDocumentacion_GetValueSafely(p_DTO, "Expediente", Nothing)
    If Not expObj Is Nothing Then
        expedienteOK = True
    End If

    payload("titulo") = "" & p_Titulo
    payload("allowEdits") = True
    payload("expedienteOK") = expedienteOK

    logs(0) = "Form_Load: titulo='" & p_Titulo & "', expedienteOK=" & CStr(expedienteOK)

    ExpedienteDocumentacion_Form_Load = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "ExpedienteDocumentacion_Form_Load: " & Err.Description
    End If
    ExpedienteDocumentacion_Form_Load = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- ExpedienteDocumentacion_EstablecerDatos -------------------------------------
' Pure-data decision for the EstablecerDatos handler: returns whether the
' EJECUTIVO-tagged controls should be enabled (admin) and whether the DTO has a
' valid Expediente. The form does the actual For-Each-Controls loop.
' Returns JSON: {ok, payload:{ejecutivosEnabled, expedienteOK}, error, logs}.
Public Function ExpedienteDocumentacion_EstablecerDatos( _
    ByVal p_DTO As Object, _
    ByVal p_EsAdmin As Boolean, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(4)
    p_Error = ""

    On Error GoTo errores

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")

    Dim expedienteOK As Boolean
    expedienteOK = False

    If p_DTO Is Nothing Then
        p_Error = "El formulario se ha abierto con parámetros insuficientes (DTO Nothing)"
        ExpedienteDocumentacion_EstablecerDatos = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim expObj As Object
    Set expObj = ExpedienteDocumentacion_GetValueSafely(p_DTO, "Expediente", Nothing)
    If Not expObj Is Nothing Then
        expedienteOK = True
    End If

    payload("ejecutivosEnabled") = p_EsAdmin
    payload("expedienteOK") = expedienteOK

    logs(0) = "EstablecerDatos: esAdmin=" & CStr(p_EsAdmin) & _
              ", expedienteOK=" & CStr(expedienteOK)

    ExpedienteDocumentacion_EstablecerDatos = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "ExpedienteDocumentacion_EstablecerDatos: " & Err.Description
    End If
    ExpedienteDocumentacion_EstablecerDatos = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- ExpedienteDocumentacion_RellenarListas --------------------------------------
' Pure-data listbox payload builder: walks the DTO's ColAnexos collection and
' returns a semicolon-separated rowSource CSV. The form assigns it to
' ListaDocumentos.RowSource (and/or iterates with lst.AddItem).
'
' Returns JSON: {ok, payload:{rowSource, count, items}, error, logs}.
Public Function ExpedienteDocumentacion_RellenarListas( _
    ByVal p_DTO As Object, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(6)
    p_Error = ""

    On Error GoTo errores

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")

    Dim rowSource As String
    rowSource = ""

    Dim items As Object
    Set items = CreateObject("Scripting.Dictionary")

    Dim rowCount As Long
    rowCount = 0

    ' Defensive: empty DTO -> 0 rows, no error.
    If p_DTO Is Nothing Then
        logs(0) = "RellenarListas: p_DTO is Nothing -> 0 rows"
        payload("rowSource") = ""
        payload("count") = 0
        payload("items") = items
        ExpedienteDocumentacion_RellenarListas = BuildJsonPayload(True, payload, "", logs)
        Exit Function
    End If

    Dim colAnexos As Object
    Set colAnexos = ExpedienteDocumentacion_GetValueSafely(p_DTO, "ColAnexos", Nothing)

    If colAnexos Is Nothing Then
        logs(0) = "RellenarListas: p_DTO.ColAnexos is Nothing -> 0 rows"
        payload("rowSource") = ""
        payload("count") = 0
        payload("items") = items
        ExpedienteDocumentacion_RellenarListas = BuildJsonPayload(True, payload, "", logs)
        Exit Function
    End If

    Dim k As Variant
    Dim anexo As Object
    Dim urlDocumento As String
    Dim idExp As String
    Dim nombre As String
    Dim linea As String

    For Each k In colAnexos.Keys
        If rowCount >= EXPEDIENTEDOC_MAX_ROWS Then
            logs(3) = "RellenarListas: hit MAX_ROWS=" & EXPEDIENTEDOC_MAX_ROWS & ", truncating"
            Exit For
        End If

        Set anexo = colAnexos(k)
        If anexo Is Nothing Then
            ' Skip malformed entries — split guard.
        Else
            idExp = ExpedienteDocumentacion_LeerAnexoComoDic(anexo, "IDExpediente")
            nombre = ExpedienteDocumentacion_LeerAnexoComoDic(anexo, "NombreDocumento")

            If idExp = "" Then
                ' Anexo without expediente ID — show just the file name (URL is the name itself).
                urlDocumento = nombre
            Else
                urlDocumento = ExpedienteDocumentacion_LeerAnexoComoDic(anexo, "URLDocumento")
            End If

            linea = ExpedienteDocumentacion_BuildRowLine(urlDocumento, nombre)

            If Len(rowSource) > 0 Then
                rowSource = rowSource & vbCrLf
            End If
            rowSource = rowSource & linea

            items(CStr(rowCount)) = linea
            rowCount = rowCount + 1
        End If
    Next k

    payload("rowSource") = rowSource
    payload("count") = rowCount
    payload("items") = items

    logs(0) = "RellenarListas: emitted " & rowCount & " rows"
    logs(1) = "RellenarListas: rowSource length=" & Len(rowSource)

    ExpedienteDocumentacion_RellenarListas = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "ExpedienteDocumentacion_RellenarListas: " & Err.Description
    End If
    ExpedienteDocumentacion_RellenarListas = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- ExpedienteDocumentacion_AltaAnexo -------------------------------------------
' Pure-data alta path: verifies the local file exists, calls
' ExpedienteOperaciones.RegistrarAnexo, refreshes the DTO's ColAnexos.
' Returns JSON: {ok, payload:{registrado, fileExists, error}, error, logs}.
'
' p_PromptResult is part of the audit signature but unused for this path
' (AltaAnexo has no confirmation prompt — the form calls the file picker FIRST
' and passes the URL here).
Public Function ExpedienteDocumentacion_AltaAnexo( _
    ByVal p_DTO As Object, _
    ByVal p_URLLocal As String, _
    ByVal p_PromptResult As Long, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(8)
    p_Error = ""

    On Error GoTo errores

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")

    Dim fileExists As Boolean
    fileExists = False
    Dim registrado As Boolean
    registrado = False

    ' p_PromptResult is reserved for future prompts on this path; logged for trace.
    logs(0) = "AltaAnexo: urlLocal='" & p_URLLocal & "', prompt=" & CStr(p_PromptResult)

    ' Defensive: empty URL -> no-op (form should have rejected earlier).
    If Len(p_URLLocal) = 0 Then
        payload("registrado") = False
        payload("fileExists") = False
        logs(1) = "AltaAnexo: empty URL -> no-op"
        ExpedienteDocumentacion_AltaAnexo = BuildJsonPayload(True, payload, "", logs)
        Exit Function
    End If

    ' Step 1: verify file exists (fso is the project's global FileSystemObject).
    fileExists = fso.FileExists(p_URLLocal)
    payload("fileExists") = fileExists
    logs(1) = "AltaAnexo: fso.FileExists=" & CStr(fileExists)

    If Not fileExists Then
        ' Form expects to early-exit when fileExists=False. We return ok=true with
        ' the flag so the form can react without needing to check p_Error.
        payload("registrado") = False
        ExpedienteDocumentacion_AltaAnexo = BuildJsonPayload(True, payload, "", logs)
        Exit Function
    End If

    ' Defensive: DTO required for the rest of the path.
    If p_DTO Is Nothing Then
        p_Error = "AltaAnexo: p_DTO is Nothing"
        ExpedienteDocumentacion_AltaAnexo = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim expObj As Object
    Set expObj = ExpedienteDocumentacion_GetValueSafely(p_DTO, "Expediente", Nothing)
    If expObj Is Nothing Then
        p_Error = "AltaAnexo: p_DTO.Expediente is Nothing"
        ExpedienteDocumentacion_AltaAnexo = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    ' Step 2: call ExpedienteOperaciones.RegistrarAnexo.
    Dim expOp As New ExpedienteOperaciones
    Set expOp.Expediente = expObj
    Dim estadoRegistro As String
    estadoRegistro = expOp.RegistrarAnexo(p_URLLocal:=p_URLLocal, p_Error:=p_Error)
    If p_Error <> "" Then
        ExpedienteDocumentacion_AltaAnexo = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    logs(2) = "AltaAnexo: RegistrarAnexo estado='" & estadoRegistro & "'"

    ' Step 3: refresh the DTO's ColAnexos only if state was "1" (success).
    If estadoRegistro = "1" Then
        ' Real Expediente.Anexos is a Property Get; this call may raise in tests.
        ' Use split guards (no `And` short-circuit on anexosRefresh).
        On Error Resume Next
        Dim anexosRefresh As Object
        Set anexosRefresh = expObj.Anexos
        Dim refreshErr As Long
        refreshErr = Err.Number
        On Error GoTo errores

        If refreshErr = 0 Then
            If Not anexosRefresh Is Nothing Then
                p_DTO("ColAnexos") = anexosRefresh
                registrado = True
                logs(3) = "AltaAnexo: refresh OK"
            Else
                logs(3) = "AltaAnexo: anexosRefresh is Nothing (test stub or no Anexos)"
                registrado = True
            End If
        Else
            logs(3) = "AltaAnexo: refresh raised err " & CStr(refreshErr) & " (test stub)"
            registrado = True
        End If
    Else
        logs(3) = "AltaAnexo: estado='" & estadoRegistro & "' (no refresh)"
    End If

    payload("registrado") = registrado
    payload("fileExists") = fileExists

    ExpedienteDocumentacion_AltaAnexo = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "ExpedienteDocumentacion_AltaAnexo: " & Err.Description
    End If
    ExpedienteDocumentacion_AltaAnexo = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- ExpedienteDocumentacion_EliminarAnexo ----------------------------------------
' Pure-data baja path: checks admin auth, processes the confirmation prompt,
' calls ExpedienteOperaciones.EliminarAnexo, refreshes ColAnexos.
'
' p_PromptResult convention (matches modSuministradorHelper.Eliminar_Borrar):
'   - EXPEDIENTEDOC_PROMPT_ASK (=0): form should ask via MsgBox BEFORE calling.
'     The helper returns promptAccepted=False and eliminado=False (no-op).
'   - non-zero: the helper uses the value directly. vbYes (6) -> accept, vbNo (7) -> cancel.
'
' Returns JSON: {ok, payload:{eliminado, autorizado, promptAccepted, error}, error, logs}.
Public Function ExpedienteDocumentacion_EliminarAnexo( _
    ByVal p_DTO As Object, _
    ByVal p_NombreDocumento As String, _
    ByVal p_EsAdmin As Boolean, _
    ByVal p_PromptResult As Long, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(8)
    p_Error = ""

    On Error GoTo errores

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")

    Dim autorizado As Boolean
    autorizado = p_EsAdmin

    Dim promptAccepted As Boolean
    promptAccepted = False

    Dim eliminado As Boolean
    eliminado = False

    logs(0) = "EliminarAnexo: nombre='" & p_NombreDocumento & _
              "', esAdmin=" & CStr(p_EsAdmin) & _
              ", prompt=" & CStr(p_PromptResult)

    ' Authorization gate first: non-admin -> no-op.
    If Not autorizado Then
        payload("eliminado") = False
        payload("autorizado") = False
        payload("promptAccepted") = False
        p_Error = "Operacion no autorizada"
        logs(1) = "EliminarAnexo: no autorizado -> rejected"
        ExpedienteDocumentacion_EliminarAnexo = BuildJsonPayload(False, payload, p_Error, logs)
        Exit Function
    End If

    ' Prompt gate: if form has not answered yet, return promptAccepted=False so the
    ' form knows to show MsgBox and re-call with the answer.
    If p_PromptResult = EXPEDIENTEDOC_PROMPT_ASK Then
        payload("eliminado") = False
        payload("autorizado") = True
        payload("promptAccepted") = False
        logs(1) = "EliminarAnexo: prompt pending (form should ask via MsgBox)"
        ExpedienteDocumentacion_EliminarAnexo = BuildJsonPayload(True, payload, "", logs)
        Exit Function
    End If

    ' Process the injected answer.
    If p_PromptResult = EXPEDIENTEDOC_PROMPT_YES Then
        promptAccepted = True
    ElseIf p_PromptResult = EXPEDIENTEDOC_PROMPT_NO Then
        promptAccepted = False
    Else
        ' Unknown prompt result — treat as cancel for safety.
        promptAccepted = False
        logs(2) = "EliminarAnexo: unknown prompt result " & CStr(p_PromptResult) & " -> cancel"
    End If

    payload("autorizado") = True
    payload("promptAccepted") = promptAccepted

    If Not promptAccepted Then
        payload("eliminado") = False
        logs(2) = "EliminarAnexo: prompt rejected -> no-op"
        ExpedienteDocumentacion_EliminarAnexo = BuildJsonPayload(True, payload, "", logs)
        Exit Function
    End If

    ' Defensive: DTO and entity required for the DAO call.
    If p_DTO Is Nothing Then
        p_Error = "EliminarAnexo: p_DTO is Nothing"
        ExpedienteDocumentacion_EliminarAnexo = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim expObj As Object
    Set expObj = ExpedienteDocumentacion_GetValueSafely(p_DTO, "Expediente", Nothing)
    If expObj Is Nothing Then
        p_Error = "EliminarAnexo: p_DTO.Expediente is Nothing"
        ExpedienteDocumentacion_EliminarAnexo = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    ' Defensive: empty nombre -> no-op.
    If Len(p_NombreDocumento) = 0 Then
        payload("eliminado") = False
        logs(2) = "EliminarAnexo: empty nombre -> no-op"
        ExpedienteDocumentacion_EliminarAnexo = BuildJsonPayload(True, payload, "", logs)
        Exit Function
    End If

    ' Call ExpedienteOperaciones.EliminarAnexo.
    Dim expOp As New ExpedienteOperaciones
    Set expOp.Expediente = expObj
    expOp.EliminarAnexo p_NombreDocumento:=p_NombreDocumento, p_Error:=p_Error
    If p_Error <> "" Then
        ExpedienteDocumentacion_EliminarAnexo = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    ' Refresh ColAnexos (best-effort; tolerant of test stubs).
    On Error Resume Next
    Dim anexosRefresh As Object
    Set anexosRefresh = expObj.Anexos
    Dim refreshErr As Long
    refreshErr = Err.Number
    On Error GoTo errores

    ' Split guards (no `And` short-circuit on anexosRefresh).
    If refreshErr = 0 Then
        If Not anexosRefresh Is Nothing Then
            p_DTO("ColAnexos") = anexosRefresh
        End If
    End If

    eliminado = True
    logs(3) = "EliminarAnexo: EliminarAnexo OK"

    payload("eliminado") = eliminado
    ExpedienteDocumentacion_EliminarAnexo = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "ExpedienteDocumentacion_EliminarAnexo: " & Err.Description
    End If
    ExpedienteDocumentacion_EliminarAnexo = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function