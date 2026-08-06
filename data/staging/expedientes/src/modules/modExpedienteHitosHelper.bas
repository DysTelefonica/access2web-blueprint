Attribute VB_Name = "modExpedienteHitosHelper"
Option Compare Database
Option Explicit

' modExpedienteHitosHelper — REWORK (2026-06-26)
' Pure-data helpers for Form_FormExpedienteHitos (pestaña de hitos del expediente).
'
' Implements access-vba-e2e-methodology rule #1 (forms are thin UI wiring),
' rule #2 (honest helper signatures), rule #5 (per-module prefix: ExpedienteHitos_*),
' rule #9 (no `ByRef p_Form`).
'
' Shape contract: form wraps m_ObjExpedienteDTOActivo into Dictionary with
' keys: "Expediente", "ColHitos".
'
' Anti-pattern removed:
'   - On Error GoTo errores + Err.Raise 1000 in handlers
'   - VBA.DoEvents / DoCmd.Hourglass blocks inside helper logic
'   - pregunta = MsgBox(...) with undeclared var
'
' vba-access §10.1 declaration ordering: Private consts first, Public atoms last.

' === Module-level constants ===================================================

Private Const EXPEDIENTEHITOS_FIELD_SEP As String = ";"
Private Const EXPEDIENTEHITOS_HEADERS As String = "Fecha;Descripcion;Fecha G.;Importe"
Private Const EXPEDIENTEHITOS_MAX_ROWS As Long = 5000
Private Const VBA_ERROR_OBJECT_REQUIRED As Long = 424
Private Const VBA_ERROR_MEMBER_NOT_SUPPORTED As Long = 438


' === Local helpers ============================================================

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

    Dim okJson As String
    If p_Ok Then
        okJson = "true"
    Else
        okJson = "false"
    End If

    BuildJsonPayload = "{""ok"":" & okJson & _
                       ",""value"":null" & _
                       ",""payload"":" & payloadJson & _
                       ",""error"":" & errorJson & _
                       ",""logs"":" & logsJson & "}"
End Function

Private Function GetDTOSafely(ByVal p_DTO As Object, ByVal p_Key As String) As Object
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String
    Dim errHelpFile As String
    Dim errHelpContext As Long

    On Error GoTo HandleError
    If p_DTO Is Nothing Then Exit Function
    If Not p_DTO.Exists(p_Key) Then Exit Function
    If Not IsObject(p_DTO(p_Key)) Then Exit Function
    Set GetDTOSafely = p_DTO(p_Key)
    Exit Function

HandleError:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    errHelpFile = Err.HelpFile
    errHelpContext = Err.HelpContext

    If errNumber = VBA_ERROR_OBJECT_REQUIRED Then Exit Function
    If errNumber = VBA_ERROR_MEMBER_NOT_SUPPORTED Then Exit Function

    Err.Raise errNumber, errSource, errDescription, errHelpFile, errHelpContext
End Function


' === Public API ================================================================

' --- ExpedienteHitos_AltaHito ---------------------------------------------------
Public Function ExpedienteHitos_AltaHito( _
    ByVal p_DTO As Object, _
    ByVal p_FechaHito As Variant, _
    ByVal p_Descripcion As String, _
    ByVal p_FechaGarantiaHito As Variant, _
    ByVal p_Importe As Variant, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(5)
    p_Error = ""

    On Error GoTo errores

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")

    ' Validate FechaHito is a date.
    If Not IsDate(p_FechaHito) Then
        p_Error = "La fecha del hito es obligatoria"
        ExpedienteHitos_AltaHito = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim expObj As Object
    Set expObj = GetDTOSafely(p_DTO, "Expediente")
    If expObj Is Nothing Then
        p_Error = "AltaHito: p_DTO.Expediente is Nothing"
        ExpedienteHitos_AltaHito = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    ' Build the Hito class instance.
    Dim m_Hito As New ExpedienteHito
    m_Hito.DESCRIPCION = "" & p_Descripcion
    m_Hito.FechaHito = CDate(p_FechaHito)
    m_Hito.FechaGarantiaHito = "" & p_FechaGarantiaHito
    m_Hito.Importe = "" & p_Importe

    Dim m_ExpOp As New ExpedienteOperaciones
    Set m_ExpOp.Expediente = expObj
    Dim estadoRegistro As String
    estadoRegistro = m_ExpOp.RegistrarHito(p_Hito:=m_Hito, p_Error:=p_Error)
    If p_Error <> "" Then
        ExpedienteHitos_AltaHito = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim colRefreshed As Boolean
    colRefreshed = False
    If estadoRegistro = "1" Then
        On Error Resume Next
        Dim colNew As Object
        Set colNew = expObj.Hitos
        Dim refreshErr As Long
        refreshErr = Err.Number
        On Error GoTo errores
        If refreshErr = 0 Then
            If Not colNew Is Nothing Then
                p_DTO("ColHitos") = colNew
                colRefreshed = True
            End If
        End If
        payload("registrado") = True
    Else
        payload("registrado") = False
    End If
    payload("colRefreshed") = colRefreshed

    ExpedienteHitos_AltaHito = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "AltaHito: " & Err.Description
    End If
    ExpedienteHitos_AltaHito = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- ExpedienteHitos_EliminarHito -----------------------------------------------
Public Function ExpedienteHitos_EliminarHito( _
    ByVal p_DTO As Object, _
    ByVal p_FechaHito As String, _
    ByVal p_EsAdmin As Boolean, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(5)
    p_Error = ""

    On Error GoTo errores

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")

    Dim fecha As String
    fecha = Trim$("" & p_FechaHito)

    If Len(fecha) = 0 Then
        payload("eliminado") = False
        ExpedienteHitos_EliminarHito = BuildJsonPayload(True, payload, "", logs)
        Exit Function
    End If

    If Not p_EsAdmin Then
        p_Error = "Operacion no autorizada"
        ExpedienteHitos_EliminarHito = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim expObj As Object
    Set expObj = GetDTOSafely(p_DTO, "Expediente")
    If expObj Is Nothing Then
        p_Error = "EliminarHito: p_DTO.Expediente is Nothing"
        ExpedienteHitos_EliminarHito = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim m_ExpOp As New ExpedienteOperaciones
    Set m_ExpOp.Expediente = expObj
    m_ExpOp.EliminarHito p_FechaHito:=fecha, p_Error:=p_Error
    If p_Error <> "" Then
        ExpedienteHitos_EliminarHito = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim colRefreshed As Boolean
    colRefreshed = False
    On Error Resume Next
    Dim colNew As Object
    Set colNew = expObj.Hitos
    Dim refreshErr As Long
    refreshErr = Err.Number
    On Error GoTo errores
    If refreshErr = 0 Then
        If Not colNew Is Nothing Then
            p_DTO("ColHitos") = colNew
            colRefreshed = True
        End If
    End If

    payload("eliminado") = True
    payload("colRefreshed") = colRefreshed

    ExpedienteHitos_EliminarHito = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "EliminarHito: " & Err.Description
    End If
    ExpedienteHitos_EliminarHito = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- ExpedienteHitos_RellenarListaHitos -----------------------------------------
Public Function ExpedienteHitos_RellenarListaHitos( _
    ByVal p_DTO As Object, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(4)
    p_Error = ""

    On Error GoTo errores

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")

    Dim col As Object
    Set col = GetDTOSafely(p_DTO, "ColHitos")

    Dim rowSource As String
    rowSource = EXPEDIENTEHITOS_HEADERS
    Dim rowCount As Long
    rowCount = 0

    If col Is Nothing Then
        payload("rowSource") = rowSource
        payload("count") = 0
        ExpedienteHitos_RellenarListaHitos = BuildJsonPayload(True, payload, "", logs)
        Exit Function
    End If

    Dim k As Variant
    Dim hito As Object
    Dim fechaH As String
    Dim descH As String
    Dim fechaG As String
    Dim importeStr As String

    For Each k In col.keys
        If rowCount >= EXPEDIENTEHITOS_MAX_ROWS Then Exit For
        Set hito = col(k)
        If hito Is Nothing Then
            ' skip
        Else
            fechaH = "" & hito.FechaHito
            descH = "" & hito.DESCRIPCION
            fechaG = "" & hito.FechaGarantiaHito
            If IsNumeric(hito.Importe) Then
                importeStr = Format(hito.Importe, "#,##0.00 EUR")
            Else
                importeStr = "" & hito.Importe
            End If
            rowSource = rowSource & vbCrLf & _
                fechaH & EXPEDIENTEHITOS_FIELD_SEP & _
                Replace(descH, EXPEDIENTEHITOS_FIELD_SEP, ":") & EXPEDIENTEHITOS_FIELD_SEP & _
                fechaG & EXPEDIENTEHITOS_FIELD_SEP & _
                Replace(importeStr, EXPEDIENTEHITOS_FIELD_SEP, ":")
            rowCount = rowCount + 1
        End If
    Next k

    payload("rowSource") = rowSource
    payload("count") = rowCount
    ExpedienteHitos_RellenarListaHitos = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "RellenarListaHitos: " & Err.Description
    End If
    ExpedienteHitos_RellenarListaHitos = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- ExpedienteHitos_RellenarListas ---------------------------------------------
Public Function ExpedienteHitos_RellenarListas( _
    ByVal p_DTO As Object, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(4)
    p_Error = ""

    On Error GoTo errores

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")

    Dim helperResult As String
    helperResult = ExpedienteHitos_RellenarListaHitos(p_DTO, p_Error)
    If p_Error <> "" Then
        ExpedienteHitos_RellenarListas = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(helperResult)
    If Not parsed Is Nothing Then
        If parsed.Exists("payload") Then
            Set payload = parsed("payload")
        End If
    End If

    ExpedienteHitos_RellenarListas = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "RellenarListas: " & Err.Description
    End If
    ExpedienteHitos_RellenarListas = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- ExpedienteHitos_EstablecerDatos -------------------------------------------
Public Function ExpedienteHitos_EstablecerDatos( _
    ByVal p_DTO As Object, _
    ByVal p_EsAdmin As Boolean, _
    ByVal p_AbiertoParaEditar As Boolean, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(4)
    p_Error = ""

    On Error GoTo errores

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")

    Dim expObj As Object
    Set expObj = GetDTOSafely(p_DTO, "Expediente")
    If expObj Is Nothing Then
        p_Error = "El formulario se ha abierto con parametros insuficientes"
        ExpedienteHitos_EstablecerDatos = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim ejecutivosEnabled As Boolean
    ejecutivosEnabled = False
    If p_AbiertoParaEditar And p_EsAdmin Then ejecutivosEnabled = True

    payload("ejecutivosEnabled") = ejecutivosEnabled
    payload("expedienteOK") = True

    ExpedienteHitos_EstablecerDatos = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "EstablecerDatos: " & Err.Description
    End If
    ExpedienteHitos_EstablecerDatos = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- ExpedienteHitos_Form_Load -------------------------------------------------
Public Function ExpedienteHitos_Form_Load( _
    ByVal p_DTO As Object, _
    ByVal p_EsAdmin As Boolean, _
    ByVal p_AbiertoParaEditar As Boolean, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(4)
    p_Error = ""

    On Error GoTo errores

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")

    Dim expObj As Object
    Set expObj = GetDTOSafely(p_DTO, "Expediente")
    If expObj Is Nothing Then
        p_Error = "El formulario se ha abierto con parametros insuficientes"
        ExpedienteHitos_Form_Load = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim ejecutivoEnabled As Boolean
    ejecutivoEnabled = False
    If p_AbiertoParaEditar And p_EsAdmin Then ejecutivoEnabled = True

    payload("expedienteOK") = True
    payload("ejecutivoEnabled") = ejecutivoEnabled
    payload("hideComandoRegistrar") = True

    ExpedienteHitos_Form_Load = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "Form_Load: " & Err.Description
    End If
    ExpedienteHitos_Form_Load = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

