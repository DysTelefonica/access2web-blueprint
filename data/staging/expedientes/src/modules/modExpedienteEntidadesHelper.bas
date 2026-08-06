Attribute VB_Name = "modExpedienteEntidadesHelper"
Option Compare Database
Option Explicit

' modExpedienteEntidadesHelper — REWORK (2026-06-26)
' Pure-data helpers for Form_FormExpedienteEntidades
' (pestaña de entidades del expediente: Comerciales / CPVs / Lugares /
'  PECAL / RACs / Responsables / Anualidades).
'
' Implements access-vba-e2e-methodology rule #1 (forms are thin UI wiring),
' rule #2 (honest helper signatures — accept only the data the helper needs),
' rule #5 (per-module prefix on Public names: ExpedienteEntidades_*),
' and rule #9 (no `ByRef p_Form` — helpers MUST NOT receive Form objects).
'
' Shape contract:
'   The form wraps m_ObjExpedienteDTOActivo into a Scripting.Dictionary
'   with these keys: "Expediente", "ColComerciales", "ColCPVs",
'   "ColLugaresEjecucion", "ColPECALES", "ColRACs", "ColResponsables",
'   "ColAnualidades". The helper accepts p_DTO As Object (Dictionary).
'
' Anti-pattern removed (was in pre-PR-9 baseline):
'   - every handler had On Error GoTo errores + Err.Raise 1000 (control flow)
'   - every handler had VBA.DoEvents / DoCmd.Hourglass blocks inline
'   - `pregunta = MsgBox(...)` with undeclared `pregunta` (does NOT compile under Option Explicit)
'   - helpers/handlers reached into m_ObjExpedienteDTOActivo global directly
'
' New design (~25 helpers, all pure-data, no Form refs):
'   Alta (5):  AltaComercial/CPV/Lugar/PECAL/RAC + AltaAnualidad + AltaResponsable
'   Eliminar (7): EliminarComercial/CPV/Lugar/PECAL/RAC + EliminarAnualidad + EliminarResponsable
'   Mutaciones responsables (2): CambiarJP / CambiarEnvio
'   Rellenar (8): RellenarListas + 7 per-listbox
'   Form-level (2): Form_Load / EstablecerDatos
'   Misc (1): CambiarLineaListaResponsables
'
' UI orchestration that stays in the form (rule #1):
'   - 5 ComandoElegirX_Click (close + open form hijo + set m_FormX WithEvents)
'   - 6 m_FormX_Seleccionar WithEvents subscribers (delegate to Alta helper)
'   - 7 ComandoAltaX_Click + 7 ComandoEliminarX_Click (delegate + UI guard)
'   - 2 ComandoCopiarX_Click (CopiarAlPortapapeles + MsgBox confirmation)
'   - 2 ComandoLimpiarX_Click (Me.X = Null — pure UI)
'   - ListaResponsables_Click (enable/disable buttons)
'   - Form_Load (orchestrator)
'
' Telefonica D&S convention (vba-access §1.4.1): every Public Function ends with
' `Optional ByRef p_Error As String` as the LAST parameter.
'
' vba-access §10.1 declaration ordering: all Private Const / Private Function at top,
' Public Function atoms after. No mid-module consts.
'
' vba-access §1.6.1: split guards (no IIf/And short-circuit on the same object).

' === Module-level constants (all at top per vba-access §10.1) =====================

' Field separator inside rowSource CSVs.
Private Const EXPEDIENTEENT_FIELD_SEP As String = ";"

' Headers for the 4-col Responsables listbox.
Private Const EXPEDIENTEENT_HEADERS_RESPONSABLES As String = "ID;Nombre;JP;Aviso"

' Headers for the 2-col Anualidades listbox.
Private Const EXPEDIENTEENT_HEADERS_ANUALIDADES As String = "Año;Presupuesto"

' Maximum rows emitted per rowSource.
Private Const EXPEDIENTEENT_MAX_ROWS As Long = 5000


' === Local helpers (all at top per vba-access §10.1) =============================

' --- BuildJsonPayload ------------------------------------------------------------
' Wraps a Dictionary payload in the canonical JSON envelope.
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

' --- SafeDictMember --------------------------------------------------------------
' Like GetDTOSafely but returns Object (not Variant) so the caller can `Set`
' the result without triggering Let-assignment fall-through to a default
' property. Used by ExpedienteEntidades_CargarEntidadesTabla where each
' collection must come out as a real Dictionary reference for downstream
' iteration. Returns Nothing when p_DTO is missing the key or p_DTO itself
' is Nothing.
Private Function SafeDictMember( _
    ByVal p_DTO As Object, _
    ByVal p_Key As String _
) As Object
    On Error GoTo EH
    If p_DTO Is Nothing Then Exit Function
    If Not p_DTO.Exists(p_Key) Then Exit Function
    Set SafeDictMember = p_DTO(p_Key)
    Exit Function
EH:
    Set SafeDictMember = Nothing
End Function

' --- GetDTOSafely ----------------------------------------------------------------
' Returns p_DTO(p_Key) safely. Returns p_Default if p_DTO is Nothing,
' missing key, or accessing the key raised an error.
Private Function GetDTOSafely( _
    ByVal p_DTO As Object, _
    ByVal p_Key As String, _
    ByVal p_Default As Variant _
) As Variant
    On Error GoTo EH
    If p_DTO Is Nothing Then
        GetDTOSafely = p_Default
        Exit Function
    End If
    If Not p_DTO.Exists(p_Key) Then
        GetDTOSafely = p_Default
        Exit Function
    End If
    GetDTOSafely = p_DTO(p_Key)
    Exit Function
EH:
    GetDTOSafely = p_Default
End Function

Private Function GetMemberSafely( _
    ByVal p_Source As Object, _
    ByVal p_Name As String _
) As Variant
    On Error GoTo EH
    If p_Source Is Nothing Then Exit Function
    If TypeName(p_Source) = "Dictionary" Then
        If p_Source.Exists(p_Name) Then
            GetMemberSafely = p_Source(p_Name)
            Exit Function
        End If
    End If
    GetMemberSafely = CallByName(p_Source, p_Name, VbGet)
    Exit Function
EH:
    GetMemberSafely = Empty
End Function

Private Function BuildEntidadTablaItem( _
    ByVal p_IDEntidad As String, _
    ByVal p_Nombre As String, _
    ByVal p_Tipo As String _
) As Object
    Dim item As Object
    Set item = CreateObject("Scripting.Dictionary")
    item("IDEntidad") = Trim$(p_IDEntidad)
    item("Nombre") = p_Nombre
    item("Tipo") = p_Tipo
    Set BuildEntidadTablaItem = item
End Function

Private Sub AddEntidadTablaItem( _
    ByVal p_Entidades As Object, _
    ByVal p_IDEntidad As String, _
    ByVal p_Nombre As String, _
    ByVal p_Tipo As String _
)
    Dim idEntidad As String
    idEntidad = Trim$(p_IDEntidad)
    If Len(idEntidad) = 0 Then Exit Sub
    If Not p_Entidades.Exists(idEntidad) Then
        p_Entidades.Add idEntidad, BuildEntidadTablaItem(idEntidad, p_Nombre, p_Tipo)
    End If
End Sub

Private Sub AddColeccionEntidadesTabla( _
    ByVal p_Entidades As Object, _
    ByVal p_Col As Object, _
    ByVal p_Tipo As String, _
    ByVal p_IDProp As String, _
    ByVal p_NombreProp As String _
)
    If p_Col Is Nothing Then Exit Sub

    Dim key As Variant
    Dim item As Object
    For Each key In p_Col.Keys
        Set item = p_Col(key)
        If Not item Is Nothing Then
            AddEntidadTablaItem p_Entidades, _
                CStr(GetMemberSafely(item, p_IDProp)), _
                CStr(GetMemberSafely(item, p_NombreProp)), _
                p_Tipo
        End If
    Next key
End Sub

Private Sub AddResponsablesEntidadesTabla( _
    ByVal p_Entidades As Object, _
    ByVal p_Col As Object _
)
    If p_Col Is Nothing Then Exit Sub

    Dim key As Variant
    Dim item As Object
    Dim idUsuario As String
    Dim nombre As String
    Dim usuario As Object
    For Each key In p_Col.Keys
        Set item = p_Col(key)
        If Not item Is Nothing Then
            idUsuario = CStr(GetMemberSafely(item, "IdUsuario"))
            nombre = CStr(GetMemberSafely(item, "Nombre"))
            If Len(nombre) = 0 Then
                On Error Resume Next
                Set usuario = CallByName(item, "USUARIO", VbGet)
                If Err.Number = 0 Then nombre = CStr(GetMemberSafely(usuario, "Nombre"))
                Set usuario = Nothing
                Err.Clear
                On Error GoTo 0
            End If
            AddEntidadTablaItem p_Entidades, idUsuario, nombre, "Responsable"
        End If
    Next key
End Sub

' --- BuildRowLine2 ----------------------------------------------------------------
' 2-col rowSource line: "ID;Value", with sanitization of embedded semicolons.
Private Function BuildRowLine2( _
    ByVal p_ID As String, _
    ByVal p_Value As String _
) As String
    Dim safeValue As String
    safeValue = Replace(p_Value, EXPEDIENTEENT_FIELD_SEP, ":")
    BuildRowLine2 = p_ID & EXPEDIENTEENT_FIELD_SEP & safeValue
End Function

' --- BuildRowLine4 ----------------------------------------------------------------
' 4-col rowSource line: "ID;Nombre;JP;Aviso".
Private Function BuildRowLine4( _
    ByVal p_ID As String, _
    ByVal p_Nombre As String, _
    ByVal p_JP As String, _
    ByVal p_Aviso As String _
) As String
    BuildRowLine4 = p_ID & EXPEDIENTEENT_FIELD_SEP & _
                    Replace(p_Nombre, EXPEDIENTEENT_FIELD_SEP, ":") & EXPEDIENTEENT_FIELD_SEP & _
                    Replace(p_JP, EXPEDIENTEENT_FIELD_SEP, ":") & EXPEDIENTEENT_FIELD_SEP & _
                    Replace(p_Aviso, EXPEDIENTEENT_FIELD_SEP, ":")
End Function


' === Public API =================================================================

Public Function ExpedienteEntidades_CargarEntidadesTabla( _
    ByVal p_Tabla As Helper_ExpedienteEntidadesTabla, _
    ByVal p_DTO As Object, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    Dim stage As String
    logs = TestingCore_InitLogs(4)
    p_Error = ""
    logs(0) = "start"

    On Error GoTo errores

    stage = "validate-input"
    logs(1) = stage
    If p_Tabla Is Nothing Then
        p_Error = "CargarEntidadesTabla: " & stage & ": p_Tabla is Nothing"
        ExpedienteEntidades_CargarEntidadesTabla = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    stage = "resolve-exp-start"
    logs(2) = stage
    Dim expObj As Object
    stage = "resolve-exp-call"
    If p_DTO Is Nothing Then
        p_Error = "CargarEntidadesTabla: " & stage & ": p_DTO is Nothing"
        ExpedienteEntidades_CargarEntidadesTabla = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If
    Set expObj = p_DTO("Expediente")
    stage = "resolve-exp-check"
    If expObj Is Nothing Then
        p_Error = "CargarEntidadesTabla: " & stage & ": p_DTO.Expediente is Nothing"
        ExpedienteEntidades_CargarEntidadesTabla = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim idExpediente As String
    idExpediente = CStr(expObj("IDExpediente"))
    If Len(Trim$(idExpediente)) = 0 Or Not IsNumeric(idExpediente) Then
        p_Error = "CargarEntidadesTabla: " & stage & ": IDExpediente is required"
        ExpedienteEntidades_CargarEntidadesTabla = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    stage = "build-collections"
    logs(3) = stage
    Dim entidades As Object
    Set entidades = CreateObject("Scripting.Dictionary")
    entidades.CompareMode = TextCompare

    ' Resolve each collection directly from p_DTO via Set — do NOT route
    ' through GetDTOSafely here because its Variant return cannot reliably
    ' carry an Object reference out via Let assignment (error 91 surfaces
    ' either at the Set site or at the consumer). The shape contract is
    ' already validated above (p_DTO Is Nothing + p_DTO("Expediente")), so
    ' we can safely dereference p_DTO here.
    Dim colComerciales As Object
    Dim colCPVs As Object
    Dim colLugares As Object
    Dim colPECALES As Object
    Dim colRACs As Object
    Dim colResponsables As Object
    Set colComerciales = SafeDictMember(p_DTO, "ColComerciales")
    Set colCPVs = SafeDictMember(p_DTO, "ColCPVs")
    Set colLugares = SafeDictMember(p_DTO, "ColLugaresEjecucion")
    Set colPECALES = SafeDictMember(p_DTO, "ColPECALES")
    Set colRACs = SafeDictMember(p_DTO, "ColRACs")
    Set colResponsables = SafeDictMember(p_DTO, "ColResponsables")

    stage = "populate-comerciales"
    AddColeccionEntidadesTabla entidades, colComerciales, "Comercial", "IDComercial", "Comercial"
    stage = "populate-cpvs"
    AddColeccionEntidadesTabla entidades, colCPVs, "CPV", "IDCPV", "CPV"
    stage = "populate-lugares"
    AddColeccionEntidadesTabla entidades, colLugares, "Lugar", "IDLugarEjecucion", "LugarEjecucion"
    stage = "populate-pecals"
    AddColeccionEntidadesTabla entidades, colPECALES, "PECAL", "IDPECAL", "PECAL"
    stage = "populate-racs"
    AddColeccionEntidadesTabla entidades, colRACs, "RAC", "IDRAC", "RAC"
    stage = "populate-responsables"
    AddResponsablesEntidadesTabla entidades, colResponsables

    stage = "init-and-load"
    logs(0) = stage
    p_Tabla.Init CLng(idExpediente), p_Error
    If p_Error <> "" Then Err.Raise 1000

    stage = "cargar-collection"
    p_Tabla.Cargar entidades, p_Error
    If p_Error <> "" Then Err.Raise 1000

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")
    payload("IDExpediente") = CLng(idExpediente)
        payload("count") = entidades.Count
        ExpedienteEntidades_CargarEntidadesTabla = BuildJsonPayload(True, payload, "", logs)
        Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "CargarEntidadesTabla: " & stage & ": line=" & Erl & ": " & Err.Description
    End If
    ExpedienteEntidades_CargarEntidadesTabla = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- ExpedienteEntidades_AltaComercial --------------------------------------------
' Pure-data alta Comercial: receives p_DTO + p_IDComercial, calls
' constructor.getComercial + ExpedienteOperaciones.RegistrarComercial, refreshes
' ColComerciales if state == "1". Returns JSON: {ok, payload:{registrado,colRefreshed},...}
Public Function ExpedienteEntidades_AltaComercial( _
    ByVal p_DTO As Object, _
    ByVal p_IDComercial As String, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(6)
    p_Error = ""

    On Error GoTo errores

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")

    Dim idCom As String
    idCom = Trim$("" & p_IDComercial)

    ' Empty ID is a no-op (form should have rejected earlier).
    If Len(idCom) = 0 Then
        payload("registrado") = False
        payload("colRefreshed") = False
        logs(0) = "AltaComercial: empty id -> no-op"
        ExpedienteEntidades_AltaComercial = BuildJsonPayload(True, payload, "", logs)
        Exit Function
    End If

    ' Defensive: DTO + Expediente required.
    Dim expObj As Object
    Set expObj = GetDTOSafely(p_DTO, "Expediente", Nothing)
    If expObj Is Nothing Then
        p_Error = "AltaComercial: p_DTO.Expediente is Nothing"
        ExpedienteEntidades_AltaComercial = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    ' Lookup comercial via constructor.
    Dim com As Object
    Set com = constructor.getComercial(p_IDComercial:=idCom, p_Error:=p_Error)
    If p_Error <> "" Then
        ExpedienteEntidades_AltaComercial = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If
    If com Is Nothing Then
        p_Error = "No aparece ese comercial en el repositorio"
        ExpedienteEntidades_AltaComercial = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    ' Call DAO: ExpedienteOperaciones.RegistrarComercial.
    Dim m_ExpOp As New ExpedienteOperaciones
    Set m_ExpOp.Expediente = expObj
    Dim estadoRegistro As String
    estadoRegistro = m_ExpOp.RegistrarComercial(p_IDComercial:=com.IDComercial, p_Error:=p_Error)
    If p_Error <> "" Then
        ExpedienteEntidades_AltaComercial = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    logs(1) = "AltaComercial: estadoRegistro='" & estadoRegistro & "'"

    Dim colRefreshed As Boolean
    colRefreshed = False
    If estadoRegistro = "1" Then
        ' Refresh p_DTO("ColComerciales") from expObj.Comerciales (best-effort).
        On Error Resume Next
        Dim colNew As Object
        Set colNew = expObj.Comerciales
        Dim refreshErr As Long
        refreshErr = Err.Number
        On Error GoTo errores
        If refreshErr = 0 Then
            If Not colNew Is Nothing Then
                p_DTO("ColComerciales") = colNew
                colRefreshed = True
            End If
        End If
        payload("registrado") = True
    Else
        payload("registrado") = False
    End If
    payload("colRefreshed") = colRefreshed

    ExpedienteEntidades_AltaComercial = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "AltaComercial: " & Err.Description
    End If
    ExpedienteEntidades_AltaComercial = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- ExpedienteEntidades_EliminarComercial ----------------------------------------
' Pure-data baja Comercial.
Public Function ExpedienteEntidades_EliminarComercial( _
    ByVal p_DTO As Object, _
    ByVal p_IDComercial As String, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(6)
    p_Error = ""

    On Error GoTo errores

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")

    Dim idCom As String
    idCom = Trim$("" & p_IDComercial)

    If Len(idCom) = 0 Then
        payload("eliminado") = False
        ExpedienteEntidades_EliminarComercial = BuildJsonPayload(True, payload, "", logs)
        Exit Function
    End If

    Dim expObj As Object
    Set expObj = GetDTOSafely(p_DTO, "Expediente", Nothing)
    If expObj Is Nothing Then
        p_Error = "EliminarComercial: p_DTO.Expediente is Nothing"
        ExpedienteEntidades_EliminarComercial = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim com As Object
    Set com = constructor.getComercial(p_IDComercial:=idCom, p_Error:=p_Error)
    If p_Error <> "" Then
        ExpedienteEntidades_EliminarComercial = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If
    If com Is Nothing Then
        p_Error = "Seleccione un comercial de la lista"
        ExpedienteEntidades_EliminarComercial = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim m_ExpOp As New ExpedienteOperaciones
    Set m_ExpOp.Expediente = expObj
    m_ExpOp.EliminarComercial p_IDComercial:=com.IDComercial, p_Error:=p_Error
    If p_Error <> "" Then
        ExpedienteEntidades_EliminarComercial = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim colRefreshed As Boolean
    colRefreshed = False
    On Error Resume Next
    Dim colNew As Object
    Set colNew = expObj.Comerciales
    Dim refreshErr As Long
    refreshErr = Err.Number
    On Error GoTo errores
    If refreshErr = 0 Then
        If Not colNew Is Nothing Then
            p_DTO("ColComerciales") = colNew
            colRefreshed = True
        End If
    End If

    payload("eliminado") = True
    payload("colRefreshed") = colRefreshed

    ExpedienteEntidades_EliminarComercial = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "EliminarComercial: " & Err.Description
    End If
    ExpedienteEntidades_EliminarComercial = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- ExpedienteEntidades_AltaCPV --------------------------------------------------
Public Function ExpedienteEntidades_AltaCPV( _
    ByVal p_DTO As Object, _
    ByVal p_IDCPV As String, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(6)
    p_Error = ""

    On Error GoTo errores

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")

    Dim idCPV As String
    idCPV = Trim$("" & p_IDCPV)

    If Len(idCPV) = 0 Then
        payload("registrado") = False
        ExpedienteEntidades_AltaCPV = BuildJsonPayload(True, payload, "", logs)
        Exit Function
    End If

    Dim expObj As Object
    Set expObj = GetDTOSafely(p_DTO, "Expediente", Nothing)
    If expObj Is Nothing Then
        p_Error = "AltaCPV: p_DTO.Expediente is Nothing"
        ExpedienteEntidades_AltaCPV = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim cpv As Object
    Set cpv = constructor.getCPV(p_IDCPV:=idCPV, p_Error:=p_Error)
    If p_Error <> "" Then
        ExpedienteEntidades_AltaCPV = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If
    If cpv Is Nothing Then
        p_Error = "Se ha elegido un CPV no registrado previamente"
        ExpedienteEntidades_AltaCPV = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim m_ExpOp As New ExpedienteOperaciones
    Set m_ExpOp.Expediente = expObj
    Dim estadoRegistro As String
    estadoRegistro = m_ExpOp.RegistrarCPV(p_IDCPV:=cpv.IDCPV, p_Error:=p_Error)
    If p_Error <> "" Then
        ExpedienteEntidades_AltaCPV = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim colRefreshed As Boolean
    colRefreshed = False
    If estadoRegistro = "1" Then
        On Error Resume Next
        Dim colNew As Object
        Set colNew = expObj.CPVs
        Dim refreshErr As Long
        refreshErr = Err.Number
        On Error GoTo errores
        If refreshErr = 0 Then
            If Not colNew Is Nothing Then
                p_DTO("ColCPVs") = colNew
                colRefreshed = True
            End If
        End If
        payload("registrado") = True
    Else
        payload("registrado") = False
    End If
    payload("colRefreshed") = colRefreshed

    ExpedienteEntidades_AltaCPV = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "AltaCPV: " & Err.Description
    End If
    ExpedienteEntidades_AltaCPV = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- ExpedienteEntidades_EliminarCPV ----------------------------------------------
Public Function ExpedienteEntidades_EliminarCPV( _
    ByVal p_DTO As Object, _
    ByVal p_IDCPV As String, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(6)
    p_Error = ""

    On Error GoTo errores

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")

    Dim idCPV As String
    idCPV = Trim$("" & p_IDCPV)

    If Len(idCPV) = 0 Then
        payload("eliminado") = False
        ExpedienteEntidades_EliminarCPV = BuildJsonPayload(True, payload, "", logs)
        Exit Function
    End If

    Dim expObj As Object
    Set expObj = GetDTOSafely(p_DTO, "Expediente", Nothing)
    If expObj Is Nothing Then
        p_Error = "EliminarCPV: p_DTO.Expediente is Nothing"
        ExpedienteEntidades_EliminarCPV = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim cpv As Object
    Set cpv = constructor.getCPV(p_IDCPV:=idCPV, p_Error:=p_Error)
    If p_Error <> "" Then
        ExpedienteEntidades_EliminarCPV = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If
    If cpv Is Nothing Then
        p_Error = "Seleccione un CPV de la lista"
        ExpedienteEntidades_EliminarCPV = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim m_ExpOp As New ExpedienteOperaciones
    Set m_ExpOp.Expediente = expObj
    m_ExpOp.EliminarCPV p_IDCPV:=cpv.IDCPV, p_Error:=p_Error
    If p_Error <> "" Then
        ExpedienteEntidades_EliminarCPV = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim colRefreshed As Boolean
    colRefreshed = False
    On Error Resume Next
    Dim colNew As Object
    Set colNew = expObj.CPVs
    Dim refreshErr As Long
    refreshErr = Err.Number
    On Error GoTo errores
    If refreshErr = 0 Then
        If Not colNew Is Nothing Then
            p_DTO("ColCPVs") = colNew
            colRefreshed = True
        End If
    End If

    payload("eliminado") = True
    payload("colRefreshed") = colRefreshed

    ExpedienteEntidades_EliminarCPV = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "EliminarCPV: " & Err.Description
    End If
    ExpedienteEntidades_EliminarCPV = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- ExpedienteEntidades_AltaLugar -----------------------------------------------
Public Function ExpedienteEntidades_AltaLugar( _
    ByVal p_DTO As Object, _
    ByVal p_IDLugar As String, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(6)
    p_Error = ""

    On Error GoTo errores

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")

    Dim idL As String
    idL = Trim$("" & p_IDLugar)

    If Len(idL) = 0 Then
        payload("registrado") = False
        ExpedienteEntidades_AltaLugar = BuildJsonPayload(True, payload, "", logs)
        Exit Function
    End If

    Dim expObj As Object
    Set expObj = GetDTOSafely(p_DTO, "Expediente", Nothing)
    If expObj Is Nothing Then
        p_Error = "AltaLugar: p_DTO.Expediente is Nothing"
        ExpedienteEntidades_AltaLugar = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim lugar As Object
    Set lugar = constructor.getLugarEjecucion(p_IDLugarEjecucion:=idL, p_Error:=p_Error)
    If p_Error <> "" Then
        ExpedienteEntidades_AltaLugar = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If
    If lugar Is Nothing Then
        p_Error = "Se ha elegido un Lugar de Ejecucion que parece no estar registrado previamente"
        ExpedienteEntidades_AltaLugar = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim m_ExpOp As New ExpedienteOperaciones
    Set m_ExpOp.Expediente = expObj
    Dim estadoRegistro As String
    estadoRegistro = m_ExpOp.RegistrarLugarEjecucion(p_IDLugar:=lugar.IDLugarEjecucion, p_Error:=p_Error)
    If p_Error <> "" Then
        ExpedienteEntidades_AltaLugar = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim colRefreshed As Boolean
    colRefreshed = False
    If estadoRegistro = "1" Then
        On Error Resume Next
        Dim colNew As Object
        Set colNew = expObj.LugaresEjecucion
        Dim refreshErr As Long
        refreshErr = Err.Number
        On Error GoTo errores
        If refreshErr = 0 Then
            If Not colNew Is Nothing Then
                p_DTO("ColLugaresEjecucion") = colNew
                colRefreshed = True
            End If
        End If
        payload("registrado") = True
    Else
        payload("registrado") = False
    End If
    payload("colRefreshed") = colRefreshed

    ExpedienteEntidades_AltaLugar = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "AltaLugar: " & Err.Description
    End If
    ExpedienteEntidades_AltaLugar = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- ExpedienteEntidades_EliminarLugar -------------------------------------------
Public Function ExpedienteEntidades_EliminarLugar( _
    ByVal p_DTO As Object, _
    ByVal p_IDLugar As String, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(6)
    p_Error = ""

    On Error GoTo errores

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")

    Dim idL As String
    idL = Trim$("" & p_IDLugar)

    If Len(idL) = 0 Then
        payload("eliminado") = False
        ExpedienteEntidades_EliminarLugar = BuildJsonPayload(True, payload, "", logs)
        Exit Function
    End If

    Dim expObj As Object
    Set expObj = GetDTOSafely(p_DTO, "Expediente", Nothing)
    If expObj Is Nothing Then
        p_Error = "EliminarLugar: p_DTO.Expediente is Nothing"
        ExpedienteEntidades_EliminarLugar = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim lugar As Object
    Set lugar = constructor.getLugarEjecucion(p_IDLugarEjecucion:=idL, p_Error:=p_Error)
    If p_Error <> "" Then
        ExpedienteEntidades_EliminarLugar = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If
    If lugar Is Nothing Then
        p_Error = "Seleccione un Lugar de Ejecucion de la lista"
        ExpedienteEntidades_EliminarLugar = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim m_ExpOp As New ExpedienteOperaciones
    Set m_ExpOp.Expediente = expObj
    m_ExpOp.EliminarLugarEjecucion p_IDLugar:=lugar.IDLugarEjecucion, p_Error:=p_Error
    If p_Error <> "" Then
        ExpedienteEntidades_EliminarLugar = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim colRefreshed As Boolean
    colRefreshed = False
    On Error Resume Next
    Dim colNew As Object
    Set colNew = expObj.LugaresEjecucion
    Dim refreshErr As Long
    refreshErr = Err.Number
    On Error GoTo errores
    If refreshErr = 0 Then
        If Not colNew Is Nothing Then
            p_DTO("ColLugaresEjecucion") = colNew
            colRefreshed = True
        End If
    End If

    payload("eliminado") = True
    payload("colRefreshed") = colRefreshed

    ExpedienteEntidades_EliminarLugar = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "EliminarLugar: " & Err.Description
    End If
    ExpedienteEntidades_EliminarLugar = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- ExpedienteEntidades_AltaPECAL -----------------------------------------------
Public Function ExpedienteEntidades_AltaPECAL( _
    ByVal p_DTO As Object, _
    ByVal p_IDPECAL As String, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(6)
    p_Error = ""

    On Error GoTo errores

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")

    Dim idP As String
    idP = Trim$("" & p_IDPECAL)

    If Len(idP) = 0 Then
        payload("registrado") = False
        ExpedienteEntidades_AltaPECAL = BuildJsonPayload(True, payload, "", logs)
        Exit Function
    End If

    Dim expObj As Object
    Set expObj = GetDTOSafely(p_DTO, "Expediente", Nothing)
    If expObj Is Nothing Then
        p_Error = "AltaPECAL: p_DTO.Expediente is Nothing"
        ExpedienteEntidades_AltaPECAL = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim pecal As Object
    Set pecal = constructor.getPecal(p_IDPEcal:=idP, p_Error:=p_Error)
    If p_Error <> "" Then
        ExpedienteEntidades_AltaPECAL = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If
    If pecal Is Nothing Then
        p_Error = "Se ha elegido una PECAL que parece no estar registrada previamente"
        ExpedienteEntidades_AltaPECAL = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim m_ExpOp As New ExpedienteOperaciones
    Set m_ExpOp.Expediente = expObj
    Dim estadoRegistro As String
    estadoRegistro = m_ExpOp.RegistrarPECAL(p_IDPEcal:=pecal.IDPECAL, p_Error:=p_Error)
    If p_Error <> "" Then
        ExpedienteEntidades_AltaPECAL = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim colRefreshed As Boolean
    colRefreshed = False
    If estadoRegistro = "1" Then
        On Error Resume Next
        Dim colNew As Object
        Set colNew = expObj.PECALES
        Dim refreshErr As Long
        refreshErr = Err.Number
        On Error GoTo errores
        If refreshErr = 0 Then
            If Not colNew Is Nothing Then
                p_DTO("ColPECALES") = colNew
                colRefreshed = True
            End If
        End If
        payload("registrado") = True
    Else
        payload("registrado") = False
    End If
    payload("colRefreshed") = colRefreshed

    ExpedienteEntidades_AltaPECAL = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "AltaPECAL: " & Err.Description
    End If
    ExpedienteEntidades_AltaPECAL = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- ExpedienteEntidades_EliminarPECAL -------------------------------------------
Public Function ExpedienteEntidades_EliminarPECAL( _
    ByVal p_DTO As Object, _
    ByVal p_IDPECAL As String, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(6)
    p_Error = ""

    On Error GoTo errores

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")

    Dim idP As String
    idP = Trim$("" & p_IDPECAL)

    If Len(idP) = 0 Then
        payload("eliminado") = False
        ExpedienteEntidades_EliminarPECAL = BuildJsonPayload(True, payload, "", logs)
        Exit Function
    End If

    Dim expObj As Object
    Set expObj = GetDTOSafely(p_DTO, "Expediente", Nothing)
    If expObj Is Nothing Then
        p_Error = "EliminarPECAL: p_DTO.Expediente is Nothing"
        ExpedienteEntidades_EliminarPECAL = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim pecal As Object
    Set pecal = constructor.getPecal(p_IDPEcal:=idP, p_Error:=p_Error)
    If p_Error <> "" Then
        ExpedienteEntidades_EliminarPECAL = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If
    If pecal Is Nothing Then
        p_Error = "Seleccione una PECAL de la lista"
        ExpedienteEntidades_EliminarPECAL = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim m_ExpOp As New ExpedienteOperaciones
    Set m_ExpOp.Expediente = expObj
    m_ExpOp.EliminarPECAL p_IDPEcal:=pecal.IDPECAL, p_Error:=p_Error
    If p_Error <> "" Then
        ExpedienteEntidades_EliminarPECAL = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim colRefreshed As Boolean
    colRefreshed = False
    On Error Resume Next
    Dim colNew As Object
    Set colNew = expObj.PECALES
    Dim refreshErr As Long
    refreshErr = Err.Number
    On Error GoTo errores
    If refreshErr = 0 Then
        If Not colNew Is Nothing Then
            p_DTO("ColPECALES") = colNew
            colRefreshed = True
        End If
    End If

    payload("eliminado") = True
    payload("colRefreshed") = colRefreshed

    ExpedienteEntidades_EliminarPECAL = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "EliminarPECAL: " & Err.Description
    End If
    ExpedienteEntidades_EliminarPECAL = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- ExpedienteEntidades_AltaRAC -------------------------------------------------
Public Function ExpedienteEntidades_AltaRAC( _
    ByVal p_DTO As Object, _
    ByVal p_IDRAC As String, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(6)
    p_Error = ""

    On Error GoTo errores

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")

    Dim idR As String
    idR = Trim$("" & p_IDRAC)

    If Len(idR) = 0 Then
        payload("registrado") = False
        ExpedienteEntidades_AltaRAC = BuildJsonPayload(True, payload, "", logs)
        Exit Function
    End If

    Dim expObj As Object
    Set expObj = GetDTOSafely(p_DTO, "Expediente", Nothing)
    If expObj Is Nothing Then
        p_Error = "AltaRAC: p_DTO.Expediente is Nothing"
        ExpedienteEntidades_AltaRAC = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim rac As Object
    Set rac = constructor.getRAC(p_IDRAC:=idR, p_Error:=p_Error)
    If p_Error <> "" Then
        ExpedienteEntidades_AltaRAC = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If
    If rac Is Nothing Then
        p_Error = "Se ha elegido un RAC que parece no estar registrado previamente"
        ExpedienteEntidades_AltaRAC = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim m_ExpOp As New ExpedienteOperaciones
    Set m_ExpOp.Expediente = expObj
    Dim estadoRegistro As String
    estadoRegistro = m_ExpOp.RegistrarRAC(p_IDRAC:=rac.IDRAC, p_Error:=p_Error)
    If p_Error <> "" Then
        ExpedienteEntidades_AltaRAC = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim colRefreshed As Boolean
    colRefreshed = False
    If estadoRegistro = "1" Then
        On Error Resume Next
        Dim colNew As Object
        Set colNew = expObj.RACs
        Dim refreshErr As Long
        refreshErr = Err.Number
        On Error GoTo errores
        If refreshErr = 0 Then
            If Not colNew Is Nothing Then
                p_DTO("ColRACs") = colNew
                colRefreshed = True
            End If
        End If
        payload("registrado") = True
    Else
        payload("registrado") = False
    End If
    payload("colRefreshed") = colRefreshed

    ExpedienteEntidades_AltaRAC = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "AltaRAC: " & Err.Description
    End If
    ExpedienteEntidades_AltaRAC = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- ExpedienteEntidades_EliminarRAC ---------------------------------------------
Public Function ExpedienteEntidades_EliminarRAC( _
    ByVal p_DTO As Object, _
    ByVal p_IDRAC As String, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(6)
    p_Error = ""

    On Error GoTo errores

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")

    Dim idR As String
    idR = Trim$("" & p_IDRAC)

    If Len(idR) = 0 Then
        payload("eliminado") = False
        ExpedienteEntidades_EliminarRAC = BuildJsonPayload(True, payload, "", logs)
        Exit Function
    End If

    Dim expObj As Object
    Set expObj = GetDTOSafely(p_DTO, "Expediente", Nothing)
    If expObj Is Nothing Then
        p_Error = "EliminarRAC: p_DTO.Expediente is Nothing"
        ExpedienteEntidades_EliminarRAC = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim rac As Object
    Set rac = constructor.getRAC(p_IDRAC:=idR, p_Error:=p_Error)
    If p_Error <> "" Then
        ExpedienteEntidades_EliminarRAC = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If
    If rac Is Nothing Then
        p_Error = "Seleccione un RAC de la lista"
        ExpedienteEntidades_EliminarRAC = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim m_ExpOp As New ExpedienteOperaciones
    Set m_ExpOp.Expediente = expObj
    m_ExpOp.EliminarRAC p_IDRAC:=rac.IDRAC, p_Error:=p_Error
    If p_Error <> "" Then
        ExpedienteEntidades_EliminarRAC = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim colRefreshed As Boolean
    colRefreshed = False
    On Error Resume Next
    Dim colNew As Object
    Set colNew = expObj.RACs
    Dim refreshErr As Long
    refreshErr = Err.Number
    On Error GoTo errores
    If refreshErr = 0 Then
        If Not colNew Is Nothing Then
            p_DTO("ColRACs") = colNew
            colRefreshed = True
        End If
    End If

    payload("eliminado") = True
    payload("colRefreshed") = colRefreshed

    ExpedienteEntidades_EliminarRAC = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "EliminarRAC: " & Err.Description
    End If
    ExpedienteEntidades_EliminarRAC = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function


' === Anualidades =================================================================

' --- ExpedienteEntidades_AltaAnualidad -------------------------------------------
' Pure-data alta Anualidad. Validates IsNumeric for Año and at least one BIIVA/IGIC/PSI/EXENTA.
Public Function ExpedienteEntidades_AltaAnualidad( _
    ByVal p_DTO As Object, _
    ByVal p_Año As String, _
    ByVal p_BIIVA As String, _
    ByVal p_BIIPSI As String, _
    ByVal p_BIIGIC As String, _
    ByVal p_BIEXENTA As String, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(6)
    p_Error = ""

    On Error GoTo errores

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")

    Dim año As String
    año = Trim$("" & p_Año)

    If Not IsNumeric(año) Then
        p_Error = "Se ha de incluir el anyo"
        ExpedienteEntidades_AltaAnualidad = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    ' At least one of the four importes must be numeric.
    Dim tieneImporte As Boolean
    tieneImporte = False
    If IsNumeric("" & p_BIIVA) Then tieneImporte = True
    If IsNumeric("" & p_BIIGIC) Then tieneImporte = True
    If IsNumeric("" & p_BIIPSI) Then tieneImporte = True
    If IsNumeric("" & p_BIEXENTA) Then tieneImporte = True

    If Not tieneImporte Then
        p_Error = "Se ha de incluir alguno de los cuatro importes"
        ExpedienteEntidades_AltaAnualidad = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim expObj As Object
    Set expObj = GetDTOSafely(p_DTO, "Expediente", Nothing)
    If expObj Is Nothing Then
        p_Error = "AltaAnualidad: p_DTO.Expediente is Nothing"
        ExpedienteEntidades_AltaAnualidad = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    ' Build the Anualidad class instance.
    Dim m_Anualidad As New ExpedienteAnualidad
    m_Anualidad.AÑO = año
    m_Anualidad.BIIVA = "" & p_BIIVA
    m_Anualidad.BIIGIC = "" & p_BIIGIC
    m_Anualidad.BIIPSI = "" & p_BIIPSI
    m_Anualidad.BIEXENTA = "" & p_BIEXENTA

    Dim m_ExpOp As New ExpedienteOperaciones
    Set m_ExpOp.Expediente = expObj
    Dim estadoRegistro As String
    estadoRegistro = m_ExpOp.RegistrarAnualidad(p_Anualidad:=m_Anualidad, p_Error:=p_Error)
    If p_Error <> "" Then
        ExpedienteEntidades_AltaAnualidad = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim colRefreshed As Boolean
    colRefreshed = False
    If estadoRegistro = "1" Then
        On Error Resume Next
        Dim colNew As Object
        Set colNew = expObj.Anualidades
        Dim refreshErr As Long
        refreshErr = Err.Number
        On Error GoTo errores
        If refreshErr = 0 Then
            If Not colNew Is Nothing Then
                p_DTO("ColAnualidades") = colNew
                colRefreshed = True
            End If
        End If
        payload("registrado") = True
    Else
        payload("registrado") = False
    End If
    payload("colRefreshed") = colRefreshed

    ExpedienteEntidades_AltaAnualidad = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "AltaAnualidad: " & Err.Description
    End If
    ExpedienteEntidades_AltaAnualidad = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- ExpedienteEntidades_EliminarAnualidad ---------------------------------------
' Pure-data baja Anualidad. Admin check first; if not admin, fail.
Public Function ExpedienteEntidades_EliminarAnualidad( _
    ByVal p_DTO As Object, _
    ByVal p_Año As String, _
    ByVal p_EsAdmin As Boolean, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(5)
    p_Error = ""

    On Error GoTo errores

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")

    Dim año As String
    año = Trim$("" & p_Año)

    If Len(año) = 0 Then
        payload("eliminado") = False
        ExpedienteEntidades_EliminarAnualidad = BuildJsonPayload(True, payload, "", logs)
        Exit Function
    End If

    If Not p_EsAdmin Then
        p_Error = "Operacion no autorizada"
        ExpedienteEntidades_EliminarAnualidad = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim expObj As Object
    Set expObj = GetDTOSafely(p_DTO, "Expediente", Nothing)
    If expObj Is Nothing Then
        p_Error = "EliminarAnualidad: p_DTO.Expediente is Nothing"
        ExpedienteEntidades_EliminarAnualidad = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim m_ExpOp As New ExpedienteOperaciones
    Set m_ExpOp.Expediente = expObj
    m_ExpOp.EliminarAnualidad p_Año:=año, p_Error:=p_Error
    If p_Error <> "" Then
        ExpedienteEntidades_EliminarAnualidad = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim colRefreshed As Boolean
    colRefreshed = False
    On Error Resume Next
    Dim colNew As Object
    Set colNew = expObj.Anualidades
    Dim refreshErr As Long
    refreshErr = Err.Number
    On Error GoTo errores
    If refreshErr = 0 Then
        If Not colNew Is Nothing Then
            p_DTO("ColAnualidades") = colNew
            colRefreshed = True
        End If
    End If

    payload("eliminado") = True
    payload("colRefreshed") = colRefreshed

    ExpedienteEntidades_EliminarAnualidad = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "EliminarAnualidad: " & Err.Description
    End If
    ExpedienteEntidades_EliminarAnualidad = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function


' === Responsables =================================================================

' --- ExpedienteEntidades_AltaResponsable -----------------------------------------
Public Function ExpedienteEntidades_AltaResponsable( _
    ByVal p_DTO As Object, _
    ByVal p_ResponsableNombre As String, _
    ByVal p_EsJefeProyecto As String, _
    ByVal p_CorreoSiempre As String, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(7)
    p_Error = ""

    On Error GoTo errores

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")

    Dim responsable As String
    responsable = Trim$("" & p_ResponsableNombre)

    If Len(responsable) = 0 Then
        p_Error = "Ha de seleccionar un responsable"
        ExpedienteEntidades_AltaResponsable = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim jp As String
    jp = Trim$("" & p_EsJefeProyecto)
    Dim correo As String
    correo = Trim$("" & p_CorreoSiempre)

    If Len(jp) = 0 Then
        p_Error = "Se ha de indicar si es o no Jefe de Proyecto"
        ExpedienteEntidades_AltaResponsable = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If
    If Len(correo) = 0 Then
        p_Error = "Se ha de indicar si va a recibir todas las comunicaciones por correo electronico"
        ExpedienteEntidades_AltaResponsable = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim m_Usuario As Object
    Set m_Usuario = getUsuario(p_Nombre:=responsable, p_Error:=p_Error)
    If p_Error <> "" Then
        ExpedienteEntidades_AltaResponsable = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If
    If m_Usuario Is Nothing Then
        p_Error = "Se ha de indicar el usuario"
        ExpedienteEntidades_AltaResponsable = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim expObj As Object
    Set expObj = GetDTOSafely(p_DTO, "Expediente", Nothing)
    If expObj Is Nothing Then
        p_Error = "AltaResponsable: p_DTO.Expediente is Nothing"
        ExpedienteEntidades_AltaResponsable = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim m_Responsable As New ExpedienteResponsable
    m_Responsable.IdUsuario = m_Usuario.ID
    m_Responsable.EsJefeProyecto = jp
    m_Responsable.CorreoSiempre = correo

    Dim m_ExpOp As New ExpedienteOperaciones
    Set m_ExpOp.Expediente = expObj
    Dim estadoRegistro As String
    estadoRegistro = m_ExpOp.RegistraResponsable(p_Responsable:=m_Responsable, p_Error:=p_Error)
    If p_Error <> "" Then
        ExpedienteEntidades_AltaResponsable = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim colRefreshed As Boolean
    colRefreshed = False
    If estadoRegistro = "1" Then
        On Error Resume Next
        Dim colNew As Object
        Set colNew = expObj.Responsables
        Dim refreshErr As Long
        refreshErr = Err.Number
        On Error GoTo errores
        If refreshErr = 0 Then
            If Not colNew Is Nothing Then
                p_DTO("ColResponsables") = colNew
                colRefreshed = True
            End If
        End If
        payload("registrado") = True
    Else
        payload("registrado") = False
    End If
    payload("colRefreshed") = colRefreshed

    ExpedienteEntidades_AltaResponsable = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "AltaResponsable: " & Err.Description
    End If
    ExpedienteEntidades_AltaResponsable = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- ExpedienteEntidades_EliminarResponsable -------------------------------------
Public Function ExpedienteEntidades_EliminarResponsable( _
    ByVal p_DTO As Object, _
    ByVal p_IDUsuario As String, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(6)
    p_Error = ""

    On Error GoTo errores

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")

    Dim idU As String
    idU = Trim$("" & p_IDUsuario)

    If Len(idU) = 0 Then
        payload("eliminado") = False
        ExpedienteEntidades_EliminarResponsable = BuildJsonPayload(True, payload, "", logs)
        Exit Function
    End If

    Dim expObj As Object
    Set expObj = GetDTOSafely(p_DTO, "Expediente", Nothing)
    If expObj Is Nothing Then
        p_Error = "EliminarResponsable: p_DTO.Expediente is Nothing"
        ExpedienteEntidades_EliminarResponsable = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim m_Usuario As Object
    Set m_Usuario = constructor.getUsuario(p_ID:=idU, p_Error:=p_Error)
    If p_Error <> "" Then
        ExpedienteEntidades_EliminarResponsable = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If
    If m_Usuario Is Nothing Then
        p_Error = "Seleccione un Responsable de la lista"
        ExpedienteEntidades_EliminarResponsable = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim m_ExpOp As New ExpedienteOperaciones
    Set m_ExpOp.Expediente = expObj
    m_ExpOp.EliminarResponsable p_IDUsuario:=m_Usuario.ID, p_Error:=p_Error
    If p_Error <> "" Then
        ExpedienteEntidades_EliminarResponsable = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim colRefreshed As Boolean
    colRefreshed = False
    On Error Resume Next
    Dim colNew As Object
    Set colNew = expObj.Responsables
    Dim refreshErr As Long
    refreshErr = Err.Number
    On Error GoTo errores
    If refreshErr = 0 Then
        If Not colNew Is Nothing Then
            p_DTO("ColResponsables") = colNew
            colRefreshed = True
        End If
    End If

    payload("eliminado") = True
    payload("colRefreshed") = colRefreshed

    ExpedienteEntidades_EliminarResponsable = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "EliminarResponsable: " & Err.Description
    End If
    ExpedienteEntidades_EliminarResponsable = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- ExpedienteEntidades_CambiarJP ------------------------------------------------
' Toggles JP for a responsable. p_JPFinal must be "Si" or "No".
Public Function ExpedienteEntidades_CambiarJP( _
    ByVal p_DTO As Object, _
    ByVal p_IDUsuario As String, _
    ByVal p_JPFinal As String, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(5)
    p_Error = ""

    On Error GoTo errores

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")

    Dim idU As String
    idU = Trim$("" & p_IDUsuario)
    Dim jpFinal As String
    jpFinal = Trim$("" & p_JPFinal)

    If Len(idU) = 0 Then
        p_Error = "Seleccione un elemento de la lista"
        ExpedienteEntidades_CambiarJP = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    If jpFinal <> "Si" And jpFinal <> "No" Then
        p_Error = "JPFinal invalido (debe ser 'Si' o 'No')"
        ExpedienteEntidades_CambiarJP = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim expObj As Object
    Set expObj = GetDTOSafely(p_DTO, "Expediente", Nothing)
    If expObj Is Nothing Then
        p_Error = "CambiarJP: p_DTO.Expediente is Nothing"
        ExpedienteEntidades_CambiarJP = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim m_ExpOp As New ExpedienteOperaciones
    Set m_ExpOp.Expediente = expObj
    m_ExpOp.EditarJPResponsable idU, jpFinal, p_Error:=p_Error
    If p_Error <> "" Then
        ExpedienteEntidades_CambiarJP = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    payload("editado") = True
    payload("idUsuario") = idU
    payload("jpFinal") = jpFinal

    ExpedienteEntidades_CambiarJP = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "CambiarJP: " & Err.Description
    End If
    ExpedienteEntidades_CambiarJP = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- ExpedienteEntidades_CambiarEnvio ---------------------------------------------
' Toggles envio correo for a responsable. p_EnvioFinal must be "Si" or "No".
Public Function ExpedienteEntidades_CambiarEnvio( _
    ByVal p_DTO As Object, _
    ByVal p_IDUsuario As String, _
    ByVal p_EnvioFinal As String, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(5)
    p_Error = ""

    On Error GoTo errores

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")

    Dim idU As String
    idU = Trim$("" & p_IDUsuario)
    Dim envioFinal As String
    envioFinal = Trim$("" & p_EnvioFinal)

    If Len(idU) = 0 Then
        p_Error = "Seleccione un elemento de la lista"
        ExpedienteEntidades_CambiarEnvio = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    If envioFinal <> "Si" And envioFinal <> "No" Then
        p_Error = "EnvioFinal invalido (debe ser 'Si' o 'No')"
        ExpedienteEntidades_CambiarEnvio = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim expObj As Object
    Set expObj = GetDTOSafely(p_DTO, "Expediente", Nothing)
    If expObj Is Nothing Then
        p_Error = "CambiarEnvio: p_DTO.Expediente is Nothing"
        ExpedienteEntidades_CambiarEnvio = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim m_ExpOp As New ExpedienteOperaciones
    Set m_ExpOp.Expediente = expObj
    m_ExpOp.EditarEnvioCorreoResponsable idU, envioFinal, p_Error:=p_Error
    If p_Error <> "" Then
        ExpedienteEntidades_CambiarEnvio = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    payload("editado") = True
    payload("idUsuario") = idU
    payload("envioFinal") = envioFinal

    ExpedienteEntidades_CambiarEnvio = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "CambiarEnvio: " & Err.Description
    End If
    ExpedienteEntidades_CambiarEnvio = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function


' === Rellenar Listas ===============================================================

' --- ExpedienteEntidades_RellenarListaComerciales --------------------------------
' Walks p_DTO("ColComerciales") and emits a 2-col rowSource.
' Reads each Comercial's IDComercial + Comercial as properties (or Dictionary keys).
Public Function ExpedienteEntidades_RellenarListaComerciales( _
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
    Set col = GetDTOSafely(p_DTO, "ColComerciales", Nothing)

    Dim rowSource As String
    rowSource = ""
    Dim rowCount As Long
    rowCount = 0

    If col Is Nothing Then
        payload("rowSource") = ""
        payload("count") = 0
        ExpedienteEntidades_RellenarListaComerciales = BuildJsonPayload(True, payload, "", logs)
        Exit Function
    End If

    Dim k As Variant
    Dim com As Object
    Dim idC As String
    Dim nomC As String

    For Each k In col.Keys
        If rowCount >= EXPEDIENTEENT_MAX_ROWS Then Exit For
        Set com = col(k)
        If com Is Nothing Then
            ' skip
        Else
            idC = "" & com.IDComercial
            nomC = "" & com.Comercial
            If Len(rowSource) > 0 Then rowSource = rowSource & vbCrLf
            rowSource = rowSource & BuildRowLine2(idC, nomC)
            rowCount = rowCount + 1
        End If
    Next k

    payload("rowSource") = rowSource
    payload("count") = rowCount
    ExpedienteEntidades_RellenarListaComerciales = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "RellenarListaComerciales: " & Err.Description
    End If
    ExpedienteEntidades_RellenarListaComerciales = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- ExpedienteEntidades_RellenarListaCPVs ----------------------------------------
Public Function ExpedienteEntidades_RellenarListaCPVs( _
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
    Set col = GetDTOSafely(p_DTO, "ColCPVs", Nothing)

    Dim rowSource As String
    rowSource = ""
    Dim rowCount As Long
    rowCount = 0

    If col Is Nothing Then
        payload("rowSource") = ""
        payload("count") = 0
        ExpedienteEntidades_RellenarListaCPVs = BuildJsonPayload(True, payload, "", logs)
        Exit Function
    End If

    Dim k As Variant
    Dim cpv As Object
    Dim idC As String
    Dim nomC As String

    For Each k In col.Keys
        If rowCount >= EXPEDIENTEENT_MAX_ROWS Then Exit For
        Set cpv = col(k)
        If cpv Is Nothing Then
            ' skip
        Else
            idC = "" & cpv.IDCPV
            nomC = "" & cpv.CPV
            If Len(rowSource) > 0 Then rowSource = rowSource & vbCrLf
            rowSource = rowSource & BuildRowLine2(idC, nomC)
            rowCount = rowCount + 1
        End If
    Next k

    payload("rowSource") = rowSource
    payload("count") = rowCount
    ExpedienteEntidades_RellenarListaCPVs = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "RellenarListaCPVs: " & Err.Description
    End If
    ExpedienteEntidades_RellenarListaCPVs = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- ExpedienteEntidades_RellenarListaLugares ------------------------------------
Public Function ExpedienteEntidades_RellenarListaLugares( _
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
    Set col = GetDTOSafely(p_DTO, "ColLugaresEjecucion", Nothing)

    Dim rowSource As String
    rowSource = ""
    Dim rowCount As Long
    rowCount = 0

    If col Is Nothing Then
        payload("rowSource") = ""
        payload("count") = 0
        ExpedienteEntidades_RellenarListaLugares = BuildJsonPayload(True, payload, "", logs)
        Exit Function
    End If

    Dim k As Variant
    Dim lugar As Object
    Dim idL As String
    Dim nomL As String

    For Each k In col.Keys
        If rowCount >= EXPEDIENTEENT_MAX_ROWS Then Exit For
        Set lugar = col(k)
        If lugar Is Nothing Then
            ' skip
        Else
            idL = "" & lugar.IDLugarEjecucion
            nomL = "" & lugar.LugarEjecucion
            If Len(rowSource) > 0 Then rowSource = rowSource & vbCrLf
            rowSource = rowSource & BuildRowLine2(idL, nomL)
            rowCount = rowCount + 1
        End If
    Next k

    payload("rowSource") = rowSource
    payload("count") = rowCount
    ExpedienteEntidades_RellenarListaLugares = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "RellenarListaLugares: " & Err.Description
    End If
    ExpedienteEntidades_RellenarListaLugares = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- ExpedienteEntidades_RellenarListaPECALES -------------------------------------
Public Function ExpedienteEntidades_RellenarListaPECALES( _
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
    Set col = GetDTOSafely(p_DTO, "ColPECALES", Nothing)

    Dim rowSource As String
    rowSource = ""
    Dim rowCount As Long
    rowCount = 0

    If col Is Nothing Then
        payload("rowSource") = ""
        payload("count") = 0
        ExpedienteEntidades_RellenarListaPECALES = BuildJsonPayload(True, payload, "", logs)
        Exit Function
    End If

    Dim k As Variant
    Dim pecal As Object
    Dim idP As String
    Dim nomP As String

    For Each k In col.Keys
        If rowCount >= EXPEDIENTEENT_MAX_ROWS Then Exit For
        Set pecal = col(k)
        If pecal Is Nothing Then
            ' skip
        Else
            idP = "" & pecal.IDPECAL
            nomP = "" & pecal.PECAL
            If Len(rowSource) > 0 Then rowSource = rowSource & vbCrLf
            rowSource = rowSource & BuildRowLine2(idP, nomP)
            rowCount = rowCount + 1
        End If
    Next k

    payload("rowSource") = rowSource
    payload("count") = rowCount
    ExpedienteEntidades_RellenarListaPECALES = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "RellenarListaPECALES: " & Err.Description
    End If
    ExpedienteEntidades_RellenarListaPECALES = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- ExpedienteEntidades_RellenarListaRACs ----------------------------------------
Public Function ExpedienteEntidades_RellenarListaRACs( _
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
    Set col = GetDTOSafely(p_DTO, "ColRACs", Nothing)

    Dim rowSource As String
    rowSource = ""
    Dim rowCount As Long
    rowCount = 0

    If col Is Nothing Then
        payload("rowSource") = ""
        payload("count") = 0
        ExpedienteEntidades_RellenarListaRACs = BuildJsonPayload(True, payload, "", logs)
        Exit Function
    End If

    Dim k As Variant
    Dim rac As Object
    Dim idR As String
    Dim nomR As String

    For Each k In col.Keys
        If rowCount >= EXPEDIENTEENT_MAX_ROWS Then Exit For
        Set rac = col(k)
        If rac Is Nothing Then
            ' skip
        Else
            idR = "" & rac.IDRAC
            nomR = "" & rac.RAC
            If Len(rowSource) > 0 Then rowSource = rowSource & vbCrLf
            rowSource = rowSource & BuildRowLine2(idR, nomR)
            rowCount = rowCount + 1
        End If
    Next k

    payload("rowSource") = rowSource
    payload("count") = rowCount
    ExpedienteEntidades_RellenarListaRACs = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "RellenarListaRACs: " & Err.Description
    End If
    ExpedienteEntidades_RellenarListaRACs = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- ExpedienteEntidades_RellenarListaResponsables --------------------------------
' 4-col rowSource: ID;Nombre;JP;Aviso. Headers are explicit.
Public Function ExpedienteEntidades_RellenarListaResponsables( _
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
    Set col = GetDTOSafely(p_DTO, "ColResponsables", Nothing)

    Dim rowSource As String
    rowSource = EXPEDIENTEENT_HEADERS_RESPONSABLES
    Dim rowCount As Long
    rowCount = 0

    If col Is Nothing Then
        payload("rowSource") = rowSource
        payload("count") = 0
        ExpedienteEntidades_RellenarListaResponsables = BuildJsonPayload(True, payload, "", logs)
        Exit Function
    End If

    Dim k As Variant
    Dim resp As Object
    Dim idR As String
    Dim nomR As String
    Dim jpR As String
    Dim avisoR As String

    For Each k In col.Keys
        If rowCount >= EXPEDIENTEENT_MAX_ROWS Then Exit For
        Set resp = col(k)
        If resp Is Nothing Then
            ' skip
        Else
            idR = "" & resp.IdUsuario
            nomR = "" & resp.USUARIO.Nombre
            jpR = "" & resp.EsJefeProyecto
            avisoR = "" & resp.CorreoSiempre
            rowSource = rowSource & vbCrLf & BuildRowLine4(idR, nomR, jpR, avisoR)
            rowCount = rowCount + 1
        End If
    Next k

    payload("rowSource") = rowSource
    payload("count") = rowCount
    ExpedienteEntidades_RellenarListaResponsables = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "RellenarListaResponsables: " & Err.Description
    End If
    ExpedienteEntidades_RellenarListaResponsables = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- ExpedienteEntidades_RellenarListaAnualidades ---------------------------------
' 2-col rowSource: Año;Presupuesto (formatted as #,##0.00 EUR).
Public Function ExpedienteEntidades_RellenarListaAnualidades( _
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
    Set col = GetDTOSafely(p_DTO, "ColAnualidades", Nothing)

    Dim rowSource As String
    rowSource = EXPEDIENTEENT_HEADERS_ANUALIDADES
    Dim rowCount As Long
    rowCount = 0

    If col Is Nothing Then
        payload("rowSource") = rowSource
        payload("count") = 0
        ExpedienteEntidades_RellenarListaAnualidades = BuildJsonPayload(True, payload, "", logs)
        Exit Function
    End If

    Dim k As Variant
    Dim anual As Object
    Dim año As String
    Dim presup As String

    For Each k In col.Keys
        If rowCount >= EXPEDIENTEENT_MAX_ROWS Then Exit For
        Set anual = col(k)
        If anual Is Nothing Then
            ' skip
        Else
            año = "" & anual.AÑO
            presup = Format(anual.Presupuesto, "#,##0.00 EUR")
            rowSource = rowSource & vbCrLf & BuildRowLine2(año, presup)
            rowCount = rowCount + 1
        End If
    Next k

    payload("rowSource") = rowSource
    payload("count") = rowCount
    ExpedienteEntidades_RellenarListaAnualidades = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "RellenarListaAnualidades: " & Err.Description
    End If
    ExpedienteEntidades_RellenarListaAnualidades = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- ExpedienteEntidades_RellenarListas -------------------------------------------
' Refreshes all 7 collections by calling getExpedienteX (re-fetch from DB) and returns
' JSON with all 7 rowSources. The form applies each rowSource to its listbox.
' Returns JSON: {ok, payload:{comerciales,cpvs,lugares,pecales,racs,responsables,anualidades,errors[]}, ...}
Public Function ExpedienteEntidades_RellenarListas( _
    ByVal p_DTO As Object, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(10)
    p_Error = ""

    On Error GoTo errores

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")
    Dim errors As Object
    Set errors = CreateObject("Scripting.Dictionary")

    Dim expObj As Object
    Set expObj = GetDTOSafely(p_DTO, "Expediente", Nothing)
    Dim idExp As String
    idExp = ""
    If expObj Is Nothing Then
        p_Error = "RellenarListas: p_DTO.Expediente is Nothing"
        ExpedienteEntidades_RellenarListas = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If
    idExp = "" & expObj.IDExpediente

    ' PECALES
    Dim pErrorLocal As String
    pErrorLocal = ""
    Dim newCol As Object
    Set newCol = getExpedientePECALES(idExp, pErrorLocal)
    If pErrorLocal = "" And Not newCol Is Nothing Then
        p_DTO("ColPECALES") = newCol
    Else
        errors("pecales") = pErrorLocal
    End If

    ' CPVs
    pErrorLocal = ""
    Set newCol = getExpedienteCPVS(idExp, pErrorLocal)
    If pErrorLocal = "" And Not newCol Is Nothing Then
        p_DTO("ColCPVs") = newCol
    Else
        errors("cpvs") = pErrorLocal
    End If

    ' Lugares
    pErrorLocal = ""
    Set newCol = getExpedienteLugaresEjecucion(idExp, pErrorLocal)
    If pErrorLocal = "" And Not newCol Is Nothing Then
        p_DTO("ColLugaresEjecucion") = newCol
    Else
        errors("lugares") = pErrorLocal
    End If

    ' Comerciales
    pErrorLocal = ""
    Set newCol = getExpedienteComerciales(idExp, pErrorLocal)
    If pErrorLocal = "" And Not newCol Is Nothing Then
        p_DTO("ColComerciales") = newCol
    Else
        errors("comerciales") = pErrorLocal
    End If

    ' RACs
    pErrorLocal = ""
    Set newCol = getExpedienteRACS(idExp, pErrorLocal)
    If pErrorLocal = "" And Not newCol Is Nothing Then
        p_DTO("ColRACs") = newCol
    Else
        errors("racs") = pErrorLocal
    End If

    ' Responsables
    pErrorLocal = ""
    Set newCol = getExpedienteResponsables(idExp, pErrorLocal)
    If pErrorLocal = "" And Not newCol Is Nothing Then
        p_DTO("ColResponsables") = newCol
    Else
        errors("responsables") = pErrorLocal
    End If

    ' Anualidades
    pErrorLocal = ""
    Set newCol = getExpedienteAnualidades(idExp, pErrorLocal)
    If pErrorLocal = "" And Not newCol Is Nothing Then
        p_DTO("ColAnualidades") = newCol
    Else
        errors("anualidades") = pErrorLocal
    End If

    ' Compute each rowSource by calling the per-listbox helper.
    Dim rsCom As String, rsCPV As String, rsLug As String, rsPec As String
    Dim rsRac As String, rsRes As String, rsAnu As String
    Dim helperResult As String

    helperResult = ExpedienteEntidades_RellenarListaComerciales(p_DTO, pErrorLocal)
    If JsonConverter.ParseJson(helperResult)("ok") = True Then
        rsCom = JsonConverter.ParseJson(helperResult)("payload")("rowSource")
    End If

    helperResult = ExpedienteEntidades_RellenarListaCPVs(p_DTO, pErrorLocal)
    If JsonConverter.ParseJson(helperResult)("ok") = True Then
        rsCPV = JsonConverter.ParseJson(helperResult)("payload")("rowSource")
    End If

    helperResult = ExpedienteEntidades_RellenarListaLugares(p_DTO, pErrorLocal)
    If JsonConverter.ParseJson(helperResult)("ok") = True Then
        rsLug = JsonConverter.ParseJson(helperResult)("payload")("rowSource")
    End If

    helperResult = ExpedienteEntidades_RellenarListaPECALES(p_DTO, pErrorLocal)
    If JsonConverter.ParseJson(helperResult)("ok") = True Then
        rsPec = JsonConverter.ParseJson(helperResult)("payload")("rowSource")
    End If

    helperResult = ExpedienteEntidades_RellenarListaRACs(p_DTO, pErrorLocal)
    If JsonConverter.ParseJson(helperResult)("ok") = True Then
        rsRac = JsonConverter.ParseJson(helperResult)("payload")("rowSource")
    End If

    helperResult = ExpedienteEntidades_RellenarListaResponsables(p_DTO, pErrorLocal)
    If JsonConverter.ParseJson(helperResult)("ok") = True Then
        rsRes = JsonConverter.ParseJson(helperResult)("payload")("rowSource")
    End If

    helperResult = ExpedienteEntidades_RellenarListaAnualidades(p_DTO, pErrorLocal)
    If JsonConverter.ParseJson(helperResult)("ok") = True Then
        rsAnu = JsonConverter.ParseJson(helperResult)("payload")("rowSource")
    End If

    payload("comerciales") = rsCom
    payload("cpvs") = rsCPV
    payload("lugares") = rsLug
    payload("pecales") = rsPec
    payload("racs") = rsRac
    payload("responsables") = rsRes
    payload("anualidades") = rsAnu
    payload("errors") = errors

    ExpedienteEntidades_RellenarListas = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "RellenarListas: " & Err.Description
    End If
    ExpedienteEntidades_RellenarListas = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function


' === Form-level ==================================================================

' --- ExpedienteEntidades_EstablecerDatos ------------------------------------------
' Pure-data decision: returns whether EJECUTIVO-tagged controls should be enabled
' (admin + edit) and whether the DTO has a valid Expediente. The form does the
' actual For-Each-Controls loop with Tag="EJECUTIVO".
' Returns JSON: {ok, payload:{ejecutivosEnabled, expedienteOK}, error, logs}.
Public Function ExpedienteEntidades_EstablecerDatos( _
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

    Dim expedienteOK As Boolean
    expedienteOK = False
    Dim expObj As Object
    Set expObj = GetDTOSafely(p_DTO, "Expediente", Nothing)
    If Not expObj Is Nothing Then expedienteOK = True

    ' EJECUTIVO controls enabled iff abierto-para-editar AND admin.
    Dim ejecutivosEnabled As Boolean
    ejecutivosEnabled = False
    If p_AbiertoParaEditar And p_EsAdmin Then
        ejecutivosEnabled = True
    End If

    payload("ejecutivosEnabled") = ejecutivosEnabled
    payload("expedienteOK") = expedienteOK

    logs(0) = "EstablecerDatos: esAdmin=" & CStr(p_EsAdmin) & _
              ", abiertoParaEditar=" & CStr(p_AbiertoParaEditar) & _
              ", ejecutivosEnabled=" & CStr(ejecutivosEnabled)

    ExpedienteEntidades_EstablecerDatos = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "EstablecerDatos: " & Err.Description
    End If
    ExpedienteEntidades_EstablecerDatos = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- ExpedienteEntidades_Form_Load ------------------------------------------------
' Pure-data init: validates DTO, decides enable flags, hides ComandoRegistrar.
' Returns JSON: {ok, payload:{expedienteOK, ejecutivoEnabled, hideComandoRegistrar}, error, logs}.
Public Function ExpedienteEntidades_Form_Load( _
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
    Set expObj = GetDTOSafely(p_DTO, "Expediente", Nothing)
    If expObj Is Nothing Then
        p_Error = "El formulario se ha abierto con parametros insuficientes"
        ExpedienteEntidades_Form_Load = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim ejecutivoEnabled As Boolean
    ejecutivoEnabled = False
    If p_AbiertoParaEditar And p_EsAdmin Then ejecutivoEnabled = True

    payload("expedienteOK") = True
    payload("ejecutivoEnabled") = ejecutivoEnabled
    payload("hideComandoRegistrar") = True

    logs(0) = "Form_Load: esAdmin=" & CStr(p_EsAdmin) & _
              ", abiertoParaEditar=" & CStr(p_AbiertoParaEditar)

    ExpedienteEntidades_Form_Load = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "Form_Load: " & Err.Description
    End If
    ExpedienteEntidades_Form_Load = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function


' === Misc =========================================================================

' --- ExpedienteEntidades_CambiarLineaListaResponsables ----------------------------
' Pure string op on a rowSource: replace p_LineaActual with p_LineaFinal. Returns
' the new rowSource string; the form assigns it to the listbox.
' Returns JSON: {ok, payload:{rowSource, replaced}, error, logs}.
Public Function ExpedienteEntidades_CambiarLineaListaResponsables( _
    ByVal p_RowSource As String, _
    ByVal p_LineaActual As String, _
    ByVal p_LineaFinal As String, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(4)
    p_Error = ""

    On Error GoTo errores

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")

    Dim newRowSource As String
    newRowSource = "" & p_RowSource

    Dim replaced As Boolean
    replaced = False

    If Len(p_LineaActual) > 0 Then
        If InStr(1, newRowSource, p_LineaActual) <> 0 Then
            newRowSource = Replace(newRowSource, p_LineaActual, p_LineaFinal)
            replaced = True
        End If
    End If

    payload("rowSource") = newRowSource
    payload("replaced") = replaced

    logs(0) = "CambiarLineaListaResponsables: replaced=" & CStr(replaced) & _
              ", lenAntes=" & Len(p_RowSource) & ", lenDespues=" & Len(newRowSource)

    ExpedienteEntidades_CambiarLineaListaResponsables = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "CambiarLineaListaResponsables: " & Err.Description
    End If
    ExpedienteEntidades_CambiarLineaListaResponsables = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function
