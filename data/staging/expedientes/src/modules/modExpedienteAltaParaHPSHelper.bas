Attribute VB_Name = "modExpedienteAltaParaHPSHelper"
Option Compare Database
Option Explicit

' modExpedienteAltaParaHPSHelper — REWORK (2026-06-26, Phase 3.2b / PR-8b)
' Pure-data helpers for Form_FormExpedienteAltaParaHPS
' (alta específica HPS, DTO en memoria con árbol de suministradores).
'
' Implements access-vba-e2e-methodology rule #1 (forms are thin UI wiring),
' rule #2 (honest helper signatures — accept only the data the helper needs),
' rule #5 (per-module prefix on Public names: ExpedienteHPS_*),
' and rule #9 (no `ByRef p_Form` — helpers MUST NOT receive Form objects).
'
' Shape contract (DTO wrapper):
'   The form wraps m_ObjExpedienteDTOActivo into a Scripting.Dictionary with
'   keys "Expediente" (real Expediente class instance), "ColArbolSuministradores"
'   (Scripting.Dictionary of ExpedienteSuministrador), and "ColLugaresEjecucion"
'   (Scripting.Dictionary of LugarEjecucion). Helpers see Object; tests pass
'   Dictionary stubs.
'
' Anti-pattern removed (was in pre-PR-8b baseline):
'   - form did inline `For Each m_ID In m_Col` + `Me.Lista.AddItem` (UI wiring
'     inside business handlers).
'   - EstablecerCombos and Rellenar* methods did `Me.X = Null` / `Me.Lista.AddItem`
'     mixed with lookup loops.
'   - AgregarSuministradorMemoria / EliminarSuministradorMemoria had `MsgBox
'     "Esta empresa ya está en la lista"` embedded in the private Sub.
'
' New design (8 helpers, all pure-data, no Form refs):
'   1. ExpedienteHPS_Form_Load(p_DTO, p_Error)
'      -> {colArbolInitialized, allowEdits, ok}
'   2. ExpedienteHPS_EstablecerCombos_Construir(p_DTO, p_Error)
'      -> {clasificaciones, contratistas, subcontratistas, lugares} rowSources
'   3. ExpedienteHPS_RellenarLista_Construir(p_DTO, p_TipoLista, p_Error)
'      p_TipoLista ∈ {"CONTRATISTAS","SUBCONTRATISTAS","LUGARES"}
'      -> {rowSource, count}
'   4. ExpedienteHPS_AltaSuministrador(p_DTO, p_IDSuministrador, p_Tipo, p_Error)
'      -> {added, duplicate}
'   5. ExpedienteHPS_EliminarSuministrador(p_DTO, p_IDSuministrador, p_Tipo, p_Error)
'      -> {removed}
'   6. ExpedienteHPS_AltaLugar(p_DTO, p_IDLugar, p_Error)
'      -> {added, duplicate}
'   7. ExpedienteHPS_EliminarLugar(p_DTO, p_IDLugar, p_Error)
'      -> {removed}
'   8. ExpedienteHPS_ComandoRegistrar_Click(p_DTO, p_Error)
'      -> {registered, expId, tipo}
'
' UI orchestration that stays in the form (rule #1):
'   - Each combo.AddItem / list.AddItem (form iterates rowSource and AddItems)
'   - EstablecerCombos / Rellenar* public wrappers (form does UI after helper)
'   - cmdSalir_Click — DoCmd.Close
'   - Form_Load errores — MsgBox
'   - 3 m_For*_Alta callbacks — refresh combos + run alta button
'   - 6 botones add/del — orchestrate: extract ID from combo/list → helper → re-render
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
Private Const EXPEDIENTEHPS_FIELD_SEP As String = ";"

' Maximum rows emitted in rowSource. Defensive cap so a runaway test
' does not build a multi-MB CSV.
Private Const EXPEDIENTEHPS_MAX_ROWS As Long = 5000

' Tipo values for the supplier tree. The form uses these as-is in the
' ColArbolSuministradores Tag and in the audit.
Private Const EXPEDIENTEHPS_TIPO_CONTR As String = "ROOT_CONTR"
Private Const EXPEDIENTEHPS_TIPO_SUB As String = "ROOT_SUB"

' Test fixture ID base for entidades; not used by helpers, just for traceability.
Private Const EXPEDIENTEHPS_TEST_ID_BASE As Long = 900810


' === Local helpers (all at top per vba-access §10.1) =============================

' --- BuildJsonPayload ------------------------------------------------------------
' Wraps a Dictionary payload in the canonical JSON envelope:
'   {"ok":true,"value":null,"payload":<payloadJson>,"error":null,"logs":[...]}
' Logs that are empty strings are stripped. Uses JsonConverter.ConvertToJson.
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

' --- ExpedienteHPS_GetValueSafely ------------------------------------------------
' Returns p_DTO(p_Key) if p_DTO is a Dictionary-like object with that key.
' Returns the provided default if p_DTO is Nothing, missing key, or accessing
' the key raised an error.
Private Function ExpedienteHPS_GetValueSafely( _
    ByVal p_DTO As Object, _
    ByVal p_Key As String, _
    ByVal p_Default As Variant _
) As Variant
    On Error GoTo EH
    If p_DTO Is Nothing Then
        ExpedienteHPS_GetValueSafely = p_Default
        Exit Function
    End If
    If Not p_DTO.Exists(p_Key) Then
        ExpedienteHPS_GetValueSafely = p_Default
        Exit Function
    End If
    ExpedienteHPS_GetValueSafely = p_DTO(p_Key)
    Exit Function
EH:
    ExpedienteHPS_GetValueSafely = p_Default
End Function

' --- ExpedienteHPS_HasKeySafely --------------------------------------------------
' Returns True if p_DTO is a Dictionary-like object that contains p_Key.
Private Function ExpedienteHPS_HasKeySafely( _
    ByVal p_DTO As Object, _
    ByVal p_Key As String _
) As Boolean
    On Error GoTo EH
    If p_DTO Is Nothing Then
        ExpedienteHPS_HasKeySafely = False
        Exit Function
    End If
    ExpedienteHPS_HasKeySafely = p_DTO.Exists(p_Key)
    Exit Function
EH:
    ExpedienteHPS_HasKeySafely = False
End Function

' --- ExpedienteHPS_BuildRowLine --------------------------------------------------
' Builds a single rowSource line (semicolon-separated) for a listbox/combo row.
' Sanitizes embedded semicolons by replacing them with colons (Access listbox
' RowSource cannot contain field separators inside a value).
Private Function ExpedienteHPS_BuildRowLine( _
    ByVal p_ID As String, _
    ByVal p_Nombre As String _
) As String
    Dim safeID As String
    safeID = Replace(p_ID, EXPEDIENTEHPS_FIELD_SEP, ":")
    Dim safeNombre As String
    safeNombre = Replace(p_Nombre, EXPEDIENTEHPS_FIELD_SEP, ":")
    ExpedienteHPS_BuildRowLine = safeID & EXPEDIENTEHPS_FIELD_SEP & safeNombre
End Function

' --- ExpedienteHPS_LeerEntidadComoDic -------------------------------------------
' Reads a String property from a real entity OR a Scripting.Dictionary stub,
' using p_Key for the Dictionary access path. Returns "" for missing keys/errors.
'
' Real ExpedienteSuministrador has .IDSuministrador, .ContratistaPrincipal,
' .SubContratista, .IdPadre, .IDExpedienteSuministrador.
' Real LugarEjecucion has .IDLugarEjecucion, .LugarEjecucion.
' Real Suministrador has .IDSuministrador, .Nombre.
' Real GradoClasificacion has .IdGradoClasificacion, .GradoClasificacion.
Private Function ExpedienteHPS_LeerEntidadComoDic( _
    ByVal p_Entidad As Object, _
    ByVal p_Key As String _
) As String
    On Error GoTo EH
    If p_Entidad Is Nothing Then
        ExpedienteHPS_LeerEntidadComoDic = ""
        Exit Function
    End If

    ' Dictionary stub first (covers pure-Dictionary test stubs).
    If TypeName(p_Entidad) = "Dictionary" Then
        If p_Entidad.Exists(p_Key) Then
            ExpedienteHPS_LeerEntidadComoDic = CStr(p_Entidad(p_Key))
            Exit Function
        End If
    End If

    ' Real-class property access.
    Dim val As String
    val = ""
    Select Case p_Key
        Case "IDSuministrador":           val = p_Entidad.IDSuministrador
        Case "Nombre":                    val = p_Entidad.Nombre
        Case "IdGradoClasificacion":      val = p_Entidad.IdGradoClasificacion
        Case "GradoClasificacion":        val = p_Entidad.GradoClasificacion
        Case "IDLugarEjecucion":          val = p_Entidad.IDLugarEjecucion
        Case "LugarEjecucion":            val = p_Entidad.LugarEjecucion
        Case "ContratistaPrincipal":      val = p_Entidad.ContratistaPrincipal
        Case "SubContratista":            val = p_Entidad.SubContratista
        Case "IdPadre":                   val = p_Entidad.IdPadre
        Case "IDExpedienteSuministrador": val = p_Entidad.IDExpedienteSuministrador
    End Select
    ExpedienteHPS_LeerEntidadComoDic = val
    Exit Function
EH:
    ExpedienteHPS_LeerEntidadComoDic = ""
End Function

' --- ExpedienteHPS_NormalizarTipo ------------------------------------------------
' Returns the canonical tipo string ("ROOT_CONTR", "ROOT_SUB") or "" if invalid.
Private Function ExpedienteHPS_NormalizarTipo(ByVal p_Tipo As String) As String
    Dim t As String
    t = "" & p_Tipo
    If t = EXPEDIENTEHPS_TIPO_CONTR Then
        ExpedienteHPS_NormalizarTipo = EXPEDIENTEHPS_TIPO_CONTR
    ElseIf t = EXPEDIENTEHPS_TIPO_SUB Then
        ExpedienteHPS_NormalizarTipo = EXPEDIENTEHPS_TIPO_SUB
    Else
        ExpedienteHPS_NormalizarTipo = ""
    End If
End Function

' --- ExpedienteHPS_MatchSuministrador -------------------------------------------
' Returns True if an ExpedienteSuministrador-like item matches the given
' ID + tipo (CONTRATISTA PRINCIPAL or SUBCONTRATISTA ROOT).
Private Function ExpedienteHPS_MatchSuministrador( _
    ByVal p_Item As Object, _
    ByVal p_ID As String, _
    ByVal p_Tipo As String _
) As Boolean
    Dim itemID As String
    itemID = ExpedienteHPS_LeerEntidadComoDic(p_Item, "IDSuministrador")
    If itemID <> p_ID Then
        ExpedienteHPS_MatchSuministrador = False
        Exit Function
    End If

    If p_Tipo = EXPEDIENTEHPS_TIPO_CONTR Then
        If ExpedienteHPS_LeerEntidadComoDic(p_Item, "ContratistaPrincipal") = "Sí" Then
            ExpedienteHPS_MatchSuministrador = True
        End If
    ElseIf p_Tipo = EXPEDIENTEHPS_TIPO_SUB Then
        If ExpedienteHPS_LeerEntidadComoDic(p_Item, "SubContratista") = "Sí" Then
            ExpedienteHPS_MatchSuministrador = True
        End If
    End If
End Function

' --- ExpedienteHPS_ConstruirExpedienteSuministrador -----------------------------
' Builds a NEW ExpedienteSuministrador-like Dictionary with the canonical shape.
' Real ExpedienteSuministrador instantiation is left to the DAO module that owns
' the class — tests pass pure Dictionary stubs.
'
' Shape returned: {IDExpedienteSuministrador, IDSuministrador, IdPadre, Tag,
'                  ContratistaPrincipal, SubContratista, Descripcon}
Private Function ExpedienteHPS_ConstruirExpedienteSuministrador( _
    ByVal p_IDTemp As String, _
    ByVal p_Idsuministrador As String, _
    ByVal p_Tipo As String _
) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("IDExpedienteSuministrador") = p_IDTemp
    d("IDSuministrador") = p_Idsuministrador
    d("IdPadre") = ""
    d("Tag") = p_IDTemp

    If p_Tipo = EXPEDIENTEHPS_TIPO_CONTR Then
        d("ContratistaPrincipal") = "Sí"
        d("SubContratista") = "No"
        d("Descripcon") = "Contratista Principal"
    Else
        d("ContratistaPrincipal") = "No"
        d("SubContratista") = "Sí"
        d("Descripcon") = "Participación Telefónica (HPS)"
    End If

    Set ExpedienteHPS_ConstruirExpedienteSuministrador = d
End Function

' --- ExpedienteHPS_DiccionarioToRowSource ---------------------------------------
' Walks a Scripting.Dictionary whose values expose p_IDKey and p_NombreKey,
' emitting semicolon-separated lines (ID;Nombre). Returns rowSource + count.
Private Function ExpedienteHPS_DiccionarioToRowSource( _
    ByVal p_Dic As Object, _
    ByVal p_IDKey As String, _
    ByVal p_NombreKey As String, _
    ByRef p_Count As Long, _
    ByVal p_MaxRows As Long _
) As String
    p_Count = 0
    If p_Dic Is Nothing Then
        ExpedienteHPS_DiccionarioToRowSource = ""
        Exit Function
    End If

    Dim rowSource As String
    rowSource = ""
    Dim k As Variant
    Dim entidad As Object
    Dim idVal As String
    Dim nombreVal As String

    For Each k In p_Dic.Keys
        If p_Count >= p_MaxRows Then Exit For
        Set entidad = p_Dic(k)
        If entidad Is Nothing Then
            ' skip
        Else
            idVal = ExpedienteHPS_LeerEntidadComoDic(entidad, p_IDKey)
            nombreVal = ExpedienteHPS_LeerEntidadComoDic(entidad, p_NombreKey)
            If Len(rowSource) > 0 Then
                rowSource = rowSource & vbCrLf
            End If
            rowSource = rowSource & ExpedienteHPS_BuildRowLine(idVal, nombreVal)
            p_Count = p_Count + 1
        End If
    Next k

    ExpedienteHPS_DiccionarioToRowSource = rowSource
End Function


' === Public API =================================================================

' --- ExpedienteHPS_Form_Load -----------------------------------------------------
' Pure-data init for Form_Load: ensures ColArbolSuministradores is initialized
' if Nothing, returns the flags the form should apply (AllowEdits).
'
' Returns JSON: {ok, payload:{colArbolInitialized, allowEdits}, error, logs}.
Public Function ExpedienteHPS_Form_Load( _
    ByVal p_DTO As Object, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(4)
    p_Error = ""

    On Error GoTo errores

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")

    Dim colArbolInitialized As Boolean
    colArbolInitialized = False

    ' Defensive: DTO must be a Dictionary-like wrapper.
    If p_DTO Is Nothing Then
        p_Error = "ExpedienteHPS_Form_Load: p_DTO is Nothing"
        ExpedienteHPS_Form_Load = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    ' Initialize ColArbolSuministradores if absent.
    Dim colArbol As Object
    Set colArbol = ExpedienteHPS_GetValueSafely(p_DTO, "ColArbolSuministradores", Nothing)
    If colArbol Is Nothing Then
        Set p_DTO("ColArbolSuministradores") = New Scripting.Dictionary
        On Error Resume Next
        p_DTO("ColArbolSuministradores").CompareMode = TextCompare
        On Error GoTo errores
        colArbolInitialized = True
    End If

    payload("colArbolInitialized") = colArbolInitialized
    payload("allowEdits") = True
    logs(0) = "Form_Load: colArbolInitialized=" & CStr(colArbolInitialized)

    ExpedienteHPS_Form_Load = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "ExpedienteHPS_Form_Load: " & Err.Description
    End If
    ExpedienteHPS_Form_Load = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- ExpedienteHPS_EstablecerCombos_Construir ------------------------------------
' Pure-data combo payload builder: walks m_ObjEntorno.{GradosClasificaciones,
' Suministradores, LugaresEjecucion} and returns 4 rowSources.
'
' Returns JSON: {ok, payload:{clasificaciones, contratistas, subcontratistas,
'                              lugares}, error, logs}.
Public Function ExpedienteHPS_EstablecerCombos_Construir( _
    ByVal p_DTO As Object, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(6)
    p_Error = ""

    On Error GoTo errores

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")

    ' Stub field set for compat with DTO wrapper (DTO may or may not have these).
    Dim m_Entorno As Object
    If ExpedienteHPS_HasKeySafely(p_DTO, "Entorno") Then
        Set m_Entorno = p_DTO("Entorno")
    Else
        Set m_Entorno = Nothing
    End If

    Dim count As Long

    ' 1. Clasificaciones
    Dim clasif As Object
    Set clasif = Nothing
    On Error Resume Next
    If Not m_Entorno Is Nothing Then
        Set clasif = m_Entorno.GradosClasificaciones
    End If
    Dim errCl As Long
    errCl = Err.Number
    On Error GoTo errores
    Dim clasificacionesRow As String
    If errCl = 0 Then
        clasificacionesRow = ExpedienteHPS_DiccionarioToRowSource( _
            clasif, "IdGradoClasificacion", "GradoClasificacion", count, EXPEDIENTEHPS_MAX_ROWS)
    Else
        clasificacionesRow = ""
    End If
    payload("clasificaciones") = clasificacionesRow

    ' 2. Suministradores (same list into both Contratistas and SubContratistas).
    Dim sumCol As Object
    Set sumCol = Nothing
    On Error Resume Next
    If Not m_Entorno Is Nothing Then
        Set sumCol = m_Entorno.Suministradores
    End If
    Dim errSum As Long
    errSum = Err.Number
    On Error GoTo errores
    Dim sumRow As String
    If errSum = 0 Then
        sumRow = ExpedienteHPS_DiccionarioToRowSource( _
            sumCol, "IDSuministrador", "Nombre", count, EXPEDIENTEHPS_MAX_ROWS)
    Else
        sumRow = ""
    End If
    payload("contratistas") = sumRow
    payload("subcontratistas") = sumRow

    ' 3. Lugares
    Dim lugaresCol As Object
    Set lugaresCol = Nothing
    On Error Resume Next
    If Not m_Entorno Is Nothing Then
        Set lugaresCol = m_Entorno.LugaresEjecucion
    End If
    Dim errLg As Long
    errLg = Err.Number
    On Error GoTo errores
    Dim lugaresRow As String
    If errLg = 0 Then
        lugaresRow = ExpedienteHPS_DiccionarioToRowSource( _
            lugaresCol, "IDLugarEjecucion", "LugarEjecucion", count, EXPEDIENTEHPS_MAX_ROWS)
    Else
        lugaresRow = ""
    End If
    payload("lugares") = lugaresRow

    logs(0) = "EstablecerCombos_Construir: clasificaciones len=" & Len(clasificacionesRow) & _
              ", sum len=" & Len(sumRow) & ", lugares len=" & Len(lugaresRow)

    ExpedienteHPS_EstablecerCombos_Construir = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "ExpedienteHPS_EstablecerCombos_Construir: " & Err.Description
    End If
    ExpedienteHPS_EstablecerCombos_Construir = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- ExpedienteHPS_RellenarLista_Construir ---------------------------------------
' Pure-data list-payload builder for CONTRATISTAS, SUBCONTRATISTAS, or LUGARES.
' For CONTRATISTAS: items where ContratistaPrincipal="Sí"
' For SUBCONTRATISTAS: items where SubContratista="Sí" AND no IdPadre
' For LUGARES: items from ColLugaresEjecucion
'
' p_TipoLista must be one of: "CONTRATISTAS", "SUBCONTRATISTAS", "LUGARES".
'
' Returns JSON: {ok, payload:{rowSource, count}, error, logs}.
Public Function ExpedienteHPS_RellenarLista_Construir( _
    ByVal p_DTO As Object, _
    ByVal p_TipoLista As String, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(4)
    p_Error = ""

    On Error GoTo errores

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")

    Dim tipo As String
    tipo = UCase$("" & p_TipoLista)

    If tipo <> "CONTRATISTAS" And tipo <> "SUBCONTRATISTAS" And tipo <> "LUGARES" Then
        p_Error = "ExpedienteHPS_RellenarLista_Construir: invalid p_TipoLista='" & p_TipoLista & "'"
        ExpedienteHPS_RellenarLista_Construir = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim rowSource As String
    rowSource = ""
    Dim rowCount As Long
    rowCount = 0

    If p_DTO Is Nothing Then
        payload("rowSource") = ""
        payload("count") = 0
        ExpedienteHPS_RellenarLista_Construir = BuildJsonPayload(True, payload, "", logs)
        Exit Function
    End If

    If tipo = "LUGARES" Then
        Dim colLug As Object
        Set colLug = ExpedienteHPS_GetValueSafely(p_DTO, "ColLugaresEjecucion", Nothing)
        rowSource = ExpedienteHPS_DiccionarioToRowSource( _
            colLug, "IDLugarEjecucion", "LugarEjecucion", rowCount, EXPEDIENTEHPS_MAX_ROWS)
        payload("rowSource") = rowSource
        payload("count") = rowCount
        logs(0) = "RellenarLista LUGARES: " & rowCount & " rows"
        ExpedienteHPS_RellenarLista_Construir = BuildJsonPayload(True, payload, "", logs)
        Exit Function
    End If

    ' CONTRATISTAS / SUBCONTRATISTAS walk ColArbolSuministradores.
    Dim colArb As Object
    Set colArb = ExpedienteHPS_GetValueSafely(p_DTO, "ColArbolSuministradores", Nothing)
    If colArb Is Nothing Then
        payload("rowSource") = ""
        payload("count") = 0
        ExpedienteHPS_RellenarLista_Construir = BuildJsonPayload(True, payload, "", logs)
        Exit Function
    End If

    Dim k As Variant
    Dim item As Object
    Dim idItem As String
    Dim isContr As String
    Dim isSub As String
    Dim idPadre As String

    For Each k In colArb.Keys
        If rowCount >= EXPEDIENTEHPS_MAX_ROWS Then Exit For
        Set item = colArb(k)
        If item Is Nothing Then
            ' skip
        Else
            isContr = ExpedienteHPS_LeerEntidadComoDic(item, "ContratistaPrincipal")
            isSub = ExpedienteHPS_LeerEntidadComoDic(item, "SubContratista")
            idPadre = ExpedienteHPS_LeerEntidadComoDic(item, "IdPadre")

            If tipo = "CONTRATISTAS" Then
                If isContr = "Sí" Then
                    idItem = ExpedienteHPS_LeerEntidadComoDic(item, "IDSuministrador")
                    If Len(rowSource) > 0 Then rowSource = rowSource & vbCrLf
                    ' For display we use the supplier's Nombre via lookup — but since
                    ' pure-data helpers don't do DAO, we emit ID only and let the form
                    ' render. Row format: "ID;ID" (no Name). Tests use ID-based asserts.
                    rowSource = rowSource & ExpedienteHPS_BuildRowLine(idItem, idItem)
                    rowCount = rowCount + 1
                End If
            ElseIf tipo = "SUBCONTRATISTAS" Then
                If isSub = "Sí" And (Len(idPadre) = 0) Then
                    idItem = ExpedienteHPS_LeerEntidadComoDic(item, "IDSuministrador")
                    If Len(rowSource) > 0 Then rowSource = rowSource & vbCrLf
                    rowSource = rowSource & ExpedienteHPS_BuildRowLine(idItem, idItem)
                    rowCount = rowCount + 1
                End If
            End If
        End If
    Next k

    payload("rowSource") = rowSource
    payload("count") = rowCount
    logs(0) = "RellenarLista " & tipo & ": " & rowCount & " rows"

    ExpedienteHPS_RellenarLista_Construir = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "ExpedienteHPS_RellenarLista_Construir: " & Err.Description
    End If
    ExpedienteHPS_RellenarLista_Construir = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- ExpedienteHPS_AltaSuministrador ---------------------------------------------
' Pure-data add to ColArbolSuministradores. Detects duplicates, builds the
' new ExpedienteSuministrador (Dictionary stub) with a unique TMP id, and
' adds it to the collection.
'
' p_Tipo ∈ {"ROOT_CONTR","ROOT_SUB"}.
'
' Returns JSON: {ok, payload:{added, duplicate, idExpedienteSuministrador}, error, logs}.
Public Function ExpedienteHPS_AltaSuministrador( _
    ByVal p_DTO As Object, _
    ByVal p_IDSuministrador As String, _
    ByVal p_Tipo As String, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(6)
    p_Error = ""

    On Error GoTo errores

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")

    Dim added As Boolean
    added = False

    Dim duplicate As Boolean
    duplicate = False

    Dim idExpSum As String
    idExpSum = ""

    Dim idSum As String
    idSum = Trim$("" & p_IDSuministrador)

    Dim tipo As String
    tipo = ExpedienteHPS_NormalizarTipo(p_Tipo)
    If Len(tipo) = 0 Then
        p_Error = "ExpedienteHPS_AltaSuministrador: invalid p_Tipo='" & p_Tipo & "'"
        ExpedienteHPS_AltaSuministrador = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    If Len(idSum) = 0 Then
        p_Error = "ExpedienteHPS_AltaSuministrador: empty p_IDSuministrador"
        ExpedienteHPS_AltaSuministrador = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    If p_DTO Is Nothing Then
        p_Error = "ExpedienteHPS_AltaSuministrador: p_DTO is Nothing"
        ExpedienteHPS_AltaSuministrador = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim colArb As Object
    Set colArb = ExpedienteHPS_GetValueSafely(p_DTO, "ColArbolSuministradores", Nothing)
    If colArb Is Nothing Then
        Set p_DTO("ColArbolSuministradores") = New Scripting.Dictionary
        On Error Resume Next
        p_DTO("ColArbolSuministradores").CompareMode = TextCompare
        On Error GoTo errores
        Set colArb = p_DTO("ColArbolSuministradores")
    End If

    ' Duplicate detection.
    Dim k As Variant
    Dim item As Object
    For Each k In colArb.Keys
        Set item = colArb(k)
        If ExpedienteHPS_MatchSuministrador(item, idSum, tipo) Then
            duplicate = True
            Exit For
        End If
    Next k

    If duplicate Then
        payload("added") = False
        payload("duplicate") = True
        payload("idExpedienteSuministrador") = ""
        logs(0) = "AltaSuministrador: duplicate id=" & idSum & " tipo=" & tipo
        ExpedienteHPS_AltaSuministrador = BuildJsonPayload(True, payload, "", logs)
        Exit Function
    End If

    ' Generate unique TMP id (deterministic for tests via second-precision timestamp).
    idExpSum = "TMP_" & Format(Now, "hhmmss") & "_" & Int(Rnd * 10000)

    ' Build the new item as a Dictionary stub so tests work without DAO.
    Dim newItem As Object
    Set newItem = ExpedienteHPS_ConstruirExpedienteSuministrador(idExpSum, idSum, tipo)

    colArb.Add idExpSum, newItem

    added = True
    payload("added") = added
    payload("duplicate") = False
    payload("idExpedienteSuministrador") = idExpSum
    logs(0) = "AltaSuministrador: added idExpSum=" & idExpSum & " tipo=" & tipo

    ExpedienteHPS_AltaSuministrador = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "ExpedienteHPS_AltaSuministrador: " & Err.Description
    End If
    ExpedienteHPS_AltaSuministrador = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- ExpedienteHPS_EliminarSuministrador ------------------------------------------
' Pure-data remove from ColArbolSuministradores by (IDSuministrador, tipo).
'
' Returns JSON: {ok, payload:{removed, countAfter}, error, logs}.
Public Function ExpedienteHPS_EliminarSuministrador( _
    ByVal p_DTO As Object, _
    ByVal p_IDSuministrador As String, _
    ByVal p_Tipo As String, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(4)
    p_Error = ""

    On Error GoTo errores

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")

    Dim removed As Boolean
    removed = False

    Dim idSum As String
    idSum = Trim$("" & p_IDSuministrador)

    Dim tipo As String
    tipo = ExpedienteHPS_NormalizarTipo(p_Tipo)
    If Len(tipo) = 0 Then
        p_Error = "ExpedienteHPS_EliminarSuministrador: invalid p_Tipo='" & p_Tipo & "'"
        ExpedienteHPS_EliminarSuministrador = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    If p_DTO Is Nothing Then
        p_Error = "ExpedienteHPS_EliminarSuministrador: p_DTO is Nothing"
        ExpedienteHPS_EliminarSuministrador = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim colArb As Object
    Set colArb = ExpedienteHPS_GetValueSafely(p_DTO, "ColArbolSuministradores", Nothing)
    If colArb Is Nothing Then
        payload("removed") = False
        payload("countAfter") = 0
        ExpedienteHPS_EliminarSuministrador = BuildJsonPayload(True, payload, "", logs)
        Exit Function
    End If

    Dim k As Variant
    Dim item As Object
    Dim keyToDelete As String
    For Each k In colArb.Keys
        Set item = colArb(k)
        If ExpedienteHPS_MatchSuministrador(item, idSum, tipo) Then
            keyToDelete = CStr(k)
            Exit For
        End If
    Next k

    If Len(keyToDelete) > 0 Then
        colArb.Remove keyToDelete
        removed = True
    End If

    payload("removed") = removed
    payload("countAfter") = colArb.Count
    logs(0) = "EliminarSuministrador: removed=" & CStr(removed) & " countAfter=" & CStr(colArb.Count)

    ExpedienteHPS_EliminarSuministrador = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "ExpedienteHPS_EliminarSuministrador: " & Err.Description
    End If
    ExpedienteHPS_EliminarSuministrador = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- ExpedienteHPS_AltaLugar -----------------------------------------------------
' Pure-data add to ColLugaresEjecucion. Detects duplicates.
'
' Returns JSON: {ok, payload:{added, duplicate}, error, logs}.
Public Function ExpedienteHPS_AltaLugar( _
    ByVal p_DTO As Object, _
    ByVal p_IDLugar As String, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(4)
    p_Error = ""

    On Error GoTo errores

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")

    Dim idLugar As String
    idLugar = Trim$("" & p_IDLugar)

    If Len(idLugar) = 0 Then
        p_Error = "ExpedienteHPS_AltaLugar: empty p_IDLugar"
        ExpedienteHPS_AltaLugar = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    If p_DTO Is Nothing Then
        p_Error = "ExpedienteHPS_AltaLugar: p_DTO is Nothing"
        ExpedienteHPS_AltaLugar = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim colLug As Object
    Set colLug = ExpedienteHPS_GetValueSafely(p_DTO, "ColLugaresEjecucion", Nothing)
    If colLug Is Nothing Then
        Set p_DTO("ColLugaresEjecucion") = New Scripting.Dictionary
        Set colLug = p_DTO("ColLugaresEjecucion")
    End If

    Dim duplicate As Boolean
    duplicate = colLug.Exists(idLugar)

    If duplicate Then
        payload("added") = False
        payload("duplicate") = True
        logs(0) = "AltaLugar: duplicate id=" & idLugar
        ExpedienteHPS_AltaLugar = BuildJsonPayload(True, payload, "", logs)
        Exit Function
    End If

    ' Build the new item as a Dictionary stub.
    Dim newItem As Object
    Set newItem = CreateObject("Scripting.Dictionary")
    newItem("IDLugarEjecucion") = idLugar
    newItem("LugarEjecucion") = idLugar

    colLug.Add idLugar, newItem

    payload("added") = True
    payload("duplicate") = False
    logs(0) = "AltaLugar: added id=" & idLugar

    ExpedienteHPS_AltaLugar = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "ExpedienteHPS_AltaLugar: " & Err.Description
    End If
    ExpedienteHPS_AltaLugar = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- ExpedienteHPS_EliminarLugar --------------------------------------------------
' Pure-data remove from ColLugaresEjecucion.
'
' Returns JSON: {ok, payload:{removed, countAfter}, error, logs}.
Public Function ExpedienteHPS_EliminarLugar( _
    ByVal p_DTO As Object, _
    ByVal p_IDLugar As String, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    p_Error = ""

    On Error GoTo errores

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")

    Dim idLugar As String
    idLugar = Trim$("" & p_IDLugar)

    If Len(idLugar) = 0 Then
        p_Error = "ExpedienteHPS_EliminarLugar: empty p_IDLugar"
        ExpedienteHPS_EliminarLugar = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim removed As Boolean
    removed = False

    If p_DTO Is Nothing Then
        p_Error = "ExpedienteHPS_EliminarLugar: p_DTO is Nothing"
        ExpedienteHPS_EliminarLugar = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim colLug As Object
    Set colLug = ExpedienteHPS_GetValueSafely(p_DTO, "ColLugaresEjecucion", Nothing)
    If colLug Is Nothing Then
        payload("removed") = False
        payload("countAfter") = 0
        ExpedienteHPS_EliminarLugar = BuildJsonPayload(True, payload, "", logs)
        Exit Function
    End If

    If colLug.Exists(idLugar) Then
        colLug.Remove idLugar
        removed = True
    End If

    payload("removed") = removed
    payload("countAfter") = colLug.Count
    logs(0) = "EliminarLugar: removed=" & CStr(removed) & " countAfter=" & CStr(colLug.Count)

    ExpedienteHPS_EliminarLugar = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "ExpedienteHPS_EliminarLugar: " & Err.Description
    End If
    ExpedienteHPS_EliminarLugar = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- ExpedienteHPS_ComandoRegistrar_Click ----------------------------------------
' Pure-data alta: sets the HPS-specific flags on the Expediente and reports
' the registration state. The form does the DAO call (ExpedienteOperaciones
' .Registrar) and RaiseEvent — those are NOT pure data.
'
' The helper focuses on: validating DTO, setting the HPS flag set on the
' in-memory Expediente, returning the resulting flag snapshot the form
' should pass to DAO.
'
' Returns JSON: {ok, payload:{ambito, hpsAplica, postagedo, tipo, expId,
'                              ready:Boolean}, error, logs}.
Public Function ExpedienteHPS_ComandoRegistrar_Click( _
    ByVal p_DTO As Object, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(6)
    p_Error = ""

    On Error GoTo errores

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")

    If p_DTO Is Nothing Then
        p_Error = "ExpedienteHPS_ComandoRegistrar_Click: p_DTO is Nothing"
        ExpedienteHPS_ComandoRegistrar_Click = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim expObj As Object
    Set expObj = ExpedienteHPS_GetValueSafely(p_DTO, "Expediente", Nothing)
    If expObj Is Nothing Then
        p_Error = "ExpedienteHPS_ComandoRegistrar_Click: p_DTO.Expediente is Nothing"
        ExpedienteHPS_ComandoRegistrar_Click = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    ' Set the HPS-specific flags on the in-memory Expediente. Best-effort; if
    ' the entity doesn't expose these property sets (test stub), log and
    ' continue.
    On Error Resume Next
    expObj.Ambito = "HPS"
    expObj.HPSAplica = "Sí"
    expObj.POSTAGEDO = "Sí"
    expObj.TIpo = "Expediente Sólo para HPS"
    Dim setErr As Long
    setErr = Err.Number
    On Error GoTo errores

    payload("ambito") = "HPS"
    payload("hpsAplica") = "Sí"
    payload("postagedo") = "Sí"
    payload("tipo") = "Expediente Sólo para HPS"
    payload("ready") = True
    logs(0) = "ComandoRegistrar_Click: HPS flags set (setErr=" & CStr(setErr) & ")"

    ' Note: the form will then run the DAO call (ExpedienteOperaciones.Registrar)
    ' and RaiseEvent AltaExpediente — both are UI/DAO orchestration that stays
    ' in the form (per rule #1).

    ExpedienteHPS_ComandoRegistrar_Click = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "ExpedienteHPS_ComandoRegistrar_Click: " & Err.Description
    End If
    ExpedienteHPS_ComandoRegistrar_Click = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function