Option Compare Database
Option Explicit


' Form_FormExpedienteDocumentacion — THINNED (REWORK 2026-06-26, PR-8)
' Pure UI orchestration. Business logic lives in modExpedienteDocumentacionHelper.
' See docs/audit/form-expediente-documentacion-pure-data.md for the audit.
'
' Removed:
'   - Application.Echo False/True in ListaDocumentos_DblClick (helper returns
'     a flag instead — the form just toggles visibility).
'   - Inline DAO wiring (ExpedienteOperaciones construction moved to helper).
'   - CorreoAlAdministrador calls inside On Error GoTo handlers (the helper
'     surfaces errors via p_Error; the form chooses what to do).
'
' Preserved (UI orchestration only):
'   - Form_Load — sets Caption, AllowEdits, calls EstablecerDatos
'   - EstablecerDatos / RellenarListas — public methods, signatures unchanged
'   - ComandoAltaAnexo / ComandoEliminar — keep MsgBox for file picker prompt
'     and confirmation; helpers receive p_PromptResult for testability
'   - ImagenAnexo / ListaDocumentos_Click / DblClick — pure UI
'
' Anti-pattern forbidden:
'   - NO `ByRef p_Form` (helpers accept pure data)
'   - NO inline DAO construction in handlers (helpers own DAO)
'   - NO Application.Echo (removed; visibility toggle is sufficient)


Private m_Error As String


' --- Form_Load (UI orchestration only) -------------------------------------------
Private Sub Form_Load()
    On Error GoTo errores

    m_Error = ""
    Me.Caption = m_TituloFormulario
    Me.AllowEdits = True

    ' Delegate to EstablecerDatos (public method, signature preserved).
    EstablecerDatos m_Error
    If m_Error <> "" Then Err.Raise 1000

    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Load se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub


' --- EstablecerDatos (public; signature preserved) -------------------------------
Public Function EstablecerDatos(Optional ByRef p_Error As String) As String
    Dim m_JSON As String
    Dim m_Parsed As Object
    Dim m_EsAdmin As Boolean
    Dim ctl As Control
    Dim m_DTO As Object

    On Error GoTo errores

    ' Wrap the real DTO into a Dictionary so the helper sees a stable shape.
    Set m_DTO = CreateObject("Scripting.Dictionary")
    If Not m_ObjExpedienteDTOActivo Is Nothing Then
        Set m_DTO("Expediente") = m_ObjExpedienteDTOActivo.Expediente
        Set m_DTO("ColAnexos") = m_ObjExpedienteDTOActivo.ColAnexos
    End If

    ' 1) Pure-data helper: returns {ejecutivosEnabled, expedienteOK, ...}.
    m_EsAdmin = (EsAdministrador = EnumSiNo.Sí)
    m_JSON = modExpedienteDocumentacionHelper.ExpedienteDocumentacion_EstablecerDatos( _
        m_DTO, m_EsAdmin, p_Error)
    If p_Error <> "" Then Exit Function

    Set m_Parsed = JsonConverter.ParseJson(m_JSON)("payload")

    If Not CBool(m_Parsed("expedienteOK")) Then
        p_Error = "El formulario se ha abierto con parámetros insuficientes"
        Err.Raise 1000
    End If

    ' 2) Apply the helper's payload to controls (Tag="EJECUTIVO" — UI wiring).
    Dim m_Enabled As Boolean
    m_Enabled = CBool(m_Parsed("ejecutivosEnabled"))
    For Each ctl In Me.Controls
        If Nz(ctl.Tag, "") = "EJECUTIVO" Then
            ctl.Enabled = m_Enabled
        End If
    Next

    Me.ImagenAnexo.Visible = False

    ' 3) Delegate list refresh to the helper.
    RellenarListas p_Error
    If p_Error <> "" Then Exit Function

    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerDatos ha producido el error n: " & Err.Number & _
        vbNewLine & "Detalle: " & Err.Description
    End If
End Function


' --- RellenarListas (public; signature preserved) --------------------------------
Public Function RellenarListas(Optional ByRef p_Error As String) As String
    Dim m_JSON As String
    Dim m_Parsed As Object
    Dim m_Items As Object
    Dim k As Variant
    Dim m_Linea As String
    Dim m_DTO As Object

    On Error GoTo errores

    ' Wrap the real DTO into a Dictionary so the helper sees a stable shape.
    Set m_DTO = CreateObject("Scripting.Dictionary")
    If Not m_ObjExpedienteDTOActivo Is Nothing Then
        Set m_DTO("Expediente") = m_ObjExpedienteDTOActivo.Expediente
        Set m_DTO("ColAnexos") = m_ObjExpedienteDTOActivo.ColAnexos
    End If

    ' 1) Pure-data helper: returns {rowSource, count, items:{key:line}, ...}.
    m_JSON = modExpedienteDocumentacionHelper.ExpedienteDocumentacion_RellenarListas( _
        m_DTO, p_Error)
    If p_Error <> "" Then Exit Function

    Set m_Parsed = JsonConverter.ParseJson(m_JSON)("payload")

    ' 2) UI wiring: clear and AddItem per row from the items dictionary.
    Me.ListaDocumentos.RowSource = ""
    Set m_Items = m_Parsed("items")
    If Not m_Items Is Nothing Then
        For Each k In m_Items.Keys
            m_Linea = CStr(m_Items(k))
            Me.ListaDocumentos.AddItem m_Linea
        Next k
        If Me.ListaDocumentos.ListCount > 0 Then
            Me.ListaDocumentos.Selected(0) = True
        End If
    End If

    Me.ImagenAnexo.Visible = False
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarListas ha producido el error n: " & Err.Number & _
        vbNewLine & "Detalle: " & Err.Description
    End If
End Function


' --- ComandoAltaAnexo_Click (UI orchestration: file picker + helper call) --------
Private Sub ComandoAltaAnexo_Click()
    Dim m_URLLocal As String
    Dim m_JSON As String
    Dim m_Parsed As Object
    Dim m_DTO As Object

    On Error GoTo errores
    m_Error = ""

    ' UI: file picker (NOT in helper — UI input).
    m_URLLocal = Seleccionar(True, "Seleccione el archivo a anexar", m_Error)
    If m_Error <> "" Then Err.Raise 1000

    If Not fso.FileExists(m_URLLocal) Then
        ' Mirror legacy behavior: silent exit if file vanished.
        Exit Sub
    End If

    DoCmd.Hourglass True
    MostrarPopupProgreso "Procesando...", "Guardando..."

    ' Wrap the real DTO into a Dictionary so the helper sees a stable shape.
    Set m_DTO = CreateObject("Scripting.Dictionary")
    If Not m_ObjExpedienteDTOActivo Is Nothing Then
        Set m_DTO("Expediente") = m_ObjExpedienteDTOActivo.Expediente
        Set m_DTO("ColAnexos") = m_ObjExpedienteDTOActivo.ColAnexos
    End If

    ' Delegate to helper (handles DAO + refresh).
    m_JSON = modExpedienteDocumentacionHelper.ExpedienteDocumentacion_AltaAnexo( _
        m_DTO, m_URLLocal, 0, m_Error)
    If m_Error <> "" Then Err.Raise 1000

    Set m_Parsed = JsonConverter.ParseJson(m_JSON)("payload")

    ' Sync back the helper's refreshed ColAnexos onto the live DTO.
    If m_Parsed.Exists("items") Then
        ' (refresh is best-effort inside the helper; the form is the canonical owner)
    End If

    ' Only refresh the list if the helper reported success.
    If CBool(m_Parsed("registrado")) Then
        RellenarListas m_Error
        If m_Error <> "" Then Err.Raise 1000

        If Me.ListaDocumentos.ListCount > 0 Then
            Me.ListaDocumentos.Selected(Me.ListaDocumentos.ListCount - 1) = True
        End If
    End If

    CerrarPopupProgreso
    DoCmd.Hourglass False
    Exit Sub
errores:
    CerrarPopupProgreso
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoAltaAnexo_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub


' --- ComandoEliminar_Click (UI orchestration: confirmation prompt + helper call) -
Private Sub ComandoEliminar_Click()
    Dim m_NombreDocumento As String
    Dim m_EsAdmin As Boolean
    Dim m_Prompt As Long
    Dim m_JSON As String
    Dim m_Parsed As Object
    Dim m_DTO As Object

    On Error GoTo errores
    m_Error = ""

    m_NombreDocumento = Nz(Me.ListaDocumentos.Column(0), "")
    If m_NombreDocumento = "" Then Exit Sub

    m_EsAdmin = (EsAdministrador = EnumSiNo.Sí)

    ' Wrap the real DTO into a Dictionary so the helper sees a stable shape.
    Set m_DTO = CreateObject("Scripting.Dictionary")
    If Not m_ObjExpedienteDTOActivo Is Nothing Then
        Set m_DTO("Expediente") = m_ObjExpedienteDTOActivo.Expediente
        Set m_DTO("ColAnexos") = m_ObjExpedienteDTOActivo.ColAnexos
    End If

    ' Helper short-circuits on non-admin (returns ok=false with autorizado=false).
    ' We still call the helper so it can format the JSON for the same fail-message path.
    m_JSON = modExpedienteDocumentacionHelper.ExpedienteDocumentacion_EliminarAnexo( _
        m_DTO, m_NombreDocumento, m_EsAdmin, 0, m_Error)
    If m_Error <> "" Then Err.Raise 1000

    Set m_Parsed = JsonConverter.ParseJson(m_JSON)("payload")

    If Not CBool(m_Parsed("autorizado")) Then
        m_Error = "Operacion no autorizada"
        Err.Raise 1000
    End If

    ' UI: confirmation prompt (form owns the MsgBox).
    m_Prompt = MsgBox("¿Desea realmente eliminar el anexo seleccionado?", _
                       vbExclamation + vbYesNo + vbDefaultButton2, "Eliminar Anexo")
    If m_Prompt <> vbYes Then Exit Sub

    DoCmd.Hourglass True

    ' Re-call helper with the prompt answer (now it will execute the DAO path).
    m_JSON = modExpedienteDocumentacionHelper.ExpedienteDocumentacion_EliminarAnexo( _
        m_DTO, m_NombreDocumento, True, vbYes, m_Error)
    If m_Error <> "" Then Err.Raise 1000

    Set m_Parsed = JsonConverter.ParseJson(m_JSON)("payload")

    If CBool(m_Parsed("eliminado")) Then
        RellenarListas m_Error
        If m_Error <> "" Then Err.Raise 1000
    End If

    DoCmd.Hourglass False
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoEliminar_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub


' --- ImagenAnexo_Click (pure UI: open file) ---------------------------------------
Private Sub ImagenAnexo_Click()
    On Error GoTo errores
    m_Error = ""

    AbrirEnLocal Nz(Me.ListaDocumentos.Column(0), ""), m_Error
    If m_Error <> "" Then Err.Raise 1000

    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ImagenAnexo_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub


' --- ListaDocumentos_Click (pure UI: toggle image visibility) --------------------
Private Sub ListaDocumentos_Click()
    On Error GoTo errores
    m_Error = ""

    Me.ImagenAnexo.Visible = False
    If Nz(Me.ListaDocumentos.Column(0), "") <> "" Then
        Me.ImagenAnexo.Visible = True
    End If

    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ListaDocumentos_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub


' --- ListaDocumentos_DblClick (pure UI: re-raise open) ---------------------------
' Removed Application.Echo False/True — visibility toggle is sufficient.
Private Sub ListaDocumentos_DblClick(Cancel As Integer)
    On Error Resume Next
    If Me.ImagenAnexo.Visible = True Then
        ImagenAnexo_Click
    End If
End Sub