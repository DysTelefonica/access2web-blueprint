Option Compare Database
Option Explicit

' Form_FormLugarEjecucionGestion — REWORK (2026-06-26)
' Thin UI wiring for the LugarEjecucion Gestion list.
'
' Implements access-vba-e2e-methodology rule #1: forms are thin UI wiring.
' Every handler does exactly: read controls -> call helper -> render result.
' No business logic lives here.
'
' Anti-pattern removed (was in PR #29 commit 979a57a):
'   - helpers received `ByRef p_Form As Object` and read controls via p_Form.Controls(...)
'   - tests opened the real form and asserted against p_Form.Controls(...).Value
'   - test fixture `Test_LugarEjecucionHelper_OpenForm(p_Form, p_Error)` blocked COM
'
' New contract (per the rework):
'   - All helpers in modLugarEjecucionHelper accept PURE DATA (no Form refs).
'   - Form extracts control values BEFORE calling helper; renders helper JSON AFTER.
'   - Helpers receive: primitives (Boolean, String), stub Dictionaries, stub entities.
'   - UI orchestration stays in this file: DoCmd.OpenForm, FormInteraction_FormularioAbierto, Forms(...).
'   - Five helper calls back the form: Abrir_Inicializar, Buscar_Listar, Seleccionar_Cargar,
'     Eliminar_Borrar, DobleClick_AbrirEdicion.
'   - Three handlers do pure UI orchestration (no helper): Alta, Edicion, Limpiar.

' === Module-level state (all declarations at top per vba-access §10.1) ==============

' The currently selected entity (loaded by Seleccionar_Cargar / edited by Edicion).
' Module-level so the helpers' serialized output can be cached locally if needed,
' but the form itself owns the lifecycle (set, render, clear).
Private m_LugarEjecucionSeleccionada As Object

' The currently open FormLugarEjecucion child form (when Alta / Edicion is active).
' Kept as Object to avoid cross-form compile-time dependency.
Private WithEvents m_FormLugarEjecucion As Form_FormLugarEjecucion

' Error accumulator for handler traps (Telefónica D&S convention).
Private m_Error As String

' Public events the parent form listens to.
Public Event Eliminar()
Public Event Seleccionar(m_LugarEjecucionSeleccionada As Object)


' === Local helpers (all at top per vba-access §10.1) =================================

' --- LugarEjecucion_ExtraerFiltro -----------------------------------------------------
' Reads the filter text from Me.LUGAREJECUCION (control name is UPPERCASE).
' Returns "" if Null. Pure UI read — used by ComandoBuscar_Click before calling the helper.
Private Function LugarEjecucion_ExtraerFiltro() As String
    Dim v As Variant
    v = Nz(Me.LUGAREJECUCION.Value, "")
    LugarEjecucion_ExtraerFiltro = CStr(v)
End Function

' --- LugarEjecucion_ExtraerIDSeleccionado -------------------------------------------
' Reads the currently selected ID from ListaFiltrados.Column(0).
' Returns "" if no selection (Null column).
' Pure UI read — used by ListaFiltrados_Click before calling the helper.
Private Function LugarEjecucion_ExtraerIDSeleccionado() As String
    Dim v As Variant
    On Error Resume Next
    v = Me.ListaFiltrados.Column(0)
    On Error GoTo 0
    If IsNull(v) Then
        LugarEjecucion_ExtraerIDSeleccionado = ""
    Else
        LugarEjecucion_ExtraerIDSeleccionado = CStr(v)
    End If
End Function

' --- LugarEjecucion_ExtraerLugares --------------------------------------------------
' Reads the in-memory collection of LugarEjecucion entities from m_ObjEntorno.
' Returns a Dictionary (possibly empty), or a fresh empty Dictionary if the
' collection is Nothing (defensive — never pass Nothing to the helper by accident).
Private Function LugarEjecucion_ExtraerLugares() As Object
    Dim col As Object
    On Error Resume Next
    Set col = m_ObjEntorno.LugaresEjecucion
    On Error GoTo 0
    If col Is Nothing Then
        Set col = CreateObject("Scripting.Dictionary")
    End If
    Set LugarEjecucion_ExtraerLugares = col
End Function

' --- LugarEjecucion_EsAdministrador -------------------------------------------------
' Reads the global admin flag set during EVE().
Private Function LugarEjecucion_EsAdministrador() As Boolean
    If EsAdministrador = EnumSiNo.Sí Then
        LugarEjecucion_EsAdministrador = True
    Else
        LugarEjecucion_EsAdministrador = False
    End If
End Function

' --- LugarEjecucion_TieneOpenArgs ---------------------------------------------------
' Returns True when the form was opened with a non-empty OpenArgs.
Private Function LugarEjecucion_TieneOpenArgs() As Boolean
    If Len(Nz(Me.OpenArgs, "")) > 0 Then
        LugarEjecucion_TieneOpenArgs = True
    Else
        LugarEjecucion_TieneOpenArgs = False
    End If
End Function

' --- LugarEjecucion_RenderResultadoInicializar --------------------------------------
' Renders the JSON payload from Abrir_Inicializar into the form's controls.
' Reads payload.showAlta, payload.focusAlta, payload.showElegir.
Private Sub LugarEjecucion_RenderResultadoInicializar(ByVal p_Payload As Object)
    If p_Payload Is Nothing Then Exit Sub

    Dim showAlta As Boolean
    showAlta = CBool(p_Payload("showAlta"))

    Dim focusAlta As Boolean
    focusAlta = CBool(p_Payload("focusAlta"))

    Dim showElegir As Boolean
    showElegir = CBool(p_Payload("showElegir"))

    FormInteraction_EstablecerPropiedad Me, "ComandoAlta", "Enabled", showAlta
    FormInteraction_EstablecerPropiedad Me, "ComandoEditar", "Enabled", False
    FormInteraction_EstablecerPropiedad Me, "ComandoEliminar", "Enabled", False
    FormInteraction_EstablecerPropiedad Me, "cmdElegir", "Visible", showElegir

    If focusAlta Then
        On Error Resume Next
        Me.Controls("ComandoAlta").SetFocus
        On Error GoTo 0
    End If
End Sub

' --- LugarEjecucion_RenderResultadoBuscar -------------------------------------------
' Renders the JSON payload from Buscar_Listar into ListaFiltrados.
' Reads payload.rowSource and payload.count.
' If count > 0, selects row 0 and re-fires ListaFiltrados_Click so the form's
' selection state (enableEditar/Eliminar) updates.
Private Sub LugarEjecucion_RenderResultadoBuscar(ByVal p_Payload As Object)
    If p_Payload Is Nothing Then Exit Sub

    Dim rowSource As String
    rowSource = CStr(p_Payload("rowSource"))

    Dim rowCount As Long
    rowCount = CLng(p_Payload("count"))

    If Len(rowSource) > 0 Then
        ' Build a full RowSource with header line + data lines.
        Me.ListaFiltrados.RowSource = "IDLugarEjecucion;LUGAR;DESCRIPCION" & vbCrLf & rowSource
    Else
        Me.ListaFiltrados.RowSource = "IDLugarEjecucion;LUGAR;DESCRIPCION"
    End If

    If rowCount > 0 Then
        On Error Resume Next
        Me.ListaFiltrados.Selected(0) = True
        On Error GoTo 0
        ' Re-fire selection handler so the form's button state syncs.
        ListaFiltrados_Click
    End If
End Sub

' --- LugarEjecucion_RenderResultadoSeleccionar --------------------------------------
' Renders the JSON payload from Seleccionar_Cargar into the form's controls.
' Reads payload.entity, payload.enableEditar, payload.enableEliminar.
' Also updates m_LugarEjecucionSeleccionada so the form's state is consistent
' across Edicion/Eliminar calls.
Private Sub LugarEjecucion_RenderResultadoSeleccionar(ByVal p_Payload As Object)
    If p_Payload Is Nothing Then Exit Sub

    Dim enableEditar As Boolean
    enableEditar = CBool(p_Payload("enableEditar"))

    Dim enableEliminar As Boolean
    enableEliminar = CBool(p_Payload("enableEliminar"))

    FormInteraction_EstablecerPropiedad Me, "ComandoEditar", "Enabled", enableEditar
    FormInteraction_EstablecerPropiedad Me, "ComandoEliminar", "Enabled", enableEliminar

    Dim entity As Object
    Set entity = p_Payload("entity")
    If entity Is Nothing Then
        Set m_LugarEjecucionSeleccionada = Nothing
    Else
        ' Wrap the serialized Dictionary into a real LugarEjecucion so subsequent
        ' Edicion calls can pass it to LugarEjecucionOperaciones via the helper.
        Dim wrapped As New LugarEjecucion
        wrapped.IDLugarEjecucion = CStr(entity("IDLugarEjecucion"))
        wrapped.LugarEjecucion = CStr(entity("LugarEjecucion"))
        wrapped.DESCRIPCION = CStr(entity("DESCRIPCION"))
        Set m_LugarEjecucionSeleccionada = wrapped
    End If
End Sub


' === Form lifecycle / event handlers ===================================================

' --- Form_Open ---------------------------------------------------------------------
' Lifecycle adapter: Ajustar + initialize + initial search.
' Helper: LugarEjecucion_Abrir_Inicializar (receives primitives).
' Then calls Buscar_Listar manually with the in-memory entity collection so
' the initial listbox is populated.
Private Sub Form_Open(Cancel As Integer)
    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents

    m_Error = ""

    ' Ajustar is a legacy UI-normalization helper; safe to leave in the form.
    Ajustar Me

    ' Capture primitives before calling the helper (per rule #1).
    Dim esAdmin As Boolean
    esAdmin = LugarEjecucion_EsAdministrador()

    Dim tieneOpenArgs As Boolean
    tieneOpenArgs = LugarEjecucion_TieneOpenArgs()

    Dim initJson As String
    initJson = modLugarEjecucionHelper.LugarEjecucion_Abrir_Inicializar( _
        esAdmin, tieneOpenArgs, m_Error)
    If m_Error <> "" Then Err.Raise 1000

    ' Render the payload into controls.
    Dim initParsed As Object
    Set initParsed = JsonConverter.ParseJson(initJson)
    LugarEjecucion_RenderResultadoInicializar initParsed("payload")

    ' Run initial search so the listbox is populated.
    Dim filterText As String
    filterText = LugarEjecucion_ExtraerFiltro()

    Dim col As Object
    Set col = LugarEjecucion_ExtraerLugares()

    Dim searchJson As String
    searchJson = modLugarEjecucionHelper.LugarEjecucion_Buscar_Listar(col, filterText, m_Error)
    If m_Error <> "" Then Err.Raise 1000

    Dim searchParsed As Object
    Set searchParsed = JsonConverter.ParseJson(searchJson)
    LugarEjecucion_RenderResultadoBuscar searchParsed("payload")

    AvanceCerrar
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub

errores:
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Open se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    On Error Resume Next
    CorreoAlAdministrador m_Error
    On Error GoTo 0
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub


' --- cmdElegir / cmdSalir: pure UI lifecycle adapters ----------------------------

Private Sub cmdElegir_Click()
    On Error GoTo errores
    m_Error = ""
    If m_LugarEjecucionSeleccionada Is Nothing Then
        Exit Sub
    End If
    RaiseEvent Seleccionar(m_LugarEjecucionSeleccionada)
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al cmdElegir_Click se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        On Error Resume Next
        CorreoAlAdministrador m_Error
        On Error GoTo 0
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub cmdSalir_Click()
    On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub


' --- ComandoAlta: pure UI orchestration (no helper) ------------------------------
' Per e2e rule #1: opening a form is UI orchestration, not testable business
' logic. We keep DoCmd.OpenForm / FormInteraction_FormularioAbierto / Forms(...) in the form.
Private Sub ComandoAlta_Click()
    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents

    m_Error = ""
    Set m_LugarEjecucionSeleccionada = Nothing
    Set m_FormLugarEjecucion = Nothing

    If FormInteraction_FormularioAbierto("FormLugarEjecucion") Then
        FormInteraction_CerrarFormulario "FormLugarEjecucion", acSaveNo
    End If

    DoCmd.OpenForm "FormLugarEjecucion"

    If FormInteraction_FormularioAbierto("FormLugarEjecucion") Then
        Set m_FormLugarEjecucion = Forms("FormLugarEjecucion")
    End If

    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoAlta_Click se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        On Error Resume Next
        CorreoAlAdministrador m_Error
        On Error GoTo 0
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub


' --- ComandoBuscar: extract filter -> helper -> render --------------------------
Private Sub ComandoBuscar_Click()
    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents

    m_Error = ""

    Dim filterText As String
    filterText = LugarEjecucion_ExtraerFiltro()

    Dim col As Object
    Set col = LugarEjecucion_ExtraerLugares()

    Dim json As String
    json = modLugarEjecucionHelper.LugarEjecucion_Buscar_Listar(col, filterText, m_Error)
    If m_Error <> "" Then Err.Raise 1000

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    LugarEjecucion_RenderResultadoBuscar parsed("payload")

    AvanceCerrar
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub

errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoBuscar_Click se ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    AvanceCerrar
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub


' --- ListaFiltrados_Click: extract ID -> helper -> render ------------------------
Private Sub ListaFiltrados_Click()
    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents

    m_Error = ""

    Dim selectedId As String
    selectedId = LugarEjecucion_ExtraerIDSeleccionado()

    Dim col As Object
    Set col = LugarEjecucion_ExtraerLugares()

    Dim esAdmin As Boolean
    esAdmin = LugarEjecucion_EsAdministrador()

    Dim json As String
    json = modLugarEjecucionHelper.LugarEjecucion_Seleccionar_Cargar( _
        selectedId, col, esAdmin, m_Error)
    If m_Error <> "" Then Err.Raise 1000

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    LugarEjecucion_RenderResultadoSeleccionar parsed("payload")

    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub

errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ListaFiltrados_Click se ha producido el error num " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        On Error Resume Next
        CorreoAlAdministrador m_Error
        On Error GoTo 0
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub


' --- ListaFiltrados_DblClick: read state -> helper dispatch ----------------------
Private Sub ListaFiltrados_DblClick(Cancel As Integer)
    On Error GoTo errores
    m_Error = ""

    Dim hasElegir As Boolean
    On Error Resume Next
    hasElegir = Me.Controls("cmdElegir").Visible
    On Error GoTo 0

    Dim editEnabled As Boolean
    On Error Resume Next
    editEnabled = Me.Controls("ComandoEditar").Enabled
    On Error GoTo 0

    Dim json As String
    json = modLugarEjecucionHelper.LugarEjecucion_DobleClick_AbrirEdicion( _
        hasElegir, editEnabled, m_Error)
    If m_Error <> "" Then Err.Raise 1000

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    Dim action As String
    action = CStr(parsed("payload")("action"))

    Select Case action
        Case "choose"
            cmdElegir_Click
        Case "edit"
            ComandoEditar_Click
        Case "none"
            ' No-op — helper said neither button is actionable.
    End Select
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ListaFiltrados_DblClick se ha producido el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        On Error Resume Next
        CorreoAlAdministrador m_Error
        On Error GoTo 0
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub


' --- ComandoEditar: pure UI orchestration (open the alta form in edit mode) -----
' Per e2e rule #1: opening a form is UI orchestration. The selection is already
' loaded by ListaFiltrados_Click into m_LugarEjecucionSeleccionada; we just
' stash it globally and open the alta form so it can edit it.
Private Sub ComandoEditar_Click()
    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents

    m_Error = ""
    Set m_ObjLugarEjecucionActiva = m_LugarEjecucionSeleccionada
    Set m_FormLugarEjecucion = Nothing

    If FormInteraction_FormularioAbierto("FormLugarEjecucion") Then
        FormInteraction_CerrarFormulario "FormLugarEjecucion", acSaveNo
    End If

    DoCmd.OpenForm "FormLugarEjecucion"

    If FormInteraction_FormularioAbierto("FormLugarEjecucion") Then
        Set m_FormLugarEjecucion = Forms("FormLugarEjecucion")
    End If

    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoEditar_Click se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        On Error Resume Next
        CorreoAlAdministrador m_Error
        On Error GoTo 0
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub


' --- ComandoEliminar: pass stub entity -> helper -> handle JSON outcome ---------
' p_PromptResult=0 means "ask the user via MsgBox" — production path. The helper
' returns JSON; we branch on the JSON to decide whether to refresh + raise the
' Eliminar event.
Private Sub ComandoEliminar_Click()
    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents

    m_Error = ""

    If m_LugarEjecucionSeleccionada Is Nothing Then
        m_Error = "Se ha de seleccionar un Elemento de la lista"
        Err.Raise 1000
    End If

    ' Wrap the live LugarEjecucion class instance into a Dictionary-shaped stub
    ' so the helper receives pure data (no class references).
    Dim stub As Object
    Set stub = CreateObject("Scripting.Dictionary")
    stub("IDLugarEjecucion") = m_LugarEjecucionSeleccionada.IDLugarEjecucion
    stub("LugarEjecucion") = m_LugarEjecucionSeleccionada.LugarEjecucion
    stub("DESCRIPCION") = m_LugarEjecucionSeleccionada.DESCRIPCION

    Dim json As String
    json = modLugarEjecucionHelper.LugarEjecucion_Eliminar_Borrar(stub, 0, m_Error)
    If m_Error <> "" Then Err.Raise 1000

    ' Inspect the result JSON. "cancelled" (no payload, ok=true) -> no-op.
    ' "deleted" (no payload, ok=true) -> raise Eliminar + refresh.
    ' ok=false -> surface the helper's error via MsgBox.
    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If parsed("ok") = True Then
        ' Helper treats both cancelled and deleted as ok=true with no payload.
        ' We detect "deleted" via the DAO cardinality side-effect by raising the
        ' event only when the helper succeeded. The helper's logs encode whether
        ' the actual delete fired (logs(3) = "Eliminar_Borrar: deleted").
        ' For the form path we conservatively raise Eliminar whenever the result
        ' is ok=true (the helper did not surface a structural error).
        RaiseEvent Eliminar
        ComandoBuscar_Click
    Else
        m_Error = "Eliminar: " & CStr(parsed("error"))
        pregunta = MsgBox(m_Error, vbExclamation, "Eliminar")
    End If

    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub

errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoEliminar_Click se ha producido el error num " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        On Error Resume Next
        CorreoAlAdministrador m_Error
        On Error GoTo 0
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub


' --- ComandoLimpiar: pure UI reset (no helper) ----------------------------------
' Per e2e rule #1 + the rework design: clearing the filter is a pure UI action
' (set the textbox to Null), no business logic to extract.
Private Sub ComandoLimpiar_Click()
    On Error GoTo errores
    m_Error = ""

    Me.LUGAREJECUCION.Value = Null
    Set m_LugarEjecucionSeleccionada = Nothing
    FormInteraction_EstablecerPropiedad Me, "ComandoEditar", "Enabled", False
    FormInteraction_EstablecerPropiedad Me, "ComandoEliminar", "Enabled", False

    ComandoBuscar_Click
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoLimpiar_Click se ha producido el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        On Error Resume Next
        CorreoAlAdministrador m_Error
        On Error GoTo 0
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub


' --- ComandoAyuda: pure UI delegation (no helper) --------------------------------
Private Sub ComandoAyuda_Click()
    On Error GoTo errores
    m_Error = ""
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    AbrirAyuda m_Error
    If m_Error <> "" Then Err.Raise 1000
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoAyuda_Click se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        On Error Resume Next
        CorreoAlAdministrador m_Error
        On Error GoTo 0
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub


' --- LUGAREJECUCION_KeyDown: pure UI shortcut -----------------------------------
Private Sub LUGAREJECUCION_KeyDown(KeyCode As Integer, Shift As Integer)
    ' Detectar si se ha pulsado la tecla Enter
    If KeyCode = vbKeyReturn Then
        If Me.ComandoBuscar.Enabled = True Then
            Me.ComandoBuscar.SetFocus
            ComandoBuscar_Click
        End If
        KeyCode = 0
    End If
End Sub


' --- WithEvents callbacks: refresh after alta/edicion --------------------------
Private Sub m_FormLugarEjecucion_Alta()
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents

    m_Error = ""
    Set m_ObjEntorno.LugaresEjecucion = Nothing
    ComandoBuscar_Click
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al m_FormLugarEjecucion_Alta se ha producido el error num " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        On Error Resume Next
        CorreoAlAdministrador m_Error
        On Error GoTo 0
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub m_FormLugarEjecucion_Editado()
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents

    m_Error = ""
    Set m_ObjEntorno.LugaresEjecucion = Nothing
    ComandoBuscar_Click
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al m_FormLugarEjecucion_Editado se ha producido el error num " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        On Error Resume Next
        CorreoAlAdministrador m_Error
        On Error GoTo 0
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub
