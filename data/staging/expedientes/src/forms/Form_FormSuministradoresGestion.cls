Option Compare Database
Option Explicit

' Form_FormSuministradoresGestion — REWORK (2026-06-26)
' Thin UI wiring for the Suministrador Gestion list.
'
' Implements access-vba-e2e-methodology rule #1: forms are thin UI wiring.
' Every handler does exactly: read controls -> call helper -> render result.
' No business logic lives here.
'
' Anti-pattern removed (was in PR #31 commit c86d460):
'   - helpers received `ByRef p_Form As Object` and read controls via p_Form.Controls(...)
'   - tests opened the real form and asserted against p_Form.Controls(...).Value
'   - test fixture `Test_SuministradorHelper_OpenForm(p_Form, p_Error)` blocked COM
'
' New contract (per the rework):
'   - All helpers in modSuministradorHelper accept PURE DATA (no Form refs).
'   - Form extracts control values BEFORE calling helper; renders helper JSON AFTER.
'   - Helpers receive: primitives (Boolean, String), stub Dictionaries.
'   - UI orchestration stays in this file: DoCmd.OpenForm, FormInteraction_FormularioAbierto, Forms(...).
'   - Five helper calls back the form: Abrir_Inicializar, Buscar_Listar, Seleccionar_Cargar,
'     Eliminar_Borrar, DobleClick_AbrirEdicion.
'   - Three handlers do pure UI orchestration (no helper): Alta, Edicion, Limpiar.

' === Module-level state (all declarations at top per vba-access §10.1) ==============

' The currently selected entity (loaded by Seleccionar_Cargar / edited by Edicion).
' Module-level so the helpers' serialized output can be cached locally if needed,
' but the form itself owns the lifecycle (set, render, clear).
Private m_SuministradorSeleccionada As Object

' The currently open FormSuministrador child form (when Alta / Edicion is active).
' Kept as Object to avoid cross-form compile-time dependency.
Private WithEvents m_FormSuministrador As Form_FormSuministrador

' Error accumulator for handler traps (Telefónica D&S convention).
Private m_Error As String

' Public events the parent form listens to.
Public Event Eliminar()
Public Event Seleccionar(m_SuministradorSeleccionada As Object)


' === Local helpers (all at top per vba-access §10.1) =================================

' --- Suministrador_ExtraerFiltro ----------------------------------------------
' Reads the filter text from Me.PalabraClave. Returns "" if Null.
' Pure UI read — used by ComandoBuscar_Click before calling the helper.
Private Function Suministrador_ExtraerFiltro() As String
    Dim v As Variant
    v = Nz(Me.PalabraClave.Value, "")
    Suministrador_ExtraerFiltro = CStr(v)
End Function

' --- Suministrador_ExtraerIDSeleccionado --------------------------------------
' Reads the currently selected ID from ListaFiltrados.Column(0).
' Returns "" if no selection (Null column).
Private Function Suministrador_ExtraerIDSeleccionado() As String
    Dim v As Variant
    On Error Resume Next
    v = Me.ListaFiltrados.Column(0)
    On Error GoTo 0
    If IsNull(v) Then
        Suministrador_ExtraerIDSeleccionado = ""
    Else
        Suministrador_ExtraerIDSeleccionado = CStr(v)
    End If
End Function

' --- Suministrador_ExtraerSuministradores --------------------------------------
' Reads the in-memory collection of Suministrador entities from m_ObjEntorno.
' Returns a Dictionary (possibly empty), or a fresh empty Dictionary if the
' collection is Nothing (defensive — never pass Nothing to the helper by accident).
Private Function Suministrador_ExtraerSuministradores() As Object
    Dim col As Object
    On Error Resume Next
    Set col = m_ObjEntorno.Suministradores
    On Error GoTo 0
    If col Is Nothing Then
        Set col = CreateObject("Scripting.Dictionary")
    End If
    Set Suministrador_ExtraerSuministradores = col
End Function

' --- Suministrador_EsAdministrador ---------------------------------------------
' Reads the global admin flag set during EVE().
Private Function Suministrador_EsAdministrador() As Boolean
    If EsAdministrador = EnumSiNo.Sí Then
        Suministrador_EsAdministrador = True
    Else
        Suministrador_EsAdministrador = False
    End If
End Function

' --- Suministrador_TieneOpenArgs -----------------------------------------------
Private Function Suministrador_TieneOpenArgs() As Boolean
    If Len(Nz(Me.OpenArgs, "")) > 0 Then
        Suministrador_TieneOpenArgs = True
    Else
        Suministrador_TieneOpenArgs = False
    End If
End Function

' --- Suministrador_RenderResultadoInicializar ----------------------------------
' Renders the JSON payload from Abrir_Inicializar into the form's controls.
' Reads payload.showAlta, payload.focusAlta, payload.showElegir.
Private Sub Suministrador_RenderResultadoInicializar(ByVal p_Payload As Object)
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

' --- Suministrador_RenderResultadoBuscar ---------------------------------------
' Renders the JSON payload from Buscar_Listar into ListaFiltrados.
' Reads payload.rowSource and payload.count.
' If count > 0, selects row 0 and re-fires ListaFiltrados_Click so the form's
' selection state (enableEditar/Eliminar) updates.
Private Sub Suministrador_RenderResultadoBuscar(ByVal p_Payload As Object)
    If p_Payload Is Nothing Then Exit Sub

    Dim rowSource As String
    rowSource = CStr(p_Payload("rowSource"))

    Dim rowCount As Long
    rowCount = CLng(p_Payload("count"))

    If Len(rowSource) > 0 Then
        ' Build a full RowSource with header line + data lines.
        Me.ListaFiltrados.RowSource = "IDSuministrador;CIF;Nombre" & vbCrLf & rowSource
    Else
        Me.ListaFiltrados.RowSource = "IDSuministrador;CIF;Nombre"
    End If

    If rowCount > 0 Then
        On Error Resume Next
        Me.ListaFiltrados.Selected(0) = True
        On Error GoTo 0
        ' Re-fire selection handler so the form's button state syncs.
        ListaFiltrados_Click
    End If
End Sub

' --- Suministrador_RenderResultadoSeleccionar ----------------------------------
' Renders the JSON payload from Seleccionar_Cargar into the form's controls.
' Reads payload.entity, payload.enableEditar, payload.enableEliminar.
' Also updates m_SuministradorSeleccionada so the form's state is consistent
' across Edicion/Eliminar calls.
Private Sub Suministrador_RenderResultadoSeleccionar(ByVal p_Payload As Object)
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
        Set m_SuministradorSeleccionada = Nothing
    Else
        ' Wrap the serialized Dictionary into a real Suministrador so subsequent
        ' Edicion calls can pass it to SuministradorOperaciones via the helper.
        Dim wrapped As New Suministrador
        wrapped.IDSuministrador = CStr(entity("IDSuministrador"))
        wrapped.CIF = CStr(entity("CIF"))
        wrapped.Nombre = CStr(entity("Nombre"))
        Set m_SuministradorSeleccionada = wrapped
    End If
End Sub


' === Form lifecycle / event handlers ===================================================

' --- Form_Open ---------------------------------------------------------------------
' Lifecycle adapter: Ajustar + initialize + initial search.
' Helper: Suministrador_Abrir_Inicializar (receives primitives).
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
    esAdmin = Suministrador_EsAdministrador()

    Dim tieneOpenArgs As Boolean
    tieneOpenArgs = Suministrador_TieneOpenArgs()

    Dim initJson As String
    initJson = modSuministradorHelper.Suministrador_Abrir_Inicializar( _
        esAdmin, tieneOpenArgs, m_Error)
    If m_Error <> "" Then Err.Raise 1000

    ' Render the payload into controls.
    Dim initParsed As Object
    Set initParsed = JsonConverter.ParseJson(initJson)
    Suministrador_RenderResultadoInicializar initParsed("payload")

    ' Run initial search so the listbox is populated.
    Dim filterText As String
    filterText = Suministrador_ExtraerFiltro()

    Dim col As Object
    Set col = Suministrador_ExtraerSuministradores()

    Dim searchJson As String
    searchJson = modSuministradorHelper.Suministrador_Buscar_Listar(col, filterText, m_Error)
    If m_Error <> "" Then Err.Raise 1000

    Dim searchParsed As Object
    Set searchParsed = JsonConverter.ParseJson(searchJson)
    Suministrador_RenderResultadoBuscar searchParsed("payload")

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
    If m_SuministradorSeleccionada Is Nothing Then
        Exit Sub
    End If
    RaiseEvent Seleccionar(m_SuministradorSeleccionada)
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
    Set m_SuministradorSeleccionada = Nothing
    Set m_FormSuministrador = Nothing

    If FormInteraction_FormularioAbierto("FormSuministrador") Then
        FormInteraction_CerrarFormulario "FormSuministrador", acSaveNo
    End If

    DoCmd.OpenForm "FormSuministrador"

    If FormInteraction_FormularioAbierto("FormSuministrador") Then
        Set m_FormSuministrador = Forms("FormSuministrador")
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
    filterText = Suministrador_ExtraerFiltro()

    Dim col As Object
    Set col = Suministrador_ExtraerSuministradores()

    Dim json As String
    json = modSuministradorHelper.Suministrador_Buscar_Listar(col, filterText, m_Error)
    If m_Error <> "" Then Err.Raise 1000

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    Suministrador_RenderResultadoBuscar parsed("payload")

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
    selectedId = Suministrador_ExtraerIDSeleccionado()

    Dim col As Object
    Set col = Suministrador_ExtraerSuministradores()

    Dim esAdmin As Boolean
    esAdmin = Suministrador_EsAdministrador()

    Dim json As String
    json = modSuministradorHelper.Suministrador_Seleccionar_Cargar( _
        selectedId, col, esAdmin, m_Error)
    If m_Error <> "" Then Err.Raise 1000

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    Suministrador_RenderResultadoSeleccionar parsed("payload")

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
    json = modSuministradorHelper.Suministrador_DobleClick_AbrirEdicion( _
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
' loaded by ListaFiltrados_Click into m_SuministradorSeleccionada; we just
' stash it globally and open the alta form so it can edit it.
Private Sub ComandoEditar_Click()
    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents

    m_Error = ""
    Set m_ObjSuministradorActivo = m_SuministradorSeleccionada
    Set m_FormSuministrador = Nothing

    If FormInteraction_FormularioAbierto("FormSuministrador") Then
        FormInteraction_CerrarFormulario "FormSuministrador", acSaveNo
    End If

    DoCmd.OpenForm "FormSuministrador"

    If FormInteraction_FormularioAbierto("FormSuministrador") Then
        Set m_FormSuministrador = Forms("FormSuministrador")
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

    If m_SuministradorSeleccionada Is Nothing Then
        m_Error = "Se ha de seleccionar un Elemento de la lista"
        Err.Raise 1000
    End If

    ' Wrap the live Suministrador class instance into a Dictionary-shaped stub
    ' so the helper receives pure data (no class references).
    Dim stub As Object
    Set stub = CreateObject("Scripting.Dictionary")
    stub("IDSuministrador") = m_SuministradorSeleccionada.IDSuministrador
    stub("CIF") = m_SuministradorSeleccionada.CIF
    stub("Nombre") = m_SuministradorSeleccionada.Nombre

    Dim json As String
    json = modSuministradorHelper.Suministrador_Eliminar_Borrar(stub, 0, m_Error)
    If m_Error <> "" Then Err.Raise 1000

    ' Inspect the result JSON. "cancelled" (no payload, ok=true) -> no-op.
    ' "deleted" (no payload, ok=true) -> raise Eliminar + refresh.
    ' ok=false -> surface the helper's error via MsgBox.
    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If parsed("ok") = True Then
        ' The helper treats both cancelled and deleted as ok=true with no payload.
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

    Me.PalabraClave.Value = Null
    Set m_SuministradorSeleccionada = Nothing
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


' --- Suministrador_Exit: re-run search on filter change ------------------------
Private Sub PalabraClave_Exit(Cancel As Integer)
    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents

    m_Error = ""
    ComandoBuscar_Click
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al PalabraClave_Exit se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub


' --- WithEvents callbacks: refresh after alta/edicion --------------------------
Private Sub m_FormSuministrador_Alta(p_ID As String)
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents

    m_Error = ""

    ' Invalidate the cached entity collection so the next Buscar_Listar reloads.
    Set m_ObjEntorno.Suministradores = Nothing
    ComandoBuscar_Click
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al m_FormSuministrador_Alta se ha producido el error num " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        On Error Resume Next
        CorreoAlAdministrador m_Error
        On Error GoTo 0
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub m_FormSuministrador_Editado()
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents

    m_Error = ""

    Set m_ObjEntorno.Suministradores = Nothing
    ComandoBuscar_Click
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al m_FormSuministrador_Editado se ha producido el error num " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        On Error Resume Next
        CorreoAlAdministrador m_Error
        On Error GoTo 0
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub
