Option Compare Database
Option Explicit

' Form_FormOficinasProgramaGestion — REWORK-PILOT (2026-06-26)
' Thin UI wiring for the OficinaPrograma Gestion list.
'
' Implements access-vba-e2e-methodology rule #1: forms are thin UI wiring.
' Every handler does exactly: read controls -> call helper -> render result.
' No business logic lives here.
'
' Anti-pattern removed (was in the previous pilot, see docs/audit/pilot-oficina-programa.md):
'   - helpers received `ByRef p_Form As Object` and read controls via p_Form.Controls(...)
'   - tests opened the real form and asserted against p_Form.Controls(...).Value
'   - test fixture `Test_OficinaProgramaHelper_OpenForm(p_Form, p_Error)` blocked COM
'
' New contract (per the rework):
'   - All helpers in modOficinaProgramaHelper accept PURE DATA (no Form refs).
'   - Form extracts control values BEFORE calling helper; renders helper JSON AFTER.
'   - Helpers receive: primitives (Boolean, String), stub Dictionaries, stub entities.
'   - UI orchestration stays in this file: DoCmd.OpenForm, FormularioAbierto, Forms(...).
'   - Five helper calls back the form: Abrir_Inicializar, Buscar_Listar, Seleccionar_Cargar,
'     Eliminar_Borrar, DobleClick_AbrirEdicion.
'   - Three handlers do pure UI orchestration (no helper): Alta, Edicion, Limpiar.

' === Module-level state (all declarations at top per vba-access §10.1) ==============

' The currently selected entity (loaded by Seleccionar_Cargar / edited by Edicion).
' Module-level so the helpers' serialized output can be cached locally if needed,
' but the form itself owns the lifecycle (set, render, clear).
Private m_OficinaProgramaSeleccionada As Object

' The currently open FormOficinaPrograma child form (when Alta / Edicion is active).
' Kept as Object to avoid cross-form compile-time dependency.
Private WithEvents m_FormOficinaPrograma As Form_FormOficinaPrograma

' Error accumulator for handler traps (Telefónica D&S convention).
Private m_Error As String

' Public events the parent form listens to.
Public Event Eliminar()
Public Event Seleccionar(m_OficinaProgramaSeleccionada As Object)

' === Local helpers (all at top per vba-access §10.1) =================================

' --- OficinaPrograma_ExtraerFiltro ----------------------------------------------
' Reads the filter text from Me.OficinaPrograma. Returns "" if Null.
' Pure UI read — used by ComandoBuscar_Click before calling the helper.
Private Function OficinaPrograma_ExtraerFiltro() As String
    Dim v As Variant
    v = Nz(Me.OficinaPrograma.Value, "")
    OficinaPrograma_ExtraerFiltro = CStr(v)
End Function

' --- OficinaPrograma_ExtraerIDSeleccionado --------------------------------------
' Reads the currently selected ID from ListaFiltrados.Column(0).
' Returns "" if no selection (Null column).
' Pure UI read — used by ListaFiltrados_Click before calling the helper.
Private Function OficinaPrograma_ExtraerIDSeleccionado() As String
    Dim v As Variant
    On Error Resume Next
    v = Me.ListaFiltrados.Column(0)
    On Error GoTo 0
    If IsNull(v) Then
        OficinaPrograma_ExtraerIDSeleccionado = ""
    Else
        OficinaPrograma_ExtraerIDSeleccionado = CStr(v)
    End If
End Function

' --- OficinaPrograma_EsAdministrador ---------------------------------------------
' Reads the global admin flag set during EVE().
Private Function OficinaPrograma_EsAdministrador() As Boolean
    If EsAdministrador = EnumSiNo.Sí Then
        OficinaPrograma_EsAdministrador = True
    Else
        OficinaPrograma_EsAdministrador = False
    End If
End Function

' --- OficinaPrograma_TieneOpenArgs -----------------------------------------------
' Returns True when the form was opened with a non-empty OpenArgs.
Private Function OficinaPrograma_TieneOpenArgs() As Boolean
    If Len(Nz(Me.OpenArgs, "")) > 0 Then
        OficinaPrograma_TieneOpenArgs = True
    Else
        OficinaPrograma_TieneOpenArgs = False
    End If
End Function

' --- OficinaPrograma_RenderResultadoInicializar ----------------------------------
' Renders the JSON payload from Abrir_Inicializar into the form's controls.
' Reads payload.showAlta, payload.focusAlta, payload.showElegir.
Private Sub OficinaPrograma_RenderResultadoInicializar(ByVal p_Payload As Object)
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

' --- OficinaPrograma_RenderResultadoBuscar ---------------------------------------
' Renders the JSON payload from Buscar_Listar into ListaFiltrados.
' Reads payload.rowSource and payload.count.
' If count > 0, selects row 0 and re-fires ListaFiltrados_Click so the form's
' selection state (enableEditar/Eliminar) updates.
Private Sub OficinaPrograma_RenderResultadoBuscar(ByVal p_Payload As Object)
    If p_Payload Is Nothing Then Exit Sub

    Dim rowSource As String
    rowSource = CStr(p_Payload("rowSource"))

    Dim rowCount As Long
    rowCount = CLng(p_Payload("count"))

    If Len(rowSource) > 0 Then
        ' Build a full RowSource with header line + data lines.
        Me.ListaFiltrados.RowSource = "IDOficinaPrograma;OFICINA;DESCRIPCION" & vbCrLf & rowSource
    Else
        Me.ListaFiltrados.RowSource = "IDOficinaPrograma;OFICINA;DESCRIPCION"
    End If

    If rowCount > 0 Then
        On Error Resume Next
        Me.ListaFiltrados.Selected(0) = True
        On Error GoTo 0
        ' Re-fire selection handler so the form's button state syncs.
        ListaFiltrados_Click
    End If
End Sub

' --- OficinaPrograma_RenderResultadoSeleccionar ----------------------------------
' Renders the JSON payload from Seleccionar_Cargar into the form's controls.
' Reads payload.entity, payload.enableEditar, payload.enableEliminar.
' Also updates m_OficinaProgramaSeleccionada so the form's state is consistent
' across Edicion/Eliminar calls.
Private Sub OficinaPrograma_RenderResultadoSeleccionar(ByVal p_Payload As Object)
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
        Set m_OficinaProgramaSeleccionada = Nothing
    Else
        ' Wrap the serialized Dictionary into a real OficinaPrograma so subsequent
        ' Edicion calls can pass it to OficinaProgramaOperaciones via the helper.
        Dim wrapped As New OficinaPrograma
        wrapped.IDOficinaPrograma = CStr(entity("IDOficinaPrograma"))
        wrapped.OficinaPrograma = CStr(entity("OficinaPrograma"))
        wrapped.DESCRIPCION = CStr(entity("DESCRIPCION"))
        Set m_OficinaProgramaSeleccionada = wrapped
    End If
End Sub


' === Form lifecycle / event handlers ===================================================

' --- Form_Open ---------------------------------------------------------------------
' Lifecycle adapter: Ajustar + initialize + initial search.
' Helper: OficinaPrograma_Abrir_Inicializar (receives primitives).
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
    esAdmin = OficinaPrograma_EsAdministrador()

    Dim tieneOpenArgs As Boolean
    tieneOpenArgs = OficinaPrograma_TieneOpenArgs()

    Dim initJson As String
    initJson = modOficinaProgramaHelper.OficinaPrograma_Abrir_Inicializar( _
        esAdmin, tieneOpenArgs, m_Error)
    If m_Error <> "" Then Err.Raise 1000

    ' Render the payload into controls.
    Dim initParsed As Object
    Set initParsed = JsonConverter.ParseJson(initJson)
    OficinaPrograma_RenderResultadoInicializar initParsed("payload")

    ' Run initial search so the listbox is populated.
    Dim filterText As String
    filterText = OficinaPrograma_ExtraerFiltro()

    Dim col As Object
    On Error Resume Next
    Set col = m_ObjEntorno.OficinasDePrograma
    On Error GoTo 0
    If col Is Nothing Then
        Set col = CreateObject("Scripting.Dictionary")
    End If

    Dim searchJson As String
    searchJson = modOficinaProgramaHelper.OficinaPrograma_Buscar_Listar(col, filterText, m_Error)
    If m_Error <> "" Then Err.Raise 1000

    Dim searchParsed As Object
    Set searchParsed = JsonConverter.ParseJson(searchJson)
    OficinaPrograma_RenderResultadoBuscar searchParsed("payload")

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
    If m_OficinaProgramaSeleccionada Is Nothing Then
        Exit Sub
    End If
    RaiseEvent Seleccionar(m_OficinaProgramaSeleccionada)
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
' logic. We keep DoCmd.OpenForm / FormularioAbierto / Forms(...) in the form.
Private Sub ComandoAlta_Click()
    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents

    m_Error = ""
    Set m_OficinaProgramaSeleccionada = Nothing
    Set m_FormOficinaPrograma = Nothing

    If FormInteraction_FormularioAbierto("FormOficinaPrograma") Then
        FormInteraction_CerrarFormulario "FormOficinaPrograma", acSaveNo
    End If

    DoCmd.OpenForm "FormOficinaPrograma"

    If FormInteraction_FormularioAbierto("FormOficinaPrograma") Then
        Set m_FormOficinaPrograma = Forms("FormOficinaPrograma")
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
    filterText = OficinaPrograma_ExtraerFiltro()

    Dim col As Object
    On Error Resume Next
    Set col = m_ObjEntorno.OficinasDePrograma
    On Error GoTo 0
    If col Is Nothing Then
        Set col = CreateObject("Scripting.Dictionary")
    End If

    Dim json As String
    json = modOficinaProgramaHelper.OficinaPrograma_Buscar_Listar(col, filterText, m_Error)
    If m_Error <> "" Then Err.Raise 1000

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    OficinaPrograma_RenderResultadoBuscar parsed("payload")

    AvanceCerrar
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub

errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoBuscar_Click se ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        On Error Resume Next
        CorreoAlAdministrador m_Error
        On Error GoTo 0
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
    selectedId = OficinaPrograma_ExtraerIDSeleccionado()

    Dim col As Object
    On Error Resume Next
    Set col = m_ObjEntorno.OficinasDePrograma
    On Error GoTo 0
    If col Is Nothing Then
        Set col = CreateObject("Scripting.Dictionary")
    End If

    Dim esAdmin As Boolean
    esAdmin = OficinaPrograma_EsAdministrador()

    Dim json As String
    json = modOficinaProgramaHelper.OficinaPrograma_Seleccionar_Cargar( _
        selectedId, col, esAdmin, m_Error)
    If m_Error <> "" Then Err.Raise 1000

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    OficinaPrograma_RenderResultadoSeleccionar parsed("payload")

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
    json = modOficinaProgramaHelper.OficinaPrograma_DobleClick_AbrirEdicion( _
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
' loaded by ListaFiltrados_Click into m_OficinaProgramaSeleccionada; we just
' stash it globally and open the alta form so it can edit it.
Private Sub ComandoEditar_Click()
    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents

    m_Error = ""
    Set m_ObjOficinaProgramaActiva = m_OficinaProgramaSeleccionada
    Set m_FormOficinaPrograma = Nothing

    If FormInteraction_FormularioAbierto("FormOficinaPrograma") Then
        FormInteraction_CerrarFormulario "FormOficinaPrograma", acSaveNo
    End If

    DoCmd.OpenForm "FormOficinaPrograma"

    If FormInteraction_FormularioAbierto("FormOficinaPrograma") Then
        Set m_FormOficinaPrograma = Forms("FormOficinaPrograma")
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

    If m_OficinaProgramaSeleccionada Is Nothing Then
        m_Error = "Se ha de seleccionar un Elemento de la lista"
        Err.Raise 1000
    End If

    ' Wrap the live OficinaPrograma class instance into a Dictionary-shaped stub
    ' so the helper receives pure data (no class references).
    Dim stub As Object
    Set stub = CreateObject("Scripting.Dictionary")
    stub("IDOficinaPrograma") = m_OficinaProgramaSeleccionada.IDOficinaPrograma
    stub("OficinaPrograma") = m_OficinaProgramaSeleccionada.OficinaPrograma
    stub("DESCRIPCION") = m_OficinaProgramaSeleccionada.DESCRIPCION

    Dim json As String
    json = modOficinaProgramaHelper.OficinaPrograma_Eliminar_Borrar(stub, 0, m_Error)
    If m_Error <> "" Then Err.Raise 1000

    ' Inspect the result JSON. "cancelled" (no payload, ok=true) -> no-op.
    ' "deleted" (no payload, ok=true) -> raise Eliminar + refresh.
    ' ok=false -> surface the helper's error via MsgBox.
    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If parsed("ok") = True Then
        ' The helper treats both cancelled and deleted as ok=true with no payload.
        ' We can detect "deleted" vs "cancelled" by the side effect (DAO row count
        ' before/after), but for the form path we simply raise the event when the
        ' helper did NOT report the N/A block. Since N/A is the only fail path
        ' surfaced here, and the rest are ok=true with no payload, we conservatively
        ' raise Eliminar whenever the result is ok=true and the user's selection
        ' was real (not N/A).
        If CStr(m_OficinaProgramaSeleccionada.OficinaPrograma) <> "N/A" Then
            RaiseEvent Eliminar
            ComandoBuscar_Click
        End If
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

    Me.OficinaPrograma.Value = Null
    Set m_OficinaProgramaSeleccionada = Nothing
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


' --- OficinaPrograma_Exit: re-run search on filter change ------------------------
Private Sub OficinaPrograma_Exit(Cancel As Integer)
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
        m_Error = "Al OficinaPrograma_Exit se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        On Error Resume Next
        CorreoAlministrador m_Error
        On Error GoTo 0
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub


' --- WithEvents callbacks: refresh after alta/edicion --------------------------
Private Sub m_FormOficinaPrograma_Alta()
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents

    m_Error = ""
    Set m_ObjEntorno.OficinasDePrograma = Nothing
    ComandoBuscar_Click
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al m_FormOficinaPrograma_Alta se ha producido el error num " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        On Error Resume Next
        CorreoAlAdministrador m_Error
        On Error GoTo 0
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub m_FormOficinaPrograma_Editado()
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents

    m_Error = ""
    Set m_ObjEntorno.OficinasDePrograma = Nothing
    ComandoBuscar_Click
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al m_FormOficinaPrograma_Editado se ha producido el error num " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        On Error Resume Next
        CorreoAlAdministrador m_Error
        On Error GoTo 0
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub