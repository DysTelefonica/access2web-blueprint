Option Compare Database
Option Explicit

' Form_FormUsuariosGestion — REWORK (2026-06-26)
' Thin UI wiring for the SELECT-only USUARIO Gestion form.
'
' Implements access-vba-e2e-methodology rule #1: forms are thin UI wiring.
' Every handler does exactly: read controls -> call helper -> render result.
'
' Anti-pattern removed (was in PR #31 commit c86d460):
'   - helpers received `ByRef p_Form As Object` and read controls via p_Form.Controls(...)
'   - helpers did `Application.Echo` directly
'   - tests opened the real form and asserted against p_Form.Controls(...).Value
'   - test fixture `Test_UsuariosHelper_OpenForm(p_Form, p_Error)` blocked COM
'
' New contract (per the rework):
'   - All helpers in modUsuariosHelper accept PURE DATA (no Form refs).
'   - Form extracts control values BEFORE calling helper; renders helper JSON AFTER.
'   - Five helper calls back the form: Abrir_Inicializar, Buscar_Listar, Seleccionar_Cargar,
'     Limpiar_Reset, DobleClick_AbrirDetalle.
'   - Form control names: `USUARIO` (text filter, UPPERCASE), `Activos` (dropdown).
'
' SELECT-only form: NO Alta/Editar/Eliminar helpers. SELECT-only -> DobleClick
' dispatches to "choose" only.

' === Module-level state (all declarations at top per vba-access §10.1) ==============

' The currently selected USUARIO entity (loaded by Buscar_Listar / read by
' ListaFiltrados_Click). Kept as Object so the form can wrap Dictionary output
' from the helper back into a USUARIO class for parent-form event dispatch.
Private m_UsuarioSeleccionado As Object

' Public events the parent form listens to.
Public Event Seleccionar(m_UsuarioSeleccionado As USUARIO)

' Error accumulator for handler traps (Telefónica D&S convention).
Private m_Error As String


' === Local helpers (all at top per vba-access §10.1) =================================

' --- Usuarios_ExtraerFiltro --------------------------------------------------------
' Reads the filter text from Me.USUARIO. Returns "" if Null.
' Pure UI read — used by ComandoBuscar_Click before calling the helper.
Private Function Usuarios_ExtraerFiltro() As String
    Dim v As Variant
    v = Nz(Me.USUARIO.Value, "")
    Usuarios_ExtraerFiltro = CStr(v)
End Function

' --- Usuarios_ExtraerActivos -------------------------------------------------------
' Reads the Activos dropdown value from Me.Activos. Returns "Todos" if Null/empty
' (matches legacy behavior — empty/null defaults to "Todos" in Buscar_Listar too).
Private Function Usuarios_ExtraerActivos() As String
    Dim v As Variant
    v = Nz(Me.Activos.Value, "")
    If Len(CStr(v)) = 0 Then
        Usuarios_ExtraerActivos = "Todos"
    Else
        Usuarios_ExtraerActivos = CStr(v)
    End If
End Function

' --- Usuarios_ExtraerUsuarios ------------------------------------------------------
' Reads the in-memory collection of USUARIO entities from m_ObjEntorno.
' Returns a Dictionary (possibly empty), or a fresh empty Dictionary if the
' collection is Nothing (defensive — never pass Nothing to the helper by accident).
Private Function Usuarios_ExtraerUsuarios() As Object
    Dim col As Object
    On Error Resume Next
    Set col = m_ObjEntorno.ColUsuarios
    On Error GoTo 0
    If col Is Nothing Then
        Set col = CreateObject("Scripting.Dictionary")
    End If
    Set Usuarios_ExtraerUsuarios = col
End Function

' --- Usuarios_ExtraerIDSeleccionado -----------------------------------------------
' Reads the currently selected ID from ListaFiltrados.Column(0).
' Returns "" if no selection (Null column).
Private Function Usuarios_ExtraerIDSeleccionado() As String
    Dim v As Variant
    On Error Resume Next
    v = Me.ListaFiltrados.Column(0)
    On Error GoTo 0
    If IsNull(v) Then
        Usuarios_ExtraerIDSeleccionado = ""
    Else
        Usuarios_ExtraerIDSeleccionado = CStr(v)
    End If
End Function

' --- Usuarios_EsAdministrador ------------------------------------------------------
Private Function Usuarios_EsAdministrador() As Boolean
    If EsAdministrador = EnumSiNo.Sí Then
        Usuarios_EsAdministrador = True
    Else
        Usuarios_EsAdministrador = False
    End If
End Function

' --- Usuarios_TieneOpenArgs --------------------------------------------------------
Private Function Usuarios_TieneOpenArgs() As Boolean
    If Len(Nz(Me.OpenArgs, "")) > 0 Then
        Usuarios_TieneOpenArgs = True
    Else
        Usuarios_TieneOpenArgs = False
    End If
End Function

' --- Usuarios_RenderResultadoInicializar ------------------------------------------
' Renders the JSON payload from Abrir_Inicializar into the form's controls.
' Reads payload.hasElegir.
Private Sub Usuarios_RenderResultadoInicializar(ByVal p_Payload As Object)
    If p_Payload Is Nothing Then Exit Sub

    Dim hasElegir As Boolean
    hasElegir = CBool(p_Payload("hasElegir"))

    FormInteraction_EstablecerPropiedad Me, "cmdElegir", "Visible", hasElegir
End Sub

' --- Usuarios_RenderResultadoBuscar ------------------------------------------------
' Renders the JSON payload from Buscar_Listar into ListaFiltrados.
' Reads payload.rowSource and payload.count.
' If count > 0, selects row 0 and re-fires ListaFiltrados_Click so the form's
' selection state updates.
Private Sub Usuarios_RenderResultadoBuscar(ByVal p_Payload As Object)
    If p_Payload Is Nothing Then Exit Sub

    Dim rowSource As String
    rowSource = CStr(p_Payload("rowSource"))

    Dim rowCount As Long
    rowCount = CLng(p_Payload("count"))

    If Len(rowSource) > 0 Then
        Me.ListaFiltrados.RowSource = "Id;Nombre" & vbCrLf & rowSource
    Else
        Me.ListaFiltrados.RowSource = "Id;Nombre"
    End If

    If rowCount > 0 Then
        On Error Resume Next
        Me.ListaFiltrados.Selected(0) = True
        On Error GoTo 0
        ' Re-fire selection handler so the form's selection state syncs.
        ListaFiltrados_Click
    End If
End Sub

' --- Usuarios_RenderResultadoSeleccionar ------------------------------------------
' Renders the JSON payload from Seleccionar_Cargar into the form's state.
' Reads payload.entity, payload.hasSelection.
' Wraps the serialized Dictionary into a real USUARIO class instance for parent
' event dispatch.
Private Sub Usuarios_RenderResultadoSeleccionar(ByVal p_Payload As Object)
    If p_Payload Is Nothing Then Exit Sub

    Dim hasSelection As Boolean
    hasSelection = CBool(p_Payload("hasSelection"))

    Dim entity As Object
    Set entity = p_Payload("entity")
    If hasSelection And Not entity Is Nothing Then
        ' Wrap the serialized Dictionary into a real USUARIO class instance.
        Dim wrapped As New USUARIO
        wrapped.ID = CStr(entity("ID"))
        wrapped.Nombre = CStr(entity("Nombre"))
        Set m_UsuarioSeleccionado = wrapped
    Else
        Set m_UsuarioSeleccionado = Nothing
    End If
End Sub


' === Form lifecycle / event handlers ===================================================

Private Sub Form_Open(Cancel As Integer)
    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents

    m_Error = ""

    Ajustar Me

    ' Capture primitives before calling the helper (per rule #1).
    Dim esAdmin As Boolean
    esAdmin = Usuarios_EsAdministrador()

    Dim tieneOpenArgs As Boolean
    tieneOpenArgs = Usuarios_TieneOpenArgs()

    Dim initJson As String
    initJson = modUsuariosHelper.Usuarios_Abrir_Inicializar( _
        esAdmin, tieneOpenArgs, m_Error)
    If m_Error <> "" Then Err.Raise 1000

    Dim initParsed As Object
    Set initParsed = JsonConverter.ParseJson(initJson)
    Usuarios_RenderResultadoInicializar initParsed("payload")

    ' Default Activos="Sí" (NOT "Todos" — pre-existing legacy behavior).
    FormInteraction_EstablecerValorControl Me, "Activos", "Sí"

    ' Run initial search so the listbox is populated.
    Dim filterText As String
    filterText = Usuarios_ExtraerFiltro()
    Dim activos As String
    activos = Usuarios_ExtraerActivos()
    Dim col As Object
    Set col = Usuarios_ExtraerUsuarios()

    Dim searchJson As String
    searchJson = modUsuariosHelper.Usuarios_Buscar_Listar(col, filterText, activos, m_Error)
    If m_Error <> "" Then Err.Raise 1000

    Dim searchParsed As Object
    Set searchParsed = JsonConverter.ParseJson(searchJson)
    Usuarios_RenderResultadoBuscar searchParsed("payload")

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
    If m_UsuarioSeleccionado Is Nothing Then
        Exit Sub
    End If
    RaiseEvent Seleccionar(m_UsuarioSeleccionado)
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


' --- ComandoBuscar: extract filter + activos -> helper -> render -----------------
Private Sub ComandoBuscar_Click()
    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents

    m_Error = ""

    Dim filterText As String
    filterText = Usuarios_ExtraerFiltro()
    Dim activos As String
    activos = Usuarios_ExtraerActivos()

    Dim col As Object
    Set col = Usuarios_ExtraerUsuarios()

    Dim json As String
    json = modUsuariosHelper.Usuarios_Buscar_Listar(col, filterText, activos, m_Error)
    If m_Error <> "" Then Err.Raise 1000

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    Usuarios_RenderResultadoBuscar parsed("payload")

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
    selectedId = Usuarios_ExtraerIDSeleccionado()

    Dim col As Object
    Set col = Usuarios_ExtraerUsuarios()

    Dim esAdmin As Boolean
    esAdmin = Usuarios_EsAdministrador()

    Dim json As String
    json = modUsuariosHelper.Usuarios_Seleccionar_Cargar( _
        selectedId, col, esAdmin, m_Error)
    If m_Error <> "" Then Err.Raise 1000

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    Usuarios_RenderResultadoSeleccionar parsed("payload")

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

    Dim detalleEnabled As Boolean
    On Error Resume Next
    detalleEnabled = Me.Controls("cmdElegir").Visible
    On Error GoTo 0

    Dim json As String
    json = modUsuariosHelper.Usuarios_DobleClick_AbrirDetalle( _
        detalleEnabled, m_Error)
    If m_Error <> "" Then Err.Raise 1000

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    Dim action As String
    action = CStr(parsed("payload")("action"))

    Select Case action
        Case "choose"
            cmdElegir_Click
        Case "none"
            ' No-op — helper said cmdElegir is not actionable.
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


' --- ComandoLimpiar: pure UI reset + helper ------------------------------------
Private Sub ComandoLimpiar_Click()
    On Error GoTo errores
    m_Error = ""

    Dim json As String
    json = modUsuariosHelper.Usuarios_Limpiar_Reset(m_Error)
    If m_Error <> "" Then Err.Raise 1000

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    Dim activosDefault As String
    activosDefault = CStr(parsed("payload")("activosDefault"))

    ' Apply the helper's recommendation in the form (UI responsibility).
    Me.USUARIO.Value = Null
    FormInteraction_EstablecerValorControl Me, "Activos", activosDefault
    Set m_UsuarioSeleccionado = Nothing

    ' Re-run search with the new filter.
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


' --- Activos_Change: re-run search on dropdown change -----------------------------
Private Sub Activos_Change()
    On Error Resume Next
    If Me.ComandoBuscar.Enabled = True Then
        Me.ComandoBuscar.SetFocus
        ComandoBuscar_Click
    End If
End Sub


' --- USUARIO_KeyDown: Enter key triggers search ----------------------------------
Private Sub USUARIO_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Me.ComandoBuscar.Enabled = True Then
            Me.ComandoBuscar.SetFocus
            ComandoBuscar_Click
        End If
        KeyCode = 0
    End If
End Sub


' --- ComandoAyuda: pure UI delegation (no helper) --------------------------------
Private Sub ComandoAyuda_Click()
    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents

    m_Error = ""
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
