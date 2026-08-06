Option Compare Database
Option Explicit

' Form_FormGradosClasificacionGestion — REWORK (2026-06-26)
' Thin UI wiring for the GradoClasificacion Gestion list.
'
' Implements access-vba-e2e-methodology rule #1: forms are thin UI wiring.
' Every handler does exactly: read controls -> call helper -> render result.
'
' Anti-pattern removed (was in PR #29 commit 979a57a):
'   - helpers received `ByRef p_Form As Object` and read controls via p_Form.Controls(...)
'   - tests opened the real form and asserted against p_Form.Controls(...).Value
'   - test fixture `Test_GradosClasificacionHelper_OpenForm(p_Form, p_Error)` blocked COM
'
' New contract (per the rework):
'   - All helpers in modGradosClasificacionHelper accept PURE DATA (no Form refs).
'   - Form extracts control values BEFORE calling helper; renders helper JSON AFTER.
'   - Helpers receive: primitives (Boolean, String), stub Dictionaries, stub entities.
'   - UI orchestration stays in this file: DoCmd.OpenForm, FormInteraction_FormularioAbierto, Forms(...).
'   - Five helper calls back the form: Abrir_Inicializar, Buscar_Listar, Seleccionar_Cargar,
'     Eliminar_Borrar, DobleClick_AbrirEdicion.
'   - Three handlers do pure UI orchestration (no helper): Alta, Edicion, Limpiar.

' === Module-level state (all declarations at top per vba-access §10.1) ==============

' The currently selected entity (loaded by Seleccionar_Cargar / edited by Edicion).
Private m_GradoClasificacionSeleccionado As Object

' The currently open FormGradoClasificacion child form (when Alta / Edicion is active).
Private WithEvents m_FormGradoClasificacion As Form_FormGradoClasificacion

' Error accumulator for handler traps (Telefónica D&S convention).
Private m_Error As String

' Public events the parent form listens to.
Public Event Eliminar()
Public Event Seleccionar(m_GradoClasificacionSeleccionado As Object)


' === Local helpers (all at top per vba-access §10.1) =================================

' --- GradosClasificacion_ExtraerFiltro ----------------------------------------------
' Reads the filter text from Me.GradoClasificacion (control name is PascalCase).
' Returns "" if Null.
Private Function GradosClasificacion_ExtraerFiltro() As String
    Dim v As Variant
    v = Nz(Me.GradoClasificacion.Value, "")
    GradosClasificacion_ExtraerFiltro = CStr(v)
End Function

' --- GradosClasificacion_ExtraerIDSeleccionado --------------------------------------
' Reads the currently selected ID from ListaFiltrados.Column(0).
' Returns "" if no selection.
Private Function GradosClasificacion_ExtraerIDSeleccionado() As String
    Dim v As Variant
    On Error Resume Next
    v = Me.ListaFiltrados.Column(0)
    On Error GoTo 0
    If IsNull(v) Then
        GradosClasificacion_ExtraerIDSeleccionado = ""
    Else
        GradosClasificacion_ExtraerIDSeleccionado = CStr(v)
    End If
End Function

' --- GradosClasificacion_ExtraerGrados ---------------------------------------------
' Reads the in-memory collection of GradoClasificacion entities from m_ObjEntorno.
' Returns a Dictionary (possibly empty), or a fresh empty Dictionary if the
' collection is Nothing.
Private Function GradosClasificacion_ExtraerGrados() As Object
    Dim col As Object
    On Error Resume Next
    Set col = m_ObjEntorno.GradosClasificaciones
    On Error GoTo 0
    If col Is Nothing Then
        Set col = CreateObject("Scripting.Dictionary")
    End If
    Set GradosClasificacion_ExtraerGrados = col
End Function

' --- GradosClasificacion_EsAdministrador --------------------------------------------
Private Function GradosClasificacion_EsAdministrador() As Boolean
    If EsAdministrador = EnumSiNo.Sí Then
        GradosClasificacion_EsAdministrador = True
    Else
        GradosClasificacion_EsAdministrador = False
    End If
End Function

' --- GradosClasificacion_TieneOpenArgs ----------------------------------------------
Private Function GradosClasificacion_TieneOpenArgs() As Boolean
    If Len(Nz(Me.OpenArgs, "")) > 0 Then
        GradosClasificacion_TieneOpenArgs = True
    Else
        GradosClasificacion_TieneOpenArgs = False
    End If
End Function

' --- GradosClasificacion_RenderResultadoInicializar ---------------------------------
' Renders the JSON payload from Abrir_Inicializar into the form's controls.
Private Sub GradosClasificacion_RenderResultadoInicializar(ByVal p_Payload As Object)
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

' --- GradosClasificacion_RenderResultadoBuscar --------------------------------------
' Renders the JSON payload from Buscar_Listar into ListaFiltrados.
Private Sub GradosClasificacion_RenderResultadoBuscar(ByVal p_Payload As Object)
    If p_Payload Is Nothing Then Exit Sub

    Dim rowSource As String
    rowSource = CStr(p_Payload("rowSource"))

    Dim rowCount As Long
    rowCount = CLng(p_Payload("count"))

    If Len(rowSource) > 0 Then
        Me.ListaFiltrados.RowSource = "IDGradoClasificacion;GRADO;DESCRIPCION" & vbCrLf & rowSource
    Else
        Me.ListaFiltrados.RowSource = "IDGradoClasificacion;GRADO;DESCRIPCION"
    End If

    If rowCount > 0 Then
        On Error Resume Next
        Me.ListaFiltrados.Selected(0) = True
        On Error GoTo 0
        ListaFiltrados_Click
    End If
End Sub

' --- GradosClasificacion_RenderResultadoSeleccionar --------------------------------
' Renders the JSON payload from Seleccionar_Cargar into the form's controls.
Private Sub GradosClasificacion_RenderResultadoSeleccionar(ByVal p_Payload As Object)
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
        Set m_GradoClasificacionSeleccionado = Nothing
    Else
        ' Wrap the serialized Dictionary into a real GradoClasificacion class.
        ' Entity property is `IdGradoClasificacion` (lowercase 'd').
        Dim wrapped As New GradoClasificacion
        wrapped.IdGradoClasificacion = CStr(entity("IDGradoClasificacion"))
        wrapped.GradoClasificacion = CStr(entity("GradoClasificacion"))
        wrapped.DESCRIPCION = CStr(entity("DESCRIPCION"))
        Set m_GradoClasificacionSeleccionado = wrapped
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

    Dim esAdmin As Boolean
    esAdmin = GradosClasificacion_EsAdministrador()

    Dim tieneOpenArgs As Boolean
    tieneOpenArgs = GradosClasificacion_TieneOpenArgs()

    Dim initJson As String
    initJson = modGradosClasificacionHelper.GradosClasificacion_Abrir_Inicializar( _
        esAdmin, tieneOpenArgs, m_Error)
    If m_Error <> "" Then Err.Raise 1000

    Dim initParsed As Object
    Set initParsed = JsonConverter.ParseJson(initJson)
    GradosClasificacion_RenderResultadoInicializar initParsed("payload")

    Dim filterText As String
    filterText = GradosClasificacion_ExtraerFiltro()

    Dim col As Object
    Set col = GradosClasificacion_ExtraerGrados()

    Dim searchJson As String
    searchJson = modGradosClasificacionHelper.GradosClasificacion_Buscar_Listar(col, filterText, m_Error)
    If m_Error <> "" Then Err.Raise 1000

    Dim searchParsed As Object
    Set searchParsed = JsonConverter.ParseJson(searchJson)
    GradosClasificacion_RenderResultadoBuscar searchParsed("payload")

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


Private Sub cmdElegir_Click()
    On Error GoTo errores
    m_Error = ""
    If m_GradoClasificacionSeleccionado Is Nothing Then
        Exit Sub
    End If
    RaiseEvent Seleccionar(m_GradoClasificacionSeleccionado)
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
Private Sub ComandoAlta_Click()
    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents

    m_Error = ""
    Set m_GradoClasificacionSeleccionado = Nothing
    Set m_FormGradoClasificacion = Nothing

    If FormInteraction_FormularioAbierto("FormGradoClasificacion") Then
        FormInteraction_CerrarFormulario "FormGradoClasificacion", acSaveNo
    End If

    DoCmd.OpenForm "FormGradoClasificacion"

    If FormInteraction_FormularioAbierto("FormGradoClasificacion") Then
        Set m_FormGradoClasificacion = Forms("FormGradoClasificacion")
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
    filterText = GradosClasificacion_ExtraerFiltro()

    Dim col As Object
    Set col = GradosClasificacion_ExtraerGrados()

    Dim json As String
    json = modGradosClasificacionHelper.GradosClasificacion_Buscar_Listar(col, filterText, m_Error)
    If m_Error <> "" Then Err.Raise 1000

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    GradosClasificacion_RenderResultadoBuscar parsed("payload")

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


Private Sub ListaFiltrados_Click()
    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents

    m_Error = ""

    Dim selectedId As String
    selectedId = GradosClasificacion_ExtraerIDSeleccionado()

    Dim col As Object
    Set col = GradosClasificacion_ExtraerGrados()

    Dim esAdmin As Boolean
    esAdmin = GradosClasificacion_EsAdministrador()

    Dim json As String
    json = modGradosClasificacionHelper.GradosClasificacion_Seleccionar_Cargar( _
        selectedId, col, esAdmin, m_Error)
    If m_Error <> "" Then Err.Raise 1000

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    GradosClasificacion_RenderResultadoSeleccionar parsed("payload")

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
    json = modGradosClasificacionHelper.GradosClasificacion_DobleClick_AbrirEdicion( _
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
Private Sub ComandoEditar_Click()
    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents

    m_Error = ""
    Set m_ObjGradoClasificacionActivo = m_GradoClasificacionSeleccionado
    Set m_FormGradoClasificacion = Nothing

    If FormInteraction_FormularioAbierto("FormGradoClasificacion") Then
        FormInteraction_CerrarFormulario "FormGradoClasificacion", acSaveNo
    End If

    DoCmd.OpenForm "FormGradoClasificacion"

    If FormInteraction_FormularioAbierto("FormGradoClasificacion") Then
        Set m_FormGradoClasificacion = Forms("FormGradoClasificacion")
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


Private Sub ComandoEliminar_Click()
    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents

    m_Error = ""

    If m_GradoClasificacionSeleccionado Is Nothing Then
        m_Error = "Se ha de seleccionar un Elemento de la lista"
        Err.Raise 1000
    End If

    ' Wrap the live GradoClasificacion class instance into a Dictionary-shaped stub
    ' so the helper receives pure data (no class references).
    Dim stub As Object
    Set stub = CreateObject("Scripting.Dictionary")
    stub("IDGradoClasificacion") = m_GradoClasificacionSeleccionado.IdGradoClasificacion
    stub("GradoClasificacion") = m_GradoClasificacionSeleccionado.GradoClasificacion
    stub("DESCRIPCION") = m_GradoClasificacionSeleccionado.DESCRIPCION

    Dim json As String
    json = modGradosClasificacionHelper.GradosClasificacion_Eliminar_Borrar(stub, 0, m_Error)
    If m_Error <> "" Then Err.Raise 1000

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If parsed("ok") = True Then
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
Private Sub ComandoLimpiar_Click()
    On Error GoTo errores
    m_Error = ""

    Me.GradoClasificacion.Value = Null
    Set m_GradoClasificacionSeleccionado = Nothing
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


' --- GradoClasificacion_Exit: re-run search on filter change ------------------------
Private Sub GradoClasificacion_Exit(Cancel As Integer)
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
        m_Error = "Al GradoClasificacion_Exit se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        On Error Resume Next
        CorreoAlAdministrador m_Error
        On Error GoTo 0
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub


Private Sub m_FormGradoClasificacion_Alta()
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents

    m_Error = ""
    Set m_ObjEntorno.GradosClasificaciones = Nothing
    ComandoBuscar_Click
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al m_FormGradoClasificacion_Alta se ha producido el error num " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        On Error Resume Next
        CorreoAlAdministrador m_Error
        On Error GoTo 0
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub m_FormGradoClasificacion_Editado()
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents

    m_Error = ""
    Set m_ObjEntorno.GradosClasificaciones = Nothing
    ComandoBuscar_Click
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al m_FormGradoClasificacion_Editado se ha producido el error num " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        On Error Resume Next
        CorreoAlAdministrador m_Error
        On Error GoTo 0
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub
