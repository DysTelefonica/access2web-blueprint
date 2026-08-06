Option Compare Database
Option Explicit

' Form_FormOrganoContratacion — REWORK (2026-06-26)
' Thin UI wiring for the OrganoContratacion Alta/Edición form.
'
' Implements access-vba-e2e-methodology rule #1: forms are thin UI wiring.
' Every handler does exactly: read controls -> call helper -> render result.
'
' Anti-pattern removed (was in PR #30 commit 97bdb21):
'   - helpers received `ByRef p_Form As Object` and read controls via p_Form.Controls(...)
'   - tests opened the real form and asserted against p_Form.Controls(...).Value
'   - test fixture `Test_OrganoContratacionAltaHelper_OpenForm(p_Form, p_Error)` blocked COM
'
' New contract (per the rework):
'   - All helpers in modOrganoContratacionAltaHelper accept PURE DATA (no Form refs).
'   - Form extracts control values BEFORE calling helper; renders helper JSON AFTER.
'   - Four helper calls back the form: Abrir_Inicializar, VerificarCambios,
'     Registrar, Cerrar.
'   - Helpers return JSON that the form parses via JsonConverter.ParseJson and
'     renders into controls.
'   - The form's parent sets the project-conventional global `m_ObjOrganoContratacionActivo`
'     when opening in edicion mode (via Abrir_Inicializar which sets it internally when loading).

' === Module-level state (all declarations at top per vba-access §10.1) ==============

' Public events the parent form listens to (raised after Registrar succeeds).
Public Event Alta()
Public Event Editado()

' Error accumulator for handler traps (Telefónica D&S convention).
Private m_Error As String


' === Local helpers (all at top per vba-access §10.1) =================================

' --- OrganoContratacion_ExtraerValores -----------------------------------------------
' Reads current control values into a Dictionary {OrganoContratacion, DESCRIPCION}.
' Used by VerificarCambios and Registrar before calling the helper.
Private Function OrganoContratacion_ExtraerValores() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("OrganoContratacion") = CStr(Nz(Me.OrganoContratacion.Value, ""))
    d("DESCRIPCION") = CStr(Nz(Me.DESCRIPCION.Value, ""))
    Set OrganoContratacion_ExtraerValores = d
End Function

' --- OrganoContratacion_EsAdministrador -----------------------------------------------
' Reads the global admin flag set during EVE().
Private Function OrganoContratacion_EsAdministrador() As Boolean
    If EsAdministrador = EnumSiNo.Sí Then
        OrganoContratacion_EsAdministrador = True
    Else
        OrganoContratacion_EsAdministrador = False
    End If
End Function

' --- OrganoContratacion_Modo ------------------------------------------------------------
' Returns the modo (alta | edicion) based on whether m_ObjOrganoContratacionActivo is set
' (set by the parent form's Alta/Editar handlers before opening this form).
Private Function OrganoContratacion_Modo() As String
    If m_ObjOrganoContratacionActivo Is Nothing Then
        OrganoContratacion_Modo = "alta"
    Else
        OrganoContratacion_Modo = "edicion"
    End If
End Function

' --- OrganoContratacion_IDEntidad ------------------------------------------------------
' Returns the ID of the entity being edited (only valid in edicion mode).
Private Function OrganoContratacion_IDEntidad() As String
    If m_ObjOrganoContratacionActivo Is Nothing Then
        OrganoContratacion_IDEntidad = ""
    Else
        If m_ObjOrganoContratacionActivo.IDOrganoContratacion <> "" Then
            OrganoContratacion_IDEntidad = CStr(m_ObjOrganoContratacionActivo.IDOrganoContratacion)
        Else
            OrganoContratacion_IDEntidad = ""
        End If
    End If
End Function

' --- OrganoContratacion_RenderResultadoInicializar ------------------------------------
' Renders the JSON payload from Abrir_Inicializar into the form's controls.
' Reads payload.titulo, payload.modo, payload.hasEntidad, payload.entidad.
Private Sub OrganoContratacion_RenderResultadoInicializar(ByVal p_Payload As Object)
    If p_Payload Is Nothing Then Exit Sub

    Dim titulo As String
    titulo = CStr(p_Payload("titulo"))
    Me.lblTitulo.Caption = titulo

    ' If entidad payload exists, render values into the controls.
    If p_Payload.Exists("entidad") Then
        Dim entity As Object
        Set entity = p_Payload("entidad")
        If Not entity Is Nothing Then
            If entity.Exists("OrganoContratacion") Then Me.OrganoContratacion.Value = CStr(entity("OrganoContratacion"))
            If entity.Exists("DESCRIPCION") Then Me.DESCRIPCION.Value = CStr(entity("DESCRIPCION"))
        End If
    End If

    ' Enable ComandoRegistrar based on admin flag (matches legacy behavior).
    FormInteraction_EstablecerPropiedad Me, "ComandoRegistrar", "Enabled", OrganoContratacion_EsAdministrador()
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
    Dim modo As String
    modo = OrganoContratacion_Modo()

    Dim idEntidad As String
    idEntidad = OrganoContratacion_IDEntidad()

    Dim p_Entidad As Object
    Set p_Entidad = Nothing

    Dim initJson As String
    initJson = modOrganoContratacionAltaHelper.OrganoContratacionAlta_Abrir_Inicializar(modo, idEntidad, p_Entidad, m_Error)
    If m_Error <> "" Then Err.Raise 1000

    Dim initParsed As Object
    Set initParsed = JsonConverter.ParseJson(initJson)
    If initParsed("ok") = True Then
        OrganoContratacion_RenderResultadoInicializar initParsed("payload")
    End If

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


Private Sub cmdSalir_Click()
    On Error Resume Next
    ' Clean up the project-conventional global so the next open starts clean.
    Dim cerrarJson As String
    cerrarJson = modOrganoContratacionAltaHelper.OrganoContratacionAlta_Cerrar(m_Error)
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub


Private Sub ComandoAyuda_Click()
    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents

    m_Error = ""
    AbrirAyuda m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If

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


' --- ComandoRegistrar: extract values -> verify -> register -> raise event -------
Private Sub ComandoRegistrar_Click()
    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents

    MostrarPopupProgreso "Procesando...", "Guardando..."

    m_Error = ""

    Dim valores As Object
    Set valores = OrganoContratacion_ExtraerValores()

    Dim p_Error As String
    Dim json As String
    json = modOrganoContratacionAltaHelper.OrganoContratacionAlta_Registrar(valores, p_Error)
    If p_Error <> "" Then Err.Raise 1000

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If parsed("ok") <> True Then Err.Raise 1001, , "helper returned ok=false"

    Dim modo As String
    modo = CStr(parsed("payload")("modo"))

    ' Raise the appropriate event for the parent form.
    If modo = "alta" Then
        RaiseEvent Alta
    ElseIf modo = "edicion" Then
        RaiseEvent Editado
    End If

    VBA.DoEvents
    CerrarPopupProgreso
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub

errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoRegistrar_Click se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        On Error Resume Next
        CorreoAlAdministrador m_Error
        On Error GoTo 0
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub