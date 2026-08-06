Option Compare Database
Option Explicit

' Form_FormRAC — REWORK (2026-06-26)
' Thin UI wiring for the RAC Alta/Edición form.
'
' Implements access-vba-e2e-methodology rule #1: forms are thin UI wiring.
' Every handler does exactly: read controls -> call helper -> render result.
'
' Anti-pattern removed (was in PR #29 commit 979a57a):
'   - helpers received `ByRef p_Form As Object` and read controls via p_Form.Controls(...)
'   - tests opened the real form and asserted against p_Form.Controls(...).Value
'   - test fixture `Test_RACHelper_OpenForm(p_Form, p_Error)` blocked COM
'
' New contract (per the rework):
'   - All helpers in modRACHelper accept PURE DATA (no Form refs).
'   - Form extracts control values BEFORE calling helper; renders helper JSON AFTER.
'   - Four helper calls back the form: Abrir_Inicializar, Verificar_Cambios,
'     Registrar, Cerrar.
'   - Helpers return JSON that the form parses via JsonConverter.ParseJson and
'     renders into controls.
'   - The form sets the project-conventional global `m_ObjRACActivo` when opening
'     in edicion mode (via Abrir_Inicializar which sets it internally when loading).

' === Module-level state (all declarations at top per vba-access §10.1) ==============

' Public events the parent form listens to (raised after Registrar succeeds).
Public Event Alta()
Public Event Editado()

' Error accumulator for handler traps (Telefónica D&S convention).
Private m_Error As String


' === Local helpers (all at top per vba-access §10.1) =================================

' --- RAC_ExtraerValores --------------------------------------------------------------
' Reads current control values into a Dictionary {RAC, CORREO, DESCRIPCION}.
' Used by Verificar_Cambios and Registrar before calling the helper.
Private Function RAC_ExtraerValores() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("RAC") = CStr(Nz(Me.RAC.Value, ""))
    d("CORREO") = CStr(Nz(Me.CORREO.Value, ""))
    d("DESCRIPCION") = CStr(Nz(Me.DESCRIPCION.Value, ""))
    Set RAC_ExtraerValores = d
End Function

' --- RAC_EsAdministrador -------------------------------------------------------------
' Reads the global admin flag set during EVE().
Private Function RAC_EsAdministrador() As Boolean
    If EsAdministrador = EnumSiNo.Sí Then
        RAC_EsAdministrador = True
    Else
        RAC_EsAdministrador = False
    End If
End Function

' --- RAC_Modo -------------------------------------------------------------------------
' Returns the modo (alta | edicion) based on whether m_ObjRACActivo is set
' (set by the parent form's Alta/Editar handlers before opening this form).
Private Function RAC_Modo() As String
    If m_ObjRACActivo Is Nothing Then
        RAC_Modo = "alta"
    Else
        RAC_Modo = "edicion"
    End If
End Function

' --- RAC_IDEntidad -------------------------------------------------------------------
' Returns the ID of the entity being edited (only valid in edicion mode).
Private Function RAC_IDEntidad() As String
    If m_ObjRACActivo Is Nothing Then
        RAC_IDEntidad = ""
    Else
        If m_ObjRACActivo.IDRAC <> "" Then
            RAC_IDEntidad = CStr(m_ObjRACActivo.IDRAC)
        Else
            RAC_IDEntidad = ""
        End If
    End If
End Function

' --- RAC_RenderResultadoInicializar --------------------------------------------------
' Renders the JSON payload from Abrir_Inicializar into the form's controls.
' Reads payload.titulo, payload.modo, payload.hasEntidad, payload.entidad.
Private Sub RAC_RenderResultadoInicializar(ByVal p_Payload As Object)
    If p_Payload Is Nothing Then Exit Sub

    Dim titulo As String
    titulo = CStr(p_Payload("titulo"))
    Me.lblTitulo.Caption = titulo

    ' If entidad payload exists, render values into the controls.
    If p_Payload.Exists("entidad") Then
        Dim entity As Object
        Set entity = p_Payload("entidad")
        If Not entity Is Nothing Then
            If entity.Exists("RAC") Then Me.RAC.Value = CStr(entity("RAC"))
            If entity.Exists("CORREO") Then Me.CORREO.Value = CStr(entity("CORREO"))
            If entity.Exists("DESCRIPCION") Then Me.DESCRIPCION.Value = CStr(entity("DESCRIPCION"))
        End If
    End If

    ' Enable ComandoRegistrar based on admin flag (matches legacy behavior).
    FormInteraction_EstablecerPropiedad Me, "ComandoRegistrar", "Enabled", RAC_EsAdministrador()
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
    modo = RAC_Modo()

    Dim idEntidad As String
    idEntidad = RAC_IDEntidad()

    Dim p_Entidad As Object
    Set p_Entidad = Nothing

    Dim initJson As String
    initJson = modRACHelper.RAC_Abrir_Inicializar(modo, idEntidad, p_Entidad, m_Error)
    If m_Error <> "" Then Err.Raise 1000

    Dim initParsed As Object
    Set initParsed = JsonConverter.ParseJson(initJson)
    If initParsed("ok") = True Then
        RAC_RenderResultadoInicializar initParsed("payload")
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
    cerrarJson = modRACHelper.RAC_Cerrar(m_Error)
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub


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


' --- ComandoRegistrar: extract values -> verify -> register -> raise event -------
Private Sub ComandoRegistrar_Click()
    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents

    MostrarPopupProgreso "Procesando...", "Guardando..."

    m_Error = ""

    Dim valores As Object
    Set valores = RAC_ExtraerValores()

    Dim p_Error As String
    Dim json As String
    json = modRACHelper.RAC_Registrar(valores, p_Error)
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
