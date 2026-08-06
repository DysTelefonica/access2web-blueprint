Option Compare Database
Option Explicit

' Form_FormSuministrador — REWORK (2026-06-26)
' Thin UI wiring for the Suministrador Alta/Edición form.
'
' Implements access-vba-e2e-methodology rule #1: forms are thin UI wiring.
' Every handler does exactly: read controls -> call helper -> render result.
'
' Anti-pattern removed (was in PR #31 commit c86d460):
'   - helpers received `ByRef p_Form As Object` and read controls via p_Form.Controls(...)
'   - tests opened the real form and asserted against p_Form.Controls(...).Value
'   - test fixture `Test_SuministradorAltaHelper_OpenForm(p_Form, p_Error)` blocked COM
'
' New contract (per the rework):
'   - All helpers in modSuministradorAltaHelper accept PURE DATA (no Form refs).
'   - Form extracts control values BEFORE calling helper; renders helper JSON AFTER.
'   - Four helper calls back the form: Abrir_Inicializar, VerificarCambios,
'     Registrar, Cerrar.
'   - The form's parent sets the project-conventional global `m_ObjSuministradorActivo`
'     when opening in edicion mode (via Abrir_Inicializar which sets it internally when loading).
'
' Field references: Nombre, Nemotecnico, CIF, Direccion, Ciudad, CP, TramitadoraHPS
' (7 fields), lblTitulo, ComandoRegistrar.

' === Module-level state (all declarations at top per vba-access §10.1) ==============

' Public events the parent form listens to (raised after Registrar succeeds).
' Alta(p_ID) carries the new ID; Editado() has no args.
Public Event Alta(p_ID As String)
Public Event Editado()

' Error accumulator for handler traps (Telefónica D&S convention).
Private m_Error As String

' Snapshot of the entity at form open (used by VerificarCambios). Per the
' OrganoContratacionAlta pattern, this lives in the helper's payload.entidad
' (populated by Abrir_Inicializar) and is re-passed to VerificarCambios.
' Kept as Object so the form can hand it to the helper directly without re-loading.
Private m_SuministradorAlInicio As Object


' === Local helpers (all at top per vba-access §10.1) =================================

' --- Suministrador_ExtraerValores -----------------------------------------------
' Reads current control values into a Dictionary with the 7 Suministrador fields.
' Used by ComandoRegistrar_Click before calling the helper.
Private Function Suministrador_ExtraerValores() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("Nombre") = CStr(Nz(Me.Nombre.Value, ""))
    d("Nemotecnico") = CStr(Nz(Me.Nemotecnico.Value, ""))
    d("CIF") = CStr(Nz(Me.CIF.Value, ""))
    d("Direccion") = CStr(Nz(Me.Direccion.Value, ""))
    d("Ciudad") = CStr(Nz(Me.Ciudad.Value, ""))
    d("CP") = CStr(Nz(Me.CP.Value, ""))
    d("TramitadoraHPS") = CStr(Nz(Me.TramitadoraHPS.Value, ""))
    Set Suministrador_ExtraerValores = d
End Function

' --- Suministrador_EsAdministrador -----------------------------------------------
Private Function Suministrador_EsAdministrador() As Boolean
    If EsAdministrador = EnumSiNo.Sí Then
        Suministrador_EsAdministrador = True
    Else
        Suministrador_EsAdministrador = False
    End If
End Function

' --- Suministrador_Modo ------------------------------------------------------------
' Returns the modo (alta | edicion) based on whether m_ObjSuministradorActivo is set
' (set by the parent form's Alta/Editar handlers before opening this form).
Private Function Suministrador_Modo() As String
    If m_ObjSuministradorActivo Is Nothing Then
        Suministrador_Modo = "alta"
    Else
        Suministrador_Modo = "edicion"
    End If
End Function

' --- Suministrador_IDEntidad ------------------------------------------------------
Private Function Suministrador_IDEntidad() As String
    If m_ObjSuministradorActivo Is Nothing Then
        Suministrador_IDEntidad = ""
    Else
        If m_ObjSuministradorActivo.IDSuministrador <> "" Then
            Suministrador_IDEntidad = CStr(m_ObjSuministradorActivo.IDSuministrador)
        Else
            Suministrador_IDEntidad = ""
        End If
    End If
End Function

' --- Suministrador_RenderResultadoInicializar ------------------------------------
' Renders the JSON payload from Abrir_Inicializar into the form's controls.
' Reads payload.titulo, payload.modo, payload.hasEntidad, payload.entidad.
Private Sub Suministrador_RenderResultadoInicializar(ByVal p_Payload As Object)
    If p_Payload Is Nothing Then Exit Sub

    Dim titulo As String
    titulo = CStr(p_Payload("titulo"))
    Me.lblTitulo.Caption = titulo

    ' If entidad payload exists, render values into the controls.
    If p_Payload.Exists("entidad") Then
        Dim entity As Object
        Set entity = p_Payload("entidad")
        If Not entity Is Nothing Then
            If entity.Exists("Nombre") Then Me.Nombre.Value = CStr(entity("Nombre"))
            If entity.Exists("Nemotecnico") Then Me.Nemotecnico.Value = CStr(entity("Nemotecnico"))
            If entity.Exists("CIF") Then Me.CIF.Value = CStr(entity("CIF"))
            If entity.Exists("Direccion") Then Me.Direccion.Value = CStr(entity("Direccion"))
            If entity.Exists("Ciudad") Then Me.Ciudad.Value = CStr(entity("Ciudad"))
            If entity.Exists("CP") Then Me.CP.Value = CStr(entity("CP"))
            If entity.Exists("TramitadoraHPS") Then Me.TramitadoraHPS.Value = CStr(entity("TramitadoraHPS"))
        End If
    End If

    ' Enable ComandoRegistrar based on admin flag (matches legacy behavior).
    FormInteraction_EstablecerPropiedad Me, "ComandoRegistrar", "Enabled", Suministrador_EsAdministrador()
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
    modo = Suministrador_Modo()

    Dim idEntidad As String
    idEntidad = Suministrador_IDEntidad()

    Set m_SuministradorAlInicio = Nothing

    Dim initJson As String
    initJson = modSuministradorAltaHelper.SuministradorAlta_Abrir_Inicializar( _
        modo, idEntidad, m_SuministradorAlInicio, m_Error)
    If m_Error <> "" Then Err.Raise 1000

    Dim initParsed As Object
    Set initParsed = JsonConverter.ParseJson(initJson)
    If initParsed("ok") = True Then
        Suministrador_RenderResultadoInicializar initParsed("payload")
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
    cerrarJson = modSuministradorAltaHelper.SuministradorAlta_Cerrar(m_Error)
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
    Set valores = Suministrador_ExtraerValores()

    Dim p_Error As String
    Dim json As String
    json = modSuministradorAltaHelper.SuministradorAlta_Registrar(valores, p_Error)
    If p_Error <> "" Then Err.Raise 1000

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If parsed("ok") <> True Then Err.Raise 1001, , "helper returned ok=false"

    Dim modo As String
    modo = CStr(parsed("payload")("modo"))

    ' Raise the appropriate event for the parent form.
    If modo = "alta" Then
        ' Pass the resulting ID to the Alta event (matches legacy signature).
        Dim resultingId As String
        resultingId = ""
        If m_ObjSuministradorActivo Is Nothing Then
            ' alta path: ID is in payload.
            If parsed("payload").Exists("id") Then
                resultingId = CStr(parsed("payload")("id"))
            End If
        Else
            ' edicion path: ID was already on m_ObjSuministradorActivo.
            If m_ObjSuministradorActivo.IDSuministrador <> "" Then
                resultingId = CStr(m_ObjSuministradorActivo.IDSuministrador)
            End If
        End If
        RaiseEvent Alta(resultingId)
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
