Option Compare Database
Option Explicit

' Form_FormOficinaPrograma — REWORK (2026-06-26)
' Thin UI wiring for the OficinaPrograma Alta/Edición form.
'
' Implements access-vba-e2e-methodology rule #1: forms are thin UI wiring.
' Every handler does exactly: read controls -> call helper -> render result.
'
' Anti-pattern removed (was in PR #31 commit c86d460):
'   - helpers received `ByRef p_Form As Object` and read controls via p_Form.Controls(...)
'   - tests opened the real form and asserted against p_Form.Controls(...).Value
'   - test fixture `Test_OficinaProgramaAltaHelper_OpenForm(p_Form, p_Error)` blocked COM
'
' New contract (per the rework):
'   - All helpers in modOficinaProgramaAltaHelper accept PURE DATA (no Form refs).
'   - Form extracts control values BEFORE calling helper; renders helper JSON AFTER.
'   - Four helper calls back the form: Abrir_Inicializar, VerificarCambios,
'     Registrar, Cerrar.
'   - The form's parent sets the project-conventional global `m_ObjOficinaProgramaActiva`
'     when opening in edicion mode (via Abrir_Inicializar which sets it internally when loading).
'
' Form controls: `OficinaPrograma`, `DESCRIPCION`, `lblTitulo`, `ComandoRegistrar`.

' === Module-level state (all declarations at top per vba-access §10.1) ==============

' Public events the parent form listens to (raised after Registrar succeeds).
Public Event Alta()
Public Event Editado()

' Error accumulator for handler traps (Telefónica D&S convention).
Private m_Error As String

' Snapshot of the entity at form open (populated by Abrir_Inicializar in edicion mode).
' Re-passed to VerificarCambios / Cerrar.
Private m_OficinaProgramaAlInicio As Object


' === Local helpers (all at top per vba-access §10.1) =================================

' --- OficinaPrograma_ExtraerValores -----------------------------------------------
' Reads current control values into a Dictionary {OficinaPrograma, DESCRIPCION}.
Private Function OficinaPrograma_ExtraerValores() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("OficinaPrograma") = CStr(Nz(Me.OficinaPrograma.Value, ""))
    d("DESCRIPCION") = CStr(Nz(Me.DESCRIPCION.Value, ""))
    Set OficinaPrograma_ExtraerValores = d
End Function

' --- OficinaPrograma_EsAdministrador ----------------------------------------------
Private Function OficinaPrograma_EsAdministrador() As Boolean
    If EsAdministrador = EnumSiNo.Sí Then
        OficinaPrograma_EsAdministrador = True
    Else
        OficinaPrograma_EsAdministrador = False
    End If
End Function

' --- OficinaPrograma_Modo ----------------------------------------------------------
Private Function OficinaPrograma_Modo() As String
    If m_ObjOficinaProgramaActiva Is Nothing Then
        OficinaPrograma_Modo = "alta"
    Else
        OficinaPrograma_Modo = "edicion"
    End If
End Function

' --- OficinaPrograma_IDEntidad -----------------------------------------------------
Private Function OficinaPrograma_IDEntidad() As String
    If m_ObjOficinaProgramaActiva Is Nothing Then
        OficinaPrograma_IDEntidad = ""
    Else
        If m_ObjOficinaProgramaActiva.IDOficinaPrograma <> "" Then
            OficinaPrograma_IDEntidad = CStr(m_ObjOficinaProgramaActiva.IDOficinaPrograma)
        Else
            OficinaPrograma_IDEntidad = ""
        End If
    End If
End Function

' --- OficinaPrograma_RenderResultadoInicializar -----------------------------------
Private Sub OficinaPrograma_RenderResultadoInicializar(ByVal p_Payload As Object)
    If p_Payload Is Nothing Then Exit Sub

    Dim titulo As String
    titulo = CStr(p_Payload("titulo"))
    Me.lblTitulo.Caption = titulo

    ' If entidad payload exists, render values into the controls.
    If p_Payload.Exists("entidad") Then
        Dim entity As Object
        Set entity = p_Payload("entidad")
        If Not entity Is Nothing Then
            If entity.Exists("OficinaPrograma") Then Me.OficinaPrograma.Value = CStr(entity("OficinaPrograma"))
            If entity.Exists("DESCRIPCION") Then Me.DESCRIPCION.Value = CStr(entity("DESCRIPCION"))
        End If
    End If

    ' Enable ComandoRegistrar based on admin flag (matches legacy behavior).
    FormInteraction_EstablecerPropiedad Me, "ComandoRegistrar", "Enabled", OficinaPrograma_EsAdministrador()
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
    modo = OficinaPrograma_Modo()

    Dim idEntidad As String
    idEntidad = OficinaPrograma_IDEntidad()

    Set m_OficinaProgramaAlInicio = Nothing

    Dim initJson As String
    initJson = modOficinaProgramaAltaHelper.OficinaProgramaAlta_Abrir_Inicializar( _
        modo, idEntidad, m_OficinaProgramaAlInicio, m_Error)
    If m_Error <> "" Then Err.Raise 1000

    Dim initParsed As Object
    Set initParsed = JsonConverter.ParseJson(initJson)
    If initParsed("ok") = True Then
        OficinaPrograma_RenderResultadoInicializar initParsed("payload")
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
    cerrarJson = modOficinaProgramaAltaHelper.OficinaProgramaAlta_Cerrar(m_Error)
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
    Set valores = OficinaPrograma_ExtraerValores()

    Dim p_Error As String
    Dim json As String
    json = modOficinaProgramaAltaHelper.OficinaProgramaAlta_Registrar(valores, p_Error)
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
