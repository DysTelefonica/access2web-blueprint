Option Compare Database
Option Explicit

' Form_FormComercial — REWORK (2026-06-26)
' Thin UI wiring for the Comercial Alta/Edición form.
'
' Implements access-vba-e2e-methodology rule #1: forms are thin UI wiring.
' Every handler does exactly: read controls -> call helper -> render result.
'
' Anti-pattern removed:
'   - helpers received `ByRef p_Form As Object` and read controls via p_Form.Controls(...)
'   - tests opened the real form and asserted against p_Form.Controls(...).Value
'   - test fixture `Test_ComercialAltaHelper_OpenForm(p_Form, p_Error)` blocked COM
'
' New contract:
'   - All helpers in modComercialAltaHelper accept PURE DATA (no Form refs).
'   - Four helper calls back the form: Abrir_Inicializar, VerificarCambios,
'     Registrar, Cerrar.
'   - The form's parent sets the project-conventional global `m_ObjComercialActivo`
'     when opening in edicion mode.

' === Module-level state (all declarations at top per vba-access §10.1) ==============

' Public events the parent form listens to (raised after Registrar succeeds).
Public Event Alta()
Public Event Editado()

' Error accumulator for handler traps (Telefónica D&S convention).
Private m_Error As String


' === Local helpers (all at top per vba-access §10.1) =================================

' --- Comercial_ExtraerValores -------------------------------------------------------
' Reads current control values into a Dictionary {Comercial, DESCRIPCION}.
Private Function Comercial_ExtraerValores() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("Comercial") = CStr(Nz(Me.Comercial.Value, ""))
    d("DESCRIPCION") = CStr(Nz(Me.DESCRIPCION.Value, ""))
    Set Comercial_ExtraerValores = d
End Function

' --- Comercial_EsAdministrador -----------------------------------------------------
Private Function Comercial_EsAdministrador() As Boolean
    If EsAdministrador = EnumSiNo.Sí Then
        Comercial_EsAdministrador = True
    Else
        Comercial_EsAdministrador = False
    End If
End Function

' --- Comercial_Modo -----------------------------------------------------------------
Private Function Comercial_Modo() As String
    If m_ObjComercialActivo Is Nothing Then
        Comercial_Modo = "alta"
    Else
        Comercial_Modo = "edicion"
    End If
End Function

' --- Comercial_IDEntidad ------------------------------------------------------------
Private Function Comercial_IDEntidad() As String
    If m_ObjComercialActivo Is Nothing Then
        Comercial_IDEntidad = ""
    Else
        If m_ObjComercialActivo.IDComercial <> "" Then
            Comercial_IDEntidad = CStr(m_ObjComercialActivo.IDComercial)
        Else
            Comercial_IDEntidad = ""
        End If
    End If
End Function

' --- Comercial_RenderResultadoInicializar -------------------------------------------
Private Sub Comercial_RenderResultadoInicializar(ByVal p_Payload As Object)
    If p_Payload Is Nothing Then Exit Sub

    Dim titulo As String
    titulo = CStr(p_Payload("titulo"))
    Me.lblTitulo.Caption = titulo

    If p_Payload.Exists("entidad") Then
        Dim entity As Object
        Set entity = p_Payload("entidad")
        If Not entity Is Nothing Then
            If entity.Exists("Comercial") Then Me.Comercial.Value = CStr(entity("Comercial"))
            If entity.Exists("DESCRIPCION") Then Me.DESCRIPCION.Value = CStr(entity("DESCRIPCION"))
        End If
    End If

    FormInteraction_EstablecerPropiedad Me, "ComandoRegistrar", "Enabled", Comercial_EsAdministrador()
End Sub


' === Form lifecycle / event handlers ===================================================

Private Sub Form_Open(Cancel As Integer)
    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents

    m_Error = ""
    Ajustar Me

    Dim modo As String
    modo = Comercial_Modo()

    Dim idEntidad As String
    idEntidad = Comercial_IDEntidad()

    Dim p_Entidad As Object
    Set p_Entidad = Nothing

    Dim initJson As String
    initJson = modComercialAltaHelper.ComercialAlta_Abrir_Inicializar(modo, idEntidad, p_Entidad, m_Error)
    If m_Error <> "" Then Err.Raise 1000

    Dim initParsed As Object
    Set initParsed = JsonConverter.ParseJson(initJson)
    If initParsed("ok") = True Then
        Comercial_RenderResultadoInicializar initParsed("payload")
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
    Dim cerrarJson As String
    cerrarJson = modComercialAltaHelper.ComercialAlta_Cerrar(m_Error)
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


' --- ComandoRegistrar: extract values -> register -> raise event -----------------
Private Sub ComandoRegistrar_Click()
    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents

    MostrarPopupProgreso "Procesando...", "Guardando..."

    m_Error = ""

    Dim valores As Object
    Set valores = Comercial_ExtraerValores()

    Dim p_Error As String
    Dim json As String
    json = modComercialAltaHelper.ComercialAlta_Registrar(valores, p_Error)
    If p_Error <> "" Then Err.Raise 1000

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If parsed("ok") <> True Then Err.Raise 1001, , "helper returned ok=false"

    Dim modo As String
    modo = CStr(parsed("payload")("modo"))

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
