Option Compare Database
Option Explicit

' Form_FormExpediente — THINNED (REWORK 2026-06-26, Phase 3.2b / PR-8b)
' Pure UI orchestration. Business logic lives in modExpedienteHelper.
' See docs/audit/form-expediente-pure-data.md for the audit.
'
' Removed:
'   - Inline Caption/AllowEdits decision in Form_Load (helper computes).
'   - Inline label-setting + button-visibility in EstablecerDatos.
'   - Inline dirty-state check in GuardarPestanaActivaAntesDeNavegar (helper decides).
'
' Preserved (UI orchestration only):
'   - 7 tabX_Click handlers (orchestrate save-before-navigate).
'   - ComandoActualizarCompleto_Click — DAO call.
'   - ComandoRegistrar_Click — save + popup.
'   - ComandoAyuda_Click — AbrirAyuda.
'   - cmdSalir_Click — save + close.
'   - Form_Load — backend init + DAO lookup + form wiring.
'   - Public method EstablecerDatos (signature preserved).
'   - Public events AltaExpediente, ExpedienteEditado.

Public Event AltaExpediente(m_expediente As Expediente)
Public Event ExpedienteEditado(m_expediente As Expediente)

Private m_Error As String
Private m_DicEstadoPestanas As Scripting.Dictionary
Private m_IDExpedienteAbrir As String

Private Function GuardarPestanaActivaSiEsGeneralOFechas( _
                                    ByVal p_RequierePestanaGuardable As Boolean, _
                                    Optional ByRef p_Error As String _
                                    ) As Boolean
    Dim m_form As Form
    Dim m_IDExpediente As String
    Dim m_NombrePestana As String
    Dim m_HayCambios As Boolean
    Dim m_DebeGuardar As Boolean
    Dim m_JSON As String
    Dim m_Parsed As Object

    On Error GoTo errores

    p_Error = ""

    Set m_form = getFormPestanaActiva(Me, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If

    m_NombrePestana = NombrePestanaDesdeSourceObject(m_form.Name)
    m_HayCambios = HayCambiosFormularioPestana(m_form, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If

    ' Delegate the must-save decision to the helper.
    m_JSON = modExpedienteHelper.Expediente_DebeGuardarPestana( _
        m_NombrePestana, CBool(AbiertoParaEditar), CBool(Me.AllowEdits), m_HayCambios, p_Error)
    If p_Error <> "" Then Err.Raise 1000
    Set m_Parsed = JsonConverter.ParseJson(m_JSON)("payload")
    m_DebeGuardar = CBool(m_Parsed("debeGuardar"))

    If Not m_DebeGuardar Then
        If p_RequierePestanaGuardable Then
            p_Error = "Nada que guardar"
            Err.Raise 1000
        End If
        Exit Function
    End If

    If m_ObjExpedienteDTOActivo Is Nothing Then
        p_Error = "No se ha establecido el expediente activo"
        Err.Raise 1000
    End If
    If m_ObjExpedienteDTOActivo.Expediente Is Nothing Then
        p_Error = "No se ha establecido el expediente activo"
        Err.Raise 1000
    End If

    m_IDExpediente = m_ObjExpedienteDTOActivo.Expediente.IDExpediente
    GuardarPestanaActiva frmPadre:=Me, p_IDExpediente:=m_IDExpediente, p_Error:=p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If

    ActualizarSnapshotFormularioPestana frm:=m_form, p_Error:=p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If

    GuardarPestanaActivaSiEsGeneralOFechas = True

    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método GuardarPestanaActivaSiEsGeneralOFechas ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Private Sub GuardarPestanaActivaAntesDeNavegar()
    Dim m_JSON As String
    Dim m_Parsed As Object

    On Error GoTo errores

    m_Error = ""

    ' Delegate to helper to decide whether to attempt save.
    m_JSON = modExpedienteHelper.Expediente_Tab_Seleccionar_Guardar( _
        CBool(AbiertoParaEditar), CBool(Me.AllowEdits), m_Error)
    If m_Error <> "" Then Err.Raise 1000
    Set m_Parsed = JsonConverter.ParseJson(m_JSON)("payload")
    If Not CBool(m_Parsed("intentarGuardar")) Then
        Exit Sub
    End If

    GuardarPestanaActivaSiEsGeneralOFechas p_RequierePestanaGuardable:=False, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If

    Exit Sub

errores:
    If Err.Number <> 1000 Then
        m_Error = "El método GuardarPestanaActivaAntesDeNavegar ha devuelto el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    MsgBox m_Error, vbCritical, "Error"
End Sub

Private Sub tabGeneral_Click()
    GuardarPestanaActivaAntesDeNavegar
End Sub

Private Sub tabFechas_Click()
    GuardarPestanaActivaAntesDeNavegar
End Sub

Private Sub tabEntidades_Click()
    GuardarPestanaActivaAntesDeNavegar
End Sub

Private Sub tabSuministradores_Click()
    GuardarPestanaActivaAntesDeNavegar
End Sub

Private Sub tabAnexos_Click()
    GuardarPestanaActivaAntesDeNavegar
End Sub

Private Sub tabHitos_Click()
    GuardarPestanaActivaAntesDeNavegar
End Sub

Private Sub tabModificados_Click()
    GuardarPestanaActivaAntesDeNavegar
End Sub

Private Sub ComandoActualizarCompleto_Click()
    Dim m_CaptionOriginal As String
    Dim m_ExpOp As ExpedienteOperaciones
    On Error GoTo errores

    m_CaptionOriginal = Me.Caption
    If AntiSpamIsOperationInProgress(Me, m_Error) Then Exit Sub
    AntiSpamEnterOperation Me, m_CaptionOriginal, "ComandoRegistrar", "ComandoActualizarCompleto"

    m_Error = ""
    DoCmd.Hourglass True

    Set m_ExpOp = New ExpedienteOperaciones
    With m_ExpOp
        Set .Expediente = m_ObjExpedienteDTOActivo.Expediente
        .RegistrarExpEntidades p_Error:=m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End With

    DoCmd.Hourglass False
    AntiSpamExitOperation Me, m_CaptionOriginal, m_Error
    Exit Sub
errores:
     AntiSpamExitOperation Me, m_CaptionOriginal, m_Error
     If Err.Number <> 1000 Then
        m_Error = "El método ComandoActualizarCompleto_Click ha devuelto el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoAyuda_Click()
    On Error GoTo errores
    m_Error = ""
    DoCmd.Hourglass True
    AbrirAyuda m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    DoCmd.Hourglass False
    Exit Sub
errores:
     If Err.Number <> 1000 Then
        m_Error = "El método ComandoAyudaFormulario_Click ha devuelto el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoRegistrar_Click()
    Dim m_CaptionOriginal As String
    On Error GoTo errores

    m_CaptionOriginal = Me.Caption
    If AntiSpamIsOperationInProgress(Me, m_Error) Then Exit Sub
    AntiSpamEnterOperation Me, m_CaptionOriginal, "ComandoRegistrar", "ComandoActualizarCompleto"

    m_Error = ""

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents

    MostrarPopupProgreso "Guardando...", "Guardando expediente..."

    GuardarPestanaActivaSiEsGeneralOFechas p_RequierePestanaGuardable:=True, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If

    CerrarPopupProgreso
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents

    AntiSpamExitOperation Me, m_CaptionOriginal, m_Error
    Exit Sub

errores:
    CerrarPopupProgreso
    AntiSpamExitOperation Me, m_CaptionOriginal, m_Error
    If Err.Number <> 1000 Then
        m_Error = "El método ComandoRegistrar_Click ha devuelto el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    MsgBox m_Error, vbCritical, "Error"
End Sub
Private Sub cmdSalir_Click()
    On Error GoTo errores
    m_Error = ""
    If AbiertoParaEditar = False Then
        DoCmd.Close acForm, Me.Name, acSaveNo
        Exit Sub
    End If
    If Me.AllowEdits = False Then
        DoCmd.Close acForm, Me.Name, acSaveNo
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    GuardarPestanaActivaSiEsGeneralOFechas p_RequierePestanaGuardable:=False, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    DoCmd.Close acForm, Me.Name, acSaveNo
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents

    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al cmdSalir_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")

End Sub

Public Sub Form_Load()

    Dim m_Tipo As String
    Dim m_BackendActivo As String
    Dim m_JSON As String
    Dim m_Parsed As Object
    Dim m_HasExpediente As Boolean

    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""

    ' Leer configuración de backend antes de cualquier operación de datos
    LeeConfiguracionLocal m_Error
    If m_Error <> "" Then
        MsgBox "No se pudo inicializar la configuración de backend: " & m_Error, vbCritical, "EXPE-DYS - Formulario expediente"
        Exit Sub
    End If

    m_BackendActivo = Nz(Application.TempVars("BackendActivo"), "")
    If m_BackendActivo = "" Then
        Err.Raise 1000, , "No se pudo identificar el entorno de backend activo (BackendActivo)."
    End If

    ' Determine hasExpediente for helper signature.
    m_HasExpediente = (Not m_ObjExpedienteDTOActivo Is Nothing) And _
                      (Not m_ObjExpedienteDTOActivo.Expediente Is Nothing)

    ' Delegate the Form_Load init decision to the helper.
    m_JSON = modExpedienteHelper.Expediente_Form_Load_Init( _
        Nz(Me.OpenArgs(), ""), (EsAdministrador = EnumSiNo.Sí), m_BackendActivo, m_HasExpediente, m_Error)
    If m_Error <> "" Then Err.Raise 1000

    Set m_Parsed = JsonConverter.ParseJson(m_JSON)("payload")

    ' Apply the helper's decisions.
    If CBool(m_Parsed("abiertoParaEditar")) Then
        AbiertoParaEditar = True
    Else
        AbiertoParaEditar = False
    End If

    Me.InsideHeight = CLng(m_Parsed("insideHeight"))
    Me.InsideWidth = CLng(m_Parsed("insideWidth"))
    Me.ComandoRegistrar.Enabled = CBool(m_Parsed("comandoRegistrarEnabled"))
    Me.AllowEdits = CBool(m_Parsed("allowEdits"))
    Me.Caption = CStr(m_Parsed("caption"))

    If AbiertoParaEditar = True Then
        If EsAdministrador = EnumSiNo.Sí Then
            Me.ComandoRegistrar.Enabled = True
            Me.AllowEdits = True
        End If
    Else
        Me.ComandoRegistrar.Enabled = False
        Me.AllowEdits = False
    End If

    EstablecerDatos m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If

    ' Extraer ID de OpenArgs (formato "ID|Fechas" o "ID")
    Dim m_Args As String
    m_Args = Nz(Me.OpenArgs(), "")
    If InStr(m_Args, "|") > 0 Then
        m_IDExpedienteAbrir = Split(m_Args, "|")(0)
    ElseIf m_Args <> "" And m_Args <> "Sololectura" Then
        m_IDExpedienteAbrir = m_Args
    End If

    ' Si m_ObjExpedienteDTOActivo no está setado (ej: abierto via AbrirExpedienteEnFechas),
    ' cargarlo desde el ID extraído de OpenArgs
    If m_IDExpedienteAbrir <> "" And m_ObjExpedienteDTOActivo Is Nothing Then
        Set m_ObjExpedienteDTOActivo = EstablecerDTODeExpediente( _
            p_Expediente:=constructor.getExpediente(p_IDExpediente:=m_IDExpedienteAbrir, p_Error:=m_Error), _
            p_ParaCopia:=EnumSiNo.No, _
            p_Error:=m_Error)
        If m_Error <> "" Then Err.Raise 1000
        ' Recargar controles con los datos del DTO recién cargado
        Dim m_formActivo As Form
        Set m_formActivo = getFormPestanaActiva(Me, m_Error)
        If m_Error = "" Then
            If m_formActivo.Name = "FormExpedienteGeneral" Then
                SetFormGeneralFromExpediente p_IDExpediente:=m_IDExpedienteAbrir, frm:=m_formActivo, p_Error:=m_Error
            End If
        End If
        If m_Error <> "" Then Err.Raise 1000
    End If

    ' Navegar a Fechas si OpenArgs lo indica
    NavegarAPestanaSiSolicitada Me, Nz(Me.OpenArgs(), ""), m_Error
    If m_Error <> "" Then
        ' Loggear o mostrar warning — no bloqueante pero al menos visible en debug
        Debug.Print "WARNING NavegarAPestanaSiSolicitada: " & m_Error
    End If

    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Load se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Public Function EstablecerDatos(Optional ByRef p_Error As String) As String

    Dim m_JSON As String
    Dim m_Parsed As Object
    Dim m_DTO As Object
    On Error GoTo errores

    If m_ObjExpedienteDTOActivo Is Nothing Then
        p_Error = "No se sabe para qué tipo de registro se ha abierto"
        Err.Raise 1000
    End If
    If m_ObjExpedienteDTOActivo.Expediente Is Nothing Then
        p_Error = "No se sabe para qué tipo de registro se ha abierto"
        Err.Raise 1000
    End If
    With m_ObjExpedienteDTOActivo.Expediente
        If .TIpo = "" Then
            p_Error = "No se sabe para qué tipo de registro se ha abierto"
            Err.Raise 1000
        End If
        If .Derivable = Empty Then
            p_Error = "No se sabe para qué tipo de registro se ha abierto"
            Err.Raise 1000
        End If
    End With

    ' Wrap the live DTO into a Dictionary so the helper sees a stable shape.
    Set m_DTO = CreateObject("Scripting.Dictionary")
    Set m_DTO("Expediente") = m_ObjExpedienteDTOActivo.Expediente

    ' Delegate label-setting + button-visibility to the helper.
    m_JSON = modExpedienteHelper.Expediente_EstablecerDatos(m_DTO, p_Error)
    If p_Error <> "" Then Exit Function

    Set m_Parsed = JsonConverter.ParseJson(m_JSON)("payload")

    Me.lblTitulo.Caption = CStr(m_Parsed("lblTitulo"))

    If CBool(m_Parsed("lblUltimaModificacionVisible")) Then
        Me.lblUltimaModificacion.Visible = True
        Me.lblUltimaModificacion.Caption = CStr(m_Parsed("lblUltimaModificacion"))
    Else
        Me.lblUltimaModificacion.Visible = False
    End If

    If CBool(m_Parsed("comandoActualizarCompletoVisible")) Then
        Me.ComandoActualizarCompleto.Visible = True
    Else
        Me.ComandoActualizarCompleto.Visible = False
    End If

    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerDatos a devuelto el error: " & vbNewLine & Err.Description
    End If
End Function