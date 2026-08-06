Option Compare Database
Option Explicit

' Form_FormExpedienteCambioTipo — THINNED (REWORK 2026-06-26, Phase 3.2b / PR-8b)
' Pure UI orchestration. Business logic lives in modExpedienteCambioTipoHelper.
' See docs/audit/form-expediente-cambio-tipo-pure-data.md for the audit.
'
' Removed:
'   - 5-branch dispatch inline in ComandoRegistrar_Click (helper computes flags).
'   - Inline padre validation (helper returns valido + error).
'   - Tipo_AfterUpdate inline Locked/Enabled toggling (helper returns flags).
'
' Preserved (UI orchestration only):
'   - Public method EstablecerDatos (signature preserved) — populates the cmb.
'   - cmdSalir_Click — DoCmd.Close.
'   - ComandoAyuda_Click — AbrirAyuda.
'   - ComandoBuscar_Click — DoCmd.OpenForm "FormExpedientesParaCambioTipo".
'   - Form_Open — Ajustar Me + EstablecerDatos.
'   - m_Form_Seleccionado — UI pass-through.
'   - Tipo_AfterUpdate — applies Locked/Enabled (helper returns the flags).
'   - Public event CambioRealizado.

Private WithEvents m_form As Form_FormExpedientesParaCambioTipo
Public Event CambioRealizado()
Private m_TipoActual As EnumTipoExpediente
Private m_Error As String

Private Sub cmdSalir_Click()

    On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Sub ComandoAyuda_Click()
    On Error GoTo errores
    m_Error = ""
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
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
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoBuscar_Click()
    On Error GoTo errores
    m_Error = ""
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    If FormularioAbierto("FormExpedientesParaCambioTipo") Then
        DoCmd.Close acForm, "FormExpedientesParaCambioTipo"
    End If

    DoCmd.OpenForm "FormExpedientesParaCambioTipo"
    If FormularioAbierto("FormExpedientesParaCambioTipo") Then
        Set m_form = Forms("FormExpedientesParaCambioTipo")
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
     If Err.Number <> 1000 Then
        m_Error = "Al ComandoBuscar_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        On Error Resume Next
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoRegistrar_Click()
    Dim m_TipoElegido As String
    Dim m_expediente As Expediente
    Dim m_ExpedienteOp As ExpedienteOperaciones
    Dim m_ExpPadre As Expediente
    Dim m_IDPadre As String
    Dim m_JSON As String
    Dim m_Parsed As Object
    Dim m_Flags As Object
    Dim m_Valido As Boolean
    Dim m_ValidationErr As String
    Dim m_EsPadreAM As Boolean
    Dim m_EsPadreLote As Boolean

    On Error GoTo errores
    m_Error = ""
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    MostrarPopupProgreso "Procesando...", "Guardando..."

    m_TipoElegido = Nz(Me.TIpo, "")

    ' Validate padre via helper. For tipos that require padre, load it
    ' first so the helper can validate without doing DAO itself.
    m_Valido = True
    m_ValidationErr = ""
    m_IDPadre = Nz(Me.IDExpedientePadre, "")

    If m_TipoElegido = "Convertir a Lote" Then
        If Len(m_IDPadre) = 0 Then
            m_Valido = False
            m_ValidationErr = "Para Convertir en Lote, hemos de rellenar el ID del Acuerdo Marco"
        Else
            Set m_ExpPadre = constructor.getExpediente(p_IDExpediente:=m_IDPadre, p_Error:=m_Error)
            If m_Error <> "" Then Err.Raise 1000
            If m_ExpPadre Is Nothing Then
                m_Valido = False
                m_ValidationErr = "El IDExpediente indicado no se encuentra registrado"
            Else
                m_EsPadreAM = (m_ExpPadre.EsAM = "Sí")
                m_EsPadreLote = (m_ExpPadre.EsLote = "Sí")
                m_JSON = modExpedienteCambioTipoHelper.ExpedienteCambioTipo_ValidarPadre( _
                    m_TipoElegido, m_IDPadre, m_EsPadreAM, m_EsPadreLote, m_Error)
                If m_Error <> "" Then Err.Raise 1000
                Set m_Parsed = JsonConverter.ParseJson(m_JSON)("payload")
                If Not CBool(m_Parsed("valido")) Then
                    m_Valido = False
                    m_ValidationErr = CStr(m_Parsed("error"))
                End If
            End If
        End If
    ElseIf m_TipoElegido = "Convertir a Basado" Then
        If Len(m_IDPadre) = 0 Then
            m_Valido = False
            m_ValidationErr = "Para Convertir en Basado, hemos de indicar un IDExpediente padre de tipo Acuerdo Marco o Lote"
        Else
            Set m_ExpPadre = constructor.getExpediente(p_IDExpediente:=m_IDPadre, p_Error:=m_Error)
            If m_Error <> "" Then Err.Raise 1000
            If m_ExpPadre Is Nothing Then
                m_Valido = False
                m_ValidationErr = "El IDExpediente indicado no se encuentra registrado"
            Else
                m_EsPadreAM = (m_ExpPadre.EsAM = "Sí")
                m_EsPadreLote = (m_ExpPadre.EsLote = "Sí")
                m_JSON = modExpedienteCambioTipoHelper.ExpedienteCambioTipo_ValidarPadre( _
                    m_TipoElegido, m_IDPadre, m_EsPadreAM, m_EsPadreLote, m_Error)
                If m_Error <> "" Then Err.Raise 1000
                Set m_Parsed = JsonConverter.ParseJson(m_JSON)("payload")
                If Not CBool(m_Parsed("valido")) Then
                    m_Valido = False
                    m_ValidationErr = CStr(m_Parsed("error"))
                End If
            End If
        End If
    ElseIf m_TipoElegido = "Convertir a Expediente Independiente" Then
        ' Must NOT have padre
        If Len(m_IDPadre) > 0 Then
            m_Valido = False
            m_ValidationErr = "Para Convertir en Expediente Independiente, no hemos de indicar un IDExpediente padre"
        End If
    End If

    If Not m_Valido Then
        m_Error = m_ValidationErr
        Err.Raise 1000
    End If

    ' Compute the flag set via helper (5-branch dispatch).
    m_JSON = modExpedienteCambioTipoHelper.ExpedienteCambioTipo_ComandoRegistrar( _
        m_ObjExpedienteActivo, m_TipoElegido, m_IDPadre, m_Error)
    If m_Error <> "" Then Err.Raise 1000

    Set m_Parsed = JsonConverter.ParseJson(m_JSON)("payload")
    Set m_Flags = m_Parsed("flagsApplied")

    ' Apply the flag set to m_ObjExpedienteActivo.
    With m_ObjExpedienteActivo
        .IDExpedientePadre = CStr(m_Flags("IDExpedientePadre"))
        .EsAM = CStr(m_Flags("EsAM"))
        .EsLote = CStr(m_Flags("EsLote"))
        .EsExpediente = CStr(m_Flags("EsExpediente"))
        .EsBasado = CStr(m_Flags("EsBasado"))
    End With

    ' Persist via ExpedienteOperaciones.
    Set m_ExpedienteOp = New ExpedienteOperaciones
    With m_ExpedienteOp
        Set .Expediente = m_ObjExpedienteActivo
        .RegistrarCambioTipo p_Error:=m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End With
    RaiseEvent CambioRealizado

    VBA.DoEvents
    CerrarPopupProgreso
    VBA.DoEvents
    DoCmd.Close acForm, Me.Name, acSaveNo
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
     If Err.Number <> 1000 Then
        m_Error = "Al ComandoCambiarTipo_Click se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        On Error Resume Next
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub Form_Open(Cancel As Integer)

    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Ajustar Me


    Me.Caption = m_TituloFormulario
    If m_ObjExpedienteActivo Is Nothing Then
        m_Error = "El formulario se ha abierto con parámetros insuficientes"
        Err.Raise 1000
    End If
    EstablecerDatos m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Open se ha producido el error num " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        On Error Resume Next
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")

End Sub
Public Function EstablecerDatos(Optional ByRef p_Error As String) As String

    Dim cmb As ComboBox
    Dim m_JSON As String
    Dim m_Parsed As Object
    Dim m_Opciones As Object
    Dim m_Key As Variant

    On Error GoTo errores

    If m_ObjExpedienteActivo Is Nothing Then
        p_Error = "El formulario se ha abierto con parámetros insuficientes"
        Err.Raise 1000
    End If

    ' Delegate tipo detection + options list to the helper.
    m_JSON = modExpedienteCambioTipoHelper.ExpedienteCambioTipo_EstablecerDatos( _
        m_ObjExpedienteActivo, p_Error)
    If p_Error <> "" Then Exit Function

    Set m_Parsed = JsonConverter.ParseJson(m_JSON)("payload")

    ' UI wiring: populate the combo box from the helper's opciones Dictionary.
    Me.IDExpedientePadre.Locked = True
    Me.IDExpedientePadre.Enabled = False
    Set cmb = Me.TIpo
    cmb.RowSource = ""
    With m_ObjExpedienteActivo
        m_TipoActual = .TipoCalculado
        Me.TIPOACTUAL = .TipoCalculadoTexto
    End With

    Set m_Opciones = m_Parsed("opciones")
    If Not m_Opciones Is Nothing Then
        For Each m_Key In m_Opciones.Keys
            cmb.AddItem CStr(m_Opciones(m_Key))
        Next m_Key
    End If

    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerDatos ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If

End Function

Private Sub m_Form_Seleccionado(m_expediente As Expediente)
    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Me.IDExpedientePadre = m_expediente.IDExpediente
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al m_Form_Seleccionado se ha producido el error num " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        On Error Resume Next
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub Tipo_AfterUpdate()

    Dim m_TipoElegido As String
    Dim m_JSON As String
    Dim m_Parsed As Object
    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Me.IDExpedientePadre.Locked = True
    Me.IDExpedientePadre.Enabled = False

    m_TipoElegido = Nz(Me.TIpo, "")
    If m_TipoElegido = "" Then
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If

    ' Delegate to helper for the Locked/Enabled decision.
    m_JSON = modExpedienteCambioTipoHelper.ExpedienteCambioTipo_Tipo_AfterUpdate( _
        m_TipoElegido, m_Error)
    If m_Error <> "" Then Err.Raise 1000

    Set m_Parsed = JsonConverter.ParseJson(m_JSON)("payload")
    Me.IDExpedientePadre.Locked = CBool(m_Parsed("idPadreLocked"))
    Me.IDExpedientePadre.Enabled = CBool(m_Parsed("idPadreEnabled"))

    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al Tipo_AfterUpdate se ha producido el error num " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        On Error Resume Next
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub