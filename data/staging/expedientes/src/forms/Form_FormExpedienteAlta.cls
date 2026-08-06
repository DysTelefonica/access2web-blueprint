Option Compare Database
Option Explicit

' Form_FormExpedienteAlta — THINNED (REWORK 2026-06-26, Phase 3.2b / PR-8b)
' Pure UI orchestration. Business logic lives in modExpedienteAltaHelper.
' See docs/audit/form-expediente-alta-pure-data.md for the audit.
'
' Removed:
'   - Inline 4-branch visibility/ordinal logic in EstablecerDatos (helper computes).
'   - Inline TIpo->{EsAM,EsLote,EsExpediente} translation in EstablecerDatosConTipo.
'   - Inline Ambito->HPSAplica setting in Ambito_AfterUpdate.
'   - Inline ordinal conflict text building in Ordinal_AfterUpdate.
'
' Preserved (UI orchestration only):
'   - 4 WithEvents handlers (m_Form*_Seleccionar) — pass-through to controls
'   - 4 ComandoElegir*_Click — DoCmd.OpenForm for selection forms
'   - 3 ComandoLimpiar*_Click — Me.IdX = Null
'   - ComandoVerSharepoint_Click — Ejecutar
'   - AccesoSharepoint_Exit — enable/disable button
'   - ComandoRegistrar_Click — DAO call (ExpedienteOperaciones.RegistrarAlta)
'   - Form_Load, Form_Unload
'   - Public method EstablecerDatos (signature preserved)
'   - Public event Alta


Private WithEvents m_FormOC As Form_FormOrganoContratacionGestion
Private WithEvents m_FormClasificacion As Form_FormGradosClasificacionGestion
Private WithEvents m_FormEjercito As Form_FormEjercitosGestion
Private WithEvents m_FormOficinaPrograma As Form_FormOficinasProgramaGestion
Private m_TipoEnum As EnumTipoExpediente
Private m_IDExpedientePadre As String
Public Event Alta(m_ObjExpediente As Expediente)
Private m_Error As String
Private m_SnapshotInicial As String


Private Sub cmdSalir_Click()
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub



Private Sub ComandoLimpiarClasificacion_Click()
    Me.IdGradoClasificacion = Null
End Sub

Private Sub ComandoLimpiarOC_Click()
    Me.IDOrganoContratacion = Null
End Sub

Private Sub ComandoLimpiarOficinaPrograma_Click()
    Me.IDOficinaPrograma = Null
End Sub

Private Sub ComandoRegistrar_Click()
    Dim m_ExpOp As ExpedienteOperaciones
    Dim m_IDExpediente As String
    Dim m_FormTipo As Form_FormExpedienteAltaTipo
    Dim m_CaptionOriginal As String
    Dim m_JSON As String
    Dim m_Parsed As Object

    On Error GoTo errores

    m_CaptionOriginal = Me.Caption
    If AntiSpamIsOperationInProgress(Me, m_Error) Then Exit Sub
    AntiSpamEnterOperation Me, m_CaptionOriginal, "ComandoRegistrar"

    DoCmd.Hourglass True
    m_Error = ""

    Set m_ExpOp = New ExpedienteOperaciones
    SetDTOFromGeneral Me, m_ObjExpedienteDTOActivo.Expediente, m_Error
    If m_Error <> "" Then Err.Raise 1000

    m_ExpOp.RegistrarAlta m_ObjExpedienteDTOActivo, p_Error:=m_Error
    If m_Error <> "" Then Err.Raise 1000

    m_IDExpediente = m_ObjExpedienteDTOActivo.Expediente.IDExpediente

    RaiseEvent Alta(m_ObjExpedienteDTOActivo.Expediente)

    ' Notify AltaTipo if open — wrapper helper for the structural decision.
    m_JSON = modExpedienteAltaHelper.ExpedienteAlta_NotificarAltaTipo( _
        FormularioAbierto("FormExpedienteAltaTipo"), m_Error)
    If m_Error <> "" Then Err.Raise 1000
    Set m_Parsed = JsonConverter.ParseJson(m_JSON)("payload")
    If CBool(m_Parsed("notified")) Then
        Set m_FormTipo = Forms("FormExpedienteAltaTipo")
        m_FormTipo.ReplicarAltaEvento m_ObjExpedienteDTOActivo.Expediente
    Else
        Debug.Print "WARNING: " & CStr(m_Parsed("reason"))
    End If

    DoCmd.Close acForm, Me.Name, acSaveNo

    AntiSpamExitOperation Me, m_CaptionOriginal, m_Error
    Exit Sub
errores:
    AntiSpamExitOperation Me, m_CaptionOriginal, m_Error
    m_Error = "Al Registrar se ha producido el error: " & Err.Description
    DoCmd.Hourglass False
    MsgBox m_Error, vbCritical, "Error"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim m_SnapshotActual As String
    Dim m_Respuesta As Integer
    On Error GoTo errores

    m_Error = ""

    ' Capturar estado actual
    m_SnapshotActual = getSnapshot(Me, m_Error)

    ' Si el formulario quedó vacío (sin cambios), cerrar sin aviso
    If m_SnapshotActual = "" Or m_SnapshotActual = m_SnapshotInicial Then
        Exit Sub
    End If

    ' Hay cambios sin guardar ? preguntar
    m_Respuesta = MsgBox("Hay cambios sin guardar. ¿Quieres salir sin guardar?", _
                        vbYesNo + vbExclamation, "Cambios sin guardar")
    If m_Respuesta = vbNo Then
        Cancel = True
        Exit Sub
    End If

    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Unload se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    MsgBox m_Error, vbCritical, "Error"
End Sub
Private Sub AccesoSharepoint_Exit(Cancel As Integer)
    If Nz(Me.AccesoSharepoint, "") <> "" Then
        Me.ComandoVerSharepoint.Enabled = True
    Else
        Me.ComandoVerSharepoint.Enabled = False
    End If
End Sub

Private Sub Ambito_AfterUpdate()
    Dim m_JSON As String
    Dim m_Parsed As Object
    Dim m_HPSAplicaActual As String

    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    m_HPSAplicaActual = Nz(Me.HPSAplica, "")

    ' Delegate the HPS-flag decision to the helper.
    m_JSON = modExpedienteAltaHelper.ExpedienteAlta_Ambito_AfterUpdate( _
        Nz(Me.Ambito, ""), m_Error)
    If m_Error <> "" Then Err.Raise 1000

    Set m_Parsed = JsonConverter.ParseJson(m_JSON)("payload")
    Dim m_NewHpsAplica As String
    m_NewHpsAplica = CStr(m_Parsed("hpsAplica"))
    If Len(m_NewHpsAplica) > 0 Then
        Me.HPSAplica = m_NewHpsAplica
    End If

    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al Ambito_AfterUpdate se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoElegitClasificacion_Click()
    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If FormularioAbierto("FormGradosClasificacionGestion") Then
        DoCmd.Close acForm, "FormGradosClasificacionGestion"
    End If
    Set m_ObjGradoClasificacionActivo = Nothing
    DoCmd.OpenForm "FormGradosClasificacionGestion", OpenArgs:="PARA SELECCIONAR"
    If FormularioAbierto("FormGradosClasificacionGestion") Then
        Set m_FormClasificacion = Forms("FormGradosClasificacionGestion")
    End If

    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoElegitClasificacion_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoElegirOficinaPrograma_Click()
    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If FormularioAbierto("FormOficinasProgramaGestion") Then
        DoCmd.Close acForm, "FormOficinasProgramaGestion"
    End If
    Set m_ObjOficinaProgramaActiva = Nothing
    DoCmd.OpenForm "FormOficinasProgramaGestion", OpenArgs:="PARA SELECCIONAR"
    If FormularioAbierto("FormOficinasProgramaGestion") Then
        Set m_FormOficinaPrograma = Forms("FormOficinasProgramaGestion")
    End If

    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoElegirOficinaPrograma_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoElegirEjercito_Click()
    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If FormularioAbierto("FormEjercitosGestion") Then
        DoCmd.Close acForm, "FormEjercitosGestion"
    End If
    Set m_ObjEjercitoActivo = Nothing
    DoCmd.OpenForm "FormEjercitosGestion", OpenArgs:="PARA SELECCIONAR"
    If FormularioAbierto("FormEjercitosGestion") Then
        Set m_FormEjercito = Forms("FormEjercitosGestion")
    End If

    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoElegirEjercito_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub


Private Sub ComandoElegirOC_Click()
    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If FormularioAbierto("FormOrganoContratacionGestion") Then
        DoCmd.Close acForm, "FormOrganoContratacionGestion"
    End If
    Set m_ObjOrganoContratacionActivo = Nothing
    DoCmd.OpenForm "FormOrganoContratacionGestion", OpenArgs:="PARA SELECCIONAR"
    If FormularioAbierto("FormOrganoContratacionGestion") Then
        Set m_FormOC = Forms("FormOrganoContratacionGestion")
    End If

    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoElegirOC_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoRellenarNemotecnico_Click()
    Dim m_Padre As Expediente
    Dim m_Ordinal As String
    Dim m_EjercitoDesc As String
    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""

    Set m_Padre = m_ObjExpedienteDTOActivo.Expediente.ExpedientePadre
    If Not m_Padre Is Nothing Then
        m_Ordinal = CalcularOrdinalSiguiente(m_Padre, m_ObjExpedienteDTOActivo.Expediente.TIPOEnum, m_Error)
        If m_Error <> "" Then Err.Raise 1000

        m_EjercitoDesc = ""
        If Not m_ObjExpedienteDTOActivo.Expediente.Ejercito Is Nothing Then
            m_EjercitoDesc = m_ObjExpedienteDTOActivo.Expediente.Ejercito.Ejercito
        End If
        Me.Nemotecnico = CalcularNemotecnicoCalculado(m_Padre, m_Ordinal, _
            m_ObjExpedienteDTOActivo.Expediente.TIPOEnum, m_EjercitoDesc, m_Error)
        If m_Error <> "" Then Err.Raise 1000
    End If

    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoRellenarNemotecnico_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoVerSharepoint_Click()
    If Nz(Me.AccesoSharepoint, "") <> "" Then
        Ejecutar 1, "open", Nz(Me.AccesoSharepoint, ""), "", "", 1
    End If

End Sub


Private Sub Form_Load()
    Dim m_Argumento As String
    Dim dato As Variant
    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    m_Argumento = Nz(Me.OpenArgs(), "")
    If InStr(m_Argumento, "|") = 0 Then
        m_Error = "El formulario de alta se ha abierto sin conocer el tipo"
        Err.Raise 1000
    End If
    dato = Split(m_Argumento, "|")
    m_IDExpedientePadre = CStr(dato(0))
    If Not IsNumeric(dato(1)) Then
        m_Error = "El formulario de alta se ha abierto sin conocer el tipo"
        Err.Raise 1000
    End If
    m_TipoEnum = CLng(dato(1))
    EstablecerCombos Me, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    ' Recuperar el DTO que dejó FormExpedienteAltaTipo
    Set m_ObjExpedienteDTOActivo = New ExpedienteDTO

    With m_ObjExpedienteDTOActivo
        Set .Expediente = New Expediente
        .Expediente.IDExpedientePadre = m_IDExpedientePadre

        SetDatosDeExpedienteTipo .Expediente, m_TipoEnum, m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        If IsNumeric(m_IDExpedientePadre) Then
            HeredarDatosDePadre Me, m_IDExpedientePadre, m_TipoEnum, m_Error
            If m_Error <> "" Then
                Err.Raise 1000
            End If
        End If
        If .Expediente.TIpo <> "" Then
            Me.lblTitulo.Caption = "ALTA DE " & .Expediente.TIpo
        Else
            Me.lblTitulo.Caption = "ALTA DE EXPEDIENTE"
        End If
        Me.TIpo = .Expediente.TIpo
    End With

    Me.AllowEdits = True

    ' Capturar estado inicial para detectar cambios al cerrar
    m_SnapshotInicial = getSnapshot(Me, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If

    ' EstablecerDatos PRIMERO para fijar visibilidad de ordinal
    ' antes de Ajustar que puede pisar valores del diseñador
    EstablecerDatos m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If

    ' Ajustar AL FINAL para no sobrescribir visibilidad ya fijada
    Ajustar Me

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
End Sub


Public Function EstablecerDatos(Optional ByRef p_Error As String) As String

    Dim m_DTO As Object
    Dim m_JSON As String
    Dim m_Parsed As Object
    Dim m_OrdinalValue As String

    On Error GoTo errores

    Debug.Print "=== EstablecerDatos === m_TipoEnum:", m_TipoEnum, " m_IDExpedientePadre:", m_IDExpedientePadre

    If m_ObjExpedienteDTOActivo Is Nothing Then
        p_Error = "El formulario se ha abierto con parámetros insuficientes"
        Err.Raise 1000
    End If
    If m_ObjExpedienteDTOActivo.Expediente Is Nothing Then
        p_Error = "El formulario se ha abierto con parámetros insuficientes"
        Err.Raise 1000
    End If

    ' Wrap the live DTO into a Dictionary so the helper sees a stable shape.
    Set m_DTO = CreateObject("Scripting.Dictionary")
    Set m_DTO("Expediente") = m_ObjExpedienteDTOActivo.Expediente

    ' Delegate visibility/ordinal decision to the helper.
    m_JSON = modExpedienteAltaHelper.ExpedienteAlta_EstablecerDatos( _
        m_DTO, CLng(m_TipoEnum), CStr(m_IDExpedientePadre), p_Error)
    If p_Error <> "" Then Exit Function

    Set m_Parsed = JsonConverter.ParseJson(m_JSON)("payload")

    ' Apply the helper's decision to the form's controls.
    Me.lblNumero.Visible = CBool(m_Parsed("lblNumeroVisible"))
    Me.Ordinal.Visible = CBool(m_Parsed("ordinalVisible"))
    Me.ComandoRellenarNemotecnico.Visible = CBool(m_Parsed("comandoRellenarVisible"))

    With m_ObjExpedienteDTOActivo.Expediente
        Me.TIpo = .TIpo
        Me.Nemotecnico = .Nemotecnico

        m_OrdinalValue = CStr(m_Parsed("ordinalValue"))
        If CBool(m_Parsed("ordinalVisible")) And Len(m_OrdinalValue) > 0 Then
            Me.Ordinal = m_OrdinalValue
        End If

        If Len(CStr(m_Parsed("ordinalCaption"))) > 0 Then
            Me.lblNumero.Caption = CStr(m_Parsed("ordinalCaption"))
        End If
    End With

    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerDatos ha producido el error n: " & Err.Number & _
        vbNewLine & "Detalle: " & Err.Description
    End If

End Function

Private Sub m_FormClasificacion_Seleccionar(m_GradoClasificacionSeleccionado As GradoClasificacion)
     On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If Not m_GradoClasificacionSeleccionado Is Nothing Then
        Me.IdGradoClasificacion = m_GradoClasificacionSeleccionado.GradoClasificacion
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al m_FormClasificacion_Seleccionar se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub m_FormEjercito_Seleccionar(m_Ejercito As Ejercito)
    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""


    If Not m_Ejercito Is Nothing Then
        Me.IDEjercito = m_Ejercito.Ejercito

    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al m_FormEjercito_Seleccionar se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub m_FormOC_Seleccionar(m_OrganoContratacionSeleccionado As OrganoContratacion)
    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""


    If Not m_OrganoContratacionSeleccionado Is Nothing Then
        Me.IDOrganoContratacion = m_OrganoContratacionSeleccionado.OrganoContratacion

    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al m_FormOC_Seleccionar se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub



Private Sub m_FormOficinaPrograma_Seleccionar(m_OficinaProgramaSeleccionada As OficinaPrograma)
    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""

    If Not m_OficinaProgramaSeleccionada Is Nothing Then
        Me.IDOficinaPrograma = m_OficinaProgramaSeleccionada.OficinaPrograma

    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al m_FormOficinaPrograma_Seleccionar se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub Ordinal_AfterUpdate()

    Dim m_ExpedienteOrdinalUsado As EnumSiNo
    Dim m_Comprobar As Boolean
    Dim m_ExpedienteOtro As Object
    Dim m_textoExp As String
    Dim m_JSON As String
    Dim m_Parsed As Object
    Dim m_StubConf As Object
    On Error GoTo errores


    m_Error = ""
    If Not IsNumeric(Nz(Me.Ordinal, "")) Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    If m_ObjExpedienteDTOActivo.Expediente.IDExpedientePadre = "" Then
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If


    ' DAO lookup for ordinal conflict — the helper takes the result as a stub.
    Set m_ExpedienteOtro = ExpedienteOrdinalUsado( _
                p_IDExpedientePadre:=m_ObjExpedienteDTOActivo.Expediente.IDExpedientePadre, _
                p_Ordinal:=Me.Ordinal, p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If

    If m_ExpedienteOtro Is Nothing Then
        ' No conflict.
        Set m_StubConf = Nothing
    Else
        ' Convert the result to a Dictionary stub for the helper.
        Set m_StubConf = CreateObject("Scripting.Dictionary")
        On Error Resume Next
        m_StubConf("CodExp") = m_ExpedienteOtro.CodExp
        m_StubConf("Nemotecnico") = m_ExpedienteOtro.Nemotecnico
        m_StubConf("Titulo") = m_ExpedienteOtro.Titulo
        On Error GoTo errores
    End If

    ' Delegate conflict text formatting to the helper.
    m_JSON = modExpedienteAltaHelper.ExpedienteAlta_Ordinal_AfterUpdate( _
        CStr(Nz(Me.Ordinal, "")), _
        CStr(m_ObjExpedienteDTOActivo.Expediente.IDExpedientePadre), _
        m_StubConf, m_Error)
    If m_Error <> "" Then Err.Raise 1000

    Set m_Parsed = JsonConverter.ParseJson(m_JSON)("payload")
    If CBool(m_Parsed("showWarning")) Then
        pregunta = MsgBox(CStr(m_Parsed("conflictoTexto")), vbExclamation, "Ordinal repetido")
    End If


    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al Ordinal_AfterUpdate se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub