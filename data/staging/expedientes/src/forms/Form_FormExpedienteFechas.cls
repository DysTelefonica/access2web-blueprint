Option Compare Database
Option Explicit

' Form_FormExpedienteFechas — THINNED (Phase 3.3 / PR-9)
' Pure date/load/unload decisions live in modExpedienteFechasHelper.bas.

Private m_Error As String
Public m_SnapshotAlInicio As String
Dim m_IDExpediente As String

Private Function ParsePayload(ByVal p_Json As String) As Object
    On Error GoTo EH
    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(p_Json)
    If parsed Is Nothing Then Set ParsePayload = Nothing: Exit Function
    If parsed.Exists("payload") Then Set ParsePayload = parsed("payload") Else Set ParsePayload = Nothing
    Exit Function
EH:
    Set ParsePayload = Nothing
End Function

Private Sub AplicarEstadoCarga(ByVal p_Json As String)
    Dim payload As Object
    Set payload = ParsePayload(p_Json)
    If payload Is Nothing Then Exit Sub
    If payload.Exists("allowEdits") Then Me.AllowEdits = CBool(payload("allowEdits"))
    If payload.Exists("parentRegistrarVisible") Then FormInteraction_EstablecerPropiedad Me.Parent, "ComandoRegistrar", "Visible", CBool(payload("parentRegistrarVisible"))
    If payload.Exists("parentRegistrarEnabled") Then FormInteraction_EstablecerPropiedad Me.Parent, "ComandoRegistrar", "Enabled", CBool(payload("parentRegistrarEnabled"))
End Sub

Private Sub ComandoFECHAADJUDICACION_Click()
    Me.FECHAADJUDICACION = Date
End Sub

Private Sub ComandoFECHACERTIFICACION_Click()
    Dim fechaFin As Variant
    Dim errMsg As String

    Me.FECHACERTIFICACION = Date
    modExpedienteFechasHelper.ExpedienteFechas_CalcularFechaFinGarantia Me.GARANTIAMESES, Me.FECHACERTIFICACION, Me.FechaFinContrato, fechaFin, errMsg
    Me.FechaFinGarantia = fechaFin
End Sub

Private Sub ComandoFECHADESESTIMADA_Click()
    Me.FECHADESESTIMADA = Date
End Sub

Private Sub ComandoFECHAFIRMACONTRATO_Click()
    Me.FECHAFIRMACONTRATO = Date
End Sub

Private Sub ComandoFECHAINICIOLICITACION_Click()
    Me.FECHAINICIOLICITACION = Date
End Sub

Private Sub ComandoFECHAPREOFERTA_Click()
    Me.FECHAPREOFERTA = Date
End Sub

Private Sub ComandoFECHAOFERTA_Click()
    Me.FECHAOFERTA = Date
End Sub

Private Sub ComandoFECHAPERDIDA_Click()
    Me.FECHAPERDIDA = Date
End Sub

Private Sub FECHACERTIFICACION_BeforeUpdate(Cancel As Integer)
    Dim p_LocalResultado As Long
    Dim fechaFin As Variant
    Dim errMsg As String

    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    modExpedienteFechasHelper.ExpedienteFechas_CalcularFechaFinGarantia Me.GARANTIAMESES, Me.FECHACERTIFICACION, Me.FechaFinContrato, fechaFin, errMsg
    Me.FechaFinGarantia = fechaFin
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al Meses_Exit se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    FormInteraction_Mensaje(p_Texto:=m_Error, p_Estilo:=vbCritical, p_Resultado:=p_LocalResultado)
End Sub

Private Sub FechaFinContrato_BeforeUpdate(Cancel As Integer)
    Dim p_LocalResultado As Long
    Dim fechaFin As Variant
    Dim errMsg As String

    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    modExpedienteFechasHelper.ExpedienteFechas_CalcularFechaFinGarantia Me.GARANTIAMESES, Me.FECHACERTIFICACION, Me.FechaFinContrato, fechaFin, errMsg
    Me.FechaFinGarantia = fechaFin
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al Meses_Exit se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    FormInteraction_Mensaje(p_Texto:=m_Error, p_Estilo:=vbCritical, p_Resultado:=p_LocalResultado)
End Sub

Private Sub Form_Load()
    Dim p_LocalResultado As Long

    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Me.AllowEdits = False
    AplicarEstadoCarga modExpedienteFechasHelper.ExpedienteFechas_FormLoadState(AbiertoParaEditar, EsAdministrador = EnumSiNo.Sí, m_Error)
    If m_Error <> "" Then Err.Raise 1000
    Me.ComandoFECHAINICIOLICITACION.Enabled = Me.AllowEdits
    Me.ComandoFECHAPREOFERTA.Enabled = Me.AllowEdits
    Me.ComandoFECHAOFERTA.Enabled = Me.AllowEdits
    Me.ComandoFECHADESESTIMADA.Enabled = Me.AllowEdits
    Me.ComandoFECHAPERDIDA.Enabled = False
    Me.ComandoFECHAADJUDICACION.Enabled = Me.AllowEdits
    Me.ComandoFECHAFIRMACONTRATO.Enabled = Me.AllowEdits
    Me.ComandoFECHACERTIFICACION.Enabled = Me.AllowEdits
    m_IDExpediente = m_ObjExpedienteDTOActivo.Expediente.IDExpediente
    SetFormFechasFromExpediente p_IDExpediente:=m_IDExpediente, frm:=Me, p_Error:=m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If

    If Me.AllowEdits = True Then
        ' Create snapshot for change detection (only in EDIT mode)
        m_SnapshotAlInicio = getSnapshot(frm:=Me, p_Error:=m_Error)
        If m_Error <> "" Then
            Err.Raise 100
        End If
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
    FormInteraction_Mensaje(p_Texto:=m_Error, p_Estilo:=vbCritical, p_Resultado:=p_LocalResultado)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim p_LocalResultado As Long
    Dim m_Snapshot As String

    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If Me.AllowEdits = False Then
        Exit Sub
    End If
    m_Snapshot = getSnapshot(frm:=Me, p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Dim decisionJson As String
    decisionJson = modExpedienteFechasHelper.ExpedienteFechas_UnloadDecision(Me.AllowEdits, m_SnapshotAlInicio, m_Snapshot, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Dim payload As Object
    Set payload = ParsePayload(decisionJson)
    If payload Is Nothing Then
        Exit Sub
    End If
    If Not CBool(payload("guardar")) Then
        Exit Sub
    End If
    GuardarPestana frm:=Me, p_IDExpediente:=m_IDExpediente, p_Error:=m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If

    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Unload se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    FormInteraction_Mensaje(p_Texto:=m_Error, p_Estilo:=vbCritical, p_Resultado:=p_LocalResultado)
End Sub

Private Sub GARANTIAMESES_BeforeUpdate(Cancel As Integer)
    Dim p_LocalResultado As Long
    Dim fechaFin As Variant
    Dim errMsg As String

    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    modExpedienteFechasHelper.ExpedienteFechas_CalcularFechaFinGarantia Me.GARANTIAMESES, Me.FECHACERTIFICACION, Me.FechaFinContrato, fechaFin, errMsg
    Me.FechaFinGarantia = fechaFin
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al Meses_Exit se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    FormInteraction_Mensaje(p_Texto:=m_Error, p_Estilo:=vbCritical, p_Resultado:=p_LocalResultado)
End Sub

