Option Compare Database
Option Explicit

' Form_FormExpedienteModificados — THINNED (Phase 3.3 / PR-9)
' Logic lives in modExpedienteModificadosHelper.bas.

Private WithEvents m_form As Form_FormModificado
Private m_ModificadoSeleccionado As ExpedienteModificado
Private m_Error As String

Private Function BuildDTO() As Object
    Dim dto As Object
    Set dto = CreateObject("Scripting.Dictionary")
    If m_ObjExpedienteDTOActivo Is Nothing Then
        dto("Expediente") = Nothing
        dto("ColModificados") = Nothing
    Else
        dto("Expediente") = m_ObjExpedienteDTOActivo.Expediente
        dto("ColModificados") = m_ObjExpedienteDTOActivo.ColModificados
    End If
    Set BuildDTO = dto
End Function

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

Private Sub ApplyRowSource(ByVal p_Json As String)
    Dim payload As Object
    Set payload = ParsePayload(p_Json)
    If payload Is Nothing Then Exit Sub
    If payload.Exists("rowSource") Then Me.ListaModificados.RowSource = "" & payload("rowSource")
End Sub

Private Sub ApplyFormLoadState(ByVal p_Json As String)
    Dim payload As Object
    Set payload = ParsePayload(p_Json)
    If payload Is Nothing Then Exit Sub
    If payload.Exists("allowEdits") Then Me.AllowEdits = CBool(payload("allowEdits"))
    If payload.Exists("parentRegistrarVisible") Then Me.Parent.Controls("ComandoRegistrar").Visible = CBool(payload("parentRegistrarVisible"))
End Sub

Private Sub ComandoEditar_Click()
    On Error GoTo errores
    DoCmd.Hourglass True
    m_Error = ""
    If EsAdministrador <> EnumSiNo.Sí Then
        m_Error = "Operacion no autorizada"
        Err.Raise 1000
    End If
    If m_ModificadoSeleccionado Is Nothing Then
        Set m_ModificadoSeleccionado = constructor.getExpedienteModificado(p_IDExpedienteModificado:=Nz(Me.ListaModificados.Column(0), ""), p_Error:=m_Error)
        If m_Error <> "" Then Err.Raise 1000
        If m_ModificadoSeleccionado Is Nothing Then
            m_Error = "Seleccione un elemento de la lista"
            Err.Raise 1000
        End If
    End If
    Set m_ObjModificadoActivo = m_ModificadoSeleccionado
    If FormularioAbierto("FormModificado") Then DoCmd.Close acForm, "FormModificado", acSaveNo
    DoCmd.OpenForm "FormModificado"
    If FormularioAbierto("FormModificado") Then Set m_form = Forms("FormModificado")
    DoCmd.Hourglass False
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoEditar_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    MsgBox m_Error, vbCritical, "Error"
End Sub

Private Sub ComandoRegistrar_Click()
    On Error GoTo errores
    DoCmd.Hourglass True
    MostrarPopupProgreso "Procesando...", "Guardando..."
    m_Error = ""
    Dim result As String
    result = modExpedienteModificadosHelper.ExpedienteModificados_RegistrarModificado(BuildDTO(), Nz(Me.DESCRIPCION, ""), Nz(Me.FechaFinModificado, ""), Nz(Me.FechaFirmaModificado, ""), Nz(Me.NModificado, ""), m_Error)
    If m_Error <> "" Then Err.Raise 1000
    Set m_ObjExpedienteDTOActivo.ColModificados = m_ObjExpedienteDTOActivo.Expediente.Modificados
    RellenarListaModificados m_Error
    If m_Error <> "" Then Err.Raise 1000
    If Me.ListaModificados.ListCount > 0 Then Me.ListaModificados.Selected(Me.ListaModificados.ListCount - 1) = True
    Me.DESCRIPCION = Null
    Me.FechaFirmaModificado = Null
    Me.FechaFinModificado = Null
    Me.NModificado = m_ObjExpedienteDTOActivo.Expediente.NModificadoCalculado
    CerrarPopupProgreso
    DoCmd.Hourglass False
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoRegistrar_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    CerrarPopupProgreso
    DoCmd.Hourglass False
    MsgBox m_Error, vbCritical, "Error"
End Sub

Private Sub ComandoEliminar_Click()
    On Error GoTo errores
    DoCmd.Hourglass True
    m_Error = ""
    Dim result As String
    result = modExpedienteModificadosHelper.ExpedienteModificados_EliminarModificado(BuildDTO(), Nz(Me.ListaModificados.Column(0), ""), EsAdministrador = EnumSiNo.Sí, m_Error)
    If m_Error <> "" Then Err.Raise 1000
    Set m_ObjExpedienteDTOActivo.ColModificados = m_ObjExpedienteDTOActivo.Expediente.Modificados
    RellenarListaModificados m_Error
    If m_Error <> "" Then Err.Raise 1000
    Me.NModificado = m_ObjExpedienteDTOActivo.Expediente.NModificadoCalculado
    DoCmd.Hourglass False
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoEliminar_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    MsgBox m_Error, vbCritical, "Error"
End Sub

Private Sub Form_Load()
    On Error GoTo errores
    DoCmd.Hourglass True
    m_Error = ""
    ApplyFormLoadState modExpedienteModificadosHelper.ExpedienteModificados_FormLoadState(AbiertoParaEditar, EsAdministrador = EnumSiNo.Sí, m_Error)
    If m_Error <> "" Then Err.Raise 1000
    EstablecerDatos m_Error
    If m_Error <> "" Then Err.Raise 1000
    DoCmd.Hourglass False
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Load se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    MsgBox m_Error, vbCritical, "Error"
End Sub

Private Sub m_Form_Editado()
    On Error GoTo errores
    DoCmd.Hourglass True
    m_Error = ""
    Set m_ObjExpedienteDTOActivo.Expediente.Modificados = Nothing
    Set m_ObjExpedienteDTOActivo.ColModificados = m_ObjExpedienteDTOActivo.Expediente.Modificados
    RellenarListaModificados m_Error
    If m_Error <> "" Then Err.Raise 1000
    Set m_ModificadoSeleccionado = Nothing
    DoCmd.Hourglass False
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al m_Form_Editado se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    MsgBox m_Error, vbCritical, "Error"
End Sub

Public Function RellenarListaModificados(Optional ByRef p_Error As String) As String
    On Error GoTo errores
    Set m_ObjExpedienteDTOActivo.ColModificados = m_ObjExpedienteDTOActivo.Expediente.Modificados
    Dim result As String
    result = modExpedienteModificadosHelper.ExpedienteModificados_RellenarListaModificados(BuildDTO(), p_Error)
    If p_Error <> "" Then Err.Raise 1000
    ApplyRowSource result
    If Me.ListaModificados.ListCount > 0 Then Me.ListaModificados.Selected(0) = True
    Exit Function
errores:
    p_Error = "RellenarListaModificados: " & Err.Description
End Function

Public Function RellenarListas(Optional ByRef p_Error As String) As String
    RellenarListaModificados p_Error
End Function

Public Function EstablecerDatos(Optional ByRef p_Error As String) As String
    Dim ctl As Control
    On Error GoTo errores
    If m_ObjExpedienteDTOActivo.Expediente Is Nothing Then
        p_Error = "El formulario se ha abierto con parámetros insuficientes"
        Err.Raise 1000
    End If
    For Each ctl In Me.Controls
        If Nz(ctl.Tag, "") = "EJECUTIVO" Then ctl.Enabled = Me.AllowEdits
    Next
    RellenarListas p_Error
    If p_Error <> "" Then Err.Raise 1000
    Me.NModificado = m_ObjExpedienteDTOActivo.Expediente.NModificadoCalculado
    Exit Function
errores:
    p_Error = "EstablecerDatos: " & Err.Description
End Function
