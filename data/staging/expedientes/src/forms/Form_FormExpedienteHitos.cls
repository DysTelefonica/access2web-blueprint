Option Compare Database
Option Explicit

' Form_FormExpedienteHitos — THINNED (2026-06-26)
' Logic lives in modExpedienteHitosHelper.bas.

Private m_Error As String

Private Function BuildDTO() As Object
    Dim dto As Object
    Set dto = CreateObject("Scripting.Dictionary")
    If m_ObjExpedienteDTOActivo Is Nothing Then
        dto("Expediente") = Nothing
        dto("ColHitos") = Nothing
    Else
        dto("Expediente") = m_ObjExpedienteDTOActivo.Expediente
        dto("ColHitos") = m_ObjExpedienteDTOActivo.ColHitos
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

Private Sub Form_Load()
    On Error GoTo errores
    DoCmd.Hourglass True
    Me.AllowEdits = False
    Me.Parent.Controls("ComandoRegistrar").Visible = False
    If AbiertoParaEditar = True Then
        If EsAdministrador = EnumSiNo.Sí Then Me.AllowEdits = True
    End If
    EstablecerDatos m_Error
    If m_Error <> "" Then Err.Raise 1000
    DoCmd.Hourglass False
    Exit Sub
errores:
    DoCmd.Hourglass False
    MsgBox "Al Form_Load: " & Err.Description, vbCritical, "Error"
End Sub

Public Function EstablecerDatos(Optional ByRef p_Error As String) As String
    Dim ctl As Control
    On Error GoTo errores
    For Each ctl In Me.Controls
        If Nz(ctl.Tag, "") = "EJECUTIVO" Then
            If AbiertoParaEditar = True Then
                If EsAdministrador = EnumSiNo.Sí Then ctl.Enabled = True Else ctl.Enabled = False
            Else
                ctl.Enabled = False
            End If
        End If
    Next
    Dim result As String
    result = modExpedienteHitosHelper.ExpedienteHitos_RellenarListas(BuildDTO(), p_Error)
    If p_Error <> "" Then Err.Raise 1000
    Dim payload As Object
    Set payload = ParsePayload(result)
    If Not payload Is Nothing Then
        If payload.Exists("rowSource") Then Me.ListaHitos.RowSource = "" & payload("rowSource")
    End If
    Exit Function
errores:
    p_Error = "EstablecerDatos: " & Err.Description
End Function

Public Function RellenarListas(Optional ByRef p_Error As String) As String
    Dim result As String
    On Error GoTo errores
    result = modExpedienteHitosHelper.ExpedienteHitos_RellenarListas(BuildDTO(), p_Error)
    If p_Error <> "" Then Err.Raise 1000
    Dim payload As Object
    Set payload = ParsePayload(result)
    If Not payload Is Nothing Then
        If payload.Exists("rowSource") Then Me.ListaHitos.RowSource = "" & payload("rowSource")
    End If
    Exit Function
errores:
    p_Error = "RellenarListas: " & Err.Description
End Function

Public Function RellenarListaHitos(Optional ByRef p_Error As String) As String
    Dim result As String
    On Error GoTo errores
    Set m_ObjExpedienteDTOActivo.ColHitos = m_ObjExpedienteDTOActivo.Expediente.Hitos
    result = modExpedienteHitosHelper.ExpedienteHitos_RellenarListaHitos(BuildDTO(), p_Error)
    If p_Error <> "" Then Err.Raise 1000
    Dim payload As Object
    Set payload = ParsePayload(result)
    If Not payload Is Nothing Then
        If payload.Exists("rowSource") Then Me.ListaHitos.RowSource = "" & payload("rowSource")
    End If
    Exit Function
errores:
    p_Error = "RellenarListaHitos: " & Err.Description
End Function

Private Sub ComandoRegistrar_Click()
    On Error GoTo errores
    DoCmd.Hourglass True
    MostrarPopupProgreso "Procesando...", "Guardando..."
    Dim result As String
    result = modExpedienteHitosHelper.ExpedienteHitos_AltaHito( _
        BuildDTO(), Nz(Me.FechaHito, ""), Nz(Me.DESCRIPCION, ""), _
        Nz(Me.FechaGarantiaHito, ""), Nz(Me.Importe, ""), m_Error)
    If m_Error <> "" Then Err.Raise 1000
    RellenarListaHitos m_Error
    If m_Error <> "" Then Err.Raise 1000
    Me.ListaHitos.Selected(Me.ListaHitos.ListCount - 1) = True
    Me.DESCRIPCION = Null
    Me.FechaHito = Null
    Me.FechaGarantiaHito = Null
    Me.Importe = Null
    CerrarPopupProgreso
    DoCmd.Hourglass False
    Exit Sub
errores:
    CerrarPopupProgreso
    DoCmd.Hourglass False
    MsgBox "Alta Hito: " & m_Error, vbCritical, "Error"
End Sub

Private Sub ComandoEliminar_Click()
    On Error GoTo errores
    DoCmd.Hourglass True
    Dim fecha As String
    fecha = Nz(Me.ListaHitos.Column(0), "")
    If Len(fecha) = 0 Then
        DoCmd.Hourglass False
        Exit Sub
    End If
    Dim result As String
    result = modExpedienteHitosHelper.ExpedienteHitos_EliminarHito( _
        BuildDTO(), fecha, EsAdministrador = EnumSiNo.Sí, m_Error)
    If m_Error <> "" Then Err.Raise 1000
    RellenarListaHitos m_Error
    If m_Error <> "" Then Err.Raise 1000
    DoCmd.Hourglass False
    Exit Sub
errores:
    DoCmd.Hourglass False
    MsgBox "Eliminar Hito: " & m_Error, vbCritical, "Error"
End Sub