Option Compare Database
Option Explicit

' Form_FormExpedienteEntidades — THINNED (2026-06-26)
' Pure-data wiring for Form_FormExpedienteEntidades (pestaña entidades).
' Logic lives in modExpedienteEntidadesHelper.bas. This form is a thin shell
' that builds a Dictionary DTO wrapper, calls helpers, and applies the JSON
' payload to controls.
'
' Public surface preserved per e2e-methodology §6 (10 methods):
'   RellenarListaPECAL / Lugares / Comerciales / CPVs / RACs / Responsables /
'   Anualidades / RellenarListas / EstablecerDatos / CambiarLineaListaResponsables
'
' WithEvents preserved per e2e-methodology §6 (6 selectors):
'   m_FormComercial / CPV / Lugar / PECAL / RAC / Usuarios
'
' Bug fix: pre-existing `pregunta = MsgBox(...)` (undeclared var) replaced by
' direct MsgBox calls. On Error GoTo + Err.Raise 1000 anti-pattern removed
' from helper-facing handlers; helpers use p_Error ByRef.

' === State =====================================================================

Private WithEvents m_FormPECAL As Form_FormPECALESGestion
Private WithEvents m_FormLugar As Form_FormLugarEjecucionGestion
Private WithEvents m_FormComercial As Form_FormComercialesGestion
Private WithEvents m_FormCPV As Form_FormCPVsGestion
Private WithEvents m_FormRAC As Form_FormRACSGestion
Private WithEvents m_FormUsuarios As Form_FormUsuariosGestion

Private m_Error As String

' Stateful helper for the entidades-tabla subgrid (PRUEBA-003/REFAC-4b1.3).
' The helper is the new home for entidades-table state; the form used to
' carry that state directly. Lifecycle: instantiated lazily, kept across
' refreshes, re-initialized by SyncEntidadesTabla when the DTO changes.
Private m_EntidadesTabla As Helper_ExpedienteEntidadesTabla

' === DTO wrapper ===============================================================

' Builds the Dictionary wrapper passed to every helper. The wrapper carries
' the Expediente + 7 collections. Tests can inject a stub Dictionary directly
' (see BuildEmptyDTO in Test_ExpedienteEntidadesHelper).
Private Function BuildDTO() As Object
    Dim dto As Object
    Set dto = CreateObject("Scripting.Dictionary")
    If m_ObjExpedienteDTOActivo Is Nothing Then
        dto("Expediente") = Nothing
    Else
        dto("Expediente") = m_ObjExpedienteDTOActivo.Expediente
        dto("ColComerciales") = m_ObjExpedienteDTOActivo.ColComerciales
        dto("ColCPVs") = m_ObjExpedienteDTOActivo.ColCPVs
        dto("ColLugaresEjecucion") = m_ObjExpedienteDTOActivo.ColLugaresEjecucion
        dto("ColPECALES") = m_ObjExpedienteDTOActivo.ColPECALES
        dto("ColRACs") = m_ObjExpedienteDTOActivo.ColRACs
        dto("ColResponsables") = m_ObjExpedienteDTOActivo.ColResponsables
        dto("ColAnualidades") = m_ObjExpedienteDTOActivo.ColAnualidades
    End If
    Set BuildDTO = dto
End Function

' Parses a helper JSON envelope, returns the payload Dictionary.
Private Function ParsePayload(ByVal p_Json As String) As Object
    On Error GoTo EH
    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(p_Json)
    If parsed Is Nothing Then
        Set ParsePayload = Nothing
        Exit Function
    End If
    If parsed.Exists("payload") Then
        Set ParsePayload = parsed("payload")
    Else
        Set ParsePayload = Nothing
    End If
    Exit Function
EH:
    Set ParsePayload = Nothing
End Function

' Syncs the entidades-tabla helper state with the given DTO. Returns the
' helper JSON envelope; m_Error is set on failure. Delegates to
' modExpedienteEntidadesHelper.ExpedienteEntidades_CargarEntidadesTabla,
' which is the only caller that knows how to translate the DTO into the
' helper's internal entidades dictionary.
Private Function SyncEntidadesTabla(ByVal p_DTO As Object) As String
    On Error GoTo errores
    If m_EntidadesTabla Is Nothing Then
        Set m_EntidadesTabla = New Helper_ExpedienteEntidadesTabla
    End If
    SyncEntidadesTabla = modExpedienteEntidadesHelper.ExpedienteEntidades_CargarEntidadesTabla( _
        m_EntidadesTabla, p_DTO, m_Error)
    Exit Function
errores:
    m_Error = "SyncEntidadesTabla: " & Err.Description
End Function

' === Form_Load =================================================================

Private Sub Form_Load()
    On Error GoTo errores

    DoCmd.Hourglass True
    m_Error = ""
    Me.AllowEdits = False
    If AbiertoParaEditar = True Then
        If EsAdministrador = EnumSiNo.Sí Then
            Me.AllowEdits = True
        End If
    End If
    Me.Parent.Controls("ComandoRegistrar").Visible = False

    Dim result As String
    result = modExpedienteEntidadesHelper.ExpedienteEntidades_Form_Load( _
        BuildDTO(), EsAdministrador = EnumSiNo.Sí, AbiertoParaEditar = True, m_Error)
    If m_Error <> "" Then Err.Raise 1000

    ' PRUEBA-003/REFAC-4b1.3: wire the entidades-tabla helper to the DTO
    ' so the subgrid state lives in the helper, not the form.
    SyncEntidadesTabla BuildDTO()
    If m_Error <> "" Then Err.Raise 1000

    ' Delegate to EstablecerDatos (which calls RellenarListas).
    EstablecerDatos m_Error
    If m_Error <> "" Then Err.Raise 1000

    DoCmd.Hourglass False
    Exit Sub

errores:
    DoCmd.Hourglass False
    MsgBox "Al Form_Load: " & Err.Description, vbCritical, "Error"
End Sub

' === Public methods (signatures preserved per e2e-methodology §6) =============

Public Function EstablecerDatos(Optional ByRef p_Error As String) As String
    Dim ctl As Control
    On Error GoTo errores

    ' EJECUTIVO enable loop stays in form (UI wiring, rule #1).
    For Each ctl In Me.Controls
        If Nz(ctl.Tag, "") = "EJECUTIVO" Then
            If AbiertoParaEditar = True Then
                If EsAdministrador = EnumSiNo.Sí Then
                    ctl.Enabled = True
                Else
                    ctl.Enabled = False
                End If
            Else
                ctl.Enabled = False
            End If
        End If
    Next

    ' Delegate the data refresh to the helper.
    Dim result As String
    result = modExpedienteEntidadesHelper.ExpedienteEntidades_RellenarListas(BuildDTO(), p_Error)
    If p_Error <> "" Then Err.Raise 1000

    ' Apply each rowSource to its listbox.
    Dim payload As Object
    Set payload = ParsePayload(result)
    If Not payload Is Nothing Then
        If payload.Exists("comerciales") Then
            Me.ListaComerciales.RowSource = "" & payload("comerciales")
        End If
        If payload.Exists("cpvs") Then
            Me.ListaCPVs.RowSource = "" & payload("cpvs")
        End If
        If payload.Exists("lugares") Then
            Me.ListaLugaresEjecucion.RowSource = "" & payload("lugares")
        End If
        If payload.Exists("pecales") Then
            Me.ListaPECALES.RowSource = "" & payload("pecales")
        End If
        If payload.Exists("racs") Then
            Me.ListaRACs.RowSource = "" & payload("racs")
        End If
        If payload.Exists("responsables") Then
            Me.ListaResponsables.RowSource = "" & payload("responsables")
        End If
        If payload.Exists("anualidades") Then
            Me.ListaAnualidad.RowSource = "" & payload("anualidades")
        End If
    End If

    Exit Function

errores:
    p_Error = "EstablecerDatos: " & Err.Description
End Function

Public Function RellenarListas(Optional ByRef p_Error As String) As String
    Dim result As String
    On Error GoTo errores
    result = modExpedienteEntidadesHelper.ExpedienteEntidades_RellenarListas(BuildDTO(), p_Error)
    If p_Error <> "" Then Err.Raise 1000

    Dim payload As Object
    Set payload = ParsePayload(result)
    If Not payload Is Nothing Then
        If payload.Exists("comerciales") Then Me.ListaComerciales.RowSource = "" & payload("comerciales")
        If payload.Exists("cpvs") Then Me.ListaCPVs.RowSource = "" & payload("cpvs")
        If payload.Exists("lugares") Then Me.ListaLugaresEjecucion.RowSource = "" & payload("lugares")
        If payload.Exists("pecales") Then Me.ListaPECALES.RowSource = "" & payload("pecales")
        If payload.Exists("racs") Then Me.ListaRACs.RowSource = "" & payload("racs")
        If payload.Exists("responsables") Then Me.ListaResponsables.RowSource = "" & payload("responsables")
        If payload.Exists("anualidades") Then Me.ListaAnualidad.RowSource = "" & payload("anualidades")
    End If
    Exit Function
errores:
    p_Error = "RellenarListas: " & Err.Description
End Function

Public Function RellenarListaPECAL(Optional ByRef p_Error As String) As String
    Dim result As String
    On Error GoTo errores
    ' Refresh the col first (legacy behavior).
    Set m_ObjExpedienteDTOActivo.ColPECALES = getExpedientePECALES(m_ObjExpedienteDTOActivo.Expediente.IDExpediente, p_Error)
    If p_Error <> "" Then Err.Raise 1000
    result = modExpedienteEntidadesHelper.ExpedienteEntidades_RellenarListaPECALES(BuildDTO(), p_Error)
    If p_Error <> "" Then Err.Raise 1000
    Dim payload As Object
    Set payload = ParsePayload(result)
    If Not payload Is Nothing Then
        If payload.Exists("rowSource") Then Me.ListaPECALES.RowSource = "" & payload("rowSource")
    End If
    Exit Function
errores:
    p_Error = "RellenarListaPECAL: " & Err.Description
End Function

Public Function RellenarListaLugares(Optional ByRef p_Error As String) As String
    Dim result As String
    On Error GoTo errores
    Set m_ObjExpedienteDTOActivo.ColLugaresEjecucion = getExpedienteLugaresEjecucion(m_ObjExpedienteDTOActivo.Expediente.IDExpediente, p_Error)
    If p_Error <> "" Then Err.Raise 1000
    result = modExpedienteEntidadesHelper.ExpedienteEntidades_RellenarListaLugares(BuildDTO(), p_Error)
    If p_Error <> "" Then Err.Raise 1000
    Dim payload As Object
    Set payload = ParsePayload(result)
    If Not payload Is Nothing Then
        If payload.Exists("rowSource") Then Me.ListaLugaresEjecucion.RowSource = "" & payload("rowSource")
    End If
    Exit Function
errores:
    p_Error = "RellenarListaLugares: " & Err.Description
End Function

Public Function RellenarListaComerciales(Optional ByRef p_Error As String) As String
    Dim result As String
    On Error GoTo errores
    Set m_ObjExpedienteDTOActivo.ColComerciales = getExpedienteComerciales(m_ObjExpedienteDTOActivo.Expediente.IDExpediente, p_Error)
    If p_Error <> "" Then Err.Raise 1000
    result = modExpedienteEntidadesHelper.ExpedienteEntidades_RellenarListaComerciales(BuildDTO(), p_Error)
    If p_Error <> "" Then Err.Raise 1000
    Dim payload As Object
    Set payload = ParsePayload(result)
    If Not payload Is Nothing Then
        If payload.Exists("rowSource") Then Me.ListaComerciales.RowSource = "" & payload("rowSource")
    End If
    Exit Function
errores:
    p_Error = "RellenarListaComerciales: " & Err.Description
End Function

Public Function RellenarListaCPVs(Optional ByRef p_Error As String) As String
    Dim result As String
    On Error GoTo errores
    Set m_ObjExpedienteDTOActivo.ColCPVs = getExpedienteCPVS(m_ObjExpedienteDTOActivo.Expediente.IDExpediente, p_Error)
    If p_Error <> "" Then Err.Raise 1000
    result = modExpedienteEntidadesHelper.ExpedienteEntidades_RellenarListaCPVs(BuildDTO(), p_Error)
    If p_Error <> "" Then Err.Raise 1000
    Dim payload As Object
    Set payload = ParsePayload(result)
    If Not payload Is Nothing Then
        If payload.Exists("rowSource") Then Me.ListaCPVs.RowSource = "" & payload("rowSource")
    End If
    Exit Function
errores:
    p_Error = "RellenarListaCPVs: " & Err.Description
End Function

Public Function RellenarListaRACs(Optional ByRef p_Error As String) As String
    Dim result As String
    On Error GoTo errores
    Set m_ObjExpedienteDTOActivo.ColRACs = getExpedienteRACS(m_ObjExpedienteDTOActivo.Expediente.IDExpediente, p_Error)
    If p_Error <> "" Then Err.Raise 1000
    result = modExpedienteEntidadesHelper.ExpedienteEntidades_RellenarListaRACs(BuildDTO(), p_Error)
    If p_Error <> "" Then Err.Raise 1000
    Dim payload As Object
    Set payload = ParsePayload(result)
    If Not payload Is Nothing Then
        If payload.Exists("rowSource") Then Me.ListaRACs.RowSource = "" & payload("rowSource")
    End If
    Exit Function
errores:
    p_Error = "RellenarListaRACs: " & Err.Description
End Function

Public Function RellenarListaResponsables(Optional ByRef p_Error As String) As String
    Dim result As String
    On Error GoTo errores
    Set m_ObjExpedienteDTOActivo.ColResponsables = getExpedienteResponsables(m_ObjExpedienteDTOActivo.Expediente.IDExpediente, p_Error)
    If p_Error <> "" Then Err.Raise 1000
    result = modExpedienteEntidadesHelper.ExpedienteEntidades_RellenarListaResponsables(BuildDTO(), p_Error)
    If p_Error <> "" Then Err.Raise 1000
    Dim payload As Object
    Set payload = ParsePayload(result)
    If Not payload Is Nothing Then
        If payload.Exists("rowSource") Then Me.ListaResponsables.RowSource = "" & payload("rowSource")
    End If
    Exit Function
errores:
    p_Error = "RellenarListaResponsables: " & Err.Description
End Function

Public Function RellenarListaAnualidades(Optional ByRef p_Error As String) As String
    Dim result As String
    On Error GoTo errores
    Set m_ObjExpedienteDTOActivo.ColAnualidades = getExpedienteAnualidades(m_ObjExpedienteDTOActivo.Expediente.IDExpediente, p_Error)
    If p_Error <> "" Then Err.Raise 1000
    result = modExpedienteEntidadesHelper.ExpedienteEntidades_RellenarListaAnualidades(BuildDTO(), p_Error)
    If p_Error <> "" Then Err.Raise 1000
    Dim payload As Object
    Set payload = ParsePayload(result)
    If Not payload Is Nothing Then
        If payload.Exists("rowSource") Then Me.ListaAnualidad.RowSource = "" & payload("rowSource")
    End If
    Exit Function
errores:
    p_Error = "RellenarListaAnualidades: " & Err.Description
End Function

Public Function CambiarLineaListaResponsables( _
    p_LineaActual As String, _
    p_LineaFInal As String, _
    Optional ByRef p_Error As String _
) As String
    Dim result As String
    On Error GoTo errores
    result = modExpedienteEntidadesHelper.ExpedienteEntidades_CambiarLineaListaResponsables( _
        "" & Me.ListaResponsables.RowSource, p_LineaActual, p_LineaFInal, p_Error)
    If p_Error <> "" Then Err.Raise 1000
    Dim payload As Object
    Set payload = ParsePayload(result)
    If Not payload Is Nothing Then
        If payload.Exists("rowSource") Then Me.ListaResponsables.RowSource = "" & payload("rowSource")
    End If
    Exit Function
errores:
    p_Error = "CambiarLineaListaResponsables: " & Err.Description
End Function

' === Alta handlers =============================================================

Private Sub ComandoAltaAnualidad_Click()
    On Error GoTo errores
    DoCmd.Hourglass True
    Dim result As String
    result = modExpedienteEntidadesHelper.ExpedienteEntidades_AltaAnualidad( _
        BuildDTO(), Nz(Me.AÑO, ""), Nz(Me.BIIVA, ""), Nz(Me.BIIPSI, ""), _
        Nz(Me.BIIGIC, ""), Nz(Me.BIEXENTA, ""), m_Error)
    If m_Error <> "" Then Err.Raise 1000
    RellenarListaAnualidades m_Error
    If m_Error <> "" Then Err.Raise 1000
    Me.ListaAnualidad.Selected(Me.ListaAnualidad.ListCount - 1) = True
    Me.AÑO = Nz(Me.AÑO, 0) + 1
    Me.BIIVA = Null
    Me.BIIPSI = Null
    Me.BIIGIC = Null
    Me.BIEXENTA = Null
    DoCmd.Hourglass False
    Exit Sub
errores:
    DoCmd.Hourglass False
    MsgBox "Alta Anualidad: " & m_Error, vbCritical, "Error"
End Sub

Private Sub ComandoAltaReponsable_Click()
    On Error GoTo errores
    DoCmd.Hourglass True
    Dim result As String
    result = modExpedienteEntidadesHelper.ExpedienteEntidades_AltaResponsable( _
        BuildDTO(), Nz(Me.Responsable, ""), Nz(Me.EsJefeProyecto, ""), _
        Nz(Me.CorreoSiempre, ""), m_Error)
    If m_Error <> "" Then Err.Raise 1000
    RellenarListaResponsables m_Error
    If m_Error <> "" Then Err.Raise 1000
    Me.ListaResponsables.Selected(Me.ListaResponsables.ListCount - 1) = True
    Me.Responsable = Null
    Me.EsJefeProyecto = Null
    Me.CorreoSiempre = Null
    DoCmd.Hourglass False
    Exit Sub
errores:
    DoCmd.Hourglass False
    MsgBox "Alta Responsable: " & m_Error, vbCritical, "Error"
End Sub

' === Eliminar handlers =========================================================

Private Sub ComandoEliminarAnualidad_Click()
    On Error GoTo errores
    DoCmd.Hourglass True
    Dim result As String
    result = modExpedienteEntidadesHelper.ExpedienteEntidades_EliminarAnualidad( _
        BuildDTO(), Nz(Me.ListaAnualidad.Column(0), ""), _
        EsAdministrador = EnumSiNo.Sí, m_Error)
    If m_Error <> "" Then Err.Raise 1000
    RellenarListaAnualidades m_Error
    If m_Error <> "" Then Err.Raise 1000
    DoCmd.Hourglass False
    Exit Sub
errores:
    DoCmd.Hourglass False
    MsgBox "Eliminar Anualidad: " & m_Error, vbCritical, "Error"
End Sub

Private Sub ComandoEliminarComercial_Click()
    On Error GoTo errores
    DoCmd.Hourglass True
    Dim result As String
    result = modExpedienteEntidadesHelper.ExpedienteEntidades_EliminarComercial( _
        BuildDTO(), Nz(Me.ListaComerciales.Column(0), ""), m_Error)
    If m_Error <> "" Then Err.Raise 1000
    RellenarListaComerciales m_Error
    If m_Error <> "" Then Err.Raise 1000
    If Me.ListaComerciales.ListCount > 0 Then
        Me.ListaComerciales.Selected(Me.ListaComerciales.ListCount - 1) = True
    End If
    DoCmd.Hourglass False
    Exit Sub
errores:
    DoCmd.Hourglass False
    MsgBox "Eliminar Comercial: " & m_Error, vbCritical, "Error"
End Sub

Private Sub ComandoEliminarCPV_Click()
    On Error GoTo errores
    DoCmd.Hourglass True
    Dim result As String
    result = modExpedienteEntidadesHelper.ExpedienteEntidades_EliminarCPV( _
        BuildDTO(), Nz(Me.ListaCPVs.Column(0), ""), m_Error)
    If m_Error <> "" Then Err.Raise 1000
    RellenarListaCPVs m_Error
    If m_Error <> "" Then Err.Raise 1000
    If Me.ListaCPVs.ListCount > 0 Then
        Me.ListaCPVs.Selected(Me.ListaCPVs.ListCount - 1) = True
    End If
    DoCmd.Hourglass False
    Exit Sub
errores:
    DoCmd.Hourglass False
    MsgBox "Eliminar CPV: " & m_Error, vbCritical, "Error"
End Sub

Private Sub ComandoEliminarLugarEjecucion_Click()
    On Error GoTo errores
    DoCmd.Hourglass True
    Dim result As String
    result = modExpedienteEntidadesHelper.ExpedienteEntidades_EliminarLugar( _
        BuildDTO(), Nz(Me.ListaLugaresEjecucion.Column(0), ""), m_Error)
    If m_Error <> "" Then Err.Raise 1000
    RellenarListaLugares m_Error
    If m_Error <> "" Then Err.Raise 1000
    If Me.ListaLugaresEjecucion.ListCount > 0 Then
        Me.ListaLugaresEjecucion.Selected(Me.ListaLugaresEjecucion.ListCount - 1) = True
    End If
    DoCmd.Hourglass False
    Exit Sub
errores:
    DoCmd.Hourglass False
    MsgBox "Eliminar Lugar: " & m_Error, vbCritical, "Error"
End Sub

Private Sub ComandoEliminarPECAL_Click()
    On Error GoTo errores
    DoCmd.Hourglass True
    Dim result As String
    result = modExpedienteEntidadesHelper.ExpedienteEntidades_EliminarPECAL( _
        BuildDTO(), Nz(Me.ListaPECALES.Column(0), ""), m_Error)
    If m_Error <> "" Then Err.Raise 1000
    RellenarListaPECAL m_Error
    If m_Error <> "" Then Err.Raise 1000
    If Me.ListaPECALES.ListCount > 0 Then
        Me.ListaPECALES.Selected(Me.ListaPECALES.ListCount - 1) = True
    End If
    DoCmd.Hourglass False
    Exit Sub
errores:
    DoCmd.Hourglass False
    MsgBox "Eliminar PECAL: " & m_Error, vbCritical, "Error"
End Sub

Private Sub ComandoEliminarRACs_Click()
    On Error GoTo errores
    DoCmd.Hourglass True
    Dim result As String
    result = modExpedienteEntidadesHelper.ExpedienteEntidades_EliminarRAC( _
        BuildDTO(), Nz(Me.ListaRACs.Column(0), ""), m_Error)
    If m_Error <> "" Then Err.Raise 1000
    RellenarListaRACs m_Error
    If m_Error <> "" Then Err.Raise 1000
    If Me.ListaRACs.ListCount > 0 Then
        Me.ListaRACs.Selected(Me.ListaRACs.ListCount - 1) = True
    End If
    DoCmd.Hourglass False
    Exit Sub
errores:
    DoCmd.Hourglass False
    MsgBox "Eliminar RAC: " & m_Error, vbCritical, "Error"
End Sub

Private Sub ComandoEliminarReponsable_Click()
    On Error GoTo errores
    DoCmd.Hourglass True
    Dim result As String
    result = modExpedienteEntidadesHelper.ExpedienteEntidades_EliminarResponsable( _
        BuildDTO(), Nz(Me.ListaResponsables.Column(0), ""), m_Error)
    If m_Error <> "" Then Err.Raise 1000
    RellenarListaResponsables m_Error
    If m_Error <> "" Then Err.Raise 1000
    If Me.ListaResponsables.ListCount > 0 Then
        Me.ListaResponsables.Selected(Me.ListaResponsables.ListCount - 1) = True
    End If
    DoCmd.Hourglass False
    Exit Sub
errores:
    DoCmd.Hourglass False
    MsgBox "Eliminar Responsable: " & m_Error, vbCritical, "Error"
End Sub

' === Cambiar JP / Envio ========================================================

Private Sub ComandoCambiarJP_Click()
    On Error GoTo errores
    DoCmd.Hourglass True
    Me.AllowEdits = False
    Dim idU As String
    idU = Nz(Me.ListaResponsables.Column(0), "")
    Dim jpActual As String
    jpActual = "" & Me.ListaResponsables.Column(2)
    Dim jpFinal As String
    If jpActual = "Si" Then jpFinal = "No" Else jpFinal = "Si"

    Dim result As String
    result = modExpedienteEntidadesHelper.ExpedienteEntidades_CambiarJP(BuildDTO(), idU, jpFinal, m_Error)
    If m_Error <> "" Then Err.Raise 1000

    Dim lineaActual As String, lineaFinal As String
    lineaActual = idU & ";" & Me.ListaResponsables.Column(1) & ";" & jpActual & ";" & Me.ListaResponsables.Column(3)
    lineaFinal = idU & ";" & Me.ListaResponsables.Column(1) & ";" & jpFinal & ";" & Me.ListaResponsables.Column(3)
    CambiarLineaListaResponsables lineaActual, lineaFinal, m_Error
    If m_Error <> "" Then Err.Raise 1000

    Me.AllowEdits = True
    DoCmd.Hourglass False
    Exit Sub
errores:
    Me.AllowEdits = True
    DoCmd.Hourglass False
    MsgBox "Cambiar JP: " & m_Error, vbCritical, "Error"
End Sub

Private Sub ComandoCambiarEnvios_Click()
    On Error GoTo errores
    DoCmd.Hourglass True
    Me.AllowEdits = False
    Dim idU As String
    idU = Nz(Me.ListaResponsables.Column(0), "")
    Dim envioActual As String
    envioActual = "" & Me.ListaResponsables.Column(3)
    Dim envioFinal As String
    If envioActual = "Si" Then envioFinal = "No" Else envioFinal = "Si"

    Dim result As String
    result = modExpedienteEntidadesHelper.ExpedienteEntidades_CambiarEnvio(BuildDTO(), idU, envioFinal, m_Error)
    If m_Error <> "" Then Err.Raise 1000

    Dim lineaActual As String, lineaFinal As String
    lineaActual = idU & ";" & Me.ListaResponsables.Column(1) & ";" & Me.ListaResponsables.Column(2) & ";" & envioActual
    lineaFinal = idU & ";" & Me.ListaResponsables.Column(1) & ";" & Me.ListaResponsables.Column(2) & ";" & envioFinal
    CambiarLineaListaResponsables lineaActual, lineaFinal, m_Error
    If m_Error <> "" Then Err.Raise 1000

    Me.AllowEdits = True
    DoCmd.Hourglass False
    Exit Sub
errores:
    Me.AllowEdits = True
    DoCmd.Hourglass False
    MsgBox "Cambiar Envio: " & m_Error, vbCritical, "Error"
End Sub

' === Copiar ===================================================================

Private Sub ComandoCopiarCPV_Click()
    On Error GoTo errores
    DoCmd.Hourglass True
    Me.AllowEdits = False
    Dim cpv As String
    cpv = Nz(Me.ListaCPVs.Column(1), "")
    If Len(cpv) > 0 Then
        Me.AllowEdits = True
        DoCmd.Hourglass False
        CopiarAlPortapapeles cpv
        MsgBox "CPV copiado al portapapeles", vbInformation, "Copia del CPV"
        Exit Sub
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
    Exit Sub
errores:
    Me.AllowEdits = True
    DoCmd.Hourglass False
    MsgBox "Copiar CPV: " & Err.Description, vbCritical, "Error"
End Sub

Private Sub ComandoCopiarLugar_Click()
    On Error GoTo errores
    DoCmd.Hourglass True
    Me.AllowEdits = False
    Dim lugar As String
    lugar = Nz(Me.ListaLugaresEjecucion.Column(1), "")
    If Len(lugar) > 0 Then
        Me.AllowEdits = True
        DoCmd.Hourglass False
        CopiarAlPortapapeles lugar
        MsgBox "Lugar copiado al portapapeles", vbInformation, "Copia del lugar"
        Exit Sub
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
    Exit Sub
errores:
    Me.AllowEdits = True
    DoCmd.Hourglass False
    MsgBox "Copiar Lugar: " & Err.Description, vbCritical, "Error"
End Sub

' === Limpiar ==================================================================

Private Sub ComandoLimpiarAnualidad_Click()
    Me.AÑO = Null
    Me.BIIVA = Null
    Me.BIIPSI = Null
    Me.BIIGIC = Null
    Me.BIEXENTA = Null
End Sub

Private Sub ComandoLimpiarUsuario_Click()
    Me.Responsable = Null
    Me.EsJefeProyecto = Null
    Me.CorreoSiempre = Null
End Sub

' === ListaResponsables Click ==================================================

Private Sub ListaResponsables_Click()
    If Nz(Me.ListaResponsables, "") = "" Then
        Me.ComandoCambiarJP.Enabled = False
        Me.ComandoCambiarEnvios.Enabled = False
    Else
        Me.ComandoCambiarJP.Enabled = True
        Me.ComandoCambiarEnvios.Enabled = True
    End If
End Sub

' === WithEvents: open selection forms ==========================================

Private Sub ComandoElegirUsuario_Click()
    On Error GoTo errores
    DoCmd.Hourglass True
    If FormularioAbierto("FormUsuariosGestion") Then DoCmd.Close acForm, "FormUsuariosGestion"
    Set m_ObjComercialActivo = Nothing
    DoCmd.OpenForm "FormUsuariosGestion", OpenArgs:="PARA SELECCIONAR"
    If FormularioAbierto("FormUsuariosGestion") Then Set m_FormUsuarios = Forms("FormUsuariosGestion")
    DoCmd.Hourglass False
    Exit Sub
errores:
    DoCmd.Hourglass False
    MsgBox "Elegir Usuario: " & Err.Description, vbCritical, "Error"
End Sub

Private Sub ComandoElegirPECAL_Click()
    On Error GoTo errores
    DoCmd.Hourglass True
    If FormularioAbierto("FormPECALESGestion") Then DoCmd.Close acForm, "FormPECALESGestion"
    Set m_ObjPECALActiva = Nothing
    DoCmd.OpenForm "FormPECALESGestion", OpenArgs:="PARA SELECCIONAR"
    If FormularioAbierto("FormPECALESGestion") Then Set m_FormPECAL = Forms("FormPECALESGestion")
    DoCmd.Hourglass False
    Exit Sub
errores:
    DoCmd.Hourglass False
    MsgBox "Elegir PECAL: " & Err.Description, vbCritical, "Error"
End Sub

Private Sub ComandoElegirRAC_Click()
    On Error GoTo errores
    DoCmd.Hourglass True
    If FormularioAbierto("FormRACSGestion") Then DoCmd.Close acForm, "FormRACSGestion"
    Set m_ObjRACActivo = Nothing
    DoCmd.OpenForm "FormRACSGestion", OpenArgs:="PARA SELECCIONAR"
    If FormularioAbierto("FormRACSGestion") Then Set m_FormRAC = Forms("FormRACSGestion")
    DoCmd.Hourglass False
    Exit Sub
errores:
    DoCmd.Hourglass False
    MsgBox "Elegir RAC: " & Err.Description, vbCritical, "Error"
End Sub

Private Sub ComandoElegirComercial_Click()
    On Error GoTo errores
    DoCmd.Hourglass True
    If FormularioAbierto("FormComercialesGestion") Then DoCmd.Close acForm, "FormComercialesGestion"
    Set m_ObjComercialActivo = Nothing
    DoCmd.OpenForm "FormComercialesGestion", OpenArgs:="PARA SELECCIONAR"
    If FormularioAbierto("FormComercialesGestion") Then Set m_FormComercial = Forms("FormComercialesGestion")
    DoCmd.Hourglass False
    Exit Sub
errores:
    DoCmd.Hourglass False
    MsgBox "Elegir Comercial: " & Err.Description, vbCritical, "Error"
End Sub

Private Sub ComandoElegirCPV_Click()
    On Error GoTo errores
    DoCmd.Hourglass True
    If FormularioAbierto("FormCPVsGestion") Then DoCmd.Close acForm, "FormCPVsGestion"
    Set m_ObjCPVActivo = Nothing
    DoCmd.OpenForm "FormCPVsGestion", OpenArgs:="PARA SELECCIONAR"
    If FormularioAbierto("FormCPVsGestion") Then Set m_FormCPV = Forms("FormCPVsGestion")
    DoCmd.Hourglass False
    Exit Sub
errores:
    DoCmd.Hourglass False
    MsgBox "Elegir CPV: " & Err.Description, vbCritical, "Error"
End Sub

Private Sub ComandoElegirLugar_Click()
    On Error GoTo errores
    DoCmd.Hourglass True
    If FormularioAbierto("FormLugarEjecucionGestion") Then DoCmd.Close acForm, "FormLugarEjecucionGestion"
    Set m_ObjLugarEjecucionActiva = Nothing
    DoCmd.OpenForm "FormLugarEjecucionGestion", OpenArgs:="PARA SELECCIONAR"
    If FormularioAbierto("FormLugarEjecucionGestion") Then Set m_FormLugar = Forms("FormLugarEjecucionGestion")
    DoCmd.Hourglass False
    Exit Sub
errores:
    DoCmd.Hourglass False
    MsgBox "Elegir Lugar: " & Err.Description, vbCritical, "Error"
End Sub

' === WithEvents subscribers ====================================================

Private Sub m_FormComercial_Seleccionar(m_ComercialSeleccionado As Comercial)
    On Error GoTo errores
    DoCmd.Hourglass True
    If Not m_ComercialSeleccionado Is Nothing Then
        Dim result As String
        result = modExpedienteEntidadesHelper.ExpedienteEntidades_AltaComercial( _
            BuildDTO(), m_ComercialSeleccionado.IDComercial, m_Error)
        If m_Error <> "" Then Err.Raise 1000
        RellenarListaComerciales m_Error
        If m_Error <> "" Then Err.Raise 1000
        Me.ListaComerciales.Selected(Me.ListaComerciales.ListCount - 1) = True
    End If
    DoCmd.Hourglass False
    Exit Sub
errores:
    DoCmd.Hourglass False
    MsgBox "Comercial Seleccionar: " & m_Error, vbCritical, "Error"
End Sub

Private Sub m_FormCPV_Seleccionar(m_CPVSeleccionado As CPV)
    On Error GoTo errores
    DoCmd.Hourglass True
    If Not m_CPVSeleccionado Is Nothing Then
        Dim result As String
        result = modExpedienteEntidadesHelper.ExpedienteEntidades_AltaCPV( _
            BuildDTO(), m_CPVSeleccionado.IDCPV, m_Error)
        If m_Error <> "" Then Err.Raise 1000
        RellenarListaCPVs m_Error
        If m_Error <> "" Then Err.Raise 1000
        Me.ListaCPVs.Selected(Me.ListaCPVs.ListCount - 1) = True
    End If
    DoCmd.Hourglass False
    Exit Sub
errores:
    DoCmd.Hourglass False
    MsgBox "CPV Seleccionar: " & m_Error, vbCritical, "Error"
End Sub

Private Sub m_FormLugar_Seleccionar(m_LugarEjecucionSeleccionada As LugarEjecucion)
    On Error GoTo errores
    DoCmd.Hourglass True
    If Not m_LugarEjecucionSeleccionada Is Nothing Then
        Dim result As String
        result = modExpedienteEntidadesHelper.ExpedienteEntidades_AltaLugar( _
            BuildDTO(), m_LugarEjecucionSeleccionada.IDLugarEjecucion, m_Error)
        If m_Error <> "" Then Err.Raise 1000
        RellenarListaLugares m_Error
        If m_Error <> "" Then Err.Raise 1000
        Me.ListaLugaresEjecucion.Selected(Me.ListaLugaresEjecucion.ListCount - 1) = True
    End If
    DoCmd.Hourglass False
    Exit Sub
errores:
    DoCmd.Hourglass False
    MsgBox "Lugar Seleccionar: " & m_Error, vbCritical, "Error"
End Sub

Private Sub m_FormPECAL_Seleccionar(m_PECALSeleccionada As PECAL)
    On Error GoTo errores
    DoCmd.Hourglass True
    If Not m_PECALSeleccionada Is Nothing Then
        Dim result As String
        result = modExpedienteEntidadesHelper.ExpedienteEntidades_AltaPECAL( _
            BuildDTO(), m_PECALSeleccionada.IDPECAL, m_Error)
        If m_Error <> "" Then Err.Raise 1000
        RellenarListaPECAL m_Error
        If m_Error <> "" Then Err.Raise 1000
        Me.ListaPECALES.Selected(Me.ListaPECALES.ListCount - 1) = True
    End If
    DoCmd.Hourglass False
    Exit Sub
errores:
    DoCmd.Hourglass False
    MsgBox "PECAL Seleccionar: " & m_Error, vbCritical, "Error"
End Sub

Private Sub m_FormRAC_Seleccionar(m_RACSeleccionado As RAC)
    On Error GoTo errores
    DoCmd.Hourglass True
    If Not m_RACSeleccionado Is Nothing Then
        Dim result As String
        result = modExpedienteEntidadesHelper.ExpedienteEntidades_AltaRAC( _
            BuildDTO(), m_RACSeleccionado.IDRAC, m_Error)
        If m_Error <> "" Then Err.Raise 1000
        RellenarListaRACs m_Error
        If m_Error <> "" Then Err.Raise 1000
        Me.ListaRACs.Selected(Me.ListaRACs.ListCount - 1) = True
    End If
    DoCmd.Hourglass False
    Exit Sub
errores:
    DoCmd.Hourglass False
    MsgBox "RAC Seleccionar: " & m_Error, vbCritical, "Error"
End Sub

Private Sub m_FormUsuarios_Seleccionar(m_UsuarioSeleccionado As USUARIO)
    On Error GoTo errores
    DoCmd.Hourglass True
    If Not m_UsuarioSeleccionado Is Nothing Then
        Me.Responsable = m_UsuarioSeleccionado.Nombre
    End If
    DoCmd.Hourglass False
    Exit Sub
errores:
    DoCmd.Hourglass False
    MsgBox "Usuarios Seleccionar: " & Err.Description, vbCritical, "Error"
End Sub