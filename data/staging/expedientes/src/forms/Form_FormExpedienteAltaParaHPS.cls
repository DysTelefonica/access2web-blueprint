Option Compare Database
Option Explicit

' Form_FormExpedienteAltaParaHPS — THINNED (REWORK 2026-06-26, Phase 3.2b / PR-8b)
' Pure UI orchestration. Business logic lives in modExpedienteAltaParaHPSHelper.
' See docs/audit/form-expediente-alta-para-hps-pure-data.md for the audit.
'
' Removed:
'   - Inline `For Each m_ID In m_Col` + `Me.Lista.AddItem` (moved to helper).
'   - Inline MsgBox in AgregarSuministradorMemoria ("ya está en la lista") — the
'     helper returns {added, duplicate}; form decides UI feedback.
'   - Duplicated add/del/alta logic — helpers own it.
'
' Preserved (UI orchestration only):
'   - WithEvents handlers — pass-through to UI
'   - cmdSalir_Click — DoCmd.Close
'   - 3 m_For*_Alta callbacks — refresh combos + trigger alta button
'   - 6 botones add/del — extract ID from combo/list → call helper → re-render
'   - Form_Load errores — MsgBox + DoCmd.Close
'   - Public methods (EstablecerDatos, EstablecerCombos, Rellenar*) — signature preserved

Public Event AltaExpediente(m_expediente As Expediente)

' Variables para formularios auxiliares
Private WithEvents m_FormClasificacion As Form_FormGradoClasificacion
Private WithEvents m_ForContratista As Form_FormSuministrador
Private WithEvents m_ForSubContratista As Form_FormSuministrador
Private WithEvents m_FormLugar As Form_FormLugarEjecucion

Private m_Error As String


' ==============================================================================
' INICIO Y CIERRE
' ==============================================================================

Private Sub Form_Load()
    Dim m_JSON As String
    Dim m_Parsed As Object
    Dim m_DTO As Object

    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Ajustar Me
    Me.AllowEdits = True

    ' Wrap the live DTO into a Dictionary so the helper sees a stable shape.
    Set m_DTO = CreateObject("Scripting.Dictionary")
    If Not m_ObjExpedienteDTOActivo Is Nothing Then
        If Not m_ObjExpedienteDTOActivo.Expediente Is Nothing Then
            Set m_DTO("Expediente") = m_ObjExpedienteDTOActivo.Expediente
        End If
        If Not m_ObjExpedienteDTOActivo.ColArbolSuministradores Is Nothing Then
            Set m_DTO("ColArbolSuministradores") = m_ObjExpedienteDTOActivo.ColArbolSuministradores
        End If
        If Not m_ObjExpedienteDTOActivo.ColLugaresEjecucion Is Nothing Then
            Set m_DTO("ColLugaresEjecucion") = m_ObjExpedienteDTOActivo.ColLugaresEjecucion
        End If
        If Not m_ObjEntorno Is Nothing Then
            Set m_DTO("Entorno") = m_ObjEntorno
        End If
    End If

    ' Delegate pure-data init to the helper (initializes ColArbolSuministradores if Nothing).
    m_JSON = modExpedienteAltaParaHPSHelper.ExpedienteHPS_Form_Load(m_DTO, m_Error)
    If m_Error <> "" Then Err.Raise 1000

    Set m_Parsed = JsonConverter.ParseJson(m_JSON)("payload")

    ' Sync back the live DTO's ColArbolSuministradores (helper may have initialized it).
    If Not m_ObjExpedienteDTOActivo Is Nothing Then
        If Not m_DTO Is Nothing Then
            If m_DTO.Exists("ColArbolSuministradores") Then
                If Not m_DTO("ColArbolSuministradores") Is Nothing Then
                    Set m_ObjExpedienteDTOActivo.ColArbolSuministradores = m_DTO("ColArbolSuministradores")
                End If
            End If
        End If
    End If

    ' Delegate to EstablecerDatos (public method, signature preserved).
    EstablecerDatos m_Error
    If m_Error <> "" Then Err.Raise 1000

    VBA.DoEvents
    DoCmd.Hourglass False
    Exit Sub
errores:
    DoCmd.Hourglass False
    MsgBox "Error al cargar Alta HPS: " & Err.Description, vbCritical
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Sub cmdSalir_Click()
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub


' ==============================================================================
' GESTIÓN DE DATOS Y COMBOS
' ==============================================================================

Public Function EstablecerDatos(Optional ByRef p_Error As String) As String
    On Error GoTo errores

    If m_ObjExpedienteDTOActivo Is Nothing Then
        p_Error = "Contexto perdido (DTO Nothing)"
        Err.Raise 1000
    End If

    EstablecerCombos p_Error
    If p_Error <> "" Then Err.Raise 1000

    ' Rellenar las listas visuales leyendo de la colección unificada en memoria
    RellenarListaContratistas
    RellenarListaSubContratistas
    RellenarListaLugares

    Exit Function
errores:
    p_Error = Err.Description
End Function

Public Function EstablecerCombos(Optional ByRef p_Error As String) As String
    Dim m_JSON As String
    Dim m_Parsed As Object
    Dim m_RowSource As String
    Dim m_DTO As Object

    On Error GoTo errores

    ' Wrap the live DTO into a Dictionary so the helper sees a stable shape.
    Set m_DTO = CreateObject("Scripting.Dictionary")
    If Not m_ObjExpedienteDTOActivo Is Nothing Then
        If Not m_ObjExpedienteDTOActivo.Expediente Is Nothing Then
            Set m_DTO("Expediente") = m_ObjExpedienteDTOActivo.Expediente
        End If
        If Not m_ObjExpedienteDTOActivo.ColArbolSuministradores Is Nothing Then
            Set m_DTO("ColArbolSuministradores") = m_ObjExpedienteDTOActivo.ColArbolSuministradores
        End If
        If Not m_ObjExpedienteDTOActivo.ColLugaresEjecucion Is Nothing Then
            Set m_DTO("ColLugaresEjecucion") = m_ObjExpedienteDTOActivo.ColLugaresEjecucion
        End If
        If Not m_ObjEntorno Is Nothing Then
            Set m_DTO("Entorno") = m_ObjEntorno
        End If
    End If

    ' Pure-data helper: returns 4 rowSource strings.
    m_JSON = modExpedienteAltaParaHPSHelper.ExpedienteHPS_EstablecerCombos_Construir(m_DTO, p_Error)
    If p_Error <> "" Then Exit Function

    Set m_Parsed = JsonConverter.ParseJson(m_JSON)("payload")

    ' 1. Clasificaciones
    Me.IdGradoClasificacion.RowSource = ""
    m_RowSource = CStr(m_Parsed("clasificaciones"))
    If Len(m_RowSource) > 0 Then
        Dim lines() As String
        Dim i As Long
        lines = Split(m_RowSource, vbCrLf)
        For i = LBound(lines) To UBound(lines)
            Me.IdGradoClasificacion.AddItem lines(i)
        Next i
    End If

    ' 2. Suministradores (both combos)
    Me.Contratistas.RowSource = ""
    Me.SubContratistas.RowSource = ""
    m_RowSource = CStr(m_Parsed("contratistas"))
    If Len(m_RowSource) > 0 Then
        Dim lines2() As String
        Dim j As Long
        lines2 = Split(m_RowSource, vbCrLf)
        For j = LBound(lines2) To UBound(lines2)
            Me.Contratistas.AddItem lines2(j)
            Me.SubContratistas.AddItem lines2(j)
        Next j
    End If

    ' 3. Lugares (combo de selección)
    Me.LugaresEjecucion.RowSource = ""
    m_RowSource = CStr(m_Parsed("lugares"))
    If Len(m_RowSource) > 0 Then
        Dim lines3() As String
        Dim k As Long
        lines3 = Split(m_RowSource, vbCrLf)
        For k = LBound(lines3) To UBound(lines3)
            Me.LugaresEjecucion.AddItem lines3(k)
        Next k
    End If

    Exit Function
errores:
    p_Error = Err.Description
End Function


' ==============================================================================
' RELLENO DE LISTAS VISUALES (ADAPTADOR DTO -> UI)
' ==============================================================================

Public Sub RellenarListaContratistas()
    Dim m_JSON As String
    Dim m_Parsed As Object
    Dim m_RowSource As String
    Dim m_DTO As Object

    Set m_DTO = CreateObject("Scripting.Dictionary")
    If Not m_ObjExpedienteDTOActivo Is Nothing Then
        If Not m_ObjExpedienteDTOActivo.ColArbolSuministradores Is Nothing Then
            Set m_DTO("ColArbolSuministradores") = m_ObjExpedienteDTOActivo.ColArbolSuministradores
        End If
    End If

    m_JSON = modExpedienteAltaParaHPSHelper.ExpedienteHPS_RellenarLista_Construir( _
        m_DTO, "CONTRATISTAS", m_Error)

    On Error Resume Next
    Set m_Parsed = JsonConverter.ParseJson(m_JSON)("payload")
    On Error GoTo 0

    Me.ListaContratistas.RowSource = ""
    If m_Parsed Is Nothing Then Exit Sub
    m_RowSource = CStr(m_Parsed("rowSource"))
    If Len(m_RowSource) > 0 Then
        Dim lines() As String
        Dim i As Long
        lines = Split(m_RowSource, vbCrLf)
        For i = LBound(lines) To UBound(lines)
            Me.ListaContratistas.AddItem lines(i)
        Next i
    End If
End Sub

Public Sub RellenarListaSubContratistas()
    Dim m_JSON As String
    Dim m_Parsed As Object
    Dim m_RowSource As String
    Dim m_DTO As Object

    Set m_DTO = CreateObject("Scripting.Dictionary")
    If Not m_ObjExpedienteDTOActivo Is Nothing Then
        If Not m_ObjExpedienteDTOActivo.ColArbolSuministradores Is Nothing Then
            Set m_DTO("ColArbolSuministradores") = m_ObjExpedienteDTOActivo.ColArbolSuministradores
        End If
    End If

    m_JSON = modExpedienteAltaParaHPSHelper.ExpedienteHPS_RellenarLista_Construir( _
        m_DTO, "SUBCONTRATISTAS", m_Error)

    On Error Resume Next
    Set m_Parsed = JsonConverter.ParseJson(m_JSON)("payload")
    On Error GoTo 0

    Me.ListaSubContratistas.RowSource = ""
    If m_Parsed Is Nothing Then Exit Sub
    m_RowSource = CStr(m_Parsed("rowSource"))
    If Len(m_RowSource) > 0 Then
        Dim lines() As String
        Dim i As Long
        lines = Split(m_RowSource, vbCrLf)
        For i = LBound(lines) To UBound(lines)
            Me.ListaSubContratistas.AddItem lines(i)
        Next i
    End If
End Sub

Public Sub RellenarListaLugares()
    Dim m_JSON As String
    Dim m_Parsed As Object
    Dim m_RowSource As String
    Dim m_DTO As Object

    Set m_DTO = CreateObject("Scripting.Dictionary")
    If Not m_ObjExpedienteDTOActivo Is Nothing Then
        If Not m_ObjExpedienteDTOActivo.ColLugaresEjecucion Is Nothing Then
            Set m_DTO("ColLugaresEjecucion") = m_ObjExpedienteDTOActivo.ColLugaresEjecucion
        End If
    End If

    m_JSON = modExpedienteAltaParaHPSHelper.ExpedienteHPS_RellenarLista_Construir( _
        m_DTO, "LUGARES", m_Error)

    On Error Resume Next
    Set m_Parsed = JsonConverter.ParseJson(m_JSON)("payload")
    On Error GoTo 0

    Me.ListaLugaresEjecucion.RowSource = ""
    If m_Parsed Is Nothing Then Exit Sub
    m_RowSource = CStr(m_Parsed("rowSource"))
    If Len(m_RowSource) > 0 Then
        Dim lines() As String
        Dim i As Long
        lines = Split(m_RowSource, vbCrLf)
        For i = LBound(lines) To UBound(lines)
            Me.ListaLugaresEjecucion.AddItem lines(i)
        Next i
    End If
End Sub


' ==============================================================================
' BOTONES DE ACCIÓN - CONTRATISTAS (ROOT_CONTR)
' ==============================================================================

Private Sub ComandoAltaContratista_Click()
    Dim sIDSuministrador As String
    Dim m_JSON As String
    Dim m_Parsed As Object
    Dim m_DTO As Object

    On Error GoTo errores
    sIDSuministrador = Nz(Me.Contratistas.Column(0), "")
    If sIDSuministrador = "" Then Exit Sub

    Set m_DTO = CreateObject("Scripting.Dictionary")
    If Not m_ObjExpedienteDTOActivo Is Nothing Then
        If Not m_ObjExpedienteDTOActivo.ColArbolSuministradores Is Nothing Then
            Set m_DTO("ColArbolSuministradores") = m_ObjExpedienteDTOActivo.ColArbolSuministradores
        End If
    End If

    m_JSON = modExpedienteAltaParaHPSHelper.ExpedienteHPS_AltaSuministrador( _
        m_DTO, sIDSuministrador, "ROOT_CONTR", m_Error)
    If m_Error <> "" Then Err.Raise 1000

    Set m_Parsed = JsonConverter.ParseJson(m_JSON)("payload")
    If CBool(m_Parsed("duplicate")) Then
        pregunta = MsgBox("Esta empresa ya está en la lista.", vbExclamation)
    End If

    RellenarListaContratistas
    Me.Contratistas = Null
    Exit Sub
errores:
    MsgBox "Error: " & Err.Description
End Sub

Private Sub ComandoEliminarContratista_Click()
    Dim sIDSuministrador As String
    Dim m_DTO As Object

    On Error GoTo errores
    sIDSuministrador = Nz(Me.ListaContratistas.Column(0), "")
    If sIDSuministrador = "" Then Exit Sub

    Set m_DTO = CreateObject("Scripting.Dictionary")
    If Not m_ObjExpedienteDTOActivo Is Nothing Then
        If Not m_ObjExpedienteDTOActivo.ColArbolSuministradores Is Nothing Then
            Set m_DTO("ColArbolSuministradores") = m_ObjExpedienteDTOActivo.ColArbolSuministradores
        End If
    End If

    m_DTO("ColArbolSuministradores") = m_ObjExpedienteDTOActivo.ColArbolSuministradores

    modExpedienteAltaParaHPSHelper.ExpedienteHPS_EliminarSuministrador _
        m_DTO, sIDSuministrador, "ROOT_CONTR", m_Error
    If m_Error <> "" Then Err.Raise 1000

    RellenarListaContratistas
    Exit Sub
errores:
    MsgBox "Error: " & Err.Description
End Sub


' ==============================================================================
' BOTONES DE ACCIÓN - SUBCONTRATISTAS (ROOT_SUB)
' ==============================================================================

Private Sub ComandoAltaSubContratista_Click()
    Dim sIDSuministrador As String
    Dim m_JSON As String
    Dim m_Parsed As Object
    Dim m_DTO As Object

    On Error GoTo errores
    sIDSuministrador = Nz(Me.SubContratistas.Column(0), "")
    If sIDSuministrador = "" Then Exit Sub

    Set m_DTO = CreateObject("Scripting.Dictionary")
    If Not m_ObjExpedienteDTOActivo Is Nothing Then
        If Not m_ObjExpedienteDTOActivo.ColArbolSuministradores Is Nothing Then
            Set m_DTO("ColArbolSuministradores") = m_ObjExpedienteDTOActivo.ColArbolSuministradores
        End If
    End If

    m_JSON = modExpedienteAltaParaHPSHelper.ExpedienteHPS_AltaSuministrador( _
        m_DTO, sIDSuministrador, "ROOT_SUB", m_Error)
    If m_Error <> "" Then Err.Raise 1000

    Set m_Parsed = JsonConverter.ParseJson(m_JSON)("payload")
    If CBool(m_Parsed("duplicate")) Then
        pregunta = MsgBox("Esta empresa ya está en la lista.", vbExclamation)
    End If

    RellenarListaSubContratistas
    Me.SubContratistas = Null
    Exit Sub
errores:
    MsgBox "Error: " & Err.Description
End Sub

Private Sub ComandoEliminarSubContratista_Click()
    Dim sIDSuministrador As String
    Dim m_DTO As Object

    On Error GoTo errores
    sIDSuministrador = Nz(Me.ListaSubContratistas.Column(0), "")
    If sIDSuministrador = "" Then Exit Sub

    Set m_DTO = CreateObject("Scripting.Dictionary")
    If Not m_ObjExpedienteDTOActivo Is Nothing Then
        If Not m_ObjExpedienteDTOActivo.ColArbolSuministradores Is Nothing Then
            Set m_DTO("ColArbolSuministradores") = m_ObjExpedienteDTOActivo.ColArbolSuministradores
        End If
    End If

    modExpedienteAltaParaHPSHelper.ExpedienteHPS_EliminarSuministrador _
        m_DTO, sIDSuministrador, "ROOT_SUB", m_Error
    If m_Error <> "" Then Err.Raise 1000

    RellenarListaSubContratistas
    Exit Sub
errores:
    MsgBox "Error: " & Err.Description
End Sub


' ==============================================================================
' EVENTOS UX / HELPERS
' ==============================================================================

Private Sub Contratistas_AfterUpdate()
    ComandoAltaContratista_Click
End Sub
Private Sub SubContratistas_AfterUpdate()
    ComandoAltaSubContratista_Click
End Sub
Private Sub LugaresEjecucion_AfterUpdate()
    ComandoAltaLugarEjecucion_Click
End Sub

' Retorno de formularios auxiliares de alta rápida
Private Sub m_ForContratista_Alta(p_ID As String)
    EstablecerCombos
    Me.Contratistas = p_ID
    ComandoAltaContratista_Click
End Sub
Private Sub m_ForSubContratista_Alta(p_ID As String)
    EstablecerCombos
    Me.SubContratistas = p_ID
    ComandoAltaSubContratista_Click
End Sub
Private Sub m_FormLugar_Alta()
    EstablecerCombos
End Sub


' ==============================================================================
' BOTONES DE ACCIÓN - LUGARES (legacy API)
' ==============================================================================

Private Sub ComandoAltaLugarEjecucion_Click()
    Dim sIDLugar As String
    Dim oLugar As LugarEjecucion
    Dim m_DTO As Object
    Dim m_JSON As String
    Dim m_Parsed As Object

    sIDLugar = Nz(Me.LugaresEjecucion.Column(0), "")
    If sIDLugar = "" Then Exit Sub

    Set m_DTO = CreateObject("Scripting.Dictionary")
    If Not m_ObjExpedienteDTOActivo Is Nothing Then
        If Not m_ObjExpedienteDTOActivo.ColLugaresEjecucion Is Nothing Then
            Set m_DTO("ColLugaresEjecucion") = m_ObjExpedienteDTOActivo.ColLugaresEjecucion
        End If
    End If

    m_JSON = modExpedienteAltaParaHPSHelper.ExpedienteHPS_AltaLugar(m_DTO, sIDLugar, m_Error)
    If m_Error <> "" Then
        MsgBox "Error: " & m_Error
        Exit Sub
    End If

    Set m_Parsed = JsonConverter.ParseJson(m_JSON)("payload")
    If Not CBool(m_Parsed("added")) Then
        ' duplicate — silent exit (legacy behavior).
        Me.LugaresEjecucion = Null
        Exit Sub
    End If

    ' Sync back the live DTO's ColLugaresEjecucion.
    If Not m_ObjExpedienteDTOActivo Is Nothing Then
        If Not m_DTO Is Nothing Then
            If m_DTO.Exists("ColLugaresEjecucion") Then
                Set m_ObjExpedienteDTOActivo.ColLugaresEjecucion = m_DTO("ColLugaresEjecucion")
            End If
        End If
    End If

    RellenarListaLugares
    Me.LugaresEjecucion = Null
End Sub

Private Sub ComandoEliminarLugarEjecucion_Click()
    Dim sIDLugar As String
    Dim m_DTO As Object

    sIDLugar = Nz(Me.ListaLugaresEjecucion.Column(0), "")
    If sIDLugar = "" Then Exit Sub

    Set m_DTO = CreateObject("Scripting.Dictionary")
    If Not m_ObjExpedienteDTOActivo Is Nothing Then
        If Not m_ObjExpedienteDTOActivo.ColLugaresEjecucion Is Nothing Then
            Set m_DTO("ColLugaresEjecucion") = m_ObjExpedienteDTOActivo.ColLugaresEjecucion
        End If
    End If

    modExpedienteAltaParaHPSHelper.ExpedienteHPS_EliminarLugar m_DTO, sIDLugar, m_Error
    If m_Error <> "" Then
        MsgBox "Error: " & m_Error
        Exit Sub
    End If

    RellenarListaLugares
End Sub


' ==============================================================================
' GUARDADO
' ==============================================================================

Private Sub ComandoRegistrar_Click()
    Dim m_ExpOp As ExpedienteOperaciones
    Dim m_JSON As String
    Dim m_Parsed As Object
    Dim m_DTO As Object

    On Error GoTo errores
    m_Error = ""

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    MostrarPopupProgreso "Procesando...", "Guardando..."

    ' 1. Rellenar campos de texto del formulario al DTO
    SetDTOFromForm

    ' 2. Wrap and call the pure-data helper to set HPS flags.
    Set m_DTO = CreateObject("Scripting.Dictionary")
    If Not m_ObjExpedienteDTOActivo Is Nothing Then
        Set m_DTO("Expediente") = m_ObjExpedienteDTOActivo.Expediente
    End If

    m_JSON = modExpedienteAltaParaHPSHelper.ExpedienteHPS_ComandoRegistrar_Click(m_DTO, m_Error)
    If m_Error <> "" Then Err.Raise 1000

    Set m_Parsed = JsonConverter.ParseJson(m_JSON)("payload")
    If Not CBool(m_Parsed("ready")) Then
        m_Error = "ComandoRegistrar: helper reported not ready"
        Err.Raise 1000
    End If

    ' 3. Llamada al Registrar Transaccional (Generic)
    Set m_ExpOp = New ExpedienteOperaciones
    m_ExpOp.Registrar m_ObjExpedienteDTOActivo, Nothing, m_Error
    If m_Error <> "" Then Err.Raise 1000

    ' 4. Finalizar
    RaiseEvent AltaExpediente(m_ObjExpedienteDTOActivo.Expediente)

    VBA.DoEvents
    CerrarPopupProgreso
    VBA.DoEvents
    DoCmd.Hourglass False
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub

errores:
    VBA.DoEvents
    CerrarPopupProgreso
    VBA.DoEvents
    DoCmd.Hourglass False
    MsgBox "Error al registrar: " & m_Error, vbCritical
End Sub

Private Function SetDTOFromForm()
    With m_ObjExpedienteDTOActivo.Expediente
        .Titulo = Nz(Me.Titulo, "")
        .Nemotecnico = Nz(Me.Nemotecnico, "")
        .CodExp = Nz(Me.CodExp, "")
        .IdGradoClasificacion = Nz(Me.IdGradoClasificacion.Column(0), "")
        .Observaciones = Nz(Me.Observaciones, "")
    End With
End Function