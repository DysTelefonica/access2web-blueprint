Attribute VB_Name = "Test_GestionRiesgosViewStateBuilder"
Option Compare Database
Option Explicit

' ============================================================
' Test Battery: GestionRiesgosViewStateBuilder (issue #41)
' Pure logic only. No DB, no UI, no fixtures.
' Returns JSON {ok, value, payload, error, logs}
' ============================================================

Private Function GrvsBuildState( _
    ByVal p_EsAlta As EnumSiNo, _
    ByVal p_EsTecnico As EnumSiNo, _
    ByVal p_NombreProyecto As String, _
    ByRef p_Error As String _
) As GestionRiesgosViewState
    Dim builder As GestionRiesgosViewStateBuilder
    Set builder = New GestionRiesgosViewStateBuilder
    Set GrvsBuildState = builder.Build(p_EsAlta, p_EsTecnico, p_NombreProyecto, p_Error)
End Function

Public Function Test_GestionRiesgosViewStateBuilder_AltaNoTecnico_State() As String
    On Error GoTo EH

    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: alta, non-technician, project not known"
    logs(1) = "2. Act: Build view state"
    logs(2) = "3. Assert: alta title, disabled related nav, save enabled"

    Dim errMsg As String
    Dim state As GestionRiesgosViewState
    Set state = GrvsBuildState(EnumSiNo.Sí, EnumSiNo.No, vbNullString, errMsg)

    If errMsg <> "" Then
        Test_GestionRiesgosViewStateBuilder_AltaNoTecnico_State = BuildJsonFail("unexpected error: " & errMsg, logs)
        Exit Function
    End If
    If state.Titulo <> "ALTA DE GESTIÓN DE RIESGOS" Then
        Test_GestionRiesgosViewStateBuilder_AltaNoTecnico_State = BuildJsonFail("unexpected title: " & state.Titulo, logs)
        Exit Function
    End If
    If state.NavResponsablesEnabled Or state.NavRiesgosOfertaEnabled Or state.NavSuministradoresEnabled Then
        Test_GestionRiesgosViewStateBuilder_AltaNoTecnico_State = BuildJsonFail("alta related nav must be disabled", logs)
        Exit Function
    End If
    If Not state.NavDatosGeneralesEnabled Then
        Test_GestionRiesgosViewStateBuilder_AltaNoTecnico_State = BuildJsonFail("datos generales nav must be enabled", logs)
        Exit Function
    End If
    If Not state.PermitidoEditar Or Not state.ComandoGrabarEnabled Then
        Test_GestionRiesgosViewStateBuilder_AltaNoTecnico_State = BuildJsonFail("non-technician must be allowed to edit/save", logs)
        Exit Function
    End If

    Test_GestionRiesgosViewStateBuilder_AltaNoTecnico_State = BuildJsonOk("alta_no_tecnico_state_ok", logs)
    Exit Function

EH:
    Test_GestionRiesgosViewStateBuilder_AltaNoTecnico_State = BuildJsonFail("Test_GestionRiesgosViewStateBuilder_AltaNoTecnico_State: " & Err.Description, logs)
End Function

Public Function Test_GestionRiesgosViewStateBuilder_AltaTecnico_Error() As String
    On Error GoTo EH

    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: alta requested by technician"
    logs(1) = "2. Act: Build view state"
    logs(2) = "3. Assert: unauthorized error is surfaced"

    Dim errMsg As String
    Dim state As GestionRiesgosViewState
    Set state = GrvsBuildState(EnumSiNo.Sí, EnumSiNo.Sí, vbNullString, errMsg)

    If errMsg <> "Es un usuario no autorizado" Then
        Test_GestionRiesgosViewStateBuilder_AltaTecnico_Error = BuildJsonFail("unexpected error: " & errMsg, logs)
        Exit Function
    End If
    If state.ErrorMessage <> errMsg Then
        Test_GestionRiesgosViewStateBuilder_AltaTecnico_Error = BuildJsonFail("state must carry same error", logs)
        Exit Function
    End If

    Test_GestionRiesgosViewStateBuilder_AltaTecnico_Error = BuildJsonOk("alta_tecnico_error_ok", logs)
    Exit Function

EH:
    Test_GestionRiesgosViewStateBuilder_AltaTecnico_Error = BuildJsonFail("Test_GestionRiesgosViewStateBuilder_AltaTecnico_Error: " & Err.Description, logs)
End Function

Public Function Test_GestionRiesgosViewStateBuilder_EditWithProjectName_State() As String
    On Error GoTo EH

    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: edit mode, non-technician, active project known"
    logs(1) = "2. Act: Build view state with project name"
    logs(2) = "3. Assert: edit title and related nav enabled"

    Dim errMsg As String
    Dim state As GestionRiesgosViewState
    Set state = GrvsBuildState(EnumSiNo.No, EnumSiNo.No, "Proyecto X", errMsg)

    If errMsg <> "" Then
        Test_GestionRiesgosViewStateBuilder_EditWithProjectName_State = BuildJsonFail("unexpected error: " & errMsg, logs)
        Exit Function
    End If
    If state.Titulo <> "EDICIÓN DE LA GESTIÓN DE RIESGOS: Proyecto X" Then
        Test_GestionRiesgosViewStateBuilder_EditWithProjectName_State = BuildJsonFail("unexpected title: " & state.Titulo, logs)
        Exit Function
    End If
    If Not state.NavResponsablesEnabled Or Not state.NavRiesgosOfertaEnabled Or Not state.NavSuministradoresEnabled Then
        Test_GestionRiesgosViewStateBuilder_EditWithProjectName_State = BuildJsonFail("edit related nav must be enabled", logs)
        Exit Function
    End If
    Test_GestionRiesgosViewStateBuilder_EditWithProjectName_State = BuildJsonOk("edit_project_name_state_ok", logs)
    Exit Function

EH:
    Test_GestionRiesgosViewStateBuilder_EditWithProjectName_State = BuildJsonFail("Test_GestionRiesgosViewStateBuilder_EditWithProjectName_State: " & Err.Description, logs)
End Function

Public Function Test_GestionRiesgosViewStateBuilder_UnknownAltaMode_Error() As String
    On Error GoTo EH

    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: alta mode is Empty/unknown"
    logs(1) = "2. Act: Build view state"
    logs(2) = "3. Assert: original unknown-mode error is surfaced"

    Dim errMsg As String
    Dim state As GestionRiesgosViewState
    Set state = GrvsBuildState(Empty, EnumSiNo.No, vbNullString, errMsg)

    If errMsg <> "No se sabe si es para alta o no" Then
        Test_GestionRiesgosViewStateBuilder_UnknownAltaMode_Error = BuildJsonFail("unexpected error: " & errMsg, logs)
        Exit Function
    End If
    If state.ErrorMessage <> errMsg Then
        Test_GestionRiesgosViewStateBuilder_UnknownAltaMode_Error = BuildJsonFail("state must carry same error", logs)
        Exit Function
    End If

    Test_GestionRiesgosViewStateBuilder_UnknownAltaMode_Error = BuildJsonOk("unknown_alta_error_ok", logs)
    Exit Function

EH:
    Test_GestionRiesgosViewStateBuilder_UnknownAltaMode_Error = BuildJsonFail("Test_GestionRiesgosViewStateBuilder_UnknownAltaMode_Error: " & Err.Description, logs)
End Function

Public Function Test_GestionRiesgosViewStateBuilder_EditTecnico_DisablesEditing() As String
    On Error GoTo EH

    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: edit mode requested by technician"
    logs(1) = "2. Act: Build view state"
    logs(2) = "3. Assert: edit allowed state is disabled, save disabled"

    Dim errMsg As String
    Dim state As GestionRiesgosViewState
    Set state = GrvsBuildState(EnumSiNo.No, EnumSiNo.Sí, "Proyecto Técnico", errMsg)

    If errMsg <> "" Then
        Test_GestionRiesgosViewStateBuilder_EditTecnico_DisablesEditing = BuildJsonFail("unexpected error: " & errMsg, logs)
        Exit Function
    End If
    If state.PermitidoEditar Or state.ComandoGrabarEnabled Then
        Test_GestionRiesgosViewStateBuilder_EditTecnico_DisablesEditing = BuildJsonFail("technician edit state must disable editing/save", logs)
        Exit Function
    End If

    Test_GestionRiesgosViewStateBuilder_EditTecnico_DisablesEditing = BuildJsonOk("edit_tecnico_disabled_ok", logs)
    Exit Function

EH:
    Test_GestionRiesgosViewStateBuilder_EditTecnico_DisablesEditing = BuildJsonFail("Test_GestionRiesgosViewStateBuilder_EditTecnico_DisablesEditing: " & Err.Description, logs)
End Function

Public Function Test_GestionRiesgosViewStateBuilder_EditNoTecnico_EnablesEditing() As String
    On Error GoTo EH

    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: edit mode requested by non-technician"
    logs(1) = "2. Act: Build view state"
    logs(2) = "3. Assert: edit allowed state is enabled, save enabled"

    Dim errMsg As String
    Dim state As GestionRiesgosViewState
    Set state = GrvsBuildState(EnumSiNo.No, EnumSiNo.No, "Proyecto Calidad", errMsg)

    If errMsg <> "" Then
        Test_GestionRiesgosViewStateBuilder_EditNoTecnico_EnablesEditing = BuildJsonFail("unexpected error: " & errMsg, logs)
        Exit Function
    End If
    If Not state.PermitidoEditar Or Not state.ComandoGrabarEnabled Then
        Test_GestionRiesgosViewStateBuilder_EditNoTecnico_EnablesEditing = BuildJsonFail("non-technician edit state must enable editing/save", logs)
        Exit Function
    End If

    Test_GestionRiesgosViewStateBuilder_EditNoTecnico_EnablesEditing = BuildJsonOk("edit_no_tecnico_enabled_ok", logs)
    Exit Function

EH:
    Test_GestionRiesgosViewStateBuilder_EditNoTecnico_EnablesEditing = BuildJsonFail("Test_GestionRiesgosViewStateBuilder_EditNoTecnico_EnablesEditing: " & Err.Description, logs)
End Function
