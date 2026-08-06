Attribute VB_Name = "Test_PublicacionCalidadStatusPresenter"
Option Compare Database
Option Explicit

' ============================================================
' Test Battery: PublicacionCalidadStatusPresenter (issue #46)
' Pure logic only. No DB, no UI, no filesystem, no fixtures.
' Returns JSON {ok, value, payload, error, logs}
' ============================================================

Private Function PcspBuild( _
    ByVal p_ProyectoEnUTE As Boolean, _
    ByVal p_TieneAnexoEvidenciaUTE As Boolean, _
    ByVal p_HaySuministradores As Boolean, _
    ByVal p_EvidenciasSuministradoresCompletadas As Boolean, _
    ByVal p_EsActivo As Boolean, _
    ByVal p_FechaPreparadaParaPublicar As String, _
    ByVal p_PropuestaRechazadaPorCalidadFecha As String, _
    ByVal p_FechaPublicacion As String _
) As PublicacionCalidadStatusViewState
    Dim presenter As PublicacionCalidadStatusPresenter
    Set presenter = New PublicacionCalidadStatusPresenter

    Set PcspBuild = presenter.Build( _
            p_ProyectoEnUTE, _
            p_TieneAnexoEvidenciaUTE, _
            p_HaySuministradores, _
            p_EvidenciasSuministradoresCompletadas, _
            p_EsActivo, _
            p_FechaPreparadaParaPublicar, _
            p_PropuestaRechazadaPorCalidadFecha, _
            p_FechaPublicacion)
End Function

Private Sub PcspSetDefaultLogs(ByRef p_Logs() As String)
    p_Logs(0) = "1. Arrange: define pure publication status inputs"
    p_Logs(1) = "2. Act: presenter.Build"
    p_Logs(2) = "3. Assert: labels and command state match legacy form rules"
    p_Logs(3) = "4. Cleanup: none, pure test"
End Sub

Private Function PcspComplementadoColorRole() As String
    PcspComplementadoColorRole = "complementado"
End Function

Private Function PcspNoComplementadoColorRole() As String
    PcspNoComplementadoColorRole = "no-complementado"
End Function

Public Function Test_PublicacionCalidadStatusPresenter_UteRequiredCompleted() As String
    On Error GoTo EH

    Dim logs(0 To 3) As String
    PcspSetDefaultLogs logs

    Dim state As PublicacionCalidadStatusViewState
    Set state = PcspBuild(True, True, False, False, True, vbNullString, vbNullString, vbNullString)

    If state.UteRequeridoCaption <> "Requerido" Then
        Test_PublicacionCalidadStatusPresenter_UteRequiredCompleted = BuildJsonFail("ute_required_completed_caption", logs)
        Exit Function
    End If
    If state.UteComplementadoCaption <> "Complementado" Then
        Test_PublicacionCalidadStatusPresenter_UteRequiredCompleted = BuildJsonFail("ute_completed_caption", logs)
        Exit Function
    End If
    If state.UteComplementadoColorRole <> PcspComplementadoColorRole() Then
        Test_PublicacionCalidadStatusPresenter_UteRequiredCompleted = BuildJsonFail("ute_completed_color", logs)
        Exit Function
    End If

    Test_PublicacionCalidadStatusPresenter_UteRequiredCompleted = BuildJsonOk("ute_required_completed_ok", logs)
    Exit Function

EH:
    Test_PublicacionCalidadStatusPresenter_UteRequiredCompleted = BuildJsonFail("ute_required_completed: " & Err.Description, logs)
End Function

Public Function Test_PublicacionCalidadStatusPresenter_UteRequiredPending() As String
    On Error GoTo EH

    Dim logs(0 To 3) As String
    PcspSetDefaultLogs logs

    Dim state As PublicacionCalidadStatusViewState
    Set state = PcspBuild(True, False, False, False, True, vbNullString, vbNullString, vbNullString)

    If state.UteRequeridoCaption <> "Requerido" Then
        Test_PublicacionCalidadStatusPresenter_UteRequiredPending = BuildJsonFail("ute_required_caption", logs)
        Exit Function
    End If
    If state.UteComplementadoCaption <> "No complementado" Then
        Test_PublicacionCalidadStatusPresenter_UteRequiredPending = BuildJsonFail("ute_pending_caption", logs)
        Exit Function
    End If
    If state.UteComplementadoColorRole <> PcspNoComplementadoColorRole() Then
        Test_PublicacionCalidadStatusPresenter_UteRequiredPending = BuildJsonFail("ute_pending_color", logs)
        Exit Function
    End If

    Test_PublicacionCalidadStatusPresenter_UteRequiredPending = BuildJsonOk("ute_required_pending_ok", logs)
    Exit Function

EH:
    Test_PublicacionCalidadStatusPresenter_UteRequiredPending = BuildJsonFail("ute_required_pending: " & Err.Description, logs)
End Function

Public Function Test_PublicacionCalidadStatusPresenter_UteNotRequired() As String
    On Error GoTo EH

    Dim logs(0 To 3) As String
    PcspSetDefaultLogs logs

    Dim state As PublicacionCalidadStatusViewState
    Set state = PcspBuild(False, False, False, False, True, vbNullString, vbNullString, vbNullString)

    If state.UteRequeridoCaption <> "No Requerido" Then
        Test_PublicacionCalidadStatusPresenter_UteNotRequired = BuildJsonFail("ute_not_required_caption", logs)
        Exit Function
    End If
    If state.UteComplementadoCaption <> "N/A" Then
        Test_PublicacionCalidadStatusPresenter_UteNotRequired = BuildJsonFail("ute_not_required_na", logs)
        Exit Function
    End If
    If state.UteComplementadoColorRole <> PcspComplementadoColorRole() Then
        Test_PublicacionCalidadStatusPresenter_UteNotRequired = BuildJsonFail("ute_not_required_color", logs)
        Exit Function
    End If

    Test_PublicacionCalidadStatusPresenter_UteNotRequired = BuildJsonOk("ute_not_required_ok", logs)
    Exit Function

EH:
    Test_PublicacionCalidadStatusPresenter_UteNotRequired = BuildJsonFail("ute_not_required: " & Err.Description, logs)
End Function

Public Function Test_PublicacionCalidadStatusPresenter_SuppliersNotRequired() As String
    On Error GoTo EH

    Dim logs(0 To 3) As String
    PcspSetDefaultLogs logs

    Dim state As PublicacionCalidadStatusViewState
    Set state = PcspBuild(False, False, False, False, True, vbNullString, vbNullString, vbNullString)

    If state.SuministradoresRequeridoCaption <> "No requerido" Then
        Test_PublicacionCalidadStatusPresenter_SuppliersNotRequired = BuildJsonFail("suppliers_not_required_caption", logs)
        Exit Function
    End If
    If state.SuministradoresComplementadoCaption <> "N/A" Then
        Test_PublicacionCalidadStatusPresenter_SuppliersNotRequired = BuildJsonFail("suppliers_not_required_na", logs)
        Exit Function
    End If
    If state.SuministradoresComplementadoColorRole <> PcspComplementadoColorRole() Then
        Test_PublicacionCalidadStatusPresenter_SuppliersNotRequired = BuildJsonFail("suppliers_not_required_color", logs)
        Exit Function
    End If

    Test_PublicacionCalidadStatusPresenter_SuppliersNotRequired = BuildJsonOk("suppliers_not_required_ok", logs)
    Exit Function

EH:
    Test_PublicacionCalidadStatusPresenter_SuppliersNotRequired = BuildJsonFail("suppliers_not_required: " & Err.Description, logs)
End Function

Public Function Test_PublicacionCalidadStatusPresenter_SuppliersPending() As String
    On Error GoTo EH

    Dim logs(0 To 3) As String
    PcspSetDefaultLogs logs

    Dim state As PublicacionCalidadStatusViewState
    Set state = PcspBuild(False, False, True, False, True, vbNullString, vbNullString, vbNullString)

    If state.SuministradoresRequeridoCaption <> "Requerido" Then
        Test_PublicacionCalidadStatusPresenter_SuppliersPending = BuildJsonFail("suppliers_required_caption", logs)
        Exit Function
    End If
    If state.SuministradoresComplementadoCaption <> "No complementado" Then
        Test_PublicacionCalidadStatusPresenter_SuppliersPending = BuildJsonFail("suppliers_pending_caption", logs)
        Exit Function
    End If
    If state.SuministradoresComplementadoColorRole <> PcspNoComplementadoColorRole() Then
        Test_PublicacionCalidadStatusPresenter_SuppliersPending = BuildJsonFail("suppliers_pending_color", logs)
        Exit Function
    End If

    Test_PublicacionCalidadStatusPresenter_SuppliersPending = BuildJsonOk("suppliers_pending_ok", logs)
    Exit Function

EH:
    Test_PublicacionCalidadStatusPresenter_SuppliersPending = BuildJsonFail("suppliers_pending: " & Err.Description, logs)
End Function

Public Function Test_PublicacionCalidadStatusPresenter_SuppliersCompleted() As String
    On Error GoTo EH

    Dim logs(0 To 3) As String
    PcspSetDefaultLogs logs

    Dim state As PublicacionCalidadStatusViewState
    Set state = PcspBuild(False, False, True, True, True, vbNullString, vbNullString, vbNullString)

    If state.SuministradoresComplementadoCaption <> "Complementado" Then
        Test_PublicacionCalidadStatusPresenter_SuppliersCompleted = BuildJsonFail("suppliers_completed_caption", logs)
        Exit Function
    End If
    If state.SuministradoresComplementadoColorRole <> PcspComplementadoColorRole() Then
        Test_PublicacionCalidadStatusPresenter_SuppliersCompleted = BuildJsonFail("suppliers_completed_color", logs)
        Exit Function
    End If

    Test_PublicacionCalidadStatusPresenter_SuppliersCompleted = BuildJsonOk("suppliers_completed_ok", logs)
    Exit Function

EH:
    Test_PublicacionCalidadStatusPresenter_SuppliersCompleted = BuildJsonFail("suppliers_completed: " & Err.Description, logs)
End Function

Public Function Test_PublicacionCalidadStatusPresenter_ProposalNotPrepared() As String
    On Error GoTo EH

    Dim logs(0 To 3) As String
    PcspSetDefaultLogs logs

    Dim state As PublicacionCalidadStatusViewState
    Set state = PcspBuild(False, False, False, False, True, vbNullString, vbNullString, vbNullString)

    If Not state.ComandoPublicarEnabled Or Not state.ComandoRechazarEnabled Then
        Test_PublicacionCalidadStatusPresenter_ProposalNotPrepared = BuildJsonFail("active_buttons_enabled", logs)
        Exit Function
    End If
    If state.PublicacionCaption <> "SIN PROPUESTA TÉCNICA" Then
        Test_PublicacionCalidadStatusPresenter_ProposalNotPrepared = BuildJsonFail("proposal_not_prepared_caption", logs)
        Exit Function
    End If
    If state.PublicacionColorRole <> PcspNoComplementadoColorRole() Then
        Test_PublicacionCalidadStatusPresenter_ProposalNotPrepared = BuildJsonFail("proposal_not_prepared_color", logs)
        Exit Function
    End If
    If state.HasComandoVerRechazoVisibleDecision Then
        Test_PublicacionCalidadStatusPresenter_ProposalNotPrepared = BuildJsonFail("proposal_not_prepared_should_not_mutate_rejection_visibility", logs)
        Exit Function
    End If

    Test_PublicacionCalidadStatusPresenter_ProposalNotPrepared = BuildJsonOk("proposal_not_prepared_ok", logs)
    Exit Function

EH:
    Test_PublicacionCalidadStatusPresenter_ProposalNotPrepared = BuildJsonFail("proposal_not_prepared: " & Err.Description, logs)
End Function

Public Function Test_PublicacionCalidadStatusPresenter_WaitingQuality() As String
    On Error GoTo EH

    Dim logs(0 To 3) As String
    PcspSetDefaultLogs logs

    Dim state As PublicacionCalidadStatusViewState
    Set state = PcspBuild(False, False, False, False, True, "01/02/2026", vbNullString, vbNullString)

    If state.PublicacionCaption <> "PROPUESTA (01/02/2026) ESPERANDO A CALIDAD" Then
        Test_PublicacionCalidadStatusPresenter_WaitingQuality = BuildJsonFail("waiting_quality_caption", logs)
        Exit Function
    End If
    If state.PublicacionColorRole <> PcspComplementadoColorRole() Then
        Test_PublicacionCalidadStatusPresenter_WaitingQuality = BuildJsonFail("waiting_quality_color", logs)
        Exit Function
    End If
    If state.ComandoVerRechazoVisible Then
        Test_PublicacionCalidadStatusPresenter_WaitingQuality = BuildJsonFail("waiting_quality_rejection_hidden", logs)
        Exit Function
    End If
    If Not state.HasComandoVerRechazoVisibleDecision Then
        Test_PublicacionCalidadStatusPresenter_WaitingQuality = BuildJsonFail("waiting_quality_rejection_decision", logs)
        Exit Function
    End If

    Test_PublicacionCalidadStatusPresenter_WaitingQuality = BuildJsonOk("waiting_quality_ok", logs)
    Exit Function

EH:
    Test_PublicacionCalidadStatusPresenter_WaitingQuality = BuildJsonFail("waiting_quality: " & Err.Description, logs)
End Function

Public Function Test_PublicacionCalidadStatusPresenter_Rejected() As String
    On Error GoTo EH

    Dim logs(0 To 3) As String
    PcspSetDefaultLogs logs

    Dim state As PublicacionCalidadStatusViewState
    Set state = PcspBuild(False, False, False, False, True, "01/02/2026", "03/02/2026", vbNullString)

    If state.PublicacionCaption <> "PROPUESTA (01/02/2026) RECHAZADA (03/02/2026)" Then
        Test_PublicacionCalidadStatusPresenter_Rejected = BuildJsonFail("rejected_caption", logs)
        Exit Function
    End If
    If state.PublicacionColorRole <> PcspNoComplementadoColorRole() Then
        Test_PublicacionCalidadStatusPresenter_Rejected = BuildJsonFail("rejected_color", logs)
        Exit Function
    End If
    If Not state.ComandoVerRechazoVisible Then
        Test_PublicacionCalidadStatusPresenter_Rejected = BuildJsonFail("rejected_rejection_visible", logs)
        Exit Function
    End If
    If Not state.HasComandoVerRechazoVisibleDecision Then
        Test_PublicacionCalidadStatusPresenter_Rejected = BuildJsonFail("rejected_rejection_decision", logs)
        Exit Function
    End If

    Test_PublicacionCalidadStatusPresenter_Rejected = BuildJsonOk("rejected_ok", logs)
    Exit Function

EH:
    Test_PublicacionCalidadStatusPresenter_Rejected = BuildJsonFail("rejected: " & Err.Description, logs)
End Function

Public Function Test_PublicacionCalidadStatusPresenter_PublishedWithPreparedDate() As String
    On Error GoTo EH

    Dim logs(0 To 3) As String
    PcspSetDefaultLogs logs

    Dim state As PublicacionCalidadStatusViewState
    Set state = PcspBuild(False, False, False, False, False, "01/02/2026", vbNullString, "04/02/2026")

    If state.ComandoPublicarEnabled Or state.ComandoRechazarEnabled Then
        Test_PublicacionCalidadStatusPresenter_PublishedWithPreparedDate = BuildJsonFail("published_buttons_disabled", logs)
        Exit Function
    End If
    If state.ComandoPublicarCaption <> "Publicado" Then
        Test_PublicacionCalidadStatusPresenter_PublishedWithPreparedDate = BuildJsonFail("published_button_caption", logs)
        Exit Function
    End If
    If state.PublicacionCaption <> "PROPUESTA (01/02/2026) PUBLICADO EL 04/02/2026" Then
        Test_PublicacionCalidadStatusPresenter_PublishedWithPreparedDate = BuildJsonFail("published_caption", logs)
        Exit Function
    End If
    If state.ComandoVerRechazoVisible Then
        Test_PublicacionCalidadStatusPresenter_PublishedWithPreparedDate = BuildJsonFail("published_rejection_hidden", logs)
        Exit Function
    End If
    If Not state.HasComandoVerRechazoVisibleDecision Then
        Test_PublicacionCalidadStatusPresenter_PublishedWithPreparedDate = BuildJsonFail("published_rejection_decision", logs)
        Exit Function
    End If

    Test_PublicacionCalidadStatusPresenter_PublishedWithPreparedDate = BuildJsonOk("published_with_prepared_date_ok", logs)
    Exit Function

EH:
    Test_PublicacionCalidadStatusPresenter_PublishedWithPreparedDate = BuildJsonFail("published_with_prepared_date: " & Err.Description, logs)
End Function

Public Function Test_PublicacionCalidadStatusPresenter_PublishedWithoutPreparedDate() As String
    On Error GoTo EH

    Dim logs(0 To 3) As String
    PcspSetDefaultLogs logs

    Dim state As PublicacionCalidadStatusViewState
    Set state = PcspBuild(False, False, False, False, False, vbNullString, vbNullString, "04/02/2026")

    If Not state.PublicacionVisible Then
        Test_PublicacionCalidadStatusPresenter_PublishedWithoutPreparedDate = BuildJsonFail("published_without_prepared_visible", logs)
        Exit Function
    End If
    If state.PublicacionCaption <> "PROPUESTA PUBLICADA EL 04/02/2026" Then
        Test_PublicacionCalidadStatusPresenter_PublishedWithoutPreparedDate = BuildJsonFail("published_without_prepared_caption", logs)
        Exit Function
    End If
    If state.PublicacionColorRole <> PcspComplementadoColorRole() Then
        Test_PublicacionCalidadStatusPresenter_PublishedWithoutPreparedDate = BuildJsonFail("published_without_prepared_color", logs)
        Exit Function
    End If
    If state.ComandoVerRechazoVisible Then
        Test_PublicacionCalidadStatusPresenter_PublishedWithoutPreparedDate = BuildJsonFail("published_without_prepared_rejection_hidden", logs)
        Exit Function
    End If
    If Not state.HasComandoVerRechazoVisibleDecision Then
        Test_PublicacionCalidadStatusPresenter_PublishedWithoutPreparedDate = BuildJsonFail("published_without_prepared_rejection_decision", logs)
        Exit Function
    End If

    Test_PublicacionCalidadStatusPresenter_PublishedWithoutPreparedDate = BuildJsonOk("published_without_prepared_date_ok", logs)
    Exit Function

EH:
    Test_PublicacionCalidadStatusPresenter_PublishedWithoutPreparedDate = BuildJsonFail("published_without_prepared_date: " & Err.Description, logs)
End Function

Public Function Test_PublicacionCalidadStatusPresenter_PublishedWithoutPublicationDateHidesLabel() As String
    On Error GoTo EH

    Dim logs(0 To 3) As String
    PcspSetDefaultLogs logs

    Dim state As PublicacionCalidadStatusViewState
    Set state = PcspBuild(False, False, False, False, False, "01/02/2026", vbNullString, vbNullString)

    If state.PublicacionVisible Then
        Test_PublicacionCalidadStatusPresenter_PublishedWithoutPublicationDateHidesLabel = BuildJsonFail("published_without_publication_date_hidden", logs)
        Exit Function
    End If
    If state.ComandoVerRechazoVisible Then
        Test_PublicacionCalidadStatusPresenter_PublishedWithoutPublicationDateHidesLabel = BuildJsonFail("published_without_publication_date_rejection_hidden", logs)
        Exit Function
    End If
    If Not state.HasComandoVerRechazoVisibleDecision Then
        Test_PublicacionCalidadStatusPresenter_PublishedWithoutPublicationDateHidesLabel = BuildJsonFail("published_without_publication_date_rejection_decision", logs)
        Exit Function
    End If

    Test_PublicacionCalidadStatusPresenter_PublishedWithoutPublicationDateHidesLabel = BuildJsonOk("published_without_publication_date_hides_label_ok", logs)
    Exit Function

EH:
    Test_PublicacionCalidadStatusPresenter_PublishedWithoutPublicationDateHidesLabel = BuildJsonFail("published_without_publication_date_hides_label: " & Err.Description, logs)
End Function
